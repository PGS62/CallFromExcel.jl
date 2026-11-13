using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace CallFromExcel.AddIn
{
    // The port file, written by Julia's start_server, holds the port Julia listens on and, on
    // a second line, a secret token for the session that every request must carry (see
    // _reject_request in comms.jl). A third line holds the version of the CallFromExcel
    // Julia package serving it, which must equal this add-in's (see CheckVersion). It is named by this process's own PID - Excel's PID once
    // loaded via Excel-DNA - matching what Julia's startup script is told to expect (see
    // JuliaLaunch.BuildStartupScript's "CallFromExcel.setxlpid(...)" call), and it lives in the
    // user's own temp folder, so only this user can read the token.
    public static class JuliaPort
    {
        private const string NoConnectionMessage =
            "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. " +
            "If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see " +
            "\"CallFromExcel HTTP server listening\" printed in its window, then try again.";

        private static int _cachedPort;
        private static string _cachedToken = "";
        private static string _cachedVersion = "";

        private static string PortFilePath()
        {
            return Path.Combine(FileUtils.LocalTemp(), "Port_" + Process.GetCurrentProcess().Id + ".txt");
        }

        // Caches the port (and the token the port file holds for it). Zero forgets the
        // session and deletes the port file.
        public static void SetJuliaPort(int port)
        {
            if (port == 0)
            {
                _cachedPort = 0;
                _cachedToken = "";
                _cachedVersion = "";
                try { File.Delete(PortFilePath()); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                return;
            }
            _cachedPort = port;
            _cachedToken = TokenFromFile(port);
            _cachedVersion = VersionFromFile(port);
        }

        private static void ReadPortFile(out int port, out string token)
        {
            port = 0;
            token = "";
            try
            {
                string[] lines = FileUtils.ReadTextFile(PortFilePath()).Split('\n');
                if (int.TryParse(lines[0].Trim(), out int p) && p > 0) port = p;
                if (lines.Length > 1) token = lines[1].Trim();
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        public static int ReadJuliaPortFromFile()
        {
            ReadPortFile(out int port, out string _);
            return port;
        }

        // The token the port file currently holds for `port`, bypassing the cache - the XLL
        // and the COM server each have their own copy of this class, so after Julia is
        // relaunched one of them may still have the old session's token cached.
        public static string TokenFromFile(int port)
        {
            ReadPortFile(out int filePort, out string token);
            return filePort == port ? token : "";
        }

        // Line `index` (0-based) of the port file, if the file records `port` and has that line;
        // otherwise null, as for a line added after the Julia package that wrote the file.
        private static string LineFromFile(int port, int index)
        {
            try
            {
                string[] lines = FileUtils.ReadTextFile(PortFilePath()).Split('\n');
                if (!int.TryParse(lines[0].Trim(), out int filePort) || filePort != port) return null;
                return lines.Length > index ? lines[index].Trim() : null;
            }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
        }

        // The CallFromExcel version the port file records for `port` - "" if it records none,
        // as from a Julia package older than this check.
        private static string VersionFromFile(int port)
        {
            return LineFromFile(port, 2) ?? "";
        }

        // For the About box: whether a Julia session is listening on this Excel's port, and the
        // CallFromExcel version ("" if too old to record one) and platform its port file records.
        // A short connect timeout, so the box opens promptly when there's no session.
        internal static void RunningSession(out bool running, out string version, out bool wsl)
        {
            int port = GetJuliaPortOrZero();
            running = port > 0 && JuliaHttpClient.IsListening(port, 300);
            version = running ? VersionFromFile(port) : "";
            wsl = running && LineFromFile(port, 4) == "WSL";
        }

        // How to start Julia in a given environment: "start Julia" for Julia's default
        // environment (folder ""), with --project for a project folder. A Julia package older
        // than 0.1.2 records no folder (null), so the environment is then described rather than
        // named.
        private static string StartJulia(string folder, bool wsl)
        {
            string julia = wsl ? "start Julia under WSL" : "start Julia";
            if (folder == null) return julia + " in the environment you use with Excel";
            if (folder == "") return julia;
            return julia + " with julia --project=\"" + folder + "\"";
        }

        // This add-in's version, the AssemblyInformationalVersion that test/runtests.jl keeps
        // equal to Project.toml's.
        public static string AddInVersion
        {
            get
            {
                return typeof(JuliaPort).Assembly
                    .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "";
            }
        }

        // Throws if the Julia listening on `port` is a different version of CallFromExcel from
        // this add-in. The two must match - e.g. a change to the wire format changes both - and
        // since installing a package no longer installs the add-in, updating one without the
        // other is easy: this turns that into a clear message rather than baffling errors.
        public static void CheckVersion(int port)
        {
            string juliaVersion = port == _cachedPort ? _cachedVersion : VersionFromFile(port);
            if (juliaVersion == AddInVersion) return;
            throw MismatchError(juliaVersion, LineFromFile(port, 3), LineFromFile(port, 4) == "WSL");
        }

        // The error for a Julia running CallFromExcel `juliaVersion` ("" if too old to say)
        // in project `folder` (see StartJulia), natively or under WSL. Also used by JuliaLaunch,
        // whose start-up script checks the version before calling anything in the package that
        // an older version may lack.
        //
        // The fix depends on which is older. An older add-in is reinstalled from the package. An
        // older package is updated first - running installme() against it would downgrade the
        // add-in - and then installme() is run too, since the update may go past the add-in's
        // version. That needs a new Julia session: in one that has already loaded the package,
        // using CallFromExcel after Pkg.update keeps the old version, and installme() with it.
        // installme() only works in Julia on Windows, so under WSL an older package is just
        // updated (if that goes past the add-in, the next launch says so), and an older add-in
        // is updated from Julia on Windows.
        public static JuliaCommandException MismatchError(string juliaVersion, string folder, bool wsl)
        {
            string message = "CallFromExcel's Excel add-in (version " + AddInVersion + ") doesn't match the CallFromExcel " +
                "package Julia has loaded (" + (juliaVersion == "" ? "an older version" : "version " + juliaVersion) + "). ";
            const string Update = "using Pkg; Pkg.update(\"CallFromExcel\")";
            const string Install = "using CallFromExcel; CallFromExcel.installme()";
            const string UpdateAndInstall = Update + "; " + Install;
            bool packageOlder = juliaVersion == "" || IsOlder(juliaVersion, AddInVersion);
            if (packageOlder && wsl)
                return new JuliaCommandException(message + "Close Julia, then " + StartJulia(folder, true) + " and run: " + Update +
                    " - then launch Julia from Excel again.", Update);
            if (packageOlder)
                return new JuliaCommandException(message + "Close Excel and Julia, then " + StartJulia(folder, false) +
                    " and run: " + UpdateAndInstall, UpdateAndInstall);
            if (wsl)
                return new JuliaCommandException(message + "Close Excel, then start Julia on Windows and run: " + UpdateAndInstall,
                    UpdateAndInstall);
            return new JuliaCommandException(message + "Close Excel, then " + StartJulia(folder, false) + " and run: " + Install,
                Install);
        }

        // Whether version a is older than version b, comparing their numeric parts; false if
        // either doesn't parse.
        private static bool IsOlder(string a, string b)
        {
            return Version.TryParse(a.Split('-', '+')[0], out Version va) &&
                   Version.TryParse(b.Split('-', '+')[0], out Version vb) && va < vb;
        }

        // The token to send to the Julia listening on `port`.
        public static string TokenFor(int port)
        {
            if (port == _cachedPort && _cachedToken.Length > 0) return _cachedToken;
            return TokenFromFile(port);
        }

        // Returns the cached port, recovering from the port file on disk if the in-memory
        // cache is empty, or throws the "no connection" message.
        public static int GetJuliaPort()
        {
            int port = GetJuliaPortOrZero();
            if (port > 0) return port;
            throw new InvalidOperationException(NoConnectionMessage);
        }

        // Non-throwing variant, for JuliaCall/JuliaEval's own "please call JuliaLaunch first"
        // check.
        public static int GetJuliaPortOrZero()
        {
            if (_cachedPort != 0) return _cachedPort;
            ReadPortFile(out _cachedPort, out _cachedToken);
            _cachedVersion = _cachedPort == 0 ? "" : VersionFromFile(_cachedPort);
            return _cachedPort;
        }

        public static string NoConnectionErrorMessage
        {
            get { return NoConnectionMessage; }
        }
    }
}
