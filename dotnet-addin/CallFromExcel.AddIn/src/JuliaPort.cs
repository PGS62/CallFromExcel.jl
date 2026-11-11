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

        // The CallFromExcel version the port file records for `port` - "" if it records none,
        // as from a Julia package older than this check.
        private static string VersionFromFile(int port)
        {
            try
            {
                string[] lines = FileUtils.ReadTextFile(PortFilePath()).Split('\n');
                if (!int.TryParse(lines[0].Trim(), out int filePort) || filePort != port) return "";
                return lines.Length > 2 ? lines[2].Trim() : "";
            }
            catch (IOException) { return ""; }
            catch (UnauthorizedAccessException) { return ""; }
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
            throw new InvalidOperationException(
                "CallFromExcel's Excel add-in (version " + AddInVersion + ") doesn't match the CallFromExcel " +
                "package Julia has loaded (" + (juliaVersion == "" ? "an older version" : "version " + juliaVersion) +
                "). Close Excel, then in Julia, in the environment you use with Excel, run: " +
                "using CallFromExcel; CallFromExcel.installme()");
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
