using System;
using System.Diagnostics;
using System.IO;

namespace CallFromExcel.AddIn
{
    // The port file, written by Julia's start_server, holds the port Julia listens on and, on
    // a second line, a secret token for the session that every request must carry (see
    // _reject_request in comms.jl). It is named by this process's own PID - Excel's PID once
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
                try { File.Delete(PortFilePath()); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                return;
            }
            _cachedPort = port;
            _cachedToken = TokenFromFile(port);
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
            return _cachedPort;
        }

        public static string NoConnectionErrorMessage
        {
            get { return NoConnectionMessage; }
        }
    }
}
