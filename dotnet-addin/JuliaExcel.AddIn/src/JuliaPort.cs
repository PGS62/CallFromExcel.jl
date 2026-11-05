using System;
using System.Diagnostics;
using System.IO;

namespace JuliaExcel.AddIn
{
    // C# port of the port-caching logic in modHttp.bas (SetJuliaPort/GetJuliaPort/
    // ReadJuliaPortFromFile). The port file is named by this process's own PID -- when
    // this code eventually runs loaded into Excel via Excel-DNA, Process.GetCurrentProcess().Id
    // is Excel's own PID, matching exactly what Julia's startup script is told to expect
    // (see JuliaLaunch.BuildStartupScript's "JuliaExcel.setxlpid(...)" call) and what the
    // VBA version keys its own port file on.
    public static class JuliaPort
    {
        private const string NoConnectionMessage =
            "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. " +
            "If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see " +
            "\"JuliaExcel HTTP server listening\" printed in its window, then try again.";

        private static int _cachedPort;

        private static string PortFilePath()
        {
            return Path.Combine(FileUtils.LocalTemp(), "Port_" + Process.GetCurrentProcess().Id + ".txt");
        }

        public static void SetJuliaPort(int port)
        {
            _cachedPort = port;
            string portFile = PortFilePath();
            if (port == 0)
            {
                try { File.Delete(portFile); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
            else
            {
                FileUtils.SaveTextFile(portFile, port.ToString());
            }
        }

        public static int ReadJuliaPortFromFile()
        {
            try
            {
                string text = FileUtils.ReadTextFile(PortFilePath());
                if (int.TryParse(text, out int port) && port > 0) return port;
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            return 0;
        }

        // Port of GetJuliaPort (modHttp.bas): returns the cached port, recovering from the
        // port file on disk if the in-memory cache was reset (e.g. by an AppDomain reload),
        // or throws the same "no connection" message VBA's GetJuliaPort does.
        public static int GetJuliaPort()
        {
            if (_cachedPort != 0) return _cachedPort;
            _cachedPort = ReadJuliaPortFromFile();
            if (_cachedPort > 0) return _cachedPort;
            throw new InvalidOperationException(NoConnectionMessage);
        }

        // Non-throwing variant used by JuliaCall/JuliaEval's own "please call JuliaLaunch
        // first" check (modMain.bas's JuliaCall_LowLevel checks "GetJuliaPort() = 0"
        // directly in VBA, which works there because VBA's Err mechanism is separate from
        // a function's return value -- GetJuliaPort's C# port throws instead, so that
        // check needs its own non-throwing path here).
        public static int GetJuliaPortOrZero()
        {
            if (_cachedPort != 0) return _cachedPort;
            _cachedPort = ReadJuliaPortFromFile();
            return _cachedPort;
        }

        public static string NoConnectionErrorMessage
        {
            get { return NoConnectionMessage; }
        }
    }
}
