using System.Runtime.InteropServices;

namespace JuliaExcel.AddIn
{
    // Test harness only, for validating the JuliaLaunch/JuliaHttpClient port from
    // PowerShell/VBA before it's wired into real JuliaCall/JuliaEval XLL functions.
    [ComVisible(true)]
    [Guid("6a6e3e0a-6f3a-4b9a-8a2a-1d6f7c9a2b4e")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJuliaExcelComServer
    {
        string Launch(bool minimiseWindow, string commandLineOptions, string packages, int timeoutSeconds);
        string Post(string payload, string path);
        bool IsListening(int port);
        int GetPort();
    }

    [ComVisible(true)]
    [Guid("9d7b5c1a-2e4f-4a8b-9c3d-7e1a5b6f8c2d")]
    [ProgId("JuliaExcel.AddIn.ComServer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJuliaExcelComServer))]
    public class ComServer : IJuliaExcelComServer
    {
        public string Launch(bool minimiseWindow, string commandLineOptions, string packages, int timeoutSeconds)
        {
            return JuliaLaunch.Launch(minimiseWindow, commandLineOptions, packages, timeoutSeconds);
        }

        public string Post(string payload, string path)
        {
            return JuliaHttpClient.Post(payload, path);
        }

        public bool IsListening(int port)
        {
            return JuliaHttpClient.IsListening(port);
        }

        public int GetPort()
        {
            return JuliaPort.GetJuliaPort();
        }
    }
}
