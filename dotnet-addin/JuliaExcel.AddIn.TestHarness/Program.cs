using System;
using JuliaExcel.AddIn;

namespace JuliaExcel.AddIn.TestHarness
{
    // Standalone console harness for JuliaLaunch/JuliaHttpClient -- exercises the exact
    // same code the real add-in will use, without needing Excel or COM registration at
    // all, so iteration doesn't require closing/reopening Excel each time.
    internal static class Program
    {
        private static void Main()
        {
            Console.WriteLine("Process ID: " + System.Diagnostics.Process.GetCurrentProcess().Id);

            Console.WriteLine();
            Console.WriteLine("Calling JuliaLaunch.Launch() ...");
            Console.WriteLine("Result: " + JuliaLaunch.Launch());
            Console.WriteLine("Port: " + JuliaPort.GetJuliaPort());

            Console.WriteLine();
            Console.WriteLine("Posting '1+1' to /eval ...");
            Console.WriteLine("Result: " + JuliaHttpClient.Post("1+1", "/eval"));

            Console.WriteLine();
            Console.WriteLine("Calling JuliaLaunch.Launch() again (should detect already-running) ...");
            Console.WriteLine("Result: " + JuliaLaunch.Launch());

            Console.WriteLine();
            Console.WriteLine("Posting '2+2' to /eval ...");
            Console.WriteLine("Result: " + JuliaHttpClient.Post("2+2", "/eval"));
        }
    }
}
