namespace JuliaExcel.AddIn.TestHarness
{
    internal static class Program
    {
        private static int Main()
        {
            bool ok = WireFormatTests.RunAll();
            return ok ? 0 : 1;
        }
    }
}
