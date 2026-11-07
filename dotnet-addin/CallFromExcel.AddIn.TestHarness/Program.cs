namespace JuliaExcel.AddIn.TestHarness
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            bool selfOk = WireFormatTests.RunAll();

            bool juliaOk = true;
            if (args.Length > 0 && args[0] == "nojulia")
            {
                // skip
            }
            else
            {
                System.Console.WriteLine();
                System.Console.WriteLine("--- Julia cross-validation ---");
                juliaOk = JuliaCrossValidation.RunAll();
            }

            return (selfOk && juliaOk) ? 0 : 1;
        }
    }
}
