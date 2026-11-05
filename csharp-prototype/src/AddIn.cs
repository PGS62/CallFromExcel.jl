using ExcelDna.Integration;

namespace JuliaExcelPrototype
{
    // Toolchain-validation prototype only: not wired to Julia.
    // One trivial XLL function, to prove Excel-DNA function registration works
    // end to end and that VBA can call it via Application.Run.
    public static class JepFunctions
    {
        [ExcelFunction(Description = "Prototype XLL function (Excel-DNA toolchain test): adds two numbers.")]
        public static double JepAdd(double a, double b)
        {
            return a + b;
        }
    }
}
