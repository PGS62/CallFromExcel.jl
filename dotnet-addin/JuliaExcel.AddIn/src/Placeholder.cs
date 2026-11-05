using ExcelDna.Integration;

namespace JuliaExcel.AddIn
{
    public static class Placeholder
    {
        [ExcelFunction(Description = "Toolchain check: proves the MSBuild/Excel-DNA/ExcelDnaPack pipeline works.")]
        public static string JepToolchainCheck()
        {
            return "JuliaExcel.AddIn toolchain OK";
        }
    }
}
