using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace CallFromExcel.AddIn
{
    // Embeds the IntelliSense display server into this add-in's own .xll, so users get
    // formula-bar IntelliSense (from the [ExcelFunction]/[ExcelArgument] descriptions in
    // JuliaFunctions.cs) without separately installing/loading ExcelDna.IntelliSense*.xll.
    // Only one display server is active per Excel instance (newest version wins if another
    // add-in also embeds one) - see https://github.com/Excel-DNA/IntelliSense/wiki.
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            JuliaFunctions.CacheStringLengthLimit();
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
