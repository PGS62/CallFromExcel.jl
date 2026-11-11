using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace CallFromExcel.AddIn
{
    // The "Julia" ribbon tab. Excel-DNA finds this class because it derives from ExcelRibbon,
    // and loads it through a COM add-in it creates at load time; ComVisible(true) is needed
    // since the assembly as a whole is ComVisible(false) (AssemblyInfo.cs). Ribbon callbacks
    // run on Excel's main thread, as the worksheet functions do.
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='CallFromExcelTab' label='Julia'>
        <group id='CallFromExcelGroup' label='Julia'>
          <button id='LaunchJulia' label='Launch Julia' size='large' imageMso='MacroPlay'
                  onAction='OnLaunchJulia'
                  screentip='Launch Julia'
                  supertip='Launch a local Julia session for this Excel session, as JuliaLaunch() does with its default arguments.' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLaunchJulia(IRibbonControl control)
        {
            try
            {
                SetStatusBar("Launching Julia...");
                ShowBriefly(JuliaLaunch.Launch());
            }
            catch (Exception ex)
            {
                SetStatusBar(null);
                MessageBox.Show(ex.Message, "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private const int StatusBarSeconds = 5;

        // Kept in a field so the timer isn't garbage-collected before it fires.
        private static System.Threading.Timer _statusBarTimer;

        // Shows text in Excel's status bar, then gives the status bar back to Excel after a few
        // seconds - unless something else has put its own text there meanwhile. The wait happens
        // on a timer thread; QueueAsMacro (the C# counterpart of VBA's Application.OnTime) then
        // runs the clearing on Excel's main thread when Excel is ready, e.g. not mid-edit.
        private static void ShowBriefly(string text)
        {
            SetStatusBar(text);
            _statusBarTimer?.Dispose();
            _statusBarTimer = new System.Threading.Timer(_ => ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (GetStatusBar() as string == text) SetStatusBar(null);
            }), null, StatusBarSeconds * 1000, System.Threading.Timeout.Infinite);
        }

        // Shows text in Excel's status bar; null gives the status bar back to Excel.
        private static void SetStatusBar(string text)
        {
            object app = ExcelDnaUtil.Application;
            app.GetType().InvokeMember("StatusBar", BindingFlags.SetProperty, null, app,
                new object[] { text ?? (object)false });
        }

        // The status bar's text, or false when Excel is showing its own.
        private static object GetStatusBar()
        {
            object app = ExcelDnaUtil.Application;
            return app.GetType().InvokeMember("StatusBar", BindingFlags.GetProperty, null, app, null);
        }
    }
}
