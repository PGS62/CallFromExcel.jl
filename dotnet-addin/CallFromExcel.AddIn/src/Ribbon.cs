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
          <splitButton id='LaunchSplit' size='large'>
            <button id='LaunchJulia' label='Launch Julia' imageMso='MacroPlay'
                    onAction='OnLaunchJulia'
                    screentip='Launch Julia'
                    supertip='Launch a local Julia session for this Excel session, using the Launch Settings.' />
            <menu id='LaunchMenu'>
              <button id='LaunchSettings' label='Launch Settings...' imageMso='PropertySheet'
                      onAction='OnLaunchSettings'
                      supertip='Choose the project folder, packages and other options the Launch Julia button uses. JuliaLaunch() formulas use their own arguments instead.' />
            </menu>
          </splitButton>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLaunchJulia(IRibbonControl control)
        {
            Launch(LaunchSettings.Load());
        }

        public void OnLaunchSettings(IRibbonControl control)
        {
            try
            {
                using (var form = new LaunchSettingsForm(LaunchSettings.Load()))
                {
                    if (form.ShowDialog(new ExcelWindow()) == DialogResult.Yes) Launch(form.Settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new ExcelWindow(), ex.Message, "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void Launch(LaunchSettings s)
        {
            try
            {
                SetStatusBar("Launching Julia...");
                ShowBriefly(JuliaLaunch.Launch(s.UseLinux, s.MinimiseWindow, s.CommandLineOptions,
                                               s.Packages, s.BashStatements, s.TimeOut));
            }
            catch (Exception ex)
            {
                SetStatusBar(null);
                MessageBox.Show(new ExcelWindow(), ex.Message, "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Excel's main window, as the owner of our dialogs, so they stay in front of Excel.
        private sealed class ExcelWindow : IWin32Window
        {
            public IntPtr Handle => ExcelDnaUtil.WindowHandle;
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
