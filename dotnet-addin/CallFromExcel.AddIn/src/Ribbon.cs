using System;
using System.IO;
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
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
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
          <button id='ShowJuliaWindow' label='Show Julia Window' imageMso='WindowSwitchWindowsMenuExcel'
                  onAction='OnShowJuliaWindow'
                  supertip='Bring the Julia window serving this Excel session to the front, restoring it if it is minimised.' />
          <button id='CloseJulia' label='Close Julia' imageMso='PrintPreviewClose'
                  onAction='OnCloseJulia'
                  supertip='Close the Julia session serving this Excel session, after asking for confirmation. Variables and functions defined in it are lost.' />
        </group>
        <group id='CallFromExcelFileGroup' label='Julia File'>
          <box id='IncludeFileRow' boxStyle='horizontal'>
            <comboBox id='IncludeFileBox' label='File:' sizeString='C:\Users\username\Documents\Projects\Example\src\example.jl'
                      getText='GetIncludeText' onChange='OnIncludeTextChanged'
                      getItemCount='GetRecentCount' getItemLabel='GetRecentLabel'
                      screentip='Julia file'
                      getSupertip='GetIncludeSupertip' />
            <button id='BrowseIncludeFile' label='Browse...' showLabel='false' imageMso='FileOpen'
                    onAction='OnBrowseIncludeFile'
                    screentip='Browse'
                    supertip='Choose a Julia file to put in the box. It is not included until you click Include.' />
          </box>
          <button id='IncludeFile' label='Include' imageMso='ImportTextFile'
                  onAction='OnIncludeFile'
                  supertip='Load the file in the box into Julia, as JuliaInclude does, making the functions it defines available to JuliaCall and JuliaEval. Cells already using them are not recalculated.' />
        </group>
        <group id='CallFromExcelHelpGroup' label='Help'>
          <button id='CallFromExcelHelp' label='Help' size='large' imageMso='Help'
                  onAction='OnHelp'
                  supertip='Open the CallFromExcel documentation on GitHub.' />
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

        public void OnShowJuliaWindow(IRibbonControl control)
        {
            string title = JuliaLaunch.JuliaWindowTitle(System.Diagnostics.Process.GetCurrentProcess().Id);
            if (WindowUtils.TryGetHandleFromPartialCaption(title, out IntPtr hwnd))
            {
                WindowUtils.ShowAndActivate(hwnd);
                return;
            }
            // Running but not found by its title: e.g. a session attached with serve_xl, whose
            // window (VS Code's terminal, say) isn't titled by CallFromExcel.
            string message = JuliaFunctions.JuliaIsRunning()
                ? "Julia is running, but its window couldn't be found. If you attached your own Julia session with serve_xl, switch to it directly."
                : "Julia isn't running. Use Launch Julia to start it.";
            MessageBox.Show(new ExcelWindow(), message, "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Kept so the Julia File box can be refreshed (InvalidateControl) after Browse or Include.
        private IRibbonUI _ribbon;

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        // The Julia File box: its text is the file to include, its drop-down the recent files.
        public string GetIncludeText(IRibbonControl control) => IncludeFiles.Current;

        // The last item in the drop-down, when there is history to clear. A ribbon combo box's
        // items can only be text - no separator line, no action of their own - so choosing it
        // arrives here as the box's text changing to this label, which no file path can equal.
        private const string ClearHistoryLabel = "Clear history";

        public void OnIncludeTextChanged(IRibbonControl control, string text)
        {
            if (text == ClearHistoryLabel) IncludeFiles.ClearRecent();   // and keep the current file
            else IncludeFiles.Current = text;
            _ribbon?.InvalidateControl("IncludeFileBox");   // shows the cleaned path, e.g. without quotes
        }

        public int GetRecentCount(IRibbonControl control)
        {
            int n = IncludeFiles.Recent.Count;
            return n == 0 ? 0 : n + 1;
        }

        public string GetRecentLabel(IRibbonControl control, int index)
        {
            var recent = IncludeFiles.Recent;
            return index < recent.Count ? recent[index] : ClearHistoryLabel;
        }

        // The full path on hovering, since a long one doesn't fit in the box.
        public string GetIncludeSupertip(IRibbonControl control)
        {
            string current = IncludeFiles.Current;
            return (current == "" ? "No file chosen." : current) +
                "\n\nType or paste a path, choose a recent file from the list, or use Browse. Then click Include.";
        }

        public void OnBrowseIncludeFile(IRibbonControl control)
        {
            string current = IncludeFiles.Current;
            using (var dialog = new OpenFileDialog
            {
                Title = "Choose Julia File",
                Filter = "Julia files (*.jl)|*.jl|All files (*.*)|*.*"
            })
            {
                if (current != "" && Directory.Exists(Path.GetDirectoryName(current)))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(current);
                    dialog.FileName = Path.GetFileName(current);
                }
                if (dialog.ShowDialog(new ExcelWindow()) != DialogResult.OK) return;
                IncludeFiles.Current = dialog.FileName;
            }
            _ribbon?.InvalidateControl("IncludeFileBox");
        }

        public void OnIncludeFile(IRibbonControl control)
        {
            string fileName = IncludeFiles.Current;
            if (fileName == "")
            {
                Inform("Choose a Julia file first: type or paste its path in the File box, pick a recent file from its list, or use Browse.");
                return;
            }
            if (!File.Exists(fileName))
            {
                Inform("File not found:\n" + fileName);
                return;
            }
            if (!JuliaFunctions.JuliaIsRunning())
            {
                Inform("Julia isn't running. Use Launch Julia to start it first.");
                return;
            }
            SetStatusBar("Including " + Path.GetFileName(fileName) + "...");
            // JuliaInclude reports failures as text starting "#", as on a worksheet.
            object result = JuliaFunctions.JuliaInclude(fileName);
            if (result is string text && text.StartsWith("#", StringComparison.Ordinal))
            {
                SetStatusBar(null);
                MessageBox.Show(new ExcelWindow(), text.Substring(1), "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IncludeFiles.AddRecent(fileName);
            _ribbon?.InvalidateControl("IncludeFileBox");
            ShowBriefly("Included " + Path.GetFileName(fileName));
        }

        public void OnCloseJulia(IRibbonControl control)
        {
            int port = JuliaPort.GetJuliaPortOrZero();
            if (!JuliaFunctions.JuliaIsRunning())
            {
                Inform("Julia isn't running.");
                return;
            }
            DialogResult answer = MessageBox.Show(new ExcelWindow(),
                "Close the Julia session serving this Excel session? Variables and functions defined in it will be lost.",
                "CallFromExcel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) return;

            // Julia exits while handling the request, so no reply is expected: the error that
            // results is ignored, and whether Julia has stopped listening is checked instead.
            try { JuliaHttpClient.Post("exit()", "/eval"); } catch (Exception) { }
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            while (JuliaHttpClient.IsListening(port, 500) && stopwatch.Elapsed.TotalSeconds < 5)
            {
                System.Threading.Thread.Sleep(100);
            }
            if (JuliaHttpClient.IsListening(port, 500))
            {
                Inform("Julia didn't close. It may be busy - try again, or close its window.");
                return;
            }
            ShowBriefly("Julia closed");
        }

        private static void Inform(string message)
        {
            MessageBox.Show(new ExcelWindow(), message, "CallFromExcel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnHelp(IRibbonControl control)
        {
            try
            {
                // The README's title - the anchor follows its wording, so keep the two in step.
                System.Diagnostics.Process.Start(JuliaFunctions.HelpRoot + "callfromexceljl---call-julia-from-excel");
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
