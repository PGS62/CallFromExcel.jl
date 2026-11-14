using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace CallFromExcel.AddIn
{
    // The ribbon's Launch Settings dialog. Built in code rather than with the designer, and laid
    // out with auto-sizing table panels, so it scales with the display's DPI and text size.
    // DialogResult is Yes for Launch (settings saved, then launch), OK for Save, Cancel for Cancel.
    internal sealed class LaunchSettingsForm : Form
    {
        private readonly CheckBox _useLinux = new CheckBox { Text = "Run Julia under Linux (WSL)", AutoSize = true };
        private readonly CheckBox _minimise = new CheckBox { Text = "Minimise Julia's window", AutoSize = true };
        // A combo box whose list holds the recently used project folders, as the ribbon's Julia
        // File box holds recent files (see UpdateFolderList).
        private readonly ComboBox _projectFolder = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDown };
        private static readonly RecentList RecentFolders = new RecentList(@"Software\CallFromExcel\ProjectFolders");
        // Julia's own names for these: its default environment (@v1.x), used with no --project,
        // and a temporary environment (--project=@temp, Julia 1.11 and later), deleted on exit.
        private const string DefaultEnvironmentItem = "(Default environment)";
        private const string TemporaryEnvironmentItem = "(Temporary environment)";
        private const string TemporaryEnvironment = "@temp";
        private const string ClearHistoryItem = "Clear history";
        private string _textBeforeDropDown = "";
        private readonly Button _browse = new Button { Text = "Browse...", AutoSize = false, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left };
        private readonly TextBox _otherOptions = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox _packages = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox _bash = new TextBox { Dock = DockStyle.Fill, Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical };
        private readonly NumericUpDown _timeOut = new NumericUpDown { Minimum = 1, Maximum = 3600 };
        private readonly ToolTip _tips = new ToolTip { AutoPopDelay = 20000 };

        private const int BashRow = 5;

        // Tooltips. Line breaks are written in, as WinForms tooltips don't wrap long text.
        private const string UseLinuxTip =
            "Run Julia as a Linux process under Windows Subsystem for Linux (WSL),\n" +
            "instead of as a Windows process. Julia and CallFromExcel must then be\n" +
            "installed under WSL.";
        private const string MinimiseTip =
            "Start Julia's window minimised, so it doesn't appear over Excel.\n" +
            "Click its taskbar button to see it.";
        private const string ProjectFolderTip =
            "The Julia project to start in, given to Julia as its --project option.\n" +
            "A project is a folder with a Project.toml file, listing the packages\n" +
            "(and, with Manifest.toml, their exact versions) that a piece of work\n" +
            "uses, kept separate from your other work. Set one when the functions\n" +
            "you call from Excel rely on particular packages or versions.\n\n" +
            "Leave blank to use your default environment (e.g. @v1.13).\n" +
            "Julia also finds packages installed in your default environment, so\n" +
            "CallFromExcel can be installed either there or in the project.\n\n" +
            "@temp, (Temporary environment) in the list, starts Julia in a new,\n" +
            "empty environment, deleted when Julia exits - for trying out packages\n" +
            "without adding them to a project: Packages to load that aren't\n" +
            "installed are installed into it. Needs Julia 1.11 or later.";
        private const string OtherOptionsTip =
            "Any other Julia command line options, separated by spaces,\n" +
            "e.g. --banner=no or --threads=4. -L and --load aren't allowed.";
        private const string PackagesTip =
            "Packages to load when Julia starts, separated by commas,\n" +
            "e.g. DataFrames, CSV. They must be installed in the project\n" +
            "or in your default environment - except with (Temporary environment),\n" +
            "when any that aren't are installed into it, preferring versions\n" +
            "you already have, so they're already precompiled.";
        private const string BashTip =
            "Only when running Julia under WSL: bash statements run before Julia\n" +
            "starts, one per line, e.g. export JULIA_NUM_THREADS=4.";
        private const string TimeOutTip =
            "How many seconds to wait for Julia to start, including any package\n" +
            "precompilation. If it takes longer, Excel stops waiting, but Julia\n" +
            "carries on starting and can be used once it has finished.";
        private TableLayoutPanel _grid;
        private Label _bashLabel;

        public LaunchSettings Settings { get; }

        // Opens at the size the layout needs, which is also the smallest allowed. The Bash
        // statements row stretches (Percent), and a stretching row adds nothing to the preferred
        // height - so it's measured as a fixed-height row, at its minimum height, then restored.
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Changed in place: a RowStyle can't be removed from the collection and put back.
            RowStyle bashRow = _grid.RowStyles[BashRow];
            // Absolute, at the box's minimum height: an AutoSize row was measured at a single
            // line's height, ignoring the box's MinimumSize.
            bashRow.SizeType = SizeType.Absolute;
            bashRow.Height = _bash.MinimumSize.Height + _bash.Margin.Vertical;
            Size preferred = _grid.GetPreferredSize(Size.Empty);
            bashRow.SizeType = SizeType.Percent;
            bashRow.Height = 100;
            ClientSize = new Size(preferred.Width + Padding.Horizontal, preferred.Height + Padding.Vertical);
            MinimumSize = Size;
        }

        // The resize grip in the bottom-right corner, a hint that the dialog can be resized.
        // Drawn here because SizeGripStyle.Show had no visible effect.
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            int size = Math.Min(SystemInformation.VerticalScrollBarWidth, Padding.Right);   // within the margin, not under the layout
            var corner = new Rectangle(ClientSize.Width - size, ClientSize.Height - size, size, size);
            if (VisualStyleRenderer.IsSupported && VisualStyleRenderer.IsElementDefined(VisualStyleElement.Status.Gripper.Normal))
            {
                new VisualStyleRenderer(VisualStyleElement.Status.Gripper.Normal).DrawBackground(e.Graphics, corner);
            }
            else
            {
                ControlPaint.DrawSizeGrip(e.Graphics, BackColor, corner);
            }
        }

        public LaunchSettingsForm(LaunchSettings settings)
        {
            Settings = settings;

            Text = Ribbon.DialogTitle + " - Launch Settings";
            Icon windowIcon = JuliaIcon.WindowIcon(SystemInformation.SmallIconSize.Width);   // null: keep the default icon
            if (windowIcon != null) Icon = windowIcon;
            Font = SystemFonts.MessageBoxFont;
            // Sizes are in multiples of the font's height, which grows with the display's scaling,
            // so the dialog keeps its proportions at any DPI (fixed pixel sizes didn't scale).
            int em = Font.Height;
            AutoScaleMode = AutoScaleMode.None;
            // Resizable: the text boxes take any extra width, the Bash statements box any extra
            // height. It opens at its smallest size (see OnLoad), so it can't be shrunk to clip.
            FormBorderStyle = FormBorderStyle.Sizable;
            SizeGripStyle = SizeGripStyle.Hide;   // drawn in OnPaint instead
            ResizeRedraw = true;                  // so the grip is redrawn in the new corner
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Padding = new Padding(em);

            var grid = new TableLayoutPanel { ColumnCount = 3, RowCount = 8, Dock = DockStyle.Fill };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            for (int r = 0; r < 8; r++)
            {
                grid.RowStyles.Add(r == BashRow ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.AutoSize));
            }
            _grid = grid;
            // The narrowest the text boxes can be, which sets the dialog's starting width.
            _projectFolder.MinimumSize = new Size(22 * em, 0);

            grid.Controls.Add(_useLinux, 0, 0);
            grid.SetColumnSpan(_useLinux, 3);
            grid.Controls.Add(_minimise, 0, 1);
            grid.SetColumnSpan(_minimise, 3);
            AddRow(grid, em, 2, "Project folder:", _projectFolder, _browse, ProjectFolderTip);
            // "command line options" links to Julia's documentation of them.
            const string optionsLabel = "Other command line options:";
            const string linked = "command line options";
            var optionsLink = new LinkLabel { Text = optionsLabel, LinkArea = new LinkArea(optionsLabel.IndexOf(linked), linked.Length) };
            optionsLink.LinkClicked += (sender, e) => OpenUrl(JuliaLaunch.CommandLineDocsUrl);
            AddRow(grid, em, 3, optionsLink, _otherOptions, null, OtherOptionsTip);
            AddRow(grid, em, 4, "Packages to load:", _packages, null, PackagesTip);
            _bashLabel = AddRow(grid, em, BashRow, "Bash statements:", _bash, null, BashTip);
            _bash.MinimumSize = new Size(0, 5 * em + em / 2);
            _browse.Width = TextRenderer.MeasureText(_browse.Text, Font).Width + em;
            _timeOut.Width = 5 * em;
            AddRow(grid, em, 6, "Time-out (seconds):", _timeOut, null, TimeOutTip);
            _timeOut.Dock = DockStyle.None;

            var launch = new Button { Text = "Launch", AutoSize = true, MinimumSize = new Size(6 * em, 0), DialogResult = DialogResult.Yes };
            var save = new Button { Text = "Save", AutoSize = true, MinimumSize = new Size(6 * em, 0), DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", AutoSize = true, MinimumSize = new Size(6 * em, 0), DialogResult = DialogResult.Cancel };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, em, 0, 0)
            };
            buttons.Controls.AddRange(new Control[] { cancel, save, launch });
            grid.Controls.Add(buttons, 0, 7);
            grid.SetColumnSpan(buttons, 3);

            Controls.Add(grid);
            AcceptButton = launch;
            CancelButton = cancel;

            _tips.SetToolTip(_useLinux, UseLinuxTip);
            _tips.SetToolTip(_minimise, MinimiseTip);

            _useLinux.Checked = settings.UseLinux;
            _minimise.Checked = settings.MinimiseWindow;
            _projectFolder.Text = settings.ProjectFolder;
            _otherOptions.Text = settings.OtherOptions;
            _packages.Text = settings.Packages;
            _bash.Text = (settings.BashStatements ?? "").Replace("\n", Environment.NewLine);
            _timeOut.Value = Math.Min(Math.Max(settings.TimeOut, 1), 3600);
            UpdateEnabled();

            _useLinux.CheckedChanged += (sender, e) => UpdateEnabled();
            _browse.Click += (sender, e) => Browse();
            _projectFolder.DropDown += (sender, e) => { _textBeforeDropDown = _projectFolder.Text; UpdateFolderList(); };
            _projectFolder.SelectionChangeCommitted += (sender, e) => OnFolderItemChosen();
            UpdateFolderList();
            FormClosing += (sender, e) => { if (DialogResult != DialogResult.Cancel) Store(); };
        }

        // The project folder's list, as the ribbon's Julia File box has: "(Default environment)",
        // when there is a folder, to empty the box; "(Temporary environment)", unless that's
        // already chosen, for @temp; the recent folders; and "Clear history" last, when there is
        // history to clear.
        private void UpdateFolderList()
        {
            string text = _projectFolder.Text;   // clearing the items can clear the text too
            _projectFolder.Items.Clear();
            if (_projectFolder.Text != text) _projectFolder.Text = text;
            if (text.Trim() != "") _projectFolder.Items.Add(DefaultEnvironmentItem);
            if (text.Trim() != TemporaryEnvironment) _projectFolder.Items.Add(TemporaryEnvironmentItem);
            List<string> recent = RecentFolders.Items;
            foreach (string folder in recent) _projectFolder.Items.Add(folder);
            if (recent.Count > 0) _projectFolder.Items.Add(ClearHistoryItem);
        }

        // Choosing "(Default environment)", "(Temporary environment)" or "Clear history" acts,
        // rather than putting that text in the box. The box sets its text to the chosen item after
        // this event, hence BeginInvoke.
        private void OnFolderItemChosen()
        {
            string chosen = _projectFolder.SelectedItem as string;
            if (chosen == DefaultEnvironmentItem || chosen == TemporaryEnvironmentItem)
            {
                string folder = chosen == DefaultEnvironmentItem ? "" : TemporaryEnvironment;
                BeginInvoke((Action)(() => { _projectFolder.Text = folder; UpdateFolderList(); }));
            }
            else if (chosen == ClearHistoryItem)
            {
                RecentFolders.Clear();
                string keep = _textBeforeDropDown;
                BeginInvoke((Action)(() => { _projectFolder.Text = keep; UpdateFolderList(); }));
            }
        }

        private void OpenUrl(string url)
        {
            try
            {
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Ribbon.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // A label and its field, with the same tooltip on both - the label's still shows when the
        // field is disabled (Bash statements, without WSL).
        private Label AddRow(TableLayoutPanel grid, int em, int row, string label, Control field, Control extra, string tip)
        {
            return AddRow(grid, em, row, new Label { Text = label }, field, extra, tip);
        }

        private Label AddRow(TableLayoutPanel grid, int em, int row, Label labelControl, Control field, Control extra, string tip)
        {
            labelControl.AutoSize = true;
            labelControl.Anchor = AnchorStyles.Left;
            labelControl.Margin = new Padding(em / 5, em / 3, em * 2 / 3, em / 3);
            grid.Controls.Add(labelControl, 0, row);
            _tips.SetToolTip(labelControl, tip);
            _tips.SetToolTip(field, tip);
            // The same margins for the field and the control beside it (the Browse button), which
            // stretches to the row's height - so the two are exactly level.
            field.Margin = new Padding(em / 5, em / 3, em / 5, em / 3);
            grid.Controls.Add(field, 1, row);
            if (extra != null)
            {
                extra.Margin = new Padding(em / 3, em / 3, em / 5, em / 3);
                grid.Controls.Add(extra, 2, row);
            }
            else grid.SetColumnSpan(field, 2);
            return labelControl;
        }

        private void UpdateEnabled()
        {
            _bash.Enabled = _useLinux.Checked;
            // Greyed by colour rather than by disabling: a disabled label is a hard-to-see dark
            // grey, and shows no tooltip.
            _bashLabel.ForeColor = _useLinux.Checked ? SystemColors.ControlText : Color.FromArgb(165, 165, 165);
        }

        private void Browse()
        {
            string folder = FolderPicker.Pick(Handle, "Choose the Julia project folder", _projectFolder.Text.Trim());
            if (folder == null) return;
            _projectFolder.Text = folder;
            // Remembered straight away, even if the dialog is then cancelled: it was chosen, and
            // exists. A typed folder is remembered on Launch or Save (Store).
            RecentFolders.Add(folder);
        }

        private void Store()
        {
            Settings.UseLinux = _useLinux.Checked;
            Settings.MinimiseWindow = _minimise.Checked;
            Settings.ProjectFolder = _projectFolder.Text.Trim();
            // To the history if it exists, so a mistyped folder doesn't get there. A Linux path
            // (under WSL, e.g. /home/me/proj or @myenv) can't be checked from here, so goes too.
            string folder = Settings.ProjectFolder;
            bool windowsPath = folder.Length >= 2 && (folder[1] == ':' || folder.StartsWith(@"\\", StringComparison.Ordinal));
            // @temp has its own item, so isn't added.
            if (folder != "" && folder != TemporaryEnvironment && (!windowsPath || System.IO.Directory.Exists(folder)))
                RecentFolders.Add(folder);
            Settings.OtherOptions = _otherOptions.Text.Trim();
            Settings.Packages = _packages.Text.Trim();
            // JuliaLaunch delimits bash statements with line feeds.
            Settings.BashStatements = _bash.Text.Replace("\r\n", "\n").Trim('\n');
            Settings.TimeOut = (int)_timeOut.Value;
            Settings.Save();
        }
    }
}
