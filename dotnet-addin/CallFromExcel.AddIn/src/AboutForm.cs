using System;
using System.Drawing;
using System.Windows.Forms;

namespace CallFromExcel.AddIn
{
    // The ribbon's About box: the add-in's version, the CallFromExcel version of the Julia
    // session it's connected to (if any), and a link to the project on GitHub. A form rather
    // than a MessageBox, which can't hold a link. Styled as LaunchSettingsForm: the message-box
    // font, and sizes in multiples of its height, so it scales with the display.
    internal sealed class AboutForm : Form
    {
        public AboutForm()
        {
            Text = "About " + Ribbon.DialogTitle;
            Icon windowIcon = JuliaIcon.WindowIcon(SystemInformation.SmallIconSize.Width);   // null: keep the default icon
            if (windowIcon != null) Icon = windowIcon;
            Font = SystemFonts.MessageBoxFont;
            int em = Font.Height;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(em);

            // Placed inside the padding: a control that isn't docked ignores the form's Padding,
            // which the auto-size then only added on the right and bottom.
            var grid = new TableLayoutPanel { ColumnCount = 2, RowCount = 6, AutoSize = true, Location = new Point(em, em) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            Bitmap logo = JuliaIcon.Bitmap(3 * em);   // null if no julia.exe with an icon is found
            if (logo != null)
            {
                var picture = new PictureBox
                {
                    Image = logo,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Size = new Size(3 * em, 3 * em),
                    Margin = new Padding(0, em / 4, 2 * em, 0)
                };
                grid.Controls.Add(picture, 0, 0);
                grid.SetRowSpan(picture, 5);
            }

            var title = new Label
            {
                Text = Ribbon.DialogTitle,
                AutoSize = true,
                Font = new Font(Font.FontFamily, Font.Size * 1.5f, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, em / 4)
            };
            grid.Controls.Add(title, 1, 0);
            grid.Controls.Add(Line("Call Julia from Excel worksheets and VBA.", em), 1, 1);
            grid.Controls.Add(Line("Excel add-in: version " + JuliaPort.AddInVersion + "\n" + JuliaSessionText(), em), 1, 2);

            var link = new LinkLabel { Text = JuliaFunctions.DocsUrl, AutoSize = true, Margin = new Padding(0, em / 2, 0, 0) };
            // The documentation, as the ribbon's Help button opens.
            link.LinkClicked += (sender, e) => WebPage.Open(JuliaFunctions.DocsUrl, this);
            grid.Controls.Add(link, 1, 3);
            grid.Controls.Add(Line("Copyright (c) 2021-2026 Philip Swannell. MIT License.", em), 1, 4);

            var ok = new Button
            {
                Text = "OK",
                AutoSize = true,
                MinimumSize = new Size(6 * em, 0),
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, em, 0, 0)
            };
            grid.Controls.Add(ok, 0, 5);
            grid.SetColumnSpan(ok, 2);

            Controls.Add(grid);
            AcceptButton = ok;
            CancelButton = ok;
        }

        private static Label Line(string text, int em)
        {
            return new Label { Text = text, AutoSize = true, Margin = new Padding(0, em / 2, 0, 0) };
        }

        // The versions of the CallFromExcel package and of Julia in the session this Excel is
        // connected to, as its port file records them - read from the file, not asked of Julia,
        // so this can't hang on a busy Julia. Julia's line is left out when there's no session,
        // or its package is too old to record Julia's version.
        private static string JuliaSessionText()
        {
            JuliaPort.RunningSession(out bool running, out string version, out bool wsl, out string juliaVersion);
            if (!running) return "Julia package: Julia isn't running";
            string package = version == "" ? "an older version"
                : "version " + version + (version == JuliaPort.AddInVersion ? "" : " - doesn't match the add-in");
            string julia = juliaVersion == "" ? "" : "\nJulia: version " + juliaVersion + (wsl ? ", under WSL" : "");
            // Under WSL, said on Julia's line if there is one, else on the package's.
            if (wsl && julia == "") package += ", under WSL";
            return "Julia package: " + package + julia;
        }
    }
}
