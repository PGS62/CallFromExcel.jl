using System;
using System.Drawing;
using System.Windows.Forms;

namespace CallFromExcel.AddIn
{
    // The ribbon's error dialog for a JuliaCommandException: like an error MessageBox, plus a
    // button that copies the Julia command that fixes the problem, which a MessageBox can't
    // offer. Styled as LaunchSettingsForm: the message-box font, and sizes in multiples of its
    // height, so it scales with the display.
    internal sealed class CommandErrorForm : Form
    {
        public CommandErrorForm(JuliaCommandException error)
        {
            Text = Ribbon.DialogTitle;
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

            // Placed inside the padding: a control that isn't docked ignores the form's Padding.
            var grid = new TableLayoutPanel { ColumnCount = 2, RowCount = 2, AutoSize = true, Location = new Point(em, em) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var icon = new PictureBox
            {
                Image = SystemIcons.Error.ToBitmap(),
                SizeMode = PictureBoxSizeMode.Zoom,
                Size = new Size(2 * em, 2 * em),
                Margin = new Padding(0, 0, em, 0)
            };
            grid.Controls.Add(icon, 0, 0);

            // The message, with the command set apart on its own, in a fixed-width font: the
            // text before it ("... and run:"), the command, then any text after it (under WSL,
            // "- then launch Julia from Excel again."), as a sentence of its own.
            var text = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Margin = new Padding(0) };
            int at = error.Message.IndexOf(error.JuliaCommand, StringComparison.Ordinal);
            if (at < 0)
            {
                text.Controls.Add(Paragraph(error.Message, Font, em, 0));
            }
            else
            {
                string after = error.Message.Substring(at + error.JuliaCommand.Length).Trim().TrimStart('-', ' ');
                text.Controls.Add(Paragraph(error.Message.Substring(0, at).TrimEnd(), Font, em, 0));
                text.Controls.Add(Paragraph(error.JuliaCommand, new Font("Consolas", Font.Size), em, em / 2));
                if (after.Length > 0)
                    text.Controls.Add(Paragraph(char.ToUpperInvariant(after[0]) + after.Substring(1), Font, em, em / 2));
            }
            grid.Controls.Add(text, 1, 0);

            var copy = new Button { Text = "Copy Julia Command", AutoSize = true };
            // Keeps its width when its text changes to "Copied".
            copy.MinimumSize = new Size(TextRenderer.MeasureText(copy.Text, Font).Width + em, 0);
            copy.Click += (sender, e) =>
            {
                try
                {
                    Clipboard.SetText(error.JuliaCommand);
                    copy.Text = "Copied";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, Ribbon.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            var ok = new Button { Text = "OK", AutoSize = true, MinimumSize = new Size(6 * em, 0), DialogResult = DialogResult.OK };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, em, 0, 0)
            };
            buttons.Controls.AddRange(new Control[] { ok, copy });
            grid.Controls.Add(buttons, 0, 1);
            grid.SetColumnSpan(buttons, 2);

            Controls.Add(grid);
            AcceptButton = ok;
            CancelButton = ok;
        }

        // Wraps, as a MessageBox's text does.
        private static Label Paragraph(string text, Font font, int em, int spaceAbove)
        {
            return new Label
            {
                Text = text,
                Font = font,
                AutoSize = true,
                MaximumSize = new Size(28 * em, 0),
                Margin = new Padding(0, spaceAbove, 0, 0)
            };
        }

        // The sound an error MessageBox makes - useful here, as the dialog can open behind the
        // Julia window that has just been launched.
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            System.Media.SystemSounds.Hand.Play();
        }
    }
}
