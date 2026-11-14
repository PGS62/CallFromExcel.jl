using System;
using System.Drawing;
using System.Windows.Forms;

namespace CallFromExcel.AddIn
{
    // The ribbon's error dialog when part of the message deserves more than a MessageBox can
    // give it: a Julia command that fixes the problem (JuliaCommandException), set apart in a
    // fixed-width font with a button that copies it; or a web page that helps (HelpLinkException),
    // as a link. Otherwise like an error MessageBox, including its sound. Styled as
    // LaunchSettingsForm: the message-box font, and sizes in multiples of its height, so it
    // scales with the display.
    internal sealed class ErrorForm : Form
    {
        // The parts of a message set apart from its text: a command (copyable), other code such
        // as command line options, or a link.
        private enum Kind { Command, Code, Link }

        public static ErrorForm For(JuliaCommandException error) =>
            new ErrorForm(error.Message, new[] { Tuple.Create(error.JuliaCommand, Kind.Command) });

        public static ErrorForm For(HelpLinkException error) =>
            new ErrorForm(error.Message, error.Code == null
                ? new[] { Tuple.Create(error.Url, Kind.Link) }
                : new[] { Tuple.Create(error.Code, Kind.Code), Tuple.Create(error.Url, Kind.Link) });

        private ErrorForm(string message, Tuple<string, Kind>[] highlights)
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

            // The message, with each highlighted part on its own line, in order: e.g. the text
            // before a command ("... and run:"), the command, then any text after it (under WSL,
            // "- then launch Julia from Excel again."), as a sentence of its own.
            var text = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Margin = new Padding(0) };
            string rest = message;
            string command = null;
            foreach (Tuple<string, Kind> highlight in highlights)
            {
                int at = rest.IndexOf(highlight.Item1, StringComparison.Ordinal);
                if (at < 0) continue;
                AddSentence(text, rest.Substring(0, at), em);
                string part = highlight.Item1;
                if (highlight.Item2 == Kind.Link)
                {
                    var link = (LinkLabel)Paragraph(new LinkLabel(), part, Font, em, em / 2);
                    link.LinkClicked += (sender, e) => OpenUrl(part);
                    text.Controls.Add(link);
                }
                else
                {
                    text.Controls.Add(Paragraph(new Label(), part, new Font("Consolas", Font.Size), em, em / 2));
                    if (highlight.Item2 == Kind.Command) command = part;
                }
                rest = rest.Substring(at + part.Length);
            }
            AddSentence(text, rest, em);
            grid.Controls.Add(text, 1, 0);

            var ok = new Button { Text = "OK", AutoSize = true, MinimumSize = new Size(6 * em, 0), DialogResult = DialogResult.OK };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, em, 0, 0)
            };
            buttons.Controls.Add(ok);
            if (command != null) buttons.Controls.Add(CopyButton(command, em));
            grid.Controls.Add(buttons, 0, 1);
            grid.SetColumnSpan(buttons, 2);

            Controls.Add(grid);
            AcceptButton = ok;
            CancelButton = ok;
        }

        private Button CopyButton(string command, int em)
        {
            var copy = new Button { Text = "Copy Julia Command", AutoSize = true };
            // Keeps its width when its text changes to "Copied".
            copy.MinimumSize = new Size(TextRenderer.MeasureText(copy.Text, Font).Width + em, 0);
            copy.Click += (sender, e) =>
            {
                try
                {
                    Clipboard.SetText(command);
                    copy.Text = "Copied";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, Ribbon.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            return copy;
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

        // Text between highlighted parts, as a paragraph. After one, it's a sentence of its own:
        // the punctuation that joined it to the part before (". " or " - ") goes, and it starts
        // with a capital. Nothing if there's no text.
        private void AddSentence(FlowLayoutPanel text, string sentence, int em)
        {
            bool first = text.Controls.Count == 0;
            sentence = first ? sentence.Trim() : sentence.Trim().TrimStart('-', '.', ' ');
            if (sentence.Length == 0) return;
            if (!first) sentence = char.ToUpperInvariant(sentence[0]) + sentence.Substring(1);
            text.Controls.Add(Paragraph(new Label(), sentence, Font, em, first ? 0 : em / 2));
        }

        // Wraps, as a MessageBox's text does.
        private static Label Paragraph(Label label, string text, Font font, int em, int spaceAbove)
        {
            label.Text = text;
            label.Font = font;
            label.AutoSize = true;
            label.MaximumSize = new Size(28 * em, 0);
            label.Margin = new Padding(0, spaceAbove, 0, 0);
            return label;
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
