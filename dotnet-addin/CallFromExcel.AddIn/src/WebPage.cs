using System;
using System.Windows.Forms;

namespace CallFromExcel.AddIn
{
    // Opens a web page in the user's browser, for the ribbon's Help button and the links in our
    // dialogs. A failure (e.g. no browser set up) is shown in a message box owned by `owner`.
    internal static class WebPage
    {
        public static void Open(string url, IWin32Window owner)
        {
            try
            {
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, ex.Message, Ribbon.DialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
