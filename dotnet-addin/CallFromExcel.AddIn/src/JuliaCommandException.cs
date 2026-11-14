using System;

namespace CallFromExcel.AddIn
{
    // An error whose fix is a Julia command for the user to run, such as updating the
    // CallFromExcel package. Carries the command separately from the message, so the ribbon's
    // error dialog can offer to copy it to the clipboard (see CommandErrorForm). Elsewhere, e.g. a
    // JuliaLaunch formula, it's an ordinary error whose message includes the command.
    public sealed class JuliaCommandException : InvalidOperationException
    {
        public JuliaCommandException(string message, string juliaCommand) : base(message)
        {
            JuliaCommand = juliaCommand;
        }

        public string JuliaCommand { get; }
    }

    // An error whose message includes a web page that helps, such as Julia's documentation of
    // its command line options. The ribbon's error dialog shows the address as a link; elsewhere
    // it's an ordinary error whose message includes the address.
    public sealed class HelpLinkException : InvalidOperationException
    {
        // code: any part of the message, such as the command line options Julia was given, to
        // set apart on its own line in a fixed-width font in the dialog; null for none.
        public HelpLinkException(string message, string url, string code = null) : base(message)
        {
            Url = url;
            Code = code;
        }

        public string Url { get; }

        public string Code { get; }
    }
}
