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
}
