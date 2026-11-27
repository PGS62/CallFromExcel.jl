using System;

namespace CallFromExcel.AddIn
{
    // A call or evaluation that failed in Julia: the code threw, or its result couldn't be
    // encoded. Julia sends it as a failure ("X<message>", see WireFormatDecode), not as a string
    // result, so it can't be confused with one. The message is Julia's summary, e.g.
    // "MethodError: no method matching ... Julia REPL has more details and stacktrace".
    // On a worksheet it is shown as "#<message>!"; from VBA it is raised (see ComServer).
    public sealed class JuliaException : InvalidOperationException
    {
        public JuliaException(string message) : base(message) { }
    }

    // Julia isn't there to answer: not launched, exited, or still starting.
    public sealed class JuliaNotRunningException : InvalidOperationException
    {
        public JuliaNotRunningException(string message) : base(message) { }
    }
}
