# Debugging

Every call to `JuliaCall`/`JuliaCallVBA`/`JuliaEval`/`JuliaEvalVBA` leaves a trace behind in Julia, so a problem can be debugged at the REPL without needing to repeat the call from Excel:

* `CallFromExcel.last_question` is the expression that was evaluated - either the exact string passed to `JuliaEval`/`JuliaEvalVBA`, or, for `JuliaCall`/`JuliaCallVBA`, a reconstructed call such as `myfunction(CallFromExcel.args_from_xl...)`.
* `CallFromExcel.last_answer` is the value that was returned to Excel.
* `CallFromExcel.args_from_xl` holds the arguments most recently passed from Excel to `JuliaCall`/`JuliaCallVBA`, or the positional arguments for `JuliaCallNamedArgs`/`JuliaCallNamedArgsVBA`. For those, `CallFromExcel.kwargs_from_xl` holds the named arguments as a `NamedTuple`, and `last_question` is a call such as `myfunction(args_from_xl...; kwargs_from_xl...)`.
* [`CallFromExcel.answer_again()`](@ref answer_again) evaluates `last_question` again, and you can wrap it with `@enter` or your own debugging tools.

All of these only reflect the most recent call, so debug before making another one. Call [`CallFromExcel.display_results(true)`](@ref display_results), or tick Display Results on the ribbon, to also have each call and its result echoed to the Julia console as it happens.

To step through your code with a debugger, attach a Julia session you already have open, such as one in VS Code - see [Attaching to a running Julia session](attaching.md).
