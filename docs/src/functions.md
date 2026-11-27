# Worksheet functions

These functions can be used in worksheet formulas. A call that fails returns a string describing the failure, starting with `#` and ending with `!`. All of them except `JuliaHandleInfo` can also be called from VBA, as for example `jl.JuliaCall(...)`, where a failure instead raises an [error](@ref vba-errors) - see [Calling from VBA](vba.md). For the functions that can only be called from VBA, see [VBA functions](vba_functions.md).

|Name|Description|
|:---|:----------|
|[JuliaLaunch](@ref julialaunch)|Launches a local Julia session which listens to the current Excel session and responds to calls to `JuliaEval` etc.|
|[JuliaCall](@ref juliacall)|Call a named Julia function, passing in data from the worksheet. Returns an error string for results that cannot be displayed on a worksheet (nested arrays, dictionaries etc.). `JuliaCallVBA` lifts those restrictions.|
|[JuliaCallNamedArgs](@ref juliacallnamedargs)|Call a named Julia function, passing positional and named arguments from the worksheet. The argument `";"` marks where the named arguments start, given as name-value pairs. A named argument whose value is `";"` must be the last pair.|
|[JuliaCallHandle](@ref juliacallhandle)|Call a named Julia function and return a [handle](handles.md) to the result, which stays in Julia for other functions to use. An argument `";"` marks where named arguments start, as in `JuliaCallNamedArgs`.|
|[JuliaEval](@ref juliaeval)|Evaluate a Julia expression and return the result to an Excel worksheet, or with `ReturnHandle` TRUE, a [handle](handles.md) to the result.|
|[JuliaInclude](@ref juliainclude)|Load a Julia source file into the Julia process, to make additional functions available via `JuliaEval` and `JuliaCall`.|
|[JuliaSetVar](@ref juliasetvar)|Set a global variable in the Julia process.|
|[JuliaHandleInfo](@ref juliahandleinfo)|Information about object handles: how many are live, or recent events.|
|[JuliaIsRunning](@ref juliaisrunning)|Returns TRUE if an instance of Julia is running and "listening" to the current Excel session, or FALSE otherwise.|
|[JuliaExcelPID](@ref juliaexcelpid)|Returns the process id of the current Excel session. Only needed if more than one Excel instance is running - pass it to `serve_xl` to say which one to attach to.|

## [`JuliaLaunch`](@id julialaunch)
Launches a local Julia session which listens to the current Excel session and responds to calls to `JuliaEval` etc. Uses only its own arguments, not the ribbon's Launch Settings (see [The Julia ribbon tab](ribbon.md)).
```
JuliaLaunch([UseLinux], [MinimiseWindow], [CommandLineOptions], [Packages], [BashStatements], [TimeOut])
```

|Argument|Description|
|:-------|:----------|
|`UseLinux`|TRUE to run Julia as a Linux process under Windows Subsystem for Linux; FALSE (the default) to run as a Windows process.|
|`MinimiseWindow`|If TRUE, then the Julia session window is minimised; if FALSE (the default) then the window is sized normally.|
|`CommandLineOptions`|Command line options for Julia, e.g. `--project=C:\MyProject` or `--banner=no`. A `--project` folder must exist; `--project=@temp` starts Julia in a temporary environment, deleted when Julia exits (Julia 1.11 or later). If `CommandLineOptions` has no `--threads` option, `--threads=auto,1` is added: Julia's usual threads, plus one interactive thread, which HTTP.jl recommends for servers. To choose the number of threads yourself, give your own `--threads` option, e.g. `--threads=8,1` (the `,1` keeps the interactive thread). See [Julia's command line options](https://docs.julialang.org/en/v1/manual/command-line-interface/#command-line-switches-for-julia).|
|`Packages`|`Packages` to load, which must be installed in Julia's active environment (set by the `--project` command line option) or the default environment. Delimit multiple packages with commas. With `--project=@temp`, any that aren't installed are first installed into the temporary environment, preferring versions you already have, so they're already precompiled, and then loaded like the others.|
|`BashStatements`|Relevant only when `UseLinux` is TRUE. Bash statements executed prior to launching Julia, which can be used to set environment variables. Example: `export JULIA_PKG_DEVDIR=/mnt/c/Projects`. Delimit multiple statements with the line feed character.|
|`TimeOut`|The number of seconds to wait for Julia to fully start (including any package precompilation) before `JuliaLaunch` gives up waiting and returns an informational message rather than an error - Julia is not killed, and calling `JuliaLaunch` or `JuliaEval` again once it has finished starting will work normally. If Julia exits during start-up, e.g. because of an invalid command line option, `JuliaLaunch` reports an error straight away. `TimeOut` is optional and defaults to 30.|

**Example.** `=JuliaLaunch(, TRUE, "--threads=4,1")` launches Julia with its window minimised, four worker threads and one interactive thread; then `=JuliaEval("Threads.nthreads()")` returns 4.

## [`JuliaCall`](@id juliacall)
Call a named Julia function, passing in data from the worksheet. Returns an error string for results that cannot be displayed on a worksheet (nested arrays, dictionaries etc.). `JuliaCallVBA` lifts those restrictions.
```
JuliaCall(JuliaFunction, [Arg1], [Arg2], ...)
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function visible from the Julia REPL.|
|`Args...`|Zero or more arguments. Each may be a number, string, Boolean, empty cell, error value, array or range. A range's values are passed as Excel stores them (its `Value2`), so cells formatted as dates arrive as serial numbers, not as Julia `Date`s. An empty cell arrives as `missing`, and an error value such as `#N/A` as an [`ExcelError`](@ref) - see [Marshalling](marshalling.md) for every type.|

**Example.** With the coefficients of three linear equations in A1:C3 and their right-hand sides in D1:D3, `=JuliaCall("\", A1:C3, D1:D3)` solves them, returning the solution as a column. Julia's backslash operator solves the system directly, which is more accurate than Excel's `=MMULT(MINVERSE(A1:C3), D1:D3)`. A trailing `.` on the function name broadcasts: `=JuliaCall("bitstring.", H1:H3)` shows the 64 bits Julia stores for each number, which shows why 0.1 can't be stored exactly.

## [`JuliaCallNamedArgs`](@id juliacallnamedargs)
Call a named Julia function, passing positional and named arguments from the worksheet. The argument `";"` marks where the named arguments start, given as name-value pairs. A named argument whose value is `";"` must be the last pair.
```
JuliaCallNamedArgs(JuliaFunction, [Arg1], ..., ";", Name1, Value1, ...)
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function visible from the Julia REPL.|
|`Args...`|Zero or more positional arguments, then `";"`, then pairs of name and value, e.g. `"a",1,"b",2`. Names must be valid Julia identifiers. Each argument may be a number, string, Boolean, empty cell, array or range.|

Named arguments are what Julia's documentation and error messages call _keyword_ arguments. The `";"` mirrors Julia's own call syntax, so for example `=JuliaCallNamedArgs("sort", A1:A10, ";", "rev", TRUE)` calls `sort(A1:A10; rev=true)`. The `";"` is required, and must be followed by at least one name-value pair, so an even number of arguments. Otherwise `JuliaCallNamedArgs` returns an error saying so, rather than letting the call go ahead with misplaced arguments. To call a function without named arguments, use `JuliaCall`.

Julia finds the `";"` by searching backwards from the end of the arguments, in steps of two (since name-value pairs follow it), and stops at the first `";"` it finds. So a `";"` among the positional arguments is never mistaken for the marker: `=JuliaCallNamedArgs("split", A1, ";", ";", "keepempty", FALSE)` calls `split(A1, ";"; keepempty=false)`. The cost is that if a named argument's _value_ is `";"`, that pair must be given last, since a `";"` value in any earlier pair would be found first and taken as the marker.

A trailing `.` on the function name broadcasts over the positional arguments, as for `JuliaCall`. As with `f.(x; k=v)` in Julia, the named arguments are passed unchanged to every call rather than broadcast over.

## [`JuliaCallHandle`](@id juliacallhandle)
Call a named Julia function and return a [handle](handles.md) to the result, which stays in Julia for other functions to use. An argument `";"` marks where named arguments start, as in `JuliaCallNamedArgs`.
```
JuliaCallHandle(JuliaFunction, [Arg1], [Arg2], ..., [";", Name1, Value1, ...])
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function visible from the Julia REPL. A trailing `.` broadcasts, as for `JuliaCall`.|
|`Args...`|Zero or more positional arguments, optionally followed by `";"` and pairs of name and value for named arguments, as for `JuliaCallNamedArgs`. Handles among them are replaced by their objects.|

**Example.** With a list of values in G1:G20, `=JuliaCallHandle("Set", G1:G20)` in H1 builds a Julia `Set` once, and then `=JuliaCall("in", D1, H1)` tests whether D1's value is in the list, without searching it each time.

Unlike `JuliaCallNamedArgs`, any `";"` argument starts the named arguments, so `";"` can't be passed as a positional argument.

Unlike `JuliaCall`, an error value as an argument, such as `#N/A`, is returned rather than passed to Julia: see [errors pass through](handles.md#Errors-pass-through).

From VBA, `jl.JuliaCallHandle` takes at most 30 arguments after the function name. The handle lasts until `JuliaHandleRelease` or `JuliaHandleReleaseAll` frees it.

## [`JuliaEval`](@id juliaeval)
Evaluate a Julia expression and return the result to an Excel worksheet, or with `ReturnHandle` TRUE, a [handle](handles.md) to the result.
```
JuliaEval(JuliaExpression, [ReturnHandle])
```

|Argument|Description|
|:-------|:----------|
|`JuliaExpression`|Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.|
|`ReturnHandle`|Optional, default FALSE. TRUE returns a handle to the result, which stays in Julia while a formula uses it. The code is evaluated again only when it changes. From VBA, use `JuliaEvalHandle` instead.|

**Example.** `=JuliaEval("string(factorial(big(30)))")` returns 30! exactly, `265252859812191058636308480000000`, using Julia's arbitrary-precision integers. Excel's `=FACT(30)` shows only 15 significant digits.

## [`JuliaInclude`](@id juliainclude)
Load a Julia source file into the Julia process, to make additional functions available via `JuliaEval` and `JuliaCall`.
```
JuliaInclude(FileName)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|The full name of the file to be included.|

The ribbon's Include button does the same for the file in its File box, without needing a formula.

**Example.** If `C:\MyCode\collatz.jl` contains
```julia
collatz(n) = n == 1 ? 0 : 1 + collatz(iseven(n) ? n / 2 : 3n + 1)
```
then after `=JuliaInclude("C:\MyCode\collatz.jl")`, `=JuliaCall("collatz", 27)` returns 111: starting from 27, the Collatz sequence takes 111 steps to reach 1.

## [`JuliaSetVar`](@id juliasetvar)
Set a global variable in the Julia process.
```
JuliaSetVar(VariableName, RefersTo)
```

|Argument|Description|
|:-------|:----------|
|`VariableName`|The name of the variable to be set. Must follow Julia's [rules](https://docs.julialang.org/en/v1/manual/variables/#man-allowed-variable-names) for allowed variable names.|
|`RefersTo`|A range, number, string, Boolean, empty cell or array of such types. As for `JuliaCall`, cells formatted as dates arrive as serial numbers. When called from VBA, nested arrays are supported.|

**Example.** With prices in B1:B5, `=JuliaSetVar("prices", B1:B5)` makes them a Julia variable, and then `=JuliaEval("diff(log.(prices))")` returns their log returns as a column.

## [`JuliaHandleInfo`](@id juliahandleinfo)
Information about object handles: how many are live, or recent events.
```
JuliaHandleInfo(Type)
```

|Argument|Description|
|:-------|:----------|
|`Type`|`"count"` for the number of live handles created by worksheet formulas, and of objects Julia holds (which includes those created from VBA). `"events"` for a column of recent create and release events, newest last.|

## [`JuliaIsRunning`](@id juliaisrunning)
Returns TRUE if an instance of Julia is running and "listening" to the current Excel session, or FALSE otherwise.
```
JuliaIsRunning()
```

## [`JuliaExcelPID`](@id juliaexcelpid)
Returns the process id of the current Excel session. You'll rarely need this directly: `CallFromExcel.serve_xl()` finds the running Excel process automatically when only one is open. It's only when more than one Excel instance is running that you need it, to tell `CallFromExcel.serve_xl` explicitly which one to attach to - see [Attaching to a running Julia session](attaching.md).
```
JuliaExcelPID()
```
