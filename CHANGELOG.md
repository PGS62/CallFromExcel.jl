# Changelog

CallFromExcel is the successor to [JuliaExcel](https://github.com/PGS62/JuliaExcel.jl), and its version numbers start again at 0.1.0. The entries from JuliaExcel 2.4.1 down are JuliaExcel's, kept for the record of how it evolved, including its performance.

## [Unreleased]

### Breaking: failures raise errors in VBA

Functions called from VBA (`jl.JuliaCallVBA`, `jl.JuliaEvalVBA`, `jl.JuliaCall`, `jl.JuliaEval`, `jl.JuliaSetVar`, `jl.JuliaInclude`, `jl.JuliaLaunch` and the rest) now raise an error when they fail, rather than returning a string starting with `#`. `Err.Number` says whether the Julia code threw an exception (`vbObjectError + 1001`), Julia isn't running (`vbObjectError + 1002`) or something else went wrong (`vbObjectError + 1003`), and `Err.Description` gives the message. A string result that merely looks like an error is no longer mistaken for one. Code that tested results, such as `If Left(Result, 1) = "#" Then`, should use `On Error` instead. Worksheet formulas are unchanged. See [Calling from VBA](https://pgs62.github.io/CallFromExcel.jl/stable/vba/#Errors) in the documentation.

### Object handles

Large or structured Julia objects - a big matrix, a fitted model, a DataFrame - can now stay in Julia, with Excel holding a short string naming each one, a _handle_ such as `jl:Array#17@k3x9`. See [Object handles](https://pgs62.github.io/CallFromExcel.jl/stable/handles/) in the documentation.

- `JuliaCallHandle` calls a Julia function and returns a handle to the result. Named arguments follow `";"`, as in `JuliaCallNamedArgs`, and a trailing `.` on the function name broadcasts, as for `JuliaCall`. `JuliaEval` has a new optional argument, `ReturnHandle`, which does the same for an expression.
- A handle can be passed to `JuliaCall` and the other functions in place of its object, alone or in a range or array, and Julia receives the object itself.
- Handles use Excel's RTD (real-time data) mechanism, so an object lives while a formula that creates it exists, and is freed when that formula is edited or deleted, or its inputs change. The same function with the same arguments gives the same object, built once, however many cells use it and however often they recalculate.
- When `JuliaLaunch`, or the ribbon's Launch Julia button, finds a new Julia session, every handle's object is rebuilt in it.
- Your own worksheet functions can return handles: `@excel handle=true fit_model(x, y) = ...` makes `=jl.fit_model(A1:A100, B1:B100)` return a handle to the model, as `JuliaCallHandle` would (`expose_to_excel` takes the same option).
- A handle formula whose arguments change at every recalculation, such as `=JuliaCallHandle("sum", RANDARRAY(1000))`, gives an error saying so, rather than making Excel recalculate it, and create a new object, over and over in automatic calculation.
- From VBA, handles from cells can be passed to `jl.JuliaCall` etc., and `jl.JuliaCallHandle` and `jl.JuliaEvalHandle` create handles that last until `jl.JuliaHandleRelease` or `jl.JuliaHandleReleaseAll` frees them. These methods are at the end of the COM interface, so VBA code compiled against the earlier type library is unaffected.
- `JuliaHandleInfo("count")` shows how many handles are live, and `JuliaHandleInfo("events")` recent activity.
- 2,000 cells each creating a handle calculate in under a second (0.5 to 0.9 seconds in tests), and 2,000 `JuliaCall`s using them in about 0.3 seconds.
- Any string of the form `jl:<label>#<id>@<session>` passed to a function is now taken to be a handle, so a string that merely looks like one gives a "not found" error rather than reaching Julia.

### Your own worksheet functions

A Julia function can now be a worksheet function of its own, `=jl.birthday(A1)` rather than `=JuliaCall("birthday", A1)`. See [Your own worksheet functions](https://pgs62.github.io/CallFromExcel.jl/stable/own_functions/) in the documentation.

- Mark a function with `@excel`, or make an existing one available with `expose_to_excel`. Its Excel name is `jl.` followed by its Julia name unless you choose another, and its description and arguments' descriptions, shown by the Insert Function dialog and the tooltips as you type, come from its docstring.
- Arguments with default values are optional in Excel. Leaving out a required argument is an error naming it.
- Ranges are passed whole by default. With `broadcast=true`, or `broadcast=(:x, :y)` for some arguments, a function written for single values is broadcast over ranges in one call, and its results spill.
- With `volatile=true`, Excel recalculates the function at every calculation, as it does `NOW` and `RAND`: for functions that read state Excel can't see, such as Julia globals, files or the clock.
- The add-in registers exposed functions after `JuliaLaunch`, `JuliaInclude` and `JuliaEval`. A new ribbon button, Register Functions, does it for functions exposed from the Julia REPL, and VBA can run `Application.Run "JuliaRegisterFunctions"`.

### Fixes

- `JuliaInclude` (worksheet, VBA and the ribbon's Include button) loaded the file into the `CallFromExcel` module instead of `Main`, so the functions it defined couldn't be called by name from `JuliaCall` or `JuliaEval`. It now loads into `Main`, as `include` at the REPL does.

## [0.1.7] - 2026-11-17

### Clearer guidance when Julia can't start

When launching Julia from Excel fails because something it needs is missing, the message says what to do, and the ribbon's error dialog offers to copy the command that does it.

- Julia isn't installed on Windows: the command that installs it, from julialang.org. A Julia installed with juliaup while Excel is open is then found without restarting Excel.
- Julia is older than 1.10, which CallFromExcel doesn't support: the command that updates it, rather than a confusing error about loading or installing the package.
- Trying Julia under WSL: when WSL isn't set up, Julia isn't installed under it, or Julia there doesn't have the CallFromExcel package. The package is installed at exactly the add-in's version, which is also what a version mismatch under WSL now suggests.
- A project whose packages aren't installed for the Julia being launched - for example one used from Julia on Windows, then from Julia under WSL, which has its own package store - is now instantiated (`Pkg.instantiate()`) at launch, rather than failing with a long list of precompilation errors.
- Whether CallFromExcel and the `Packages` are installed is checked before loading them, rather than worked out from Julia's error messages.
- When Excel stops waiting for a slow start, its message says that any error will be shown in Julia's window.

### Sessions end cleanly and can't be impersonated

- When the Excel that a Julia session is serving closes (or crashes), the session stops serving it, deletes its port file and says so in its window and title. Before, it kept listening, and since Windows reuses process ids, a later Excel given the same id could silently connect to the old session and its leftover state. Julia itself stays open. Windows only: Julia under WSL can't watch a Windows process.
- Excel now checks that it's talking to its own Julia session, not to another program that has taken the port since that session ended - which, on a machine shared by several users, could be another user's. Each reply from Julia carries a secret from the port file, and a new connection is checked before any request is sent on it, so such a program never receives one. Julia also deletes its port file when it exits.

### Ribbon and installation

- Launch Settings is now its own large button on the ribbon, with a gear icon, rather than on Launch Julia's drop-down menu.
- The ribbon's File list, and the project folder list in Launch Settings, have a Browse... item, and choosing a recent file or folder that no longer exists offers to remove it from the list.
- The About dialog also shows the version of Julia, and its wording is clearer.
- The Excel add-in is listed in Windows' Settings > Apps, from where it can be uninstalled. Uninstalling removes its files, its registry entries and the ribbon's saved settings.
- UI tweaks: the error dialog's button is now "Copy Command", as some commands aren't Julia's; the tooltip for "Minimise Julia's window" mentions Show Julia.

## [0.1.6] - 2026-11-14

### Changes

- Temporary environments: with `--project=@temp`, Julia starts in a new, empty environment that it deletes when it exits, and any of `JuliaLaunch`'s `Packages` that aren't installed are installed into it - preferring versions already installed, so they're already precompiled. A convenient way to try out packages from Excel without adding them to a project. Needs Julia 1.11 or later; with Julia 1.10, `JuliaLaunch` says so, rather than Julia quietly using a permanent environment named "temp".
- Outside a temporary environment, a package in `Packages` that isn't installed gives a message naming it and its environment, and Launch Julia's error dialog offers to copy the `Pkg.add` command for it.
- `JuliaLaunch` reports a project folder (`--project`) that doesn't exist. Julia accepts one, as a new empty project, so the mistake went unnoticed.
- When Julia's process exits before start-up finishes, the message shows the command line options it was given, with a link to Julia's documentation of them - and doesn't suggest them as the cause when there were none.
- In Launch Settings, the project folder remembers up to 10 recent folders, as the ribbon's File box does (which now also remembers 10, up from 5). Its list also has "(Default environment)" and "(Temporary environment)", and "command line options" links to Julia's documentation of them.
- Fixed a long `IOError: write: operation canceled` error printed in Julia's window just after a successful launch, seen on a machine where Julia's first request took over 2 seconds to compile. Julia now handles a request of its own before telling Excel it's ready, and `JuliaLaunch` waits longer for it.
- `installme()` refuses, saying to restart Julia, when the session loaded CallFromExcel before the package was updated. It would otherwise install the add-in for the old version, or fail to find it if updating had deleted it.
- UI tweaks: Julia's window shows the Julia icon as soon as it opens; Excel shows its busy mouse pointer while Launch Julia is starting Julia; the error dialogs' sound now plays reliably; shorter tooltips for the Help and About buttons.

## [0.1.5] - 2026-11-13

### Changes

- `JuliaCall` and its relatives work with an anonymous function given as the function name, e.g. `=JuliaCall("(x->2x)", 1)`, which failed with "method too new to be called from this world context". Each such call defines a new function in Julia, so it's best not repeated many times.

## [0.1.4] - 2026-11-13

### Changes

- When Launch Julia on the ribbon fails because the Julia package is a different version from the Excel add-in, or isn't installed, the error dialog has a Copy Julia Command button, which copies the command that fixes it to the clipboard.

## [0.1.3] - 2026-11-13

### Changes

- `JuliaLaunch` checks that the Julia package is the same version as the Excel add-in before doing anything else, so a mismatch gives the usual message saying how to fix it, rather than an error from calling something an older package doesn't have (such as `UndefVarError(:set_console_icon...)`). Under WSL, the message also says where to run each step, since `installme()` only works in Julia on Windows.
- `JuliaLaunch(TRUE)` (Julia under WSL) no longer times out when a Julia on Windows, e.g. one serving another Excel, is already listening on the port Julia under WSL would choose. Julia under WSL can't see Windows' ports, so `JuliaLaunch` now picks the port.
- The Julia ribbon tab's controls are in one group, CallFromExcel.jl, rather than three.
- An About button on the Julia ribbon tab shows the add-in's version, the CallFromExcel version in the Julia session it is connected to, and a link to the project on GitHub.

## [0.1.2] - 2026-11-13

### Changes

- VBA code can use a reference to CallFromExcel (VBA editor, Tools menu -> References) and `Dim jl As New CallFromExcel.ComServer`, so the VBA editor shows each function's arguments as you type, as it did with JuliaExcel. `CreateObject("CallFromExcel.ComServer")` still works without a reference.
- From VBA, `JuliaCall`, `JuliaCallVBA`, `JuliaCallNamedArgs` and `JuliaCallNamedArgsVBA` now take at most 30 arguments after the function name. An array created with VBA's `Array` function now arrives in Julia as one array even when it's the only argument; previously it arrived as separate arguments.
- When the Excel add-in is newer than the Julia package, the version-mismatch messages (in Excel, and from `using CallFromExcel` in the REPL) now say to update the package before running `installme()`, rather than just to run `installme()`, which would downgrade the add-in. And the message in Excel says how to start Julia in the right environment: plain `julia` for Julia's default environment, or `julia --project="..."` naming the project folder Julia was launched with.

## [0.1.1] - 2026-11-12

### Changes

- Now that CallFromExcel is in Julia's General registry, the installation instructions, and `JuliaLaunch`'s message when the package isn't installed, use `Pkg.add("CallFromExcel")`.
- CallFromExcel works with HTTP.jl 2 as well as 1.
- `using CallFromExcel` in the REPL reminds you to run `CallFromExcel.installme()` if the Excel add-in isn't installed, or is a different version from the package.

## [0.1.0] - 2026-11-12

First release of CallFromExcel. If you're upgrading from JuliaExcel, see [Upgrading from JuliaExcel](https://pgs62.github.io/CallFromExcel.jl/stable/installation/#upgrading-from-juliaexcel).

### Changes

- The Excel add-in is rewritten in C#, using [Excel-DNA](https://excel-dna.net/), replacing JuliaExcel's VBA add-in `JuliaExcel.xlam`. The worksheet functions keep their names (`JuliaCall`, `JuliaEval`, `JuliaLaunch` etc.), and data now transfers between Excel and Julia two to three times faster, and about eight times faster for large DataFrames - see Performance below.
- From VBA, functions are called through an object created with `CreateObject("CallFromExcel.ComServer")`, rather than through a reference to the add-in.
- The Julia package is renamed `CallFromExcel`, so for example `JuliaExcel.serve_xl()` becomes `CallFromExcel.serve_xl()`.
- Installing the package doesn't install the Excel add-in: install it with `using CallFromExcel; CallFromExcel.installme()`, after installing and after each update. Installation shows no dialogs, needs no admin rights, and uninstalls `JuliaExcel.xlam` if present. The add-in checks that it's the same version as the Julia package it talks to, and if not, says so and how to fix it.
- A Julia tab on Excel's ribbon launches Julia (with saved settings, including a project folder), shows or closes its window, switches `display_results` on and off, and includes Julia files chosen in a file box with a list of recent files. See [The Julia ribbon tab](https://pgs62.github.io/CallFromExcel.jl/stable/ribbon/).
- `JuliaLaunch` opens Julia in a classic console window showing the Julia logo, in its title bar and taskbar button, rather than in Windows Terminal (where the default terminal is Windows Terminal, as on Windows 11).
- "Help on this function" in Excel's Insert Function dialog now opens the function's documentation on GitHub, as does clicking the function's name in its IntelliSense pop-up.
- CallFromExcel requires 64-bit Excel. JuliaExcel also supported 32-bit Excel.
- CallFromExcel requires Julia 1.10 or later. JuliaExcel required 1.7 or later, but its `JuliaLaunch` always passed `--threads=auto,1`, which Julia only accepts from 1.9.
- `JuliaLaunch` gives a clear error message, with installation instructions, if the CallFromExcel package isn't installed in Julia's active environment.

### Performance

| Test | JuliaExcel 2.4.1 | CallFromExcel 0.1.0 | Speed-up |
|------|------------------|---------------------|----------|
| Latency: `JuliaEval("1+1")` | 0.65 ms | 0.21 ms | 3.1x |
| Two-way: `JuliaCall("identity", column of 100,000 Doubles)` | 0.11 s | 0.043 s | 2.5x |
| One-way (Excel to Julia): `JuliaCall("sum", column of 100,000 Doubles)` | 0.053 s | 0.019 s | 2.8x |
| One-way (Julia to Excel): `JuliaEval("collect((1:100000).*pi)")` | 0.057 s | 0.020 s | 2.9x |
| One-way (Julia to Excel): `JuliaEval("(1:100000).*pi")` | 0.025 s | 0.017 s | 1.5x |
| Two-way: `JuliaCall("identity", column of 100,000 mixed values)` | 0.30 s | 0.085 s | 3.5x |
| One-way (Excel to Julia): `JuliaCall("length", column of 100,000 mixed values)` | 0.11 s | 0.034 s | 3.1x |
| One-way (Julia to Excel): `JuliaEval("[isodd(i) ? i*pi : string(i) for i in 1:100000]")` | 0.27 s | 0.11 s | 2.3x |
| Two-way: `JuliaCall("sort", column of 1,000,000 Doubles)` | 1.23 s | 0.42 s | 2.9x |
| One-way (Julia to Excel): `JuliaEval("DataFrame(rand(Float64,(5000,1024)),:auto)")` | 10.9 s\* | 1.34 s | 8.1x |

*Each test times the recalculation of a worksheet formula, averaged over several recalculations and two runs. The "mixed values" cycle through a string, a Boolean and a Double. All figures measured on the same machine as JuliaExcel 2.0.0's: Intel Core Ultra 9 288V, 32GB RAM. The DataFrame row is different: it times a single call to `JuliaEval` from VBA (`DataframePerformanceTest` in `workbooks/TestCallFromExcel.xlsm`), and \*its JuliaExcel figure is for version 2.1.2, the last version that test was run on.*

Since those measurements, Julia decodes large ranges of numbers passed to `JuliaCall` about seven times faster: for 1,000,000 Doubles, from about 77 ms to about 11 ms. This takes about 70 ms off `=JuliaCall("sort", A1:A1000000)`.

## JuliaExcel (predecessor), versions 2.4.1 and earlier

Released as [JuliaExcel.jl](https://github.com/PGS62/JuliaExcel.jl).

## [JuliaExcel 2.4.1] - 2026-11-01

### Bug fixes

- The add-in's VBA project is named `JuliaExcel` again. In versions 2.3.1 and 2.4.0 it was named `VBAProject` by mistake, which breaks VBA code that calls the add-in's functions as, for example, `JuliaExcel.JuliaCall(...)`, and references to the add-in from a workbook's VBA project.
- Fixed links to Julia's documentation that had moved, in the README and in `JuliaLaunch`'s error message for invalid `CommandLineOptions`.

## [JuliaExcel 2.4.0] - 2026-11-01

### Changes

- Complex numbers now return to Excel as text in the format of Excel's `COMPLEX` function, e.g. `"3+4i"`, so that Excel's `IMREAL`, `IMSUM` etc. can work with them. Previously they returned as Julia's display text, e.g. `"3 + 4im"`, which those functions don't accept. Details [here](https://pgs62.github.io/CallFromExcel.jl/stable/marshalling/#Complex-numbers-return-as-text).

## [JuliaExcel 2.3.1] - 2026-11-01

### Bug fixes

- Excel no longer changes the capitalisation of JuliaExcel's worksheet functions, e.g. turning `JuliaCall` into `juliacall`, when they're typed into a formula.

## [JuliaExcel 2.3.0] - 2026-11-01

### Changes

- `JuliaCallNamedArgs` and `JuliaCallNamedArgsVBA` now require the `";"` argument, followed by a positive even number of arguments (the name-value pairs). Previously a call without a `";"`, or with an odd number of arguments after it, quietly treated every argument as positional, so a mistake such as `=JuliaCallNamedArgs("sort", A1:A10, ";", "rev", FALSE, "foo")` returned a misleading `MethodError` for `sort`. Such calls now return an `ArgumentError` explaining what's wrong.
- JuliaExcel now requires Julia 1.7 or later.

### Bug fixes

- `serve_xl` (and the server it starts) now creates the comms folder (`%TEMP%\@JuliaExcel`) if it doesn't already exist, rather than failing to write its port file. Normally Excel creates that folder, so this only went wrong on a machine where JuliaExcel had never been used from Excel - including the GitHub Actions test runner, where the Julia tests had been failing since `serve_xl` was introduced in 2.1.0.
- GitHub Actions workflow updated to versions of `actions/checkout` and `julia-actions/setup-julia` that run on Node.js 24, in place of the deprecated Node.js 20.

## [JuliaExcel 2.2.0] - 2026-10-31

### New features

- New worksheet function `JuliaCallNamedArgs`, and VBA function `JuliaCallNamedArgsVBA`, for calling Julia functions that take named (keyword) arguments. The argument `";"` marks where the named arguments start, mirroring Julia's own call syntax, and they're given as name-value pairs: `=JuliaCallNamedArgs("sort", A1:A10, ";", "rev", TRUE)` calls `sort(A1:A10; rev=true)`. Julia finds the `";"` by searching backwards from the end of the arguments, so a `";"` among the positional arguments is never mistaken for it. A named argument whose value is `";"` must therefore be the last pair. Example [here](https://pgs62.github.io/CallFromExcel.jl/stable/functions/#juliacallnamedargs).
- New `JuliaExcel.kwargs_from_xl`, holding the named arguments most recently passed to `JuliaCallNamedArgs`/`JuliaCallNamedArgsVBA`, alongside `args_from_xl` for the positional ones.

## [JuliaExcel 2.1.3] - 2026-10-17

### Bug fixes

- Fixed `JuliaEval` of multi-line expressions, i.e. a one-column range with one Julia statement per cell. The cells were previously joined with `;` and parsed with `Meta.parse`, so a `#` comment in one cell commented out the code in every cell after it, and multi-line blocks such as `function ... end` spread across several cells could fail to parse. The cells are now joined with newlines, exactly as typed, and parsed with `Meta.parseall`, so comments and multi-line blocks behave as they would in a Julia script. `JuliaExcel.answer_again()` also now uses `Meta.parseall`.

## [JuliaExcel 2.1.2] - 2026-09-16

### Performance improvements

- Julia's HTTP responses to Excel now include an explicit `Content-Length` header. Previously, without it, HTTP.jl fell back to chunked transfer-encoding, and the VBA client's `MSXML2.ServerXMLHTTP` appears to reassemble a chunked body by repeatedly growing and copying its buffer - quadratic in the size of the response, so large results returned from `JuliaCall`/`JuliaEval` were far worse than linearly slower than small ones. Setting `Content-Length` lets it read directly into a preallocated buffer instead, restoring linear performance. For example, `JuliaEval("DataFrame(rand(Float64,(5000,1024)),:auto)")` fell from 241 seconds to 10.9 seconds - a 22x speedup - with smaller results also improving (e.g. 128 columns: 3.0s to 1.2s).

## [JuliaExcel 2.1.1] - 2026-08-25

### Bug fixes

- Fixed a `StringIndexError` when strings terminating in a multi-byte character were passed from Excel to Julia, e.g. `JuliaCall("identity", "xxx£")`. Present since v0.2.17, the first version to use the wire format for the Excel -> Julia direction.

## [JuliaExcel 2.1.0] - 2026-08-24

### New features

- New `display_results` function: switch it on to echo, in the Julia console, both the expression or function call arriving from Excel (`question>`) and the value being returned to it (`answer>`) - useful for seeing exactly what `JuliaCall`/`JuliaEval` are doing. Example [here](https://pgs62.github.io/CallFromExcel.jl/stable/#Demo).
- Clearer Julia console output when a `JuliaCall`/`JuliaEval` call fails: the error and the expression that caused it are now shown in colour, making them easier to spot among other console output.
- `JuliaExcel.last_question` and `JuliaExcel.last_answer` record the most recent expression or function call from Excel and the value returned to it, whichever of `JuliaCall`/`JuliaEval` was used. A new `JuliaExcel.answer_again()` re-evaluates `last_question`, and you can wrap it with `@enter` or your own debugging tools.
- New `JuliaExcel.serve_xl` function: attaches a Julia session that's already running (e.g. one open in VS Code) to Excel, instead of always launching a fresh process via `JuliaLaunch`. Call `JuliaExcel.serve_xl()` to attach to the single running Excel process automatically, or give the process id explicitly if more than one is running - get it from Excel's new `JuliaExcelPID()` worksheet function. Also adds `JuliaExcel.stop_server()` and `JuliaExcel.server_status()` for managing a session attached this way.
- `JuliaLaunch` and `JuliaCall`/`JuliaEval` are more resilient to the Julia session behind them changing or disappearing mid-session, e.g. after `JuliaExcel.stop_server()`, or a session attached via `JuliaExcel.serve_xl` on a different port.

## [JuliaExcel 2.0.0] - 2026-08-19

### Breaking changes

- Excel errors passed into Julia now decode to a new `ExcelError` type, not a `String`. Previously, an Excel error value (e.g. `#DIV/0!`) passed as an argument to `JuliaCall`/`JuliaEval` arrived in Julia as a string like `"#ExcelError2007!"`. It now arrives as `ExcelError(2007)` - a distinct type. This means:
  - A function like `identity` now correctly returns the *same* Excel error, round-tripping unchanged (previously it came back as literal text).
  - A function that isn't written to expect an error and tries to operate on one now fails immediately with a clear `MethodError`, rather than silently treating error text as ordinary string data.
  - Code relying on the old string representation will need updating.
- Dictionaries passed from Excel now decode to a concretely-typed `Dict{K,V}` when every key and every value share a type, rather than always `Dict{Any,Any}`. Code dispatching narrowly on `Dict{Any,Any}` should broaden to `Dict`/`AbstractDict`.

### New features

- Large numeric arrays now transfer substantially faster between Excel and Julia, in both directions - every number is packed into a fixed number of characters with no per-element bookkeeping, and the core conversion between numbers and that packed form happens as one bulk operation on both sides, rather than looping over each number individually.
- Friendlier `JuliaCall`/`JuliaEval` error messages: an expression that errors now returns a short summary to Excel (e.g. `MethodError: no method matching +(::Int64, ::String)...`) instead of a raw exception dump like `#(MethodError(+, (1, "1"), 0x000000000000979f))!`. The full error and stacktrace are printed to the Julia console.
- Julia `Range`s (`UnitRange`, `StepRange` etc.) now transfer to Excel far faster regardless of size - only the first value, step, and length cross the wire, and VBA reconstructs the array from those.
- `Byte` support: VBA's `Byte` type now round-trips to Julia's `UInt8` (previously unsupported in both directions).

### Performance

| Test | v0.2.16 | v0.2.17 | v1.0 | v2.0 |
|------|---------|---------|------|------|
| Latency: `JuliaEval("1+1")` | 7.2 ms | 6.4 ms | 1.1 ms | 1.2 ms |
| Two-way: `JuliaCall("identity", vector of 100,000 Doubles)` | 1.56 s | 0.37 s | 0.27 s | 0.10 s |
| One-way (Excel to Julia): `JuliaCall("sum", vector of 100,000 Doubles)` | 1.45 s | 0.11 s | 0.11 s | 0.054 s |
| One-way (Julia to Excel): `JuliaEval("collect((1:100000).*pi)")` | 0.31 s | 0.26 s | 0.17 s | 0.041 s |
| One-way (Julia to Excel): `JuliaEval("(1:100000).*pi")` | 0.31 s | 0.26 s | 0.17 s | 0.014 s |

*All figures measured on the same machine: Intel Core Ultra 9 288V, 32GB RAM*

## [JuliaExcel 1.0.1] - 2026-08-14

### Changes

* Simplifies installation.
* Adds a compat entry for DataFrames, to meet registration requirements.

## [JuliaExcel 1.0.0] - 2026-08-13

### Changes

This version changes how JuliaExcel works, from file-based communication between Excel and Julia (versions 0.2.17 and earlier) to a local HTTP connection: an HTTP.jl server on the Julia side, called from Excel/VBA using a component built into Windows for making HTTP requests (MSXML2.XMLHTTP.6.0).

Combined with the changes already made in v0.2.17, this reduces latency and increases throughput between Excel and Julia:

| Test | v0.2.16 | v0.2.17 | v1.0 |
|---|---|---|---|
| Latency:</br>`JuliaEval("1+1")` | 7.2 ms | 6.4 ms | 1.1 ms |
| Two-way:</br>`JuliaCall("identity", vector of 100,000 Doubles)` | 1.56 s | 0.37 s | 0.27 s |
| One-way (Excel to Julia):</br>`JuliaCall("sum", vector of 100,000 Doubles)` | 1.45 s | 0.11 s | 0.11 s |
| One-way (Julia to Excel):</br>`JuliaEval("collect((1:100000).*pi)")` | 0.31 s | 0.26 s | 0.17 s |

These figures were measured on a Windows laptop (Intel Core Ultra 9 288V, 32GB RAM). To test on your PC: set the workbook to not be an add-in (e.g. from the VBE's Immediate window, `Workbooks("JuliaExcel.xlam").IsAddin = False`), then click the "Check Performance!" button - results are printed to the VBA Immediate window (Ctrl+G in the VBE).

Other changes in this release:

  * When a call made via `JuliaCall` or `JuliaCallVBA` raises an error, the arguments passed from Excel are now saved to `JuliaExcel.args_from_xl` before the call is attempted, and the error message printed in the Julia console includes a ready-made expression for reproducing the call. So you can debug the failure directly at the Julia REPL, without needing to repeat the call from Excel.

  * `JuliaLaunch` now adds `--threads=auto,1` to `CommandLineOptions` by default, giving Julia's HTTP server its own dedicated thread rather than sharing one with other work in the session.

  * Removes the restriction on how Windows Terminal is configured (#9).

  * Removes the VBA functions `JuliaResultFile` and `JuliaUnserialiseFile` and the previously-exported Julia functions `htd` and `hts`.

## [JuliaExcel 0.2.17] - 2026-08-07

### Changes

This version has better performance when passing large arrays of data between Julia and Excel. On a test laptop:

#### v0.2.16
`JuliaCall("sum",RANDARRAY(250000))` executes in 5900 milliseconds
`JuliaCall("sum",RANDARRAY(500000))` returns an error

#### v0.2.17
`JuliaCall("sum",RANDARRAY(250000))` executes in 330 milliseconds
`JuliaCall("sum",RANDARRAY(500000))` executes in 725 milliseconds

## [JuliaExcel 0.2.16] - 2025-12-28

### Changes

- **Improved Floating-Point Precision**
Communication of 32-bit and 64-bit floating-point numbers between Excel/VBA and Julia now preserves full precision.
*Example:* Previously, `=JuliaCall("identity",SQRT(2))-SQRT(2)` returned `4.88498E-15`; it now returns `0`.
- **Improved Array Support**
Multi-dimensional arrays (up to 8 dimensions) can now be passed between VBA and Julia. Earlier versions supported only 1D and 2D arrays.
- **Improved Dictionary Handling**
Dictionaries can now be passed in both directions: VBA ↔ Julia. Previously, only Julia → VBA was supported.
- **Regional Settings Compatibility**
JuliaExcel now works correctly when the Windows decimal symbol is not `.`.

## [JuliaExcel 0.2.15] - 2025-10-30

### Changes

* Communication between Julia and Excel (which is file-based) now has retry logic for better reliability, especially in tight loops or high-frequency calls.

## [JuliaExcel 0.2.12] - 2025-10-16

### Changes

- This version is compatible with Julia 1.12.
- Julia variables of unsupported type are now returned to Excel as their string representation, via `Base.show`.
- Improved handling of strings containing characters with high code points.

## [JuliaExcel 0.2.10] - 2023-12-03

### Changes

In prior versions, two functions (`GetCurrentProcessID` and `IsWindow`) were made visible as worksheet functions when they should not have been. This version corrects that.

## [JuliaExcel 0.2.9] - 2023-09-20

### Changes

Now launches Julia in interactive mode, i.e. with the `-i` command line option. The benefit is that when using [OhMyREPL](https://github.com/KristofferC/OhMyREPL.jl), `Ctrl + R` works correctly to see command history.

## [JuliaExcel 0.2.8] - 2022-07-12

### Changes

1) When [JuliaLaunch](https://github.com/PGS62/JuliaExcel.jl#julialaunch) is called, `julia.exe` must first be located: if it's on the PATH, the first one found there is used; otherwise, we search for the most recently modified copy of `julia.exe` in the locations where the Windows Julia installer places it. This change adds further such locations, in response to #10.
2) The value returned to Excel is no longer displayed in the Julia REPL.

## [JuliaExcel 0.2.7] - 2022-03-07

### Changes

If the Julia code being called from Excel throws an exception, then a stack trace of the error is displayed in the Julia window (uses [showerror](https://docs.julialang.org/en/v1/base/io-network/#Base.showerror)). Particularly helpful when the Julia code is under development.

## [JuliaExcel 0.2.6] - 2021-12-18

### Changes

Fixes a bug in VBA method JuliaEvalVBA.

## [JuliaExcel 0.2.5] - 2021-12-15

### Changes

Now compatible with Excel 2013

## [JuliaExcel 0.2.4] - 2021-12-10

### Changes

Now possible to run Julia on Ubuntu Linux running under Windows Subsystem for Linux.

The arguments to `JuliaLaunch` have changed in a not backwards-compatible way, so any existing calls to `JuliaLaunch` in workbooks or VBA code will need to be amended when upgrading to this version.

## [JuliaExcel 0.2.3] - 2021-12-02

### Changes

Dictionaries now handled.

If `JuliaEvalVBA` and `JuliaCallVBA` evaluate on the Julia side to `result` where `typeof(result) <: AbstractDict`, then in VBA the return from the function is of type `Scripting.Dictionary`.

## [JuliaExcel 0.2.2] - 2021-11-19

### Changes

Now launches Julia with the `--threads=auto` command line option.

## [JuliaExcel 0.2.1] - 2021-11-17

The first documented release.

## [JuliaExcel 0.1.0] - 2021-11-06

First release. No documentation yet!
