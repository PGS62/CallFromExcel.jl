# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Named worksheet functions

**Problem.** Every Julia function is called through `=JuliaCall("price_swap", A1, B1)`. A demanding user wants `=PRICE_SWAP(A1, B1)`: a real worksheet function with its own name, argument names in the Function Wizard and IntelliSense, and help text. In-house bank add-ins are typically built this way. Probably the feature such users would notice most.

**Sketch.** The calling path is unchanged: a named function forwards to the same code as `JuliaCall`, with the same wire format. What's new is registration.

- **Julia side (small).** A macro or function marks what to expose, e.g. `@excel function price_swap(notional, fixed_rate, curve) ... end` with a docstring, or `CallFromExcel.expose(price_swap; name=..., description=..., args=...)`. It records the Excel name, argument names (from the method, `Base.method_argnames`), descriptions (from the docstring, kept within Excel-DNA's 253/255-character limits) and flags such as volatile in a registry. A new HTTP endpoint returns that list.
- **Add-in side (moderate).** After `JuliaLaunch`, and on request (a ribbon button, or after `JuliaInclude`/`JuliaEval` change the list), the add-in fetches the list and calls `ExcelIntegration.RegisterDelegates`. Each delegate takes the right number of `object` parameters (built with expression trees) and in effect calls `JuliaCall("price_swap", args...)`. Registration must run in a macro context (`ExcelAsyncUtil.QueueAsMacro`). Registering a name again replaces it, so redefining a function in Julia and refreshing works.
- Free once registered: Function Wizard and IntelliSense help, and `Application.Run "PRICE_SWAP"` from VBA.

**Broadcasting over broadcastable arguments.** `JuliaCall("f.", ...)` broadcasts over every argument; in Julia, a function `f(a, b, c)` whose `c` is an array is broadcast over `a` and `b` alone as `f.(a, b, Ref(c))`. Named functions can do that automatically, so the Julia author writes `f` for scalars and ranges are handled correctly and consistently, with no `.` for the user to remember.

- **Which arguments broadcast.** Declared at registration, e.g. `@excel broadcast=(:a, :b) function f(a, b, c) ... end`: the listed arguments broadcast, the rest are passed whole. This suits the common style of leaving arguments unannotated unless needed for dispatch. Where there is no list, declared types decide: an argument declared as an array type (`::AbstractVector`, `::AbstractMatrix`) is passed whole, any other argument broadcasts. So without a list or annotations everything broadcasts, as `JuliaCall("f.", ...)` does today. `broadcast=false` opts out entirely.
- **The rule per call.** An array given to a broadcasting argument is broadcast over; every other value is wrapped in `Ref` (harmless for scalars, and it protects structs that don't support broadcasting). If nothing is broadcast, it's a plain call, so the result isn't a zero-dimensional array. A column of handles given to a broadcasting argument (see Object handles in the README) broadcasts over the objects.
- **In Julia, not by Excel's lifting.** Excel can lift an add-in function over a range itself if its parameters are registered as scalar types, but it calls the function once per cell: one HTTP round trip each. Registering every parameter as a generic value and broadcasting in Julia makes one round trip and one compiled broadcast.
- **Consistent where Excel isn't.** Excel lifts a built-in function only when its parameter is registered as a scalar value, and many aren't: `=SQRT(A1:A3)` spills, but `=COMPLEX(A1:A3, B1:B3)` and `=EDATE(A1:A3, 1)` (Analysis ToolPak functions, whose parameters take references) give `#VALUE!` unless coerced with `+0`. Named Julia functions follow one rule; a surprised user can read the answer off the registration.
- **Conversions for arguments passed whole.** A single cell arrives as a scalar and an `n x 1` range as a `Matrix`; for an argument declared `::AbstractVector`, wrap or `vec` it, so the Julia author never sees the difference. With no declared type, pass what arrives.
- **Shapes.** Julia's rules match Excel's for compatible shapes: a column against a row gives a table, as `=A1:A10*B1:F1` does. For incompatible shapes (10 rows against 7) Excel pads with `#N/A`; don't copy that - give a clear error naming the two arguments.
- **Functions returning arrays.** Broadcasting them gives an array of arrays, which can't spill: a clear error, suggesting `broadcast=false`.
- **Several methods.** The method marked `@excel` defines the registered signature.
- **Optional arguments.** Trailing Julia defaults become optional Excel arguments, arriving as missing - separate from broadcasting, but part of the same registration.
- **Possible bonus for `JuliaCall`.** `JuliaCall("f.", ...)` could follow the same rule for a function with a single method whose declared types say which arguments are arrays. A change of behaviour, so separate and optional.

**Decisions to make.**

- **Workbooks opened before Julia is running.** Cells with `=PRICE_SWAP(...)` show `#NAME?`, and may not recover when the function is registered later. Add-ins that start their language runtime when Excel starts mostly avoid this; we launch on demand. Options: cache the last exposed list (per project folder) and register **stubs** when the add-in loads, which return "Julia not launched" until Julia is running and then forward normally; and/or a "launch Julia when Excel starts" setting.
- **Removing functions.** Excel-DNA can't cleanly unregister, so a removed function stays as a stub returning an error until Excel restarts. Document it.
- **Name clashes** with Excel's built-in functions or other add-ins (Excel's win). Warn, or offer an optional prefix.
- **Several projects or sessions.** Which set applies after relaunching into another environment - probably whatever the current session exposes, plus cached stubs.
- **IntelliSense** should pick up functions registered after load; confirm.

**Effort.** Roughly 1-2 weeks: the Julia macro and endpoint about a day, dynamic registration a few days, the rest on stubs and caching, `#NAME?` behaviour, tests and docs. Purely additive: `JuliaCall` users are unaffected.

**First step: a spike** (about a day). Register one function from Julia at run time and check the Function Wizard, IntelliSense, and what a saved workbook using it does when reopened before and after launch.

**Test.** Julia tests for the registry and endpoint (argument names, descriptions, length limits). VBA tests that call a registered function on a sheet and via `Application.Run`, redefine it and check the new definition is used. Julia tests for broadcasting: the `broadcast=` list, classification by declared type, `Ref` wrapping, plain call when nothing broadcasts, a column against a row, the shape-mismatch and array-of-arrays errors, the scalar-to-vector and `n x 1` conversions.

## Object handles: follow-ups

Object handles are built (README, "Object handles"). The design discussion, the spike's results and the decisions taken - including why handles are keyed by their inputs rather than by cell - are in this file's history up to commit 7a6e44f. Still open:

- **VBA conveniences.** A handle scope (`jl.HandleScopeBegin`/`End`, releasing every handle created in between). An importable VBA class module whose `Class_Terminate` calls `JuliaHandleRelease`, shipped with a standard module holding `JuliaServer()` (a cached `CreateObject("CallFromExcel.ComServer")`) and a `NewJuliaHandle` factory: VBA frees objects deterministically, so this gives automatic cleanup, except after `End`, a project reset or "End" on an error dialog, which skip `Class_Terminate` (`JuliaHandleReleaseAll` is the safety net). Rejected: a .NET proxy object freed when VBA drops it - .NET learns of that only at a garbage collection, which may be much later.
- **Separate objects for the same inputs.** Today: make the inputs differ. If needed often, a reserved named argument that `handle_create` uses in the key and drops before the call. Only if asked for.
- **"Rebuild all objects".** The macro `JuliaHandleInvalidateAll` does it; a ribbon button, or a friendlier name, would make it findable.
- **Decisions for named worksheet functions.** Whether one can declare that it returns a handle; whether a function returning an array of objects gives one handle or a spilled range of handles, one per element.
- **Not tested.** A spilled range of handles; thousands of handles over hours of editing; Excel-DNA issue #209 (RTD functions recalculated on cut and paste anywhere in a sheet).
## Untrusted workbooks can run Julia code

**Problem.** CallFromExcel's worksheet functions run arbitrary Julia code, but a workbook that uses them needs no macros, so Excel opens it without any macro-security warning. A workbook received from someone else containing `=JuliaLaunch()` and `=JuliaEval("...")` runs that code, with the user's permissions, as soon as it calculates (once the user clicks "Enable Editing", if it opened in Protected View). This is inherent to any add-in whose functions execute code, and JuliaExcel has the same exposure.

**Rejected:** asking with a dialog, once per workbook, before running anything - the user doesn't want dialogs.

**Sketches (not yet evaluated):**

- **Trusted folders setting.** A setting (e.g. a registry value or a config file next to the installed add-in) listing folders whose workbooks may call Julia, analogous to Excel's Trusted Locations. Formulas in other workbooks return an error such as `#Workbook not in a CallFromExcel trusted folder` instead of running. The add-in finds the calling workbook via `XlCall.Excel(XlCall.xlfCaller)` and its sheet reference. Needs a default that doesn't break existing users (e.g. no setting = allow everything, so it's opt-in), which limits its value.
- **Use Excel's own Trusted Locations.** Allow calls only from workbooks in a folder Excel already trusts (read from the registry under `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations`). No new setting, but most users' workbooks aren't in a trusted location, so this would need to be opt-in too.
- **Launch only by the user.** `JuliaLaunch` as a worksheet formula is what lets a workbook start Julia by itself; without it, a workbook can only use a Julia session the user started. An option could make the formula version refuse, leaving launching to VBA, a ribbon button, or `serve_xl`. Reduces exposure only while no Julia session is running.

**Test.** Depends on the approach - e.g. a workbook outside the trusted folders gets the error and Julia receives nothing.

## Document running as administrator and Constrained Language Mode

**Excel run as administrator.** Not tested; from Windows' documented behaviour:

- Worksheet functions should work: the add-in is registered in the user's own Excel settings (`HKCU\...\Excel\Options`), which an elevated Excel run by the same user reads. A Julia started by `JuliaLaunch` is elevated too.
- VBA's `CreateObject("CallFromExcel.ComServer")` probably fails ("ActiveX component can't create object"): `Install.ps1` registers the COM class per user (`HKCU\Software\Classes`), and Windows ignores per-user COM registrations in elevated processes.
- Elevating as a different account (an admin account's password typed in) runs Excel as that user, for whom the add-in isn't installed at all.
- Quick check: open Excel as administrator and run `?TypeName(CreateObject("CallFromExcel.ComServer"))` in the Immediate window.

**PowerShell Constrained Language Mode (CLM).** Where AppLocker or WDAC locks PowerShell down, scripts from user-writable folders such as `~\.julia\packages\...` run in CLM, which blocks COM objects and most .NET types (`-ExecutionPolicy Bypass` doesn't get round it).

- `CallFromExcel.installme()` runs `dotnet-addin\installer\Install.ps1`, which uses `New-Object -ComObject Excel.Application` and `[System.IO.File]::OpenRead` with a `BinaryReader` (to find Office's version and Excel's bitness) - these fail with "Cannot create type. Only core types are supported in this language mode", so installation fails. Its registry writes use plain cmdlets, which CLM allows.
- Day-to-day use doesn't involve PowerShell, except `serve_xl` under WSL, which asks `powershell.exe` for `[IO.Path]::GetTempPath()`.

**Sketches:**

- **README section** "Corporate and locked-down machines": worksheet functions work in an elevated Excel but VBA's `CreateObject` doesn't (don't run Excel elevated, or install as that account); and, for when the installer fails under CLM, what it does - copy the add-in files to `%APPDATA%\Microsoft\AddIns\CallFromExcel` and set the listed registry values - so IT can apply it by hand.
- **Make Install.ps1 CLM-safe:** use only cmdlets - read Office's version and Excel's bitness from the registry rather than via COM and the executable's header. A moderate rewrite of the installer; testing needs CLM enabled on a machine (e.g. `__PSLockdownPolicy` or a WDAC policy in a VM).