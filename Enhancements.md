# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Named worksheet functions: follow-ups

Named worksheet functions are built (documentation, "Your own worksheet functions"). The design, the spike's results and the decisions taken - the `jl.` prefix, the name checks, when to register, no broadcasting by default, why a formula typed before registration isn't re-entered for the user - are in this file's history up to commit b6b56aa. Still open:

- **Keyword arguments.** Not exposed as worksheet arguments; `JuliaCallNamedArgs` covers them. One option: trailing worksheet arguments for the keywords, named as in the method.
- **`::AbstractMatrix` from a single cell.** An argument declared `::AbstractVector` gets a vector from a cell, row or column; one declared `::AbstractMatrix` gets a single cell as a bare value, so the method doesn't match. Wrap it as a 1x1 matrix.
- **`#NAME?` for formulas typed too soon.** Documented, not fixed (decided 2026-11-25). If it bites: re-enter, after registering, the `#NAME?` cells whose formulas use a just-registered name; or register stubs at add-in load for the names last exposed (cached per project folder), which return "Julia not launched" until Julia is running.
- **Registration from VBA.** `jl.JuliaInclude` and `jl.JuliaEval` don't trigger registration, as the COM server can't queue a macro; VBA runs `Application.Run "JuliaRegisterFunctions"`. The COM server could ask the add-in to queue it.
- **A spilled range of handles.** `handle=true` (done 2026-11-26) gives one handle to the whole result, also when broadcasting. An option for one handle per element - a function returning an array of objects, or broadcast over a range - would spill a range of handles, each with its own lifetime. Only if asked for.
- **Not run by CI.** The Excel drivers in `dotnet-addin/spike-named/` (`NamedSpike.ps1`, `NamedAuto.ps1`, `NamedInspect.ps1`) are run by hand; the VBA suite covers the main paths. Descriptions from docstrings haven't been checked by eye in the Insert Function dialog and IntelliSense (the registration code that shows them is the spike's, which was).

## Object handles: follow-ups

Object handles are built (documentation, "Object handles"). The design discussion, the spike's results and the decisions taken - including why handles are keyed by their inputs rather than by cell - are in this file's history up to commit 7a6e44f. Still open:

- **VBA conveniences.** A handle scope (`jl.HandleScopeBegin`/`End`, releasing every handle created in between). An importable VBA class module whose `Class_Terminate` calls `JuliaHandleRelease`, shipped with a standard module holding `JuliaServer()` (a cached `CreateObject("CallFromExcel.ComServer")`) and a `NewJuliaHandle` factory: VBA frees objects deterministically, so this gives automatic cleanup, except after `End`, a project reset or "End" on an error dialog, which skip `Class_Terminate` (`JuliaHandleReleaseAll` is the safety net). Rejected: a .NET proxy object freed when VBA drops it - .NET learns of that only at a garbage collection, which may be much later.
- **Separate objects for the same inputs.** Today: make the inputs differ. If needed often, a reserved named argument that `handle_create` uses in the key and drops before the call. Only if asked for.
- **"Rebuild all objects".** The macro `JuliaHandleInvalidateAll` does it; a ribbon button, or a friendlier name, would make it findable.
- **Not tested.** A spilled range of handles; thousands of handles over hours of editing; Excel-DNA issue #209 (RTD functions recalculated on cut and paste anywhere in a sheet).

## VBA methods should raise errors (agreed in principle 2026-11-27, breaking)

**Problem.** The COM methods VBA calls (`jl.JuliaCallVBA`, `jl.JuliaEvalVBA`, `jl.JuliaCallNamedArgsVBA`, and the migration methods `jl.JuliaCall`, `jl.JuliaEval`, `jl.JuliaCallNamedArgs`, `jl.JuliaSetVar`, `jl.JuliaInclude`) report a failure - a Julia exception, Julia not running, a timeout - by returning a string such as `"#MethodError: ... Julia REPL has more details and stacktrace!"`. Robust VBA code must test every result for that pattern, and can't tell a failure from a Julia function that legitimately returns such a string. Only `jl.JuliaCallHandle` and `jl.JuliaEvalHandle` raise. The user's view: raising is an improvement worth breaking backward compatibility for; a pre-1.0 minor release, ideally the coming 0.2.0, is the time.

**Sketch.**
- **Tell failure from a string result on the wire, not by pattern.** `srv_call_inner`/`srv_eval_inner` know `success`; today they encode a failure as an ordinary string (comms.jl, `"#$error_desc Julia REPL has more details..."`). Mark it instead - e.g. a distinct wire tag for "call failed" carrying the message, or a header/status on the HTTP response. The worksheet path turns it into the same `#...!` string as now (cells can't raise); the COM path throws. `ComServer.CreateHandle`'s `TrimStart('#')` guess goes away too.
- **Failures on the C# side** (Julia not launched, no answer, timeout, a bad argument such as a 2-D `JuliaExpression`) throw as well, so every failure reaches `On Error`.
- **Err.Number.** A .NET exception reaches VBA with its HRESULT (`InvalidOperationException` gives -2146233079), which says nothing useful. Use one documented number, or a few (Julia exception / Julia not running / timeout), via `COMException` with `vbObjectError + n`, and the message without the leading `#` and trailing `!` as `Err.Description`. Perhaps the Julia exception type at its start, as now (`MethodError: ...`), so code can test it.
- **Excel error values are data, not failures.** A result that is, or contains, `#N/A` etc. stays a `CVErr` value.
- **Which methods.** All the COM methods, for consistency, including the migration methods - which exist so old JuliaExcel code keeps working, and that code tests for `#` strings. Alternative: only the `*VBA` methods, keeping the migration ones as they are. Decide when implementing.
- **Breaking change.** CHANGELOG entry with a before/after VBA example (`If Left(r, 1) = "#" Then` -> `On Error GoTo`); docs `vba.md`/`vba_functions.md`; VBA test suite (tests that expect `#` strings); `Application.Run` of worksheet functions is unaffected.

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

- **Documentation section** "Corporate and locked-down machines": worksheet functions work in an elevated Excel but VBA's `CreateObject` doesn't (don't run Excel elevated, or install as that account); and, for when the installer fails under CLM, what it does - copy the add-in files to `%APPDATA%\Microsoft\AddIns\CallFromExcel` and set the listed registry values - so IT can apply it by hand.
- **Make Install.ps1 CLM-safe:** use only cmdlets - read Office's version and Excel's bitness from the registry rather than via COM and the executable's header. A moderate rewrite of the installer; testing needs CLM enabled on a machine (e.g. `__PSLockdownPolicy` or a WDAC policy in a VM).