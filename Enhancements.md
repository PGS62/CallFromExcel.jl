# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Named worksheet functions: follow-ups

Named worksheet functions are built (documentation, "Your own worksheet functions"). The design, the spike's results and the decisions taken - the `jl.` prefix, the name checks, when to register, no broadcasting by default, why a formula typed before registration isn't re-entered for the user - are in this file's history up to commit b6b56aa. Still open:

- **Keyword arguments.** Not exposed as worksheet arguments; `JuliaCallNamedArgs` covers them. One option: trailing worksheet arguments for the keywords, named as in the method.
- **`::AbstractMatrix` from a single cell.** An argument declared `::AbstractVector` gets a vector from a cell, row or column; one declared `::AbstractMatrix` gets a single cell as a bare value, so the method doesn't match. Wrap it as a 1x1 matrix.
- **`#NAME?` for formulas typed too soon.** Documented, not fixed (decided 2026-11-25). If it bites: re-enter, after registering, the `#NAME?` cells whose formulas use a just-registered name; or register stubs at add-in load for the names last exposed (cached per project folder), which return "Julia not launched" until Julia is running.
- **Registration from VBA.** `jl.JuliaInclude` and `jl.JuliaEval` don't trigger registration, as the COM server can't queue a macro; VBA runs `Application.Run "JuliaRegisterFunctions"`. The COM server could ask the add-in to queue it.
- **A spilled range of handles.** `handle=true` (done 2026-11-26) gives one handle to the whole result, also when broadcasting. An option for one handle per element - a function returning an array of objects, or broadcast over a range - would spill a range of handles, each with its own lifetime. Only if asked for.
- **`threadsafe=true` (discussed 2026-11-27).** Register the function with `IsThreadSafe`, so Excel's multi-threaded recalculation can call it from several threads at once, for functions doing substantial Julia work over many independent cells. Not small: the C# thread-safety review is the main risk.
  - *Today everything is serial.* Every worksheet function is registered `IsThreadSafe = false`, so Excel calls them one at a time on its main thread, each waiting for its HTTP response; handle create/free (RTD) and VBA's COM calls are on the main thread too, and the add-in starts no background work that calls Julia. In Julia, HTTP.jl 2.16 serves each connection in its own task (`Threads.@spawn :interactive`, `_serve_conn!` in http_server.jl) and a connection's requests one after another; JuliaLaunch's `--threads=auto,1` gives one interactive thread, so all requests run on it. (With a user's own `--threads=N` and no interactive pool, `:interactive` falls back to the default pool - checked - so connections could run in parallel, but the add-in never sends two at once.) So the handles dictionary and the `last_question`/`last_answer`/`args_from_xl` globals need no lock today.
  - *Julia.* Concurrent requests would arrive on separate connections (the add-in's connection pool is already locked) but queue on the single interactive thread. For a speed-up, `srv_call_inner` would run a threadsafe function on the default pool - roughly `fetch(Threads.@spawn f(args...))` - keeping the interactive thread for receiving requests. Then: a `ReentrantLock` around the handles dictionary (threadsafe functions may take handles while other cells create or free them); don't set `last_question`/`last_answer`/`args_from_xl` for threadsafe calls, or make them per task; `display_results` echo would interleave. The user's function must itself be thread-safe - that's what the option asserts.
  - *C#.* `JuliaPort`'s cached port, tokens and version are static fields without locks; `ExcelDnaUtil.IsInFunctionWizard` and anything else that calls into Excel must be checked against what thread-safe XLL functions may call; likewise `NamedFunctions.Forward` and the encode/decode path (believed stateless).
  - *Options.* Reject with `handle=true` (RTD belongs to the main thread), as `volatile` is rejected with `handle`. `volatile=true` with it is fine. Flag in `_expose`/`@excel`, a ninth column of `exposed_functions()`, `IsThreadSafe` in the registration and the flag in the registered signature, as for `volatile`.
  - *Measure first.* For cheap functions the interactive thread's decoding/encoding may be the bottleneck and nothing gained; benchmark with modPerformance-style tests (many cells, Julia work per call from microseconds to tens of milliseconds) before committing to it.
- **Not run by CI.** The Excel drivers in `dotnet-addin/spike-named/` (`NamedSpike.ps1`, `NamedAuto.ps1`, `NamedInspect.ps1`) are run by hand; the VBA suite covers the main paths. Descriptions from docstrings haven't been checked by eye in the Insert Function dialog and IntelliSense (the registration code that shows them is the spike's, which was).

## Object handles: follow-ups

Object handles are built (documentation, "Object handles"). The design discussion, the spike's results and the decisions taken - including why handles are keyed by their inputs rather than by cell - are in this file's history up to commit 7a6e44f. Still open:

- **VBA conveniences.** A handle scope (`jl.HandleScopeBegin`/`End`, releasing every handle created in between). An importable VBA class module whose `Class_Terminate` calls `JuliaHandleRelease`, shipped with a standard module holding `JuliaServer()` (a cached `CreateObject("CallFromExcel.ComServer")`) and a `NewJuliaHandle` factory: VBA frees objects deterministically, so this gives automatic cleanup, except after `End`, a project reset or "End" on an error dialog, which skip `Class_Terminate` (`JuliaHandleReleaseAll` is the safety net). Rejected: a .NET proxy object freed when VBA drops it - .NET learns of that only at a garbage collection, which may be much later.
- **Separate objects for the same inputs.** Today: make the inputs differ. If needed often, a reserved named argument that `handle_create` uses in the key and drops before the call. Only if asked for.
- **"Rebuild all objects".** The macro `JuliaHandleInvalidateAll` does it; a ribbon button, or a friendlier name, would make it findable.
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

- **Documentation section** "Corporate and locked-down machines": worksheet functions work in an elevated Excel but VBA's `CreateObject` doesn't (don't run Excel elevated, or install as that account); and, for when the installer fails under CLM, what it does - copy the add-in files to `%APPDATA%\Microsoft\AddIns\CallFromExcel` and set the listed registry values - so IT can apply it by hand.
- **Make Install.ps1 CLM-safe:** use only cmdlets - read Office's version and Excel's bitness from the registry rather than via COM and the executable's header. A moderate rewrite of the installer; testing needs CLM enabled on a machine (e.g. `__PSLockdownPolicy` or a WDAC policy in a VM).