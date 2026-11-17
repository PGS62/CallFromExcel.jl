# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Untrusted workbooks can run Julia code

**Problem.** CallFromExcel's worksheet functions run arbitrary Julia code, but a workbook that uses them needs no macros, so Excel opens it without any macro-security warning. A workbook received from someone else containing `=JuliaLaunch()` and `=JuliaEval("...")` runs that code, with the user's permissions, as soon as it calculates (once the user clicks "Enable Editing", if it opened in Protected View). This is inherent to any add-in whose functions execute code, and JuliaExcel has the same exposure.

**Rejected:** asking with a dialog, once per workbook, before running anything - the user doesn't want dialogs.

**Sketches (not yet evaluated):**

- **Trusted folders setting.** A setting (e.g. a registry value or a config file next to the installed add-in) listing folders whose workbooks may call Julia, analogous to Excel's Trusted Locations. Formulas in other workbooks return an error such as `#Workbook not in a CallFromExcel trusted folder` instead of running. The add-in finds the calling workbook via `XlCall.Excel(XlCall.xlfCaller)` and its sheet reference. Needs a default that doesn't break existing users (e.g. no setting = allow everything, so it's opt-in), which limits its value.
- **Use Excel's own Trusted Locations.** Allow calls only from workbooks in a folder Excel already trusts (read from the registry under `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations`). No new setting, but most users' workbooks aren't in a trusted location, so this would need to be opt-in too.
- **Launch only by the user.** `JuliaLaunch` as a worksheet formula is what lets a workbook start Julia by itself; without it, a workbook can only use a Julia session the user started. An option could make the formula version refuse, leaving launching to VBA, a ribbon button, or `serve_xl`. Reduces exposure only while no Julia session is running.

**Test.** Depends on the approach - e.g. a workbook outside the trusted folders gets the error and Julia receives nothing.

## Another user's program on a freed port

**Problem.** Only matters on a machine shared by several users (terminal server, Citrix). When the Julia session ends, its port file `Port_<excelpid>.txt` stays in the user's comms folder, and its port on 127.0.0.1 is free for anyone to listen on - every user on the machine shares 127.0.0.1. If another user's program takes that port, the Excel's next `JuliaCall`/`JuliaEval` reads the leftover port file and sends it the call - arguments and expressions included - and shows whatever it replies. The session token doesn't help: it proves the caller to the server, not the server to the caller. (Other users can't read the port file, which is in the user's own TEMP, so they'd have to guess the port, but ports are allocated from a small range.)

**Sketches:**

- **Server proves itself too.** The port file gets a second secret, a server token, which Julia sends back as a response header (e.g. `X-CallFromExcel-Server`) on every reply. `JuliaHttpClient` treats a reply without the right value as "Julia not running", so a fake server's results are never shown. About 10 lines across `comms.jl` and `JuliaHttpClient.cs`/`JuliaPort.cs`. Doesn't stop the first request's contents reaching the fake server; preventing that would need a check before sending (e.g. a `/ping` handshake, an extra round trip, which could be done once per port rather than per call).
- **Delete the port file when Julia exits.** An `atexit` hook in Julia removes the port file on a normal exit (including closing its window, if Julia gets the chance to run it), so Excel finds no port file and reports Julia isn't running instead of using the freed port. Doesn't cover a crash or a killed process.
- Related: "Stop serving when Excel closes", above, removes the port file when Excel closes.

**Test.** Start a plain HTTP server on a port, write a port file pointing at it, and check that `JuliaEval` reports Julia isn't running rather than returning that server's reply.

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

## Named worksheet functions

**Problem.** Every Julia function is called through `=JuliaCall("price_swap", A1, B1)`. A demanding user wants `=PRICE_SWAP(A1, B1)`: a real worksheet function with its own name, argument names in the Function Wizard and IntelliSense, and help text. In-house bank add-ins are typically built this way. Probably the feature such users would notice most.

**Sketch.** The calling path is unchanged: a named function forwards to the same code as `JuliaCall`, with the same wire format. What's new is registration.

- **Julia side (small).** A macro or function marks what to expose, e.g. `@excel function price_swap(notional, fixed_rate, curve) ... end` with a docstring, or `CallFromExcel.expose(price_swap; name=..., description=..., args=...)`. It records the Excel name, argument names (from the method, `Base.method_argnames`), descriptions (from the docstring, kept within Excel-DNA's 253/255-character limits) and flags such as volatile in a registry. A new HTTP endpoint returns that list.
- **Add-in side (moderate).** After `JuliaLaunch`, and on request (a ribbon button, or after `JuliaInclude`/`JuliaEval` change the list), the add-in fetches the list and calls `ExcelIntegration.RegisterDelegates`. Each delegate takes the right number of `object` parameters (built with expression trees) and in effect calls `JuliaCall("price_swap", args...)`. Registration must run in a macro context (`ExcelAsyncUtil.QueueAsMacro`). Registering a name again replaces it, so redefining a function in Julia and refreshing works.
- Free once registered: Function Wizard and IntelliSense help, and `Application.Run "PRICE_SWAP"` from VBA.

**Decisions to make.**

- **Workbooks opened before Julia is running.** Cells with `=PRICE_SWAP(...)` show `#NAME?`, and may not recover when the function is registered later. Add-ins that start their language runtime when Excel starts mostly avoid this; we launch on demand. Options: cache the last exposed list (per project folder) and register **stubs** when the add-in loads, which return "Julia not launched" until Julia is running and then forward normally; and/or a "launch Julia when Excel starts" setting.
- **Removing functions.** Excel-DNA can't cleanly unregister, so a removed function stays as a stub returning an error until Excel restarts. Document it.
- **Name clashes** with Excel's built-in functions or other add-ins (Excel's win). Warn, or offer an optional prefix.
- **Several projects or sessions.** Which set applies after relaunching into another environment - probably whatever the current session exposes, plus cached stubs.
- **IntelliSense** should pick up functions registered after load; confirm.

**Effort.** Roughly 1-2 weeks: the Julia macro and endpoint about a day, dynamic registration a few days, the rest on stubs and caching, `#NAME?` behaviour, tests and docs. Purely additive: `JuliaCall` users are unaffected.

**First step: a spike** (about a day). Register one function from Julia at run time and check the Function Wizard, IntelliSense, and what a saved workbook using it does when reopened before and after launch.

**Test.** Julia tests for the registry and endpoint (argument names, descriptions, length limits). VBA tests that call a registered function on a sheet and via `Application.Run`, redefine it and check the new definition is used.

## Object handles

**Problem.** Every result comes back to Excel as data. A large or structured Julia object (a curve, a model, a DataFrame, a portfolio) must be rebuilt on every call that needs it, or be kept in a global by hand (`JuliaEval("global curve = ...")`), which doesn't recalculate with the sheet. The handle pattern common in C++ and other Excel add-ins solves this: a function returns a short string naming an object kept in the add-in's language, and other functions accept that string in place of the object. Combines well with named worksheet functions: `=BUILD_CURVE(...)` returns a handle, `=PRICE(curve, trade)` uses it.

**Sketch.**

- **Creating.** `=JuliaHandle("build_curve", A1:B20)` calls the function as `JuliaCall` does, stores the result Julia-side in a `Dict{String,Any}` (not as globals), and returns a handle such as `jl:Curve#17` (type name plus an integer id).
- **Using.** All existing functions accept handles transparently: Julia's decoder replaces any argument string that **exactly** matches `jl:<TypeName>#<integer>` **and** whose id is in the cache with the cached object. Doing it in the decoder means handles inside arrays (a column of trade handles) also work. An ordinary string can only collide by matching the pattern and an existing id, which is very unlikely.
- **Reading.** `=JuliaCall("identity", D1)` already returns the object as data; a `JuliaObject(handle)` function would be a friendlier name for it.

**Lifetime (most of the work).**

- **Recalculation.** Key each entry by the calling cell (workbook, sheet, address; Excel-DNA's `xlfCaller`). When the cell recalculates, the new object replaces the old, which Julia's GC then frees. The id increases on every recalculation so that the handle string changes and Excel recalculates dependent cells; otherwise they would keep results from the old object.
- **Deleted cells, closed workbooks.** Their entries leak until cleaned up: drop a workbook's entries when it closes (the add-in can watch `WorkbookBeforeClose`), plus `JuliaClearHandles()` or a ribbon button. Show the handle count (and perhaps memory) in the About dialog.
- **Julia restarted.** Every handle is invalid. Functions given one return a clear error, e.g. `#Handle jl:Curve#17 not found - recalculate`; `JuliaLaunch` could offer to recalculate open workbooks.
- **Reopened workbooks.** Saved cells hold handle strings whose objects no longer exist, and Excel won't recalculate them on open unless made to. Same clear error, plus an option to recalculate on launch.
- **Called from VBA.** There is no calling cell: either return the data directly, or require an explicit key argument.

**Decisions to make.** The handle format (prefix, whether to include the type name); whether `JuliaHandle` is a separate function or a flag on `JuliaCall`; whether named functions can declare that they return a handle; the cleanup triggers; whether handles from one workbook may be used in another (probably yes).

**Test.** Julia tests for the cache, the pattern matching (strings that look like handles but aren't), replacement on recalculation and clean-up. VBA tests: create a handle on a sheet, use it from another cell, recalculate and check the dependent cell recalculates, close the workbook and check the cache entry is gone.
