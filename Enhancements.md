# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Stop serving when Excel closes

**Problem.** Nothing ends the Julia session when its Excel closes, whether normally or by a crash: `JuliaLaunch` starts Julia in its own interactive console window, and nothing on the Julia side watches Excel. The server keeps listening, and the port file `Port_<excelpid>.txt` stays in the comms folder until some Excel next calls `JuliaLaunch` (which cleans up files of dead process ids). Windows reuses process ids, so a later Excel that happens to get the same id finds that port file - with a valid token - and its `JuliaEval`/`JuliaCall` silently connect to the old session and its leftover state. Rare, but confusing when it happens. JuliaExcel behaves the same way.

**Sketch.** All in Julia (`src/comms.jl`); no add-in or VBA changes. About 25 lines plus about 15 of tests.

- When `start_server` starts serving Excel process `pid` (from `JuliaLaunch` or `serve_xl`), open a handle to that process: `ccall((:OpenProcess, "kernel32"), ...)` with `SYNCHRONIZE` access (`0x00100000`). If that fails (e.g. no such process, as with the made-up pids in the tests), don't watch.
- Start a background task (`@async`/`Threads.@spawn`, as the HTTP server already runs) that sleeps a few seconds (say 5, configurable so tests can shorten it), then asks `WaitForSingleObject(handle, 0)` whether the process has finished. When it has: stop the server, delete the port file, print e.g. "Excel (process id 1234) has closed - server stopped", close the handle and end the task. The REPL stays open, so nothing the user was doing in it is lost.
- Holding a handle, rather than re-checking the id number, also copes with id reuse: the handle refers to that particular process.
- When the server is stopped or restarted for another reason, the old watcher must stand down rather than stop a newer server - e.g. compare a server generation counter, or close the handle in `stop_server`.
- Cost: one system call every few seconds; no CPU while sleeping. During a long computation on Julia's main thread the check is just delayed.
- Windows only. Under WSL Julia can't open a Windows process handle; it would have to poll `tasklist.exe`, which costs more and loses the protection against id reuse - leave WSL as it is.

**Alternative considered.** Exit Julia entirely when Excel closes. Tidier, but loses REPL work, and would also end a `serve_xl` session attached from e.g. VS Code.

**Test.** Start a short-lived dummy process standing in for Excel, attach the server to its pid, end the process, and check that within a few seconds the server has stopped and the port file is gone.

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

## WSL Julia choosing a port already used on Windows

**Problem.** `start_server` tries ports 2700, 2701, ... until a bind succeeds - but a Julia under WSL2 binds inside the Linux VM, which can't see ports in use on the Windows side. With two Excels, one served by a native Windows Julia on 2700 and one launching a WSL Julia, the WSL Julia also gets 2700 (free in Linux), and Excel's requests to `localhost:2700` reach the Windows Julia instead. Since the session token, that Julia rejects them (409), so the second Excel sees Julia as not running and `JuliaLaunch` times out; before the token, calls silently ran in the wrong session. Two WSL Julias don't clash (they share Linux's ports). Judged too contrived to fix for now: it needs native and WSL sessions running at once. Note, though, that the native session needn't still be serving an Excel: since Julia outlives its Excel (see "Stop serving when Excel closes"), a leftover Windows Julia window is enough - this happened on the developer's machine on 2026-11-09, with leftover Julias from test runs on ports 2700-2704 making a WSL `JuliaLaunch` fail.

**Sketches:**

- **Julia checks it's reachable.** After binding under WSL, Julia checks from the Windows side that `localhost:<port>` reaches itself - e.g. a `/ping` with its own token via `curl.exe`/`powershell.exe` - and moves on to the next port if not.
- **Excel chooses the port.** The add-in picks a port free on Windows and passes it to Julia in the startup script; Julia binds that port (or fails clearly). Changes the startup handshake for every launch, not just WSL.

**Test.** Hard to automate: needs a native Julia listening on a port and a WSL Julia launched at the same time.

## Document running as administrator and Constrained Language Mode

**Excel run as administrator.** Not tested; from Windows' documented behaviour:

- Worksheet functions should work: the add-in is registered in the user's own Excel settings (`HKCU\...\Excel\Options`), which an elevated Excel run by the same user reads. A Julia started by `JuliaLaunch` is elevated too.
- VBA's `CreateObject("CallFromExcel.ComServer")` probably fails ("ActiveX component can't create object"): `Install.ps1` registers the COM class per user (`HKCU\Software\Classes`), and Windows ignores per-user COM registrations in elevated processes.
- Elevating as a different account (an admin account's password typed in) runs Excel as that user, for whom the add-in isn't installed at all.
- Quick check: open Excel as administrator and run `?TypeName(CreateObject("CallFromExcel.ComServer"))` in the Immediate window.

**PowerShell Constrained Language Mode (CLM).** Where AppLocker or WDAC locks PowerShell down, scripts from user-writable folders such as `~\.julia\packages\...` run in CLM, which blocks COM objects and most .NET types (`-ExecutionPolicy Bypass` doesn't get round it).

- `Pkg.build("CallFromExcel")` runs `dotnet-addin\installer\Install.ps1`, which uses `New-Object -ComObject Excel.Application` and `[System.IO.File]::OpenRead` with a `BinaryReader` (to find Office's version and Excel's bitness) - these fail with "Cannot create type. Only core types are supported in this language mode", so installation fails. Its registry writes use plain cmdlets, which CLM allows.
- Day-to-day use doesn't involve PowerShell, except `serve_xl` under WSL, which asks `powershell.exe` for `[IO.Path]::GetTempPath()`.

**Sketches:**

- **README section** "Corporate and locked-down machines": worksheet functions work in an elevated Excel but VBA's `CreateObject` doesn't (don't run Excel elevated, or install as that account); and, for when the installer fails under CLM, what it does - copy the add-in files to `%APPDATA%\Microsoft\AddIns\CallFromExcel` and set the listed registry values - so IT can apply it by hand.
- **Make Install.ps1 CLM-safe:** use only cmdlets - read Office's version and Excel's bitness from the registry rather than via COM and the executable's header. A moderate rewrite of the installer; testing needs CLM enabled on a machine (e.g. `__PSLockdownPolicy` or a WDAC policy in a VM).

## Complex numbers with a comma decimal separator

**Problem.** Complex results are returned as text in the format of Excel's `COMPLEX` function (`encode_for_xl(::Complex)` in `src/encode.jl`), always with a point: `"1.5+2i"`. Excel's IM functions read such text using Excel's own decimal separator, so where that is a comma (most of continental Europe) any complex result with a fractional part is #NUM! to them. Verified on 2026-11-09 by switching a hidden Excel's separator to a comma:

| Formula | Point separator | Comma separator |
|---|---|---|
| `=COMPLEX(1.5,2)` | `1.5+2i` | `1,5+2i` |
| `=IMREAL("1.5+2i")` (Julia's text) | 1.5 | #NUM! |
| `=IMREAL("3+4i")` | 3 | 3 |

Workaround meanwhile: `=IMREAL(SUBSTITUTE(JuliaEval(...),".",","))`.

**Why not in Julia.** Only Excel knows the separator in effect: File > Options > Advanced > "Use system separators" can override Windows' setting, it can change during a session, and Julia under WSL or attached with `serve_xl` can't easily find out.

**Sketch.** About 5 lines changed in Julia and 25 in C#.

- Julia sends complex numbers as numbers: a new wire tag, e.g. `X` followed by the real and imaginary parts as 16-character hex each. Infinite or NaN parts still give `!2036` (#NUM!).
- The add-in's decoder formats them as text in `COMPLEX`'s format - 15 significant digits, `i` rather than `1i`, no imaginary part when it's zero - using `Application.DecimalSeparator` when `Application.UseSystemSeparators` is False, and Windows' own separator (`CultureInfo.CurrentCulture`) otherwise. The same for worksheet and VBA results.
- Ordinary strings are never touched, so a Julia string `"1.5+2i"` still comes back unchanged.

**Test.** A VBA test sets `Application.UseSystemSeparators = False` and `Application.DecimalSeparator = ","`, checks that `=JuliaEval("1.5+2im")` gives `"1,5+2i"` and that `IMREAL` of it gives 1.5, and restores both settings even if it fails (they're application settings and may persist). Existing tests cover the point-separator case.

## A ribbon tab

**Status (in progress on the `ribbon` branch).**

- Done: the Julia tab; a Launch Julia split button, launching with saved settings; the Launch Settings dialog (project folder with Browse, other options, packages, bash statements, time-out, saved in the registry, resizable, tooltips).
- To do: Include File, Show Julia Window, Close Julia and Help buttons; the Julia icon on the button and the dialog (pending the logo licence decision); README documentation of the ribbon. In particular, `JuliaLaunch`'s documentation (README and its `[ExcelFunction]` description) should say that it ignores the Launch Settings dialog, which applies only to the ribbon's Launch Julia button.

**Problem.** Julia can only be started by a `JuliaLaunch` formula, and nothing in Excel's own interface says where to find help: the Add-ins dialog shows no description for any third-party XLL. (Only "Help on this function" in the Insert Function dialog and the function name in the IntelliSense pop-up link to the README.) JuliaInXL, Julia Computing's earlier Excel add-in, had a JULIA ribbon tab with "Launch Local Julia", "Select Julia File", "Include Julia File", a `tcp://localhost:9999` connection box, "Reconnect" and "Terminate".

**Proposed buttons.**

- **Launch Julia** (large). Launches with remembered options - see below.
- **Include File.** One step: a file picker, then `JuliaInclude` on the chosen file (JuliaInXL's separate Select and Include steps add nothing).
- **Show Julia Window.** Brings Julia's console to the front, e.g. after launching minimised.
- **Close Julia.** Ends the Julia session, like JuliaInXL's "Terminate".
- **Help.** Opens the GitHub page.

Not JuliaInXL's connection box or Reconnect: they belong to its connection model, not ours (Julia writes its port to the port file).

**Decisions to make.**

- **Launch options.** A button has no arguments. Either launch with defaults, or open a small dialog for `UseLinux`, `MinimiseWindow`, `CommandLineOptions` and `Packages`, remembering the last choices (e.g. in the registry). The dialog is most of the work.
- **Ribbon versus formula.** A `JuliaLaunch` formula travels with the workbook and records how Julia should start; a ribbon click is per session. Both stay; the README would say when to use which.
- **Include from the ribbon** loads the file at once, but Excel won't recalculate cells that use the functions it defines. Either recalculate afterwards or say so in the README.
- **Close Julia** while formulas still depend on it: they will show "Please call JuliaLaunch..." on their next recalculation, which is probably right.
- **Icons.** The Julia dots are licensed CC BY-NC-SA 4.0 by Stefan Karpinski (https://github.com/JuliaLang/julia-logo-graphics). Using them in a free, open-source add-in fits the licence with attribution, the image kept under that licence separately from the MIT code. But "NonCommercial" doesn't match MIT's allowance of commercial use. Options: ask Stefan Karpinski for permission, accept the low practical risk as many Julia projects do, or design our own icon. Don't copy JuliaInXL's other icons, which are Julia Computing's artwork; Office's built-in `imageMso` icons are free to use and look native.

**Sketch.** Excel-DNA supports ribbons directly: a `CustomUI` block in `CallFromExcel.AddIn.dna`, a class deriving from `ExcelRibbon` with one handler per button, and the icon packed into the `.xll`. Most handlers call existing code (`JuliaLaunch.Launch`, the `JuliaInclude` path, `WindowUtils` for the window). Ribbon callbacks run on Excel's main thread, like our worksheet functions. Expect the design to take time to get right: layout, wording, icons, and the options dialog.

**Test.** Mostly by hand, as ribbon clicks can't easily be automated. The handlers' logic can sit in plain methods that the VBA tests call through the COM server.
