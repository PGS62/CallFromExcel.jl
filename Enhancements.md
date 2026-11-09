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

**Problem.** `start_server` tries ports 2700, 2701, ... until a bind succeeds - but a Julia under WSL2 binds inside the Linux VM, which can't see ports in use on the Windows side. With two Excels, one served by a native Windows Julia on 2700 and one launching a WSL Julia, the WSL Julia also gets 2700 (free in Linux), and Excel's requests to `localhost:2700` reach the Windows Julia instead. Since the session token, that Julia rejects them (409), so the second Excel sees Julia as not running and `JuliaLaunch` times out; before the token, calls silently ran in the wrong session. Two WSL Julias don't clash (they share Linux's ports). Judged too contrived to fix for now: it needs native and WSL sessions running at once from different Excels.

**Sketches:**

- **Julia checks it's reachable.** After binding under WSL, Julia checks from the Windows side that `localhost:<port>` reaches itself - e.g. a `/ping` with its own token via `curl.exe`/`powershell.exe` - and moves on to the next port if not.
- **Excel chooses the port.** The add-in picks a port free on Windows and passes it to Julia in the startup script; Julia binds that port (or fails clearly). Changes the startup handshake for every launch, not just WSL.

**Test.** Hard to automate: needs a native Julia listening on a port and a WSL Julia launched at the same time.
