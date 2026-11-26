# Attaching to a running Julia session

Normally [`JuliaLaunch`](@ref julialaunch) starts its own dedicated Julia process. If you'd rather use a Julia session you already have open - for example one running inside VS Code, where you can set breakpoints and step through code - you can attach that session to Excel instead:

```julia
using CallFromExcel
serve_xl()
```

With no argument, [`serve_xl`](@ref) finds the single running Excel process automatically and attaches to it, also switching `display_results` on - since seeing each call arrive is exactly what you want when attached like this for interactive debugging. Pass `false` to leave `display_results` off: `serve_xl(false)`.

If more than one Excel process is running, `serve_xl` won't guess - it throws, telling you to pass the process id explicitly instead: `serve_xl(true, pid)`, getting `pid` from Excel's [`JuliaExcelPID()`](@ref juliaexcelpid) worksheet function.

Back in Excel, `JuliaCall`/`JuliaEval`/etc. now talk to the attached session directly - there's no need to call `JuliaLaunch` at all, though if you do (e.g. from existing VBA code), it will detect the attached session and use it rather than starting a new one.

If Excel is already being served by another Julia session (e.g. one launched via `JuliaLaunch`, or attached earlier with `serve_xl` itself), `serve_xl` throws rather than starting a second, useless listener - stop the other session first (e.g. call `stop_server()` there, or close it).

Other functions for managing a session started this way:
* [`CallFromExcel.stop_server()`](@ref stop_server) stops the session listening to Excel, without closing Julia itself.
* [`CallFromExcel.server_status()`](@ref server_status) reports whether a server is currently running, which Excel process id it's serving, on which port, and the current `display_results` setting.

On the ribbon, Display Results works with an attached session, but Show Julia can't find its window, and Close Julia closes it - Julia and all - rather than just detaching it. To detach, call `stop_server()` in the session instead.

When the Excel being served closes, the session stops serving it by itself, whether it was attached with `serve_xl` or started by `JuliaLaunch`. Julia stays open, so nothing in it is lost. (Not under WSL, where Julia can't watch a Windows process.)
