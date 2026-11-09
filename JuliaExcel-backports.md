# Fixes to back-port to JuliaExcel

Bugs fixed in CallFromExcel that also exist in JuliaExcel (checked against JuliaExcel 2.4.1). JuliaExcel's code differs (VBA add-in rather than C#), so each needs re-implementing there, not copying.

## Done for JuliaExcel 2.4.2

Committed in JuliaExcel 5db8c52 (Julia-only, so no change to `JuliaExcel.xlam`), with tests:

- **Server rejects requests from web pages** (CallFromExcel f516d6a): a request with an `Origin` header, a `Host` other than 127.0.0.1/localhost[:port] (DNS rebinding), or a method other than POST is refused. Before this, any web page open while Julia was running could POST code to `127.0.0.1:<port>/eval` and have it run. Checked that the add-in's own `MSXML2.ServerXMLHTTP` requests still get through. The per-session token part of the CallFromExcel fix was not back-ported: it needs VBA changes, and only protects against other users on a shared machine.
- **Float ranges** (CallFromExcel 7023172): sent element by element, as `collect(x)`, not as first/step/length ("RF"). The add-in rebuilt elements as `first + i*step`, which differs in the last bit for decimal steps (`0:0.1:1` gave 0.30000000000000004). The VBA decoder's 'R' branch is unchanged: integer ranges still use it.
- **`setvar`** (CallFromExcel fa0311a): assigns the value itself. 2.4.1 built the assignment as text, `":(global $name = $arg)"`, so `JuliaSetVar("x", "hello")` failed with `UndefVarError: hello not defined`, and `JuliaSetVar("x", "pi")` set `x` to the number π.

## Not back-ported

| Fix | CallFromExcel commit | Where in JuliaExcel |
|---|---|---|
| `serve_xl` comms folder: on Windows use `tempdir()` (same `GetTempPath` the add-in uses, which reads TMP before TEMP) instead of `ENV["TEMP"]`; under WSL ask Windows for the temp path instead of trying hard-coded user names | f77cf72 | `src/comms.jl`, `_default_commsfolder` |
| `JuliaLaunch` CommandLineOptions checks look at options, not substrings: `--project=C:/Users/Jean-Luc` was rejected as `-L`, a path containing `-i` suppressed adding `-i`, and similarly for `--threads`. Needs a new `.xlam` | f5f2573 | `vba/JuliaExcel.xlam/VBA/modMain.bas`, `JuliaLaunch` |

## Not yet checked against JuliaExcel

Other CallFromExcel fixes that may also apply: unsigned/big/rational numbers returned as text (b9f0cbc). (Quoted PATH entries, d6c5e11, are not needed: VBA's FileExists just skips them rather than failing.)
