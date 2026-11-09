# Fixes to back-port to JuliaExcel

Bugs fixed in CallFromExcel that also exist in JuliaExcel (checked against JuliaExcel 2.4.1). JuliaExcel's code differs (VBA add-in rather than C#), so each needs re-implementing there, not copying.

## To do in JuliaExcel 2.4.2

- **`setvar`** (CallFromExcel fa0311a). Confirmed 2026-11-09 by running 2.4.1's `setvar` code: `JuliaSetVar("x", "hello")` fails with `UndefVarError: hello not defined`, and `JuliaSetVar("x", "pi")` sets `x` to the number π. Numbers and arrays, including arrays of strings, are set correctly. The cause is `src/comms.jl` building the assignment as text, `":(global $name = $arg)"`, so a single string's text becomes a variable name. Julia-only fix: assign the value itself, e.g. `Core.eval(Main, :(global $(Symbol(name)) = $arg))`. Test that a string, a string that names an existing variable (`"pi"`), a number and an array all arrive unchanged.

## Candidates

| Fix | CallFromExcel commit | Where in JuliaExcel |
|---|---|---|
| `serve_xl` comms folder: on Windows use `tempdir()` (same `GetTempPath` the add-in uses, which reads TMP before TEMP) instead of `ENV["TEMP"]`; under WSL ask Windows for the temp path instead of trying hard-coded user names | f77cf72 | `src/comms.jl`, `_default_commsfolder` |
| `JuliaLaunch` CommandLineOptions checks look at options, not substrings: `--project=C:/Users/Jean-Luc` was rejected as `-L`, a path containing `-i` suppressed adding `-i`, and similarly for `--threads` | f5f2573 | `vba/JuliaExcel.xlam/VBA/modMain.bas`, `JuliaLaunch` |
| Float ranges sent element by element (as `collect(x)`), not as first/step/length ("RF"): the add-in rebuilt elements as `first + i*step`, which differs in the last bit from Julia's own elements for decimal steps (`0:0.1:1` gave 0.30000000000000004) | 7023172 | `src/encode.jl` `encode_for_xl(::AbstractRange{Float64})`; VBA decoder 'R' branch (modUnserialise.bas) |
| Server rejects requests from web pages: any request with an `Origin` header (browsers always add one; the VBA add-in's HTTP calls never do) gets 403, as does one whose `Host` isn't 127.0.0.1/localhost[:port] (DNS rebinding). Without this, any web page visited while Julia is running can POST code to `127.0.0.1:<port>/eval` and have it run. Julia-only - the per-session token part of the CallFromExcel fix would also need VBA changes, and only adds protection from other users on a shared machine | f516d6a (`_reject_request` in comms.jl) | `src/comms.jl`, the server's request handler |

## Not yet checked against JuliaExcel

Other CallFromExcel fixes that may also apply: unsigned/big/rational numbers returned as text (b9f0cbc). (Quoted PATH entries, d6c5e11, are not needed: VBA's FileExists just skips them rather than failing.)
