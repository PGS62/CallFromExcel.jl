# Fixes to back-port to JuliaExcel

Bugs fixed in CallFromExcel that also exist in JuliaExcel (checked against JuliaExcel 2.4.1). JuliaExcel's code differs (VBA add-in rather than C#), so each needs re-implementing there, not copying.

| Fix | CallFromExcel commit | Where in JuliaExcel |
|---|---|---|
| `serve_xl` comms folder: on Windows use `tempdir()` (same `GetTempPath` the add-in uses, which reads TMP before TEMP) instead of `ENV["TEMP"]`; under WSL ask Windows for the temp path instead of trying hard-coded user names | f77cf72 | `src/comms.jl`, `_default_commsfolder` |
| `JuliaLaunch` CommandLineOptions checks look at options, not substrings: `--project=C:/Users/Jean-Luc` was rejected as `-L`, a path containing `-i` suppressed adding `-i`, and similarly for `--threads` | f5f2573 | `vba/JuliaExcel.xlam/VBA/modMain.bas`, `JuliaLaunch` |
| Float ranges sent element by element (as `collect(x)`), not as first/step/length ("RF"): the add-in rebuilt elements as `first + i*step`, which differs in the last bit from Julia's own elements for decimal steps (`0:0.1:1` gave 0.30000000000000004) | (this commit) | `src/encode.jl` `encode_for_xl(::AbstractRange{Float64})`; VBA decoder 'R' branch (modUnserialise.bas) |
| `setvar` assigns the value itself, rather than re-parsing its printed form (strings became variable names; error values failed) | fa0311a | `src/comms.jl`, `setvar` |

## Not yet checked against JuliaExcel

Other CallFromExcel fixes that may also apply: unsigned/big/rational numbers returned as text (b9f0cbc), quoted PATH entries when finding julia.exe (d6c5e11), the server's per-session token and Origin/Host checks.
