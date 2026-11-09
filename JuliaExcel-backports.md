# Fixes to back-port to JuliaExcel

Bugs fixed in CallFromExcel that also exist in JuliaExcel (checked against JuliaExcel 2.4.1). JuliaExcel's code differs (VBA add-in rather than C#), so each needs re-implementing there, not copying.

| Fix | CallFromExcel commit | Where in JuliaExcel |
|---|---|---|
| `serve_xl` comms folder: on Windows use `tempdir()` (same `GetTempPath` the add-in uses, which reads TMP before TEMP) instead of `ENV["TEMP"]`; under WSL ask Windows for the temp path instead of trying hard-coded user names | (commit that added this file) | `src/comms.jl`, `_default_commsfolder` |
| `JuliaLaunch` CommandLineOptions checks look at options, not substrings: `--project=C:/Users/Jean-Luc` was rejected as `-L`, a path containing `-i` suppressed adding `-i`, and similarly for `--threads` | f5f2573 | `vba/JuliaExcel.xlam/VBA/modMain.bas`, `JuliaLaunch` |
| `setvar` assigns the value itself, rather than re-parsing its printed form (strings became variable names; error values failed) | fa0311a | `src/comms.jl`, `setvar` |

## Not yet checked against JuliaExcel

Other CallFromExcel fixes that may also apply: unsigned/big/rational numbers returned as text (b9f0cbc), quoted PATH entries when finding julia.exe (d6c5e11), the server's per-session token and Origin/Host checks.
