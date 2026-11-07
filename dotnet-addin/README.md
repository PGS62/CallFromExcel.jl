# JuliaExcel C#/Excel-DNA add-in (dev_csharp_port branch)

Real, SDK-based (not the `csharp-prototype` folder's hand-rolled csc.exe
build) implementation of the VBA-to-C# port described in the August plan.
Lives entirely on `dev_csharp_port`; master is untouched.

## Status as of 2026-11-07

All of the August plan's steps 2-5 are done and working:

- **Step 2** (`JuliaHttpPost`/`JuliaLaunch`): `JuliaHttpClient.cs`,
  `JuliaLaunch.cs`, `JuliaExeLocation.cs`, `WindowUtils.cs`,
  `JuliaLiteral.cs`, `JuliaPort.cs`, `FileUtils.cs`. Includes launching
  Julia under WSL (`UseLinux`/`BashStatements`).
- **Step 3** (wire-format encode/decode): `WireFormatEncode.cs`,
  `WireFormatDecode.cs`, `Hex.cs`, `WireValues.cs`. Every type the VBA
  wire format supports, arbitrary array rank (1-9), the "V" bulk-Double
  fast path, the Julia-to-Excel-only Range ("R") format.
- **Step 4** (`JuliaCall`/`JuliaEval` as XLL functions): `JuliaFunctions.cs`.
  Also includes `JuliaCallNamedArgs` (`/callkw`), `JuliaIsRunning`,
  `JuliaExcelPID`, `JuliaSetVar`, `JuliaInclude`, `JuliaLaunch` itself.
- **Step 5** (`JuliaCallVBA`/`JuliaEvalVBA` via COM): `ComServer.cs`
  (also `JuliaCallNamedArgsVBA`), `WireValues.cs`'s COM-visible
  `JuliaDictionary`.

So the full `modMain.bas` worksheet/VBA-facing surface is ported. Also
done: a non-interactive, no-admin installer (`installer/Install.ps1`) and
formula-bar IntelliSense embedded in the add-in itself (replacing
`modRegister.bas`). Remaining work is in "Suggested next steps" below.

## Validated

- **55/55** VBA tests in `vba/TestJuliaExcel.xlsm` (`vba/Run-VbaTests.ps1`,
  also run by `test/runtests.jl` locally), covering both the COM and the
  worksheet/XLL surfaces. WSL: `modTest.RunWSLTest`, run by hand.
- **58/58** self-contained wire-format tests (`JuliaExcel.AddIn.TestHarness`,
  `WireFormatTests.cs`), including the exact example strings hand-transcribed
  from `modUnserialise.bas`'s own wire-format spec comment.
- **18/18** cross-validation tests against a real, freshly-launched Julia
  process (`JuliaCrossValidation.cs`) -- round-trips every type through
  Julia's own `identity()`, so this validates against Julia's actual
  `encode_for_xl`/`decode_from_xl`, not just this code's self-consistency.
- Full end-to-end real Excel testing (headless, via COM automation, since
  no one was around to keep a normal Excel session open): `JuliaLaunch`,
  `JuliaEval`, `JuliaCall` all work correctly as real worksheet functions,
  including a genuine `=JuliaCall("sort",A1#)` worksheet formula (not just
  `Application.Run`) against a 1,000,000-element `RANDARRAY`. Friendly
  "please call JuliaLaunch first" message and Julia-side errors
  (undefined function) both surface correctly rather than crashing.
- `JuliaCallVBA`/`JuliaEvalVBA` verified via real late-bound COM dispatch
  (PowerShell's dynamic COM object, same `IDispatch::Invoke` mechanism VBA
  uses): multiple separate arguments (`JuliaCallVBA("+", 3, 4)`,
  4-argument calls), Dictionary results (`Count`/`Keys`/`Items`/`Item`/
  `Exists` all correct), and nested-array results (array of arrays,
  recursing into the "V" fast path per inner array).
- `JuliaCallNamedArgs`/`JuliaCallNamedArgsVBA` verified via Julia's own
  `round(x; digits=n)`; `JuliaIsRunning`/`JuliaExcelPID` verified correct;
  `JuliaSetVar` verified by reading the variable back via `JuliaEval`.

## Performance: a real, nuanced finding worth reading before quoting a number

Earlier in this project, a hand-wired COM-method demo (`SortViaJulia`, in a
`csharp-prototype` folder now removed - see git history) measured `JuliaCall("sort",
RANDARRAY(1000000))`-equivalent at ~330-450ms, beating Excel's native
`SORT()` (~667ms). **That comparison does not hold for the real, fully-wired
XLL-based `JuliaCall`.**

Measured via both `Application.Run` and genuine worksheet-formula
recalculation (confirmed not a COM-call artifact): the real
`JuliaCall("sort", RANDARRAY(1000000))` takes **~720-770ms**. Root-caused
precisely via `JuliaCallDiag.cs`'s isolating diagnostic functions:

| Phase | Cost |
|---|---|
| Encode (C#, "V" fast path) | ~70-80ms |
| HTTP + Julia's own sort/encode/decode | ~200-230ms |
| Decode (C#, "V" fast path) | ~35-40ms |
| **Excel-DNA's own XLOPER<->CLR marshaling of the 1M-element array** (input + output, measured separately) | **~435ms** (237ms in + 200ms out) |
| **Total** | **~720-770ms** |

The marshaling cost is inherent to an XLL function accepting/returning a
generically-typed (`object`) array this large -- not specific to this
implementation, not fixable by further optimizing this project's own code,
and not present in the earlier demo because that demo returned a
strongly-typed `double[,]` via a direct COM method call, a different
(and evidently cheaper) marshaling path than a genuine Excel-DNA XLL
function return.

**Bottom line:** still a genuine ~1.7-1.9x improvement over today's VBA
`JuliaCall` (~1273-1473ms measured for the same case), but it does not
clearly beat Excel's own native `SORT()` the way the earlier demo
suggested. Worth knowing before repeating the "beats native SORT()"
claim -- it was true for the demo's calling convention, not for the real
worksheet-formula path.

## Known housekeeping items

- Testing left several orphaned `julia.exe` processes running (Julia is a
  separate process from Excel by design, so it doesn't auto-exit when a
  test's Excel instance quits) -- harmless (just memory, a few hundred MB
  each), but check Task Manager if you want to clear them. One of them
  (PID varies, the one with no corresponding recent Excel/test activity)
  is your own VS Code Julia session, left untouched throughout.
- `JuliaExcel.AddIn.ComServer` is registered for COM under
  `HKCU:\Software\Classes` (user-level, no admin needed), pointing at this
  session's build output path
  (`dotnet-addin\JuliaExcel.AddIn\bin\Release\net48\JuliaExcel.AddIn.dll`).
  If you rebuild to a different path, or on a different machine/account,
  re-register (see `vba\TestRealAddin.bas`'s own setup notes).

## How to build and test

```powershell
cd dotnet-addin\JuliaExcel.AddIn
dotnet build -c Release
```

Produces `bin\Release\net48\publish\JuliaExcel.AddIn64-packed.xll`
(load via Excel > File > Options > Add-ins > Manage Excel Add-ins >
Go... > Browse...).

For the wire-format test suite (no Excel/COM needed at all):

```powershell
cd dotnet-addin\JuliaExcel.AddIn.TestHarness
dotnet build -c Release
.\bin\Release\net48\JuliaExcel.AddIn.TestHarness.exe        # self-tests + Julia cross-validation
.\bin\Release\net48\JuliaExcel.AddIn.TestHarness.exe nojulia # self-tests only, no Julia needed
```

For manual testing from real VBA: see `vba\TestRealAddin.bas`.

## Suggested next steps

1. Remove `workbooks/`, `scripts/` (old-world demo/tooling workbooks and
   the old xlam's manual VBA test runner - `vba/JuliaExcel.xlam/` itself
   is already gone from this branch) and the rest of `vba/`'s sibling
   workbooks, once migration guidance is written and the old world is
   actually retired. As of now, `workbooks/` and `scripts/` only
   reference each other - nothing in the active build/test pipeline
   touches either any more, since `test/runtests.jl` no longer does.
2. README pass: Shortcomings section rewrite, an honest JuliaInXL
   comparison, and a migration guide covering the no-xlam architecture,
   the `JuliaExcel.JuliaXxx(...)` call-site edit pattern, the early-
   binding gotchas (a `params object[]` method like `JuliaCall` can't be
   called early-bound at all - always use late-bound `Dim j As Object`),
   and the `OneBasedArray`/0-based-`Array()`-spreading gotcha.
3. Planned breaking change (once released under the new package name,
   see below): give `JuliaLaunch` a dedicated `project` parameter
   instead of requiring `--project=...` to be embedded in
   `CommandLineOptions`.

**Resolved since the status above was written:**
- WSL/Linux Julia launch: `JuliaLaunch`'s `UseLinux`/`BashStatements`
  now work (via `wsl bash -lc ...`, so juliaup's PATH entry is picked
  up). Tested by hand with `modTest.RunWSLTest`, deliberately not part
  of `RunTests` - it hangs when run headless via `Run-VbaTests.ps1`
  (not investigated), and switching rapidly between native and WSL
  sessions can briefly leave the WSL port unreachable (a WSL2 port-
  forwarding quirk that doesn't arise at human pace). WSL's Julia needs
  this repo's dependencies instantiated once before first use.
- A real installer: `installer/Install.ps1` (in this folder) - installs
  to `%APPDATA%\Microsoft\AddIns\JuliaExcel` (no admin needed),
  uninstalls the old xlam if present, no interactive prompts. See its
  own comments for why it writes COM registration directly to
  `HKCU:\Software\Classes\` rather than shelling out to `RegAsm`.
  `src/comms.jl`'s `installme()` and `deps/build.jl` both now point at
  it instead of the old repo-root `installer/Install.ps1`, which - along
  with `csharp-prototype/`, `ExcelDNA/` and `vba/JuliaExcel.xlam/` (all
  only ever used by that old installer/prototype/xlam) - has been
  removed from this branch (still on `master`, for anyone who hasn't
  upgraded).
- `test/runtests.jl`'s local VBA testset now runs this folder's own
  `vba/Run-VbaTests.ps1` suite (the one with real worksheet/XLL-path
  coverage, see above), not the old xlam's.
- Excel-DNA was silently registering every public static method of
  every public static class in the assembly as a worksheet function
  (its own default, `ExplicitExports="false"`) - not just the intended
  API, but internal plumbing like `JuliaPort.SetJuliaPort`, directly
  callable and able to corrupt the add-in's own state from any cell.
  Fixed with `ExplicitExports="true"` in the `.dna`; the 4 genuinely-
  intentional diagnostic functions (`JuliaCallDiag.cs`) got
  `IsHidden = true` instead, keeping them callable but out of the
  Function Wizard/IntelliSense.
- IntelliSense (`modRegister.bas`'s old Function-Wizard mechanism) is
  done - `ExcelDna.IntelliSense` is now embedded directly in this add-in
  (see `src/AddIn.cs`), no separate IntelliSense add-in needed.
- ZMQ-vs-HTTP.jl transport: investigated and decided against. Fixed
  per-call latency is ~0.5-0.9ms (already negligible); for a
  1,000,000-element `sort`, the dominant cost is Julia's own wire-format
  en/decode work and genuine ~32MB hex-text data transfer, not HTTP
  protocol overhead - so a transport swap wouldn't meaningfully help
  either a "few large calls" or "many small calls" usage pattern.

**Naming/release plan:** this rewrite will be released as a new Julia
package, `CallFromExcel.jl`, not a new version of `JuliaExcel.jl` -
Julia's compat/SemVer system can't protect existing VBA users from a
breaking update (the "dependent" is invisible VBA code, not a Julia
package), and `JuliaExcel.jl`'s own registration has separately been
stalled on the Julia General registry's "no 'julia' in the name" rule
with no maintainer engagement. The existing `JuliaExcel.jl` repo stays
as-is for existing users, with its README updated to point at the new
package once released.
