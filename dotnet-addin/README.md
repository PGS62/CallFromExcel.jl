# JuliaExcel C#/Excel-DNA add-in (dev_csharp_port branch)

Real, SDK-based (not the `csharp-prototype` folder's hand-rolled csc.exe
build) implementation of the VBA-to-C# port described in the August plan.
Lives entirely on `dev_csharp_port`; master is untouched.

## Status as of 2026-11-06 (overnight autonomous session)

All of the August plan's steps 2-5 are done and working:

- **Step 2** (`JuliaHttpPost`/`JuliaLaunch`): `JuliaHttpClient.cs`,
  `JuliaLaunch.cs`, `JuliaExeLocation.cs`, `WindowUtils.cs`,
  `JuliaLiteral.cs`, `JuliaPort.cs`, `FileUtils.cs`. Windows-native only
  (WSL/Linux launch support deliberately deferred).
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

So essentially the full `modMain.bas` worksheet/VBA-facing surface is now
ported, except where explicitly deferred below.

**Not done / explicitly deferred:** WSL/Linux Julia launch; a real
installer (the COM registration below is a dev-session workaround, not a
shippable install story); `modRegister.bas`'s Function-Wizard/IntelliSense
registration (the plan's own view is this should simply disappear in
favour of Excel-DNA's declarative IntelliSense, not be ported -- not yet
wired up, but shouldn't need new logic, just `ExcelArgument`/
`ExcelFunction` attributes, which are already present on every function).

## Validated

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

Earlier in this project (see `csharp-prototype/README.md`), a hand-wired
COM-method demo (`SortViaJulia`) measured `JuliaCall("sort",
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

1. Remove folders from the repo that are no longer needed (not yet
   decided which - likely candidates: `csharp-prototype/` now that this
   folder supersedes it, and/or the old VBA xlam-based folder once
   migration guidance is written).
2. README pass: Shortcomings section rewrite, an honest JuliaInXL
   comparison, and a migration guide covering the no-xlam architecture,
   the `JuliaExcel.JuliaXxx(...)` call-site edit pattern, the early-
   binding gotchas (a `params object[]` method like `JuliaCall` can't be
   called early-bound at all - always use late-bound `Dim j As Object`),
   and the `OneBasedArray`/0-based-`Array()`-spreading gotcha.
3. WSL/Linux Julia launch support, if still wanted.
4. Planned breaking change (once released under the new package name,
   see below): give `JuliaLaunch` a dedicated `project` parameter
   instead of requiring `--project=...` to be embedded in
   `CommandLineOptions`.

**Resolved since the status above was written:**
- A real installer: `installer/Install.ps1` (in this folder) - installs
  to `%APPDATA%\Microsoft\AddIns\JuliaExcel` (no admin needed),
  uninstalls the old xlam if present, no interactive prompts. See its
  own comments for why it writes COM registration directly to
  `HKCU:\Software\Classes\` rather than shelling out to `RegAsm`.
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
