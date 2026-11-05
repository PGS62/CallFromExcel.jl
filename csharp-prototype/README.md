# C#/Excel-DNA toolchain prototype (dev_csharp_port branch)

Status: smallest-possible prototype, step 1 of the plan to port JuliaExcel's
VBA client to C#/Excel-DNA. **Not wired to Julia.** This only proves the
toolchain works: one trivial XLL function and one COM-visible class, both
callable from VBA. Lives on the `dev_csharp_port` branch; master is
untouched.

## Why two mechanisms

- `JepAdd` (in `src/AddIn.cs`) is an ordinary Excel-DNA XLL function, called
  from VBA via `Application.Run`. This is the pattern `JuliaCall`/`JuliaEval`
  would use -- but `Application.Run` can only return worksheet-safe value
  types (Double, String, arrays thereof, etc.), never a live object
  reference.
- `JepComServer` (in `src/ComServer.cs`) is a `[ComVisible(true)]` class in
  the *same* DLL, called from VBA via `CreateObject("JuliaExcelPrototype.ComServer")`.
  This is the pattern `JuliaCallVBA`/`JuliaEvalVBA` would need, since those
  can return Dictionaries/nested arrays that `Application.Run` cannot carry.

Both live in the same assembly/AppDomain, as the real port would require.

## Toolchain notes

No Visual Studio, MSBuild, .NET SDK, or `nuget.exe` is installed on this
machine, so this prototype deliberately avoids all of them:

- Compiled directly with `csc.exe` (ships with the .NET Framework runtime
  at `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe` -- no project
  file, no MSBuild).
- `lib/ExcelDna.Integration.dll`, `lib/ExcelDna.xll` and `lib/ExcelDna64.xll`
  are vendored from the `ExcelDna.AddIn` 1.9.0 NuGet package (downloaded
  directly from nuget.org and unzipped -- a `.nupkg` is just a zip), the
  same version already vendored in the repo's top-level `ExcelDNA/` folder.
  No MSBuild-based packing (`ExcelDnaPack`) is used; instead this follows
  Excel-DNA's original loader convention: rename the generic `ExcelDna64.xll`
  loader stub to match a `.dna` manifest of the same base name
  (`JuliaExcelPrototype64.xll` + `JuliaExcelPrototype64.dna`), which tells
  the loader which managed DLL to load.
- COM registration uses `RegAsm.exe` (also ships with the Framework), not
  Excel-DNA's own COM-server support -- this is plain .NET COM interop, to
  keep this first step minimal.

If Visual Studio/MSBuild/the .NET SDK get installed later, this prototype
should be replaced by a real `.csproj` using Excel-DNA's standard
MSBuild-based packing -- the manual approach here is only to unblock today's
validation step.

## Build

```powershell
.\build.ps1
```

Compiles `bin\JuliaExcelPrototype.dll` and assembles
`bin\JuliaExcelPrototype64.xll` + `.dna`. The COM-registration step
(`RegAsm /codebase`) writes to `HKLM` and needs an elevated PowerShell
window -- if you run `build.ps1` without elevation, compilation and XLL
assembly still succeed and the script tells you to re-run elevated for COM
registration.

## Test from Excel/VBA

1. In Excel: File > Options > Add-ins > Manage Excel Add-ins > Go... >
   Browse..., select `bin\JuliaExcelPrototype64.xll`.
2. Import `vba\Test.bas` into any workbook.
3. Run `TestPrototype`. It calls `JepAdd` via `Application.Run` and
   `JepComServer.HelloFromCSharp` via `CreateObject`, and shows both results
   in a message box.

## Go/no-go speed benchmark: is the port even worth it?

Before porting the whole wire format (the big-ticket item in the plan),
`src/VFormat.cs` ports *only* the already-optimized "V" fast path (bulk
Double-array encode/decode -- `TrySerialiseArrayAsV`/`BulkHexOfDoubleArray`
in `modSerialise.bas`, and the `'V'` branch of `Unserialise`/
`BulkDoublesFromHex` in `modUnserialise.bas`) to C#, exposed via
`ComServer.EncodeV`/`DecodeV`. `vba\SpeedTestCSharp.bas`'s
`CSharpVFormatSpeedTest` (run via the `RunCSharpVFormatSpeedTest` wrapper)
times this against the real, unmodified production VBA code on an identical
1,000,000-element `RANDARRAY`, interleaved (mirroring
`VFormatEncodeSpeedTest`/`VFormatDecodeSpeedTest` in the add-in's own
`modPerformance.bas`), and cross-checks that the two wire-format encodings
are byte-identical and cross-decode correctly before trusting the timings.

**Important:** it calls into the real add-in via
`Application.Run("'<path>'!modSerialise.SerialiseElement", ...)`, pinned to
the *installed* add-in's exact path (`C:\Users\Public\JuliaExcel\JuliaExcel.xlam`)
via `GetXlamWorkbook`, not a bare `"JuliaExcel.xlam"` name match. Excel's
`Workbooks` collection is keyed by name: if a second `JuliaExcel.xlam` (e.g.
this repo's own `workbooks\JuliaExcel.xlam`) is ever opened in the same
session, a bare-name lookup could silently resolve to whichever copy loaded
first -- and this prototype came within one keystroke of benchmarking
against the wrong copy of the VBA code for exactly that reason. `GetXlamWorkbook`
fails loudly instead if the loaded `JuliaExcel.xlam` isn't at the expected
installed path.

To run it: have your normal Excel session open (auto-loading the real,
installed add-in as usual) and `JuliaExcelPrototype64.xll` loaded (see
above), import `vba\SpeedTestCSharp.bas`, then run `RunCSharpVFormatSpeedTest`.

**Results (2026-11-05, 1,000,000-element `RANDARRAY`, byte-identical
encodings verified):** encode 6.8x faster (581ms VBA -> 86ms C#), decode
11.5x faster (548ms VBA -> 48ms C#). The decode number only reached that
level after finding and fixing a real bottleneck in the C# itself: an
initial version used a 3-branch-per-hex-character decode loop and only
managed ~4x; replacing it with a branch-free byte-pair lookup table (see
`VFormat.DecodeV`'s own comment) took decode from ~130ms to ~48ms. Isolating
COM marshaling from algorithm cost (`EncodeVDiag`/`DecodeVDiag`, timed
internally with a `Stopwatch`) showed marshaling the 1,000,000-element input
array across the COM boundary was *not* the dominant cost -- the per-element
algorithm was -- which is why that fix mattered so much.

`src/GeneralFormat.cs` / `ComServer.EncodeGeneral`/`DecodeGeneral` is the
same idea applied to VBA's *general* `"*"` per-element path (`SerialiseElement`'s
non-"V" fallback, `Unserialise`'s `Case 42`) -- the path Strings/mixed types
take, which has no bulk-`CopyMemory` equivalent in VBA and so is expected to
show an even larger gap. `CSharpGeneralFormatSpeedTest`
(`RunCSharpGeneralFormatSpeedTest`) runs the same byte-identical-encoding
methodology against an all-String array (`SEQUENCE(1000000)&"x"`).

## Next steps (per the August plan)

1. ~~Smallest possible prototype~~ (this).
2. Port `JuliaHttpPost`/`JuliaLaunch` (HTTP client + process launching) --
   lower-risk, mechanical.
3. Port the wire-format encode/decode -- the hard part -- against the
   current VBA implementation as ground truth.
4. Wire up `JuliaCall`/`JuliaEval` as thread-safe XLL functions; measure
   real latency against today's VBA numbers.
5. Wire up `JuliaCallVBA`/`JuliaEvalVBA` via the COM-visible class.
6. Decide, with real data, whether a Julia-side ZMQ swap is still worth it.
7. Decide on distribution/install story for the new add-in.
