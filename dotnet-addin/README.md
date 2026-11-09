# CallFromExcel's Excel add-in

The C# source of CallFromExcel's Excel add-in, built with [Excel-DNA](https://excel-dna.net/). Users never build this: `Pkg.add` installs the ready-built copy in `dist\`. This page is for anyone changing the add-in. See the [main README](../README.md) for how CallFromExcel is used and how it works.

## Folders

| Folder | Contents |
|---|---|
| `CallFromExcel.AddIn\` | The add-in. `src\JuliaFunctions.cs` holds the worksheet (XLL) functions, `src\ComServer.cs` the VBA-facing COM object (`CreateObject("CallFromExcel.ComServer")`), `src\WireFormatEncode.cs`/`WireFormatDecode.cs` the wire format shared with Julia's `src\encode.jl`/`decode.jl`, and `src\JuliaLaunch.cs` the launching of Julia, natively or under WSL. |
| `CallFromExcel.AddIn.TestHarness\` | Wire-format tests that need no Excel, plus cross-validation against a real Julia process. |
| `dist\` | The built add-in that users get, committed to git, with a hash of the source it was built from. |
| `installer\` | `Install.ps1`, run by `Pkg.build("CallFromExcel")` via `deps\build.jl`. Installs per-user to `%APPDATA%\Microsoft\AddIns\CallFromExcel`, registers the COM server under `HKCU`, and needs no admin rights. |
| `vba\` | `Run-VbaTests.ps1`, which runs the VBA test suite in `workbooks\TestCallFromExcel.xlsm`. |

## Building

```powershell
cd dotnet-addin\CallFromExcel.AddIn
dotnet build -c Release
```

A Release build also runs `Update-Dist.ps1`, which copies the add-in into `dist\` and updates `dist\source-hash.txt`. Commit `dist\` with every source change: `test/runtests.jl` fails if the source no longer matches the hash. To use a new build in Excel, close Excel and run `installer\Install.ps1`.

## Testing

- **Julia tests:** `Pkg.test()` from the repo root. On Windows with Excel installed, this also runs the VBA test suite via `vba\Run-VbaTests.ps1`, in a hidden Excel. That suite covers both the worksheet functions and the COM server, and runs against the *installed* add-in, so install your build first.
- **Wire-format tests:**
  ```powershell
  cd dotnet-addin\CallFromExcel.AddIn.TestHarness
  dotnet build -c Release
  .\bin\Release\net48\CallFromExcel.AddIn.TestHarness.exe          # self-tests + Julia cross-validation
  .\bin\Release\net48\CallFromExcel.AddIn.TestHarness.exe nojulia  # self-tests only
  ```
- **Memory soak test:** `RunSoakTest` in the test workbook's `modSoak` module, run by hand. It charts memory use over many thousands of calls.

## Design notes

- **Late binding from VBA.** Functions taking a variable number of arguments (`JuliaCall` etc.) are `params object[]` methods, which VBA can only call through an object declared `As Object`.
- **Only intended functions are exported.** The `.dna` file sets `ExplicitExports="true"`, so only methods marked `[ExcelFunction]` become worksheet functions. The diagnostic functions in `src\JuliaCallDiag.cs` are hidden from the Function Wizard and IntelliSense.
- **IntelliSense** for the worksheet functions comes from `ExcelDna.IntelliSense`, embedded in the add-in (`src\AddIn.cs`).
- **HTTP, not ZMQ.** The per-call overhead of HTTP is under a millisecond. For large calls, the cost is in encoding, decoding and Excel's own marshalling, so a different transport wouldn't help.
