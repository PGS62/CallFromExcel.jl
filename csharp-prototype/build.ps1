# Builds the Excel-DNA/COM toolchain prototype without MSBuild or the .NET SDK
# (neither is installed on this machine) -- uses csc.exe and RegAsm.exe, which
# ship with the .NET Framework runtime itself.
#
# COM registration (RegAsm /codebase) writes to HKLM and requires an elevated
# ("Run as Administrator") PowerShell session. Everything else does not.

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$bin = Join-Path $root "bin"
$framework = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
$csc = Join-Path $framework "csc.exe"
$regasm = Join-Path $framework "RegAsm.exe"

New-Item -ItemType Directory -Force -Path $bin | Out-Null

Write-Host "Compiling JuliaExcelPrototype.dll ..."
# /optimize+ /debug- matter here, not just as a minor flag: csc's default (neither
# specified) is debug-equivalent, which sets a DebuggableAttribute(DisableOptimizations)
# on the assembly -- this tells the JIT itself to skip native-code optimization even with
# no debugger attached, not just a compiler-level IL difference. An early run of the
# VFormat speed benchmark without this flag understated C#'s speed substantially.
& $csc /nologo /target:library /platform:anycpu /optimize+ /debug- `
    /out:"$bin\JuliaExcelPrototype.dll" `
    /reference:"$root\lib\ExcelDna.Integration.dll" `
    /reference:System.dll `
    "$root\src\AddIn.cs" "$root\src\ComServer.cs" "$root\src\VFormat.cs" "$root\src\GeneralFormat.cs" "$root\src\JuliaHttpClient.cs"
if ($LASTEXITCODE -ne 0) { throw "csc.exe failed with exit code $LASTEXITCODE" }

Copy-Item "$root\lib\ExcelDna.Integration.dll" "$bin\ExcelDna.Integration.dll" -Force
Copy-Item "$root\lib\ExcelDna64.xll" "$bin\JuliaExcelPrototype64.xll" -Force
Copy-Item "$root\JuliaExcelPrototype64.dna" "$bin\JuliaExcelPrototype64.dna" -Force

Write-Host "Registering JuliaExcelPrototype.dll for COM (requires elevation) ..."
& $regasm /codebase /nologo "$bin\JuliaExcelPrototype.dll"
if ($LASTEXITCODE -ne 0) {
    Write-Warning "RegAsm failed (exit $LASTEXITCODE). Re-run this script in an elevated PowerShell window to register the COM server."
} else {
    Write-Host "COM registration OK."
}

Write-Host ""
Write-Host "Build complete. In Excel: File > Options > Add-ins > Manage Excel Add-ins > Go... > Browse..."
Write-Host "and select: $bin\JuliaExcelPrototype64.xll"
Write-Host "Then import vba\Test.bas into a workbook and run TestPrototype."
