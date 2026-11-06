# Build-TestWorkbook.ps1
# (Re)builds dotnet-addin\vba\TestJuliaExcel.xlsm from its VBA source
# (dotnet-addin\vba\TestJuliaExcel.xlsm\VBA\*.bas) by importing those modules into a fresh
# workbook via COM automation - the source .bas files are what's under version control; the
# .xlsm itself is a build artifact (not committed), rebuilt from source whenever needed.
#
# Mirrors the Import-VbaModules pattern referenced in CLAUDE.md: all comments in the source
# .bas files must be ASCII-only (code points 0-127), since non-ASCII characters become
# mojibake when imported into a workbook via this kind of COM automation.
#
# Prerequisites:
#   - Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#     [x] Trust access to the VBA project object model
#
# Run from PowerShell:
#   dotnet-addin\vba\Build-TestWorkbook.ps1

param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$vbaSourceDir = Join-Path $PSScriptRoot "TestJuliaExcel.xlsm\VBA"
if ($OutputPath -eq "") {
    $OutputPath = Join-Path $PSScriptRoot "TestJuliaExcel.xlsm\TestJuliaExcel.xlsm"
}
$outputPath = $OutputPath

$modules = @("modTestUtils.bas", "modTest.bas")
foreach ($m in $modules) {
    $p = Join-Path $vbaSourceDir $m
    if (-not (Test-Path $p)) {
        Write-Error "Source module not found: $p"
        exit 1
    }
}

if (Test-Path $outputPath) {
    Write-Host "Removing existing $outputPath"
    Remove-Item $outputPath -Force
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # Remove the default worksheet's VBA has nothing to do with sheets, but ensure a known
    # sheet name exists for modTest.bas's ScratchSheet() helper to find/create "Scratch".
    $wb.Worksheets.Item(1).Name = "Main"

    # Reference: Microsoft Scripting Runtime (Dictionary tests) - well-known GUID.
    $wb.VBProject.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0) | Out-Null

    foreach ($m in $modules) {
        $p = Join-Path $vbaSourceDir $m
        Write-Host "Importing $m"
        $wb.VBProject.VBComponents.Import($p) | Out-Null
    }

    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($outputPath, 52)
    $wb.Close($false)
    Write-Host "Built $outputPath"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
