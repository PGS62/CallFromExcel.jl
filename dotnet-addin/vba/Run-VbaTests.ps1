# Run-VbaTests.ps1
# Runs the VBA test suite (modTest.RunTests in workbooks\TestCallFromExcel.xlsm) for the
# C#/Excel-DNA add-in, in a fresh, invisible Excel instance. The workbook is opened and closed
# without saving. No workbook needs to be open already, since RunTests itself registers the XLL
# (Application.RegisterXLL) and creates the COM object it needs.
#
# The workbook's VBA source lives in vba\TestCallFromExcel.xlsm\VBA. Changes to it are pushed
# into the workbook with SolumAddins\scripts\Sync-VbaModules.ps1 (the "VBA: Push" VS Code task).
#
# Prerequisites:
#   - dotnet build -c Release in dotnet-addin\CallFromExcel.AddIn (refreshes dotnet-addin\dist)
#   - CallFromExcel.ComServer registered for COM (dotnet-addin\installer\Install.ps1)
#   - Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#     [x] Trust access to the VBA project object model
#
# -CommandLineOptions is passed straight through to RunTests (and from there to JuliaLaunch).

param(
    [string]$CommandLineOptions = "",
    [string]$XlPath = ""
)

$ErrorActionPreference = "Stop"

if ($XlPath -eq "") {
    $XlPath = Join-Path $PSScriptRoot "..\..\workbooks\TestCallFromExcel.xlsm"
}
if (-not (Test-Path $XlPath)) {
    throw "Test workbook not found: $XlPath"
}
$xlPath = (Resolve-Path $XlPath).Path

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow - enable macros for this OLE-automation-opened workbook, no interactive "Enable Content" prompt possible headlessly

try {
    $wb = $excel.Workbooks.Open($xlPath)

    Write-Host ""
    Write-Host "Running RunTests (SilentMode) ..."
    if ($CommandLineOptions -ne "") {
        Write-Host "CommandLineOptions: $CommandLineOptions"
    }
    $result = $excel.Run("'$xlPath'!modTest.RunTests", $true, $CommandLineOptions)
    $log = $excel.Run("'$xlPath'!modTest.GetTestLog")

    $wb.Close($false)

    Write-Host ""
    Write-Host "--- Test log ---"
    Write-Host $log
    Write-Host "--- End test log ---"
    Write-Host ""
    if ($result) {
        Write-Host "RunTests: ALL TESTS PASSED"
        exit 0
    } else {
        Write-Host "RunTests: ONE OR MORE TESTS FAILED - see the VBA Immediate window for details"
        exit 1
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
