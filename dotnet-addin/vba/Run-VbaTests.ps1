# Run-VbaTests.ps1
# Builds (if needed) and runs the VBA test suite (modTest.RunTests) for the C#/Excel-DNA
# add-in, in a fresh, invisible Excel instance - unlike scripts\Run-VbaTests.ps1 (the old
# xlam's runner), this does not require a workbook to already be open, since RunTests itself
# registers the XLL (Application.RegisterXLL) and creates the COM object it needs.
#
# Prerequisites:
#   - dotnet build -c Release in dotnet-addin\CallFromExcel.AddIn (produces the packed .xll)
#   - CallFromExcel.ComServer registered for COM (see dotnet-addin\README.md)
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
    $XlPath = Join-Path $PSScriptRoot "TestCallFromExcel.xlsm\TestCallFromExcel.xlsm"
}
$xlPath = $XlPath
if (-not (Test-Path $xlPath)) {
    Write-Host "Workbook not found - building it first."
    & (Join-Path $PSScriptRoot "Build-TestWorkbook.ps1")
}

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
