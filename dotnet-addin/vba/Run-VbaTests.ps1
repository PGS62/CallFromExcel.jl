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
# It defaults to --project=<this repo>. Afterwards, the Excel this script started is ended if it
# hasn't exited by itself.

param(
    [string]$CommandLineOptions = "",
    [string]$XlPath = ""
)

$ErrorActionPreference = "Stop"

# By default, test this working copy's Julia code rather than whatever version of CallFromExcel the
# default Julia environment has.
if ($CommandLineOptions -eq "") {
    $CommandLineOptions = "--project=" + (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

if ($XlPath -eq "") {
    $XlPath = Join-Path $PSScriptRoot "..\..\workbooks\TestCallFromExcel.xlsm"
}
if (-not (Test-Path $XlPath)) {
    throw "Test workbook not found: $XlPath"
}
$xlPath = (Resolve-Path $XlPath).Path

# The tests' CallFromExcel.* calls go through the COM server, which Windows loads from the
# installed copy of the DLL (registered by Install.ps1), not from dist. So a rebuilt dist that
# hasn't been installed would leave most of the suite testing the old code.
$distDll = Join-Path $PSScriptRoot "..\dist\CallFromExcel.AddIn.dll"
$installedDll = Join-Path $env:APPDATA "Microsoft\AddIns\CallFromExcel\CallFromExcel.AddIn.dll"
if (-not (Test-Path $installedDll) -or
    (Get-FileHash $installedDll).Hash -ne (Get-FileHash $distDll).Hash) {
    throw "The installed CallFromExcel.AddIn.dll differs from dotnet-addin\dist. Close Excel and run dotnet-addin\installer\Install.ps1, then re-run the tests."
}

$excel = New-Object -ComObject Excel.Application
# This Excel's process id, so the finally block below can make sure it exits.
Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
$excelPid = [uint32]0
[void][Win32.User32]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref]$excelPid)
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
    # An Excel started by automation doesn't exit after Quit() while ExcelDna.IntelliSense 1.9
    # is loaded (fixed in Excel-DNA 1.10). Left running, it keeps dotnet-addin\dist's XLL open,
    # so the next build can't refresh dist. End it - this is only ever the Excel started above.
    if ($excelPid -ne 0) {
        $process = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(5000)) {
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
        }
    }
}
