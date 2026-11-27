# The ribbon's Rebuild Objects split button: (b) rebuild every handle's object
# (JuliaHandles.RebuildAll, also the macro JuliaHandleInvalidateAll), and (a) the same followed by
# a full recalculation. A script can't click the ribbon, so this runs the same code: the macro,
# then Application.CalculateFull. "Code changes" are simulated by changing a Julia global that the
# functions read. Checks automatic and manual calculation. Starts its own invisible Excel, with
# the Debug XLL and Excel's default RTD throttle (2 s).
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
$launch = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
$helpers = '=IF(LEN(Z1)>0,JuliaEval("code_version = 1; mk(n) = fill(code_version, Int(n)); plain(x) = code_version * x; ""helpers defined"""))'

$excel = New-Object -ComObject Excel.Application
Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
$excelPid = [uint32]0
[void][Win32.User32]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Retry($sb) {
    for ($i = 0; $i -lt 100; $i++) {
        try { return & $sb } catch { if ($_.Exception.Message -notmatch "rejected|0x8001010A") { throw }; Start-Sleep -Milliseconds 200 }
    }
    throw "Excel stayed busy"
}
function T($a) { Retry { $ws.Range($a).Text } }
function Show($title) { "  {0}: A1 {1} | B1 {2} | C1 {3}" -f $title, (T "A1"), (T "B1"), (T "C1") | Write-Host }
function SetVersion($v) { [void](Retry { $excel.Evaluate("JuliaEval(""global code_version = $v"")") }) }
# Waits up to $seconds for a cell's text to differ from $old.
function WaitChange($a, $old, $seconds) {
    for ($i = 0; $i -lt $seconds * 5; $i++) { if ((T $a) -ne $old) { return $true }; Start-Sleep -Milliseconds 200 }
    return $false
}

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    "RTD throttle: $($excel.RTD.ThrottleInterval) ms" | Write-Host
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    Retry { $ws.Range("Z1").Formula2 = $launch }
    Retry { $ws.Range("Z2").Formula2 = $helpers }
    for ($i = 0; $i -lt 120 -and (T "Z2") -ne "helpers defined"; $i++) { Start-Sleep -Milliseconds 500; Retry { $ws.Range("Z2").Formula2 = $helpers } }
    "  Z1 {0} | Z2 {1}" -f (T "Z1"), (T "Z2") | Write-Host
    Retry { $ws.Range("A1").Formula2 = '=JuliaCallHandle("mk",3)' }
    Retry { $ws.Range("B1").Formula2 = '=JuliaCall("sum",A1)' }
    Retry { $ws.Range("C1").Formula2 = '=JuliaCall("plain",1)' }
    [void](Retry { $excel.Calculate() })
    Show "1 start (expect B1 3, C1 1)"

    Write-Host "=== automatic calculation"
    SetVersion 2
    [void](Retry { $excel.Calculate() })
    Show "2 code changed, F9 (expect no change: B1 3, C1 1)"
    $oldA1 = T "A1"
    [void](Retry { $excel.Run("JuliaHandleInvalidateAll") })
    $changed = WaitChange "B1" "3" 8
    Show "3 (b) rebuild only (expect new A1, B1 6, C1 still 1; changed=$changed)"
    "  A1 changed: " + ($oldA1 -ne (T "A1")) | Write-Host

    SetVersion 3
    [void](Retry { $excel.Run("JuliaHandleInvalidateAll") })
    [void](Retry { $excel.CalculateFull() })
    Show "4 (a) rebuild + CalculateFull, read at once (expect B1 9, C1 3)"
    Start-Sleep -Seconds 3
    Show "4' three seconds later (expect the same)"

    Write-Host "=== manual calculation"
    $excel.Calculation = -4135
    SetVersion 4
    [void](Retry { $excel.Run("JuliaHandleInvalidateAll") })
    Start-Sleep -Seconds 4
    Show "5 (b) rebuild only, 4 s later, no F9 (B1 9 if Excel waits for F9)"
    "  CalculationState: $($excel.CalculationState) (0 done, 1 calculating, 2 pending)" | Write-Host
    [void](Retry { $excel.Calculate() })
    $changed = WaitChange "B1" "9" 6
    Show "6 then F9 (expect B1 12, C1 still 3; changed=$changed)"

    SetVersion 5
    [void](Retry { $excel.Run("JuliaHandleInvalidateAll") })
    [void](Retry { $excel.CalculateFull() })
    Show "7 (a) rebuild + CalculateFull (expect B1 15, C1 5)"
    Start-Sleep -Seconds 3
    Show "7' three seconds later (expect the same)"
    $excel.Calculation = -4105

    $ev = Retry { $excel.Evaluate('JuliaHandleInfo("count")') }
    "  handles: $ev" | Write-Host
    Retry { $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")' }
    $wb.Close($false)
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    try { $excel.Calculation = -4105 } catch {}
    try { $excel.Quit() } catch { Write-Host "Quit failed: $($_.Exception.Message)" }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($excelPid -ne 0) {
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(5000)) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
    }
}
