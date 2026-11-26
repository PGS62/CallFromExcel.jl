# Object-handles spike: Julia restarted while Excel stays open. JuliaLaunch notices the new
# session and rebuilds the handles (JuliaHandles.AfterLaunch -> JuliaHandleInvalidateAll).
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
$launch = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
# Refers to the launch cell, so it is recalculated - redefining mkrand - whenever that is, as a
# workbook's own JuliaEval/JuliaInclude cells would be.
$helpers = '=IF(LEN(Z1)>0,JuliaEval("mkrand(n)=rand(Int(n)); ""helpers defined"""))'

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
function Settle { Retry { $excel.Calculate() }; Start-Sleep -Milliseconds 300; Retry { $excel.Calculate() } }
function Ev($title) {
    Settle
    $ev = Retry { $excel.Evaluate('JuliaHandleInfo("events")') }
    Write-Host ""
    Write-Host "=== $title   [$($excel.Evaluate('JuliaHandleInfo("count")'))]"
    if ($ev -is [array]) { foreach ($e in $ev) { Write-Host "    $e" } } else { Write-Host "    $ev" }
    [void]$excel.Evaluate("JuliaHandleEventsClear()")
}
function ShowAll { foreach ($a in "A1","B1","A2","B2") { Retry { $c = $ws.Range($a); "{0} {1} -> {2}" -f $a, $c.Formula2, $c.Text | Write-Host } } }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    # Excel saves this for the user, so it is put back in finally.
    if ($null -eq $originalThrottle) { $originalThrottle = $excel.RTD.ThrottleInterval }
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    Retry { $ws.Range("Z1").Formula2 = $launch }
    Retry { $ws.Range("Z2").Formula2 = $helpers }
    Retry { $ws.Range("A1").Formula2 = '=JuliaCallHandle("mkrand",3)' }
    Retry { $ws.Range("B1").Formula2 = '=JuliaCall("sum",A1)' }
    Retry { $ws.Range("A2").Formula2 = '=JuliaCallHandle("reverse",A1)' }
    Retry { $ws.Range("B2").Formula2 = '=JuliaCall("sum",A2)' }
    ShowAll
    Ev "1 before restart"

    # 2. Close Julia and relaunch it: JuliaLaunch finds the new session and rebuilds the
    # handles itself, in dependency order, after mkrand has been redefined.
    Retry { $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")' }
    Start-Sleep -Seconds 8
    Retry { $ws.Range("Y1").ClearContents() }
    Retry { $ws.Range("Z1").Formula2 = $launch }
    "Z1 -> $(Retry { $ws.Range('Z1').Text }); Z2 -> $(Retry { $ws.Range('Z2').Text })" | Write-Host
    Settle
    ShowAll
    Ev "2 Julia relaunched: expect invalidate, then create of both in the new session"

    # 3. JuliaLaunch again while the same Julia is running: nothing to rebuild.
    Retry { $ws.Range("Z1").Formula2 = $launch }
    "Z1 -> $(Retry { $ws.Range('Z1').Text })" | Write-Host
    Settle
    ShowAll
    Ev "3 JuliaLaunch with the same Julia running: expect no events"

    Retry { $excel.CalculateFull() }
    Settle
    ShowAll
    Ev "4 after CalculateFull: expect no events"
    Retry { $ws.Range("A1:B2").ClearContents() }
    Ev "5 cleared"
    Retry { $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")' }
    $wb.Close($false)
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    try { if ($null -ne $originalThrottle) { $excel.RTD.ThrottleInterval = $originalThrottle } } catch {}
    try { $excel.Quit() } catch { Write-Host "Quit failed: $($_.Exception.Message)" }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($excelPid -ne 0) {
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(5000)) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
    }
}
