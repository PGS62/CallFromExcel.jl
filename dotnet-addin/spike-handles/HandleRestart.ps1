# Object-handles spike: Julia restarted while Excel stays open, with JuliaHandleInvalidateAll.
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
$launch = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
$helpers = '=JuliaEval("mkrand(n)=rand(Int(n)); ""helpers defined""")'

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
    $ev = Retry { $excel.Evaluate("JuliaHandleEvents()") }
    Write-Host ""
    Write-Host "=== $title   [$($excel.Evaluate("JuliaHandleCount()"))]"
    if ($ev -is [array]) { foreach ($e in $ev) { Write-Host "    $e" } } else { Write-Host "    $ev" }
    [void]$excel.Evaluate("JuliaHandleEventsClear()")
}
function ShowAll { foreach ($a in "A1","B1","A2","B2") { Retry { $c = $ws.Range($a); "{0} {1} -> {2}" -f $a, $c.Formula2, $c.Text | Write-Host } } }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Range("Z1").Formula2 = $launch
    $ws.Range("Z2").Formula2 = $helpers
    $ws.Range("A1").Formula2 = '=JuliaHandle("mkrand",3)'
    $ws.Range("B1").Formula2 = '=JuliaCall("sum",A1)'
    $ws.Range("A2").Formula2 = '=JuliaHandle("reverse",A1)'
    $ws.Range("B2").Formula2 = '=JuliaCall("sum",A2)'
    ShowAll
    Ev "1 before restart"

    $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")'
    Start-Sleep -Seconds 8
    $ws.Range("Y1").ClearContents()
    $ws.Range("Z1").Formula2 = $launch
    $ws.Range("Z2").Formula2 = $helpers
    "Z1 -> $($ws.Range('Z1').Text); Z2 -> $($ws.Range('Z2').Text)" | Write-Host
    Settle
    ShowAll
    Ev "2 Julia relaunched, no invalidation"

    [void]$excel.Run("JuliaHandleInvalidateAll")
    Settle
    ShowAll
    Ev "3 after JuliaHandleInvalidateAll + Calculate"

    Retry { $excel.CalculateFull() }
    Settle
    ShowAll
    Ev "4 after CalculateFull"
    Retry { $excel.CalculateFull() }
    Settle
    ShowAll
    Ev "4b after a second CalculateFull"

    $ws.Range("A1:B2").ClearContents()
    Ev "5 cleared"
    $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")'
    $wb.Close($false)
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    try { $excel.Quit() } catch { Write-Host "Quit failed: $($_.Exception.Message)" }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if ($excelPid -ne 0) {
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(5000)) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
    }
}
