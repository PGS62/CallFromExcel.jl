# Handle formulas with volatile arguments, e.g. =JuliaCallHandle("sum",RANDARRAY(1000)), in
# automatic calculation: each new handle's RTD update recalculates the cell, whose volatile
# arguments are then new, so another handle is created, and so on. The add-in should notice
# and return an error instead. Starts its own Excel, registers the Debug XLL, and counts the
# handles created over a few seconds.
param([int]$Seconds = 10, [int]$Throttle = -1, [switch]$Visible)
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"

$excel = New-Object -ComObject Excel.Application
Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
$excelPid = [uint32]0
[void][Win32.User32]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false

function Retry($sb) {
    for ($i = 0; $i -lt 200; $i++) {
        try { return & $sb } catch { if ($_.Exception.Message -notmatch "rejected|0x8001010A") { throw }; Start-Sleep -Milliseconds 100 }
    }
    throw "Excel stayed busy"
}
function Creates {
    $ev = Retry { $excel.Evaluate('JuliaHandleInfo("events")') }
    if ($ev -isnot [array]) { $ev = @($ev) }
    @($ev | Where-Object { "$_" -match " create " }).Count
}
function Show($ws, $addr) { Retry { $c = $ws.Range($addr); "{0} {1} -> {2}" -f $addr, $c.Formula2, $c.Text | Write-Host } }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    # Excel saves the throttle interval for the user (registry), so put it back afterwards.
    $originalThrottle = $excel.RTD.ThrottleInterval
    if ($Throttle -ge 0) { $excel.RTD.ThrottleInterval = $Throttle }
    "RTD throttle interval: $($excel.RTD.ThrottleInterval) ms" | Write-Host
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    Retry { $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")' }
    Show $ws "Z1"
    $excel.Calculation = -4105   # automatic

    foreach ($formula in @('=JuliaCallHandle("sum",RANDARRAY(1000))', '=JuliaEval("rand(" & RANDBETWEEN(1,9) & ")",TRUE)')) {
        [void](Retry { $excel.Evaluate("JuliaHandleEventsClear()") })
        Retry { $ws.Range("A1").Formula2 = $formula }
        $t0 = Get-Date
        $samples = @()
        while (((Get-Date) - $t0).TotalSeconds -lt $Seconds) { Start-Sleep -Seconds 2; $samples += Creates }
        "  creates so far, every 2s: $($samples -join ', ')" | Write-Host
        $n1 = Creates
        Start-Sleep -Seconds 2
        $n2 = Creates
        Write-Host ""
        "$formula" | Write-Host
        "  handles created in ${Seconds}s: $n1; in the next 2s: $($n2 - $n1) (0 = loop stopped)" | Write-Host
        Show $ws "A1"
        Retry { $ws.Range("A1").ClearContents() }
    }

    # Manual calculation: each recalculation makes one new handle, no loop.
    $excel.Calculation = -4135
    [void](Retry { $excel.Evaluate("JuliaHandleEventsClear()") })
    Retry { $ws.Range("A1").Formula2 = '=JuliaCallHandle("sum",RANDARRAY(1000))' }
    for ($i = 0; $i -lt 3; $i++) { Start-Sleep -Seconds 4; Retry { $excel.Calculate() } }
    Start-Sleep -Seconds 2
    Write-Host ""
    "Manual calculation, 3 recalculations 4s apart: handles created $(Creates) (expect 4)" | Write-Host
    Show $ws "A1"
}
finally {
    try { if ($null -ne $originalThrottle) { $excel.RTD.ThrottleInterval = $originalThrottle } } catch {}
    try { $wb.Close($false) } catch {}
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Start-Sleep -Seconds 2
    $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
    if ($p) { Stop-Process -Id $excelPid -Force; "ended Excel $excelPid" | Write-Host }
}
