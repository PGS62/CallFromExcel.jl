# Named-functions spike driver: starts its own invisible Excel, registers the Debug XLL, exposes
# Julia functions, registers them with the JuliaRegisterFunctions command and checks them.
# Then saves a workbook using one, and reopens it in a new Excel before and after registering.
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
$launch = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
$saved = Join-Path $env:TEMP "NamedSpike.xlsx"
if (Test-Path $saved) { Remove-Item $saved }

Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
function Retry($sb) {
    for ($i = 0; $i -lt 100; $i++) {
        try { return & $sb } catch { if ($_.Exception.Message -notmatch "rejected|0x8001010A") { throw }; Start-Sleep -Milliseconds 200 }
    }
    throw "Excel stayed busy"
}
function New-Excel {
    $x = New-Object -ComObject Excel.Application
    $p = [uint32]0
    [void][Win32.User32]::GetWindowThreadProcessId([System.IntPtr]$x.Hwnd, [ref]$p)
    $x.Visible = $false
    $x.DisplayAlerts = $false
    if (-not $x.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    return @($x, $p)
}
function Close-Excel($x, $p) {
    try { $x.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
    $pr = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($pr -and -not $pr.WaitForExit(5000)) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue }
}
function F($ws, $addr, $formula) { Retry { $ws.Range($addr).Formula2 = $formula } }
function Show($ws, $addr) { $c = Retry { $ws.Range($addr) }; "  {0} {1} -> {2}" -f $addr, (Retry { $c.Formula2 }), (Retry { $c.Text }) | Write-Host }
function Log($excel, $title) {
    $l = Retry { $excel.Evaluate("JuliaRegisterLog()") }
    Write-Host "  [$title]"
    if ($l -is [array]) { foreach ($e in $l) { Write-Host "    $e" } } else { Write-Host "    $l" }
}
$define = '=JuliaEval({"price_swap(notional, fixed_rate, years) = notional * fixed_rate * years";' +
    '"SIGMA(x) = 2x"; "f(x) = 1"; "F(x) = 2"; "abc123(x) = x + 1";' +
    '"CallFromExcel.expose_to_excel(""price_swap""; description=""Toy swap price"", argdescriptions=[""Notional"", ""Fixed rate"", ""Years""])";' +
    '"CallFromExcel.expose_to_excel(""hypot"")"; "CallFromExcel.expose_to_excel(""SIGMA"")"; "CallFromExcel.expose_to_excel(""f"")";' +
    '"CallFromExcel.expose_to_excel(""price_swap""; name=""PRICE_SWAP"")";' +
    '"join(sort(collect(keys(CallFromExcel._exposed))), "", "")"})'
$define = $define.Replace("SIGMA", [string][char]0x3C3)

$ex = New-Excel; $excel = $ex[0]; $excelPid = $ex[1]
try {
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    Write-Host "=== 1 launch, define and expose_to_excel, register"
    F $ws "Z1" $launch
    F $ws "Z2" $define
    Show $ws "Z1"; Show $ws "Z2"
    # Names expose_to_excel refuses (expected: errors): F would replace jl.f, as Excel ignores case;
    # abc123 is a cell reference.
    F $ws "Z5" '=JuliaEval("CallFromExcel.expose_to_excel(""F"")")'
    F $ws "Z6" '=JuliaEval("CallFromExcel.expose_to_excel(""abc123""; name=""abc123"")")'
    Show $ws "Z5"; Show $ws "Z6"
    Retry { [void]$excel.Run("JuliaRegisterFunctions") }
    Log $excel "register log"

    Write-Host "=== 2 calling them from cells"
    F $ws "A1" '=jl.price_swap(1000000, 0.05, 10)'
    F $ws "A2" '=jl.hypot(3, 4)'
    F $ws "A3" ('=jl.' + [char]0x3C3 + '(21)')
    F $ws "A4" '=jl.f(0)'
    F $ws "A5" '=jl.F(0)'
    F $ws "A6" '=PRICE_SWAP(100, 0.1, 2)'
    try { F $ws "A7" '=abc123(1)' } catch { Write-Host "  A7 =abc123(1) could not be entered: $($_.Exception.Message)" }
    F $ws "A8" '=jl.price_swap(A10:A12, 0.05, 10)'
    foreach ($v in @(@("A10", 100), @("A11", 200), @("A12", 300))) { $c = $ws.Range($v[0]); $val = $v[1]; Retry { $c.Value2 = $val } }
    Retry { $excel.Calculate() }
    foreach ($a in "A1","A2","A3","A4","A5","A6","A7","A8") { Show $ws $a }

    Write-Host "=== 3 Application.Run"
    try { "  Run jl.price_swap(100, 0.1, 2) -> " + (Retry { $excel.Run("jl.price_swap", 100, 0.1, 2) }) | Write-Host } catch { Write-Host "  Run failed: $($_.Exception.Message)" }

    Write-Host "=== 4 redefine in Julia, no re-registration"
    F $ws "Z3" '=JuliaEval("price_swap(notional, fixed_rate, years) = -1.0")'
    Show $ws "Z3"
    Retry { $excel.CalculateFull() }
    Show $ws "A1"

    Write-Host "=== 5 new signature, re-registered"
    F $ws "Z4" '=JuliaEval({"price_swap(notional, fixed_rate, years, spread) = notional * (fixed_rate + spread) * years"; "CallFromExcel.expose_to_excel(""price_swap""; description=""Toy swap price with spread"")"})'
    Show $ws "Z4"
    Retry { [void]$excel.Run("JuliaRegisterFunctions") }
    Log $excel "register log after re-registering"
    F $ws "B1" '=jl.price_swap(100, 0.1, 2, 0.05)'
    Retry { $excel.CalculateFull() }
    Show $ws "B1"; Show $ws "A1"

    Write-Host "=== 6 save a workbook that uses one"
    $wb2 = $excel.Workbooks.Add()
    $ws2 = $wb2.Worksheets.Item(1)
    F $ws2 "A1" '=jl.hypot(5, 12)'
    Retry { $excel.Calculate() }
    Show $ws2 "A1"
    Retry { $wb2.SaveAs($saved, 51) }
    Retry { $wb2.Close($false) }
    F $ws "Y1" '=JuliaEval("exit()")'
    Retry { $wb.Close($false) }
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    Close-Excel $excel $excelPid
}

Write-Host "=== 7 reopen it in a new Excel, before Julia is launched"
$ex = New-Excel; $excel = $ex[0]; $excelPid = $ex[1]
try {
    $excel.RTD.ThrottleInterval = 0
    $wb2 = Retry { $excel.Workbooks.Open($saved) }
    $ws2 = $wb2.Worksheets.Item(1)
    Show $ws2 "A1"
    Retry { $excel.CalculateFull() }
    Write-Host "  after CalculateFull:"; Show $ws2 "A1"

    Write-Host "=== 8 launch, expose_to_excel, register, then recalculate"
    $wbz = $excel.Workbooks.Add(); $wsz = $wbz.Worksheets.Item(1)
    F $wsz "Z1" $launch
    F $wsz "Z2" '=JuliaEval("CallFromExcel.expose_to_excel(""hypot"")")'
    Show $wsz "Z2"
    Retry { [void]$excel.Run("JuliaRegisterFunctions") }
    Log $excel "register log"
    Retry { $excel.Calculate() }
    Write-Host "  after Calculate:"; Show $ws2 "A1"
    Retry { $excel.CalculateFull() }
    Write-Host "  after CalculateFull:"; Show $ws2 "A1"
    Retry { $excel.CalculateFullRebuild() }
    Write-Host "  after CalculateFullRebuild:"; Show $ws2 "A1"
    F $wsz "Y1" '=JuliaEval("exit()")'
    Retry { $wb2.Close($false) }; Retry { $wbz.Close($false) }
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    Close-Excel $excel $excelPid
}
