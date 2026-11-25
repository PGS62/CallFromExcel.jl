# Automatic registration: after JuliaLaunch, JuliaInclude and JuliaEval, with no explicit
# JuliaRegisterFunctions. Starts its own invisible Excel and registers the Debug XLL.
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
$launch = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'

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
# Waits up to 10 s for a cell's text to stop being $not (queued registration runs when Excel is idle).
function WaitText($ws, $addr, $not) {
    for ($i = 0; $i -lt 50; $i++) { $t = Retry { $ws.Range($addr).Text }; if ($t -ne $not) { return $t }; Start-Sleep -Milliseconds 200 }
    return $t
}
$inc = Join-Path $env:TEMP "NamedAuto.jl"
Set-Content -Path $inc -Encoding utf8 -Value 'timesthree(x) = 3x; CallFromExcel.expose_to_excel("timesthree")'
$inc2 = Join-Path $env:TEMP "NamedAuto2.jl"
Set-Content -Path $inc2 -Encoding utf8 -Value @'
"""
    scale(x, factor)

Multiplies x by factor.
"""
@excel broadcast=true scale(x, factor) = x * factor

@excel broadcast=true wavg(x, w::AbstractVector) = x * sum(w) / length(w)
'@

$ex = New-Excel; $excel = $ex[0]; $excelPid = $ex[1]
try {
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    # Typed while unregistered, these stay #NAME? (expected, documented) until re-entered (step 6).
    Write-Host "=== 1 formulas entered before Julia runs"
    F $ws "A1" '=jl.addtwo(1, 2)'
    F $ws "A2" '=jl.timesthree(5)'
    Show $ws "A1"; Show $ws "A2"

    Write-Host "=== 2 launch, then JuliaEval exposes addtwo (no explicit registration)"
    F $ws "Z1" $launch
    F $ws "Z2" '=JuliaEval("addtwo(a, b) = a + b; CallFromExcel.expose_to_excel(""addtwo"")")'
    Show $ws "Z1"; Show $ws "Z2"
    # Registration is queued, running when Excel is next idle: a formula written before then
    # counts as typed while unregistered and stays #NAME?. So wait, as a script must.
    for ($i = 0; $i -lt 50 -and -not ((Retry { $excel.Evaluate("=jl.addtwo(1, 1)") }) -is [double]); $i++) { Start-Sleep -Milliseconds 200 }
    F $ws "B1" '=jl.addtwo(2, 3)'
    "  B1 (typed after) -> " + (WaitText $ws "B1" "#NAME?") + " (expect 5)" | Write-Host
    Show $ws "A1"
    Log $excel "register log"

    Write-Host "=== 3 JuliaInclude exposes timesthree"
    F $ws "Z3" ('=JuliaInclude("' + $inc + '")')
    Show $ws "Z3"
    F $ws "B2" '=jl.timesthree(4)'
    "  B2 (typed after) -> " + (WaitText $ws "B2" "#NAME?") + " (expect 12)" | Write-Host

    Write-Host "=== 4 JuliaEval that exposes nothing: no new registration"
    F $ws "Z4" '=JuliaEval("1 + 1")'
    Show $ws "Z4"
    Start-Sleep -Milliseconds 500
    Log $excel "register log (expect no new entries)"

    Write-Host "=== 5 Application.Run of the new function"
    "  Run jl.timesthree(7) -> " + (Retry { $excel.Run("jl.timesthree", 7) }) + " (expect 21)" | Write-Host

    Write-Host "=== 6 re-entering the formulas typed in step 1 (expect 3 and 15)"
    foreach ($a in "A1", "A2") { $c = $ws.Range($a); Retry { $c.Formula2 = $c.Formula2 }; Show $ws $a }

    Write-Host "=== 7 optional arguments (defaults years=10, spread=0)"
    F $ws "Z5" '=JuliaEval("swap(notional, rate, years=10, spread=0) = notional * (rate + spread) * years; CallFromExcel.expose_to_excel(""swap"")")'
    Show $ws "Z5"
    for ($i = 0; $i -lt 50 -and -not ((Retry { $excel.Evaluate("=jl.swap(1, 1)") }) -is [double]); $i++) { Start-Sleep -Milliseconds 200 }
    F $ws "C1" '=jl.swap(100, 0.1)'           # expect 100
    F $ws "C2" '=jl.swap(100, 0.1, 2)'        # expect 20
    F $ws "C3" '=jl.swap(100, 0.1, 2, 0.4)'   # expect 100
    F $ws "C4" '=jl.swap(100, 0.1, , 0.4)'    # expect "years can't be left out..."
    F $ws "C5" '=jl.swap(100)'                # expect "needs at least 2 arguments"
    F $ws "C6" '=jl.swap(100, )'              # expect "needs at least 2 arguments"
    foreach ($a in "C1", "C2", "C3", "C4", "C5", "C6") { Show $ws $a }

    Write-Host "=== 8 @excel in an included file, broadcasting over ranges"
    F $ws "Z6" ('=JuliaInclude("' + $inc2 + '")')
    Show $ws "Z6"
    for ($i = 0; $i -lt 50 -and -not ((Retry { $excel.Evaluate("=jl.scale(1, 1)") }) -is [double]); $i++) { Start-Sleep -Milliseconds 200 }
    foreach ($v in @(@("E1", 1), @("E2", 2), @("E3", 3), @("F1", 10), @("G1", 100), @("F3", 4), @("G3", 6))) { $c = $ws.Range($v[0]); $val = $v[1]; Retry { $c.Value2 = $val } }
    F $ws "H1" '=jl.scale(E1:E3, 10)'          # expect a column 10, 20, 30
    F $ws "I1" '=jl.scale(E1:E3, F1:G1)'       # expect a 3x2 table
    F $ws "K1" '=jl.scale(E1:E3, E1:E2)'       # expect "x is 3x1 and factor is 2x1"
    F $ws "L1" '=jl.wavg(E1:E3, F3:G3)'        # w whole: mean 5, expect 5, 10, 15
    F $ws "M1" '=jl.wavg(2, 7)'                # expect 14
    Retry { $excel.Calculate() }
    foreach ($a in "H1", "H2", "H3", "I1", "J1", "I3", "J3", "K1", "L1", "L3", "M1") { Show $ws $a }
    F $ws "Y1" '=JuliaEval("exit()")'
    Retry { $wb.Close($false) }
} catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    Close-Excel $excel $excelPid
}
