# Object-handles spike driver: starts its own Excel, registers the Debug XLL, runs scenarios
# and prints the add-in's handle events after each step.
param([int]$PerfCount = 2000, [switch]$Visible)
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"

$excel = New-Object -ComObject Excel.Application
Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
$excelPid = [uint32]0
[void][Win32.User32]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false

function Settle { $excel.Calculate(); Start-Sleep -Milliseconds 300; $excel.Calculate() }
function Ev($title) {
    Settle
    $ev = $excel.Evaluate("JuliaHandleEvents()")
    $cnt = $excel.Evaluate("JuliaHandleCount()")
    Write-Host ""
    Write-Host "=== $title   [$cnt]"
    if ($ev -is [array]) { foreach ($e in $ev) { Write-Host "    $e" } } else { Write-Host "    $ev" }
    [void]$excel.Evaluate("JuliaHandleEventsClear()")
}
function Show($ws, $addr) { $c = $ws.Range($addr); "{0} {1} -> {2}" -f $addr, $c.Formula2, $c.Text | Write-Host }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    # 1. Before Julia is launched: error, then retry after launch.
    $ws.Range("A1").Formula2 = '=JuliaHandle("mkrand",3)'
    Show $ws "A1"
    Ev "1a before launch"
    $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
    Show $ws "Z1"
    $ws.Range("Z2").Formula2 = '=JuliaEval("mkrand(n)=rand(Int(n)); mkones(n)=ones(Int(n)); mkzeros(n)=zeros(Int(n)); mkfill(x,n)=fill(x,Int(n)); ""helpers defined""")'
    Show $ws "Z2"
    Settle; Show $ws "A1"
    $excel.CalculateFull(); Show $ws "A1"
    Ev "1b after launch + recalc"

    # 2. Use it; same args in another cell; recalc with same args.
    $ws.Range("B1").Formula2 = '=JuliaCall("length",A1)'
    $ws.Range("B2").Formula2 = '=JuliaCall("sum",A1)'
    $ws.Range("A2").Formula2 = '=JuliaHandle("mkrand",3)'
    Show $ws "B1"; Show $ws "B2"; Show $ws "A2"
    $before = $ws.Range("B2").Text
    $excel.CalculateFull()
    "sum before/after CalculateFull: $before / $($ws.Range('B2').Text)" | Write-Host
    Ev "2 use, same args, CalculateFull"

    # 3. New inputs, formula edit, delete.
    $ws.Range("C3").Value2 = 4
    $ws.Range("A3").Formula2 = '=JuliaHandle("mkrand",C3)'
    $ws.Range("B3").Formula2 = '=JuliaCall("length",A3)'
    Show $ws "A3"; Show $ws "B3"
    Ev "3a create from C3=4"
    $ws.Range("C3").Value2 = 5
    Show $ws "A3"; Show $ws "B3"
    Ev "3b C3 changed to 5"
    $ws.Range("A3").Formula2 = '=JuliaHandle("mkones",2)'
    Show $ws "A3"; Show $ws "B3"
    Ev "3c formula edited"
    $ws.Range("A3").ClearContents()
    Show $ws "B3"
    Ev "3d cell cleared"

    # 4. Nested, concatenated, range of handles.
    $ws.Range("D1").Formula2 = '=JuliaCall("length",JuliaHandle("mkones",7))'
    Show $ws "D1"
    Ev "4a nested"
    $excel.CalculateFull()
    Ev "4b nested after CalculateFull"
    $ws.Range("D2").Formula2 = '=JuliaHandle("mkones",8)&"x"'
    Show $ws "D2"
    Ev "4c concatenated"
    $ws.Range("D3").Formula2 = '=JuliaCall("length",A1:A2)'
    $ws.Range("D4").Formula2 = '=JuliaCall("typeof",A1:A2)'
    $ws.Range("D5").Formula2 = '=JuliaEvalHandle("Dict(:a=>1, :b=>2)")'
    $ws.Range("D6").Formula2 = '=JuliaCall("length",D5)'
    Show $ws "D3"; Show $ws "D4"; Show $ws "D5"; Show $ws "D6"
    Ev "4d range of handles, eval handle"

    # 5. Spilled handles: one handle per element of an array argument?
    $ws.Range("G1").Formula2 = '=JuliaHandle("mkrand",{2;3;4})'
    Show $ws "G1"; Show $ws "G2"
    Ev "5 array argument"

    # 6. Moves, inserted rows, copies.
    $ws.Range("A1").Cut($ws.Range("H10")) | Out-Null
    Show $ws "H10"; Show $ws "B1"
    Ev "6a cut A1 -> H10"
    $ws.Rows.Item(1).Insert() | Out-Null
    Show $ws "H11"; Show $ws "B2"
    Ev "6b row inserted"
    $ws.Range("A4").Formula2 = '=JuliaHandle("mkrand",9)'
    $ws.Range("A4").Copy($ws.Range("A5")) | Out-Null
    Show $ws "A4"; Show $ws "A5"
    Ev "6c copy of a handle cell"
    $ws.Range("A4").ClearContents()
    Ev "6d one of the two cleared"
    $ws.Range("A5").ClearContents()
    Ev "6e the other cleared"

    # 7. Close a workbook holding handles.
    $wb2 = $excel.Workbooks.Add()
    $wb2.Worksheets.Item(1).Range("A1").Formula2 = '=JuliaHandle("mkzeros",11)'
    $wb2.Worksheets.Item(1).Range("A2").Formula2 = '=JuliaHandle("mkzeros",12)'
    Ev "7a second workbook"
    $wb2.Close($false)
    Ev "7b second workbook closed"
    Start-Sleep -Seconds 3
    Ev "7c 3 s later, after Calculate"
    $ws.Range("C3").Value2 = 6
    Ev "7d after an edit in the remaining workbook"
    $excel.CalculateFull()
    Ev "7e after CalculateFull"

    # 8. Performance: many handle cells, then use them, then clear.
    $n = $PerfCount
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $ws.Range("K1:K$n").Formula2 = '=JuliaHandle("mkfill",ROW(),3)'
    $excel.Calculate()
    $t1 = $sw.ElapsedMilliseconds
    "{0}: K1 {1}, K{2} {3}; {4}" -f "8a", $ws.Range("K1").Text, $n, $ws.Range("K$n").Text, $excel.Evaluate("JuliaHandleCount()") | Write-Host
    "  #N/A cells in K: {0}" -f $excel.Evaluate("SUMPRODUCT(--ISNA(K1:K$n))") | Write-Host
    $sw.Restart()
    $ws.Range("L1:L$n").Formula2 = '=JuliaCall("sum",K1)'
    $excel.Calculate()
    $t2 = $sw.ElapsedMilliseconds
    "L$n = $($ws.Range("L$n").Text)" | Write-Host
    $sw.Restart()
    $excel.CalculateFull()
    $t3 = $sw.ElapsedMilliseconds
    $sw.Restart()
    $ws.Range("K1:L$n").ClearContents()
    $excel.Calculate()
    $t4 = $sw.ElapsedMilliseconds
    $sw.Restart()
    $ws.Range("M1:M$n").Formula2 = '=JuliaCall("mkfill",ROW(),3)'
    $excel.Calculate()
    $t5 = $sw.ElapsedMilliseconds
    $ws.Range("M1:M$n").ClearContents()
    "8 perf, n=$n : create {0} ms, use {1} ms, CalculateFull {2} ms, clear {3} ms; baseline JuliaCall fill {4} ms" -f $t1, $t2, $t3, $t4, $t5 | Write-Host
    [void]$excel.Evaluate("JuliaHandleEventsClear()")
    $cnt = $excel.Evaluate("JuliaHandleCount()")
    "after clearing: [$cnt]" | Write-Host

    # 9. Julia restarted while Excel stays open.
    Show $ws "H11"
    $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")'
    Start-Sleep -Seconds 8
    $ws.Range("Y1").ClearContents()
    $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")'
    $excel.CalculateFull()
    Show $ws "Z1"; Show $ws "H11"; Show $ws "B2"
    Ev "9a after Julia restart + CalculateFull"
    Show $ws "Z2"
    $ws.Range("H11").Formula2 = '=JuliaHandle("mkrand",3)'
    Show $ws "H11"; Show $ws "B2"
    Ev "9b creating formula re-entered unchanged"
    $ws.Range("H11").Formula2 = '=JuliaHandle("mkrand",3+0)'
    Show $ws "H11"; Show $ws "B2"
    Ev "9c creating formula edited (same value)"

    # 10. From VBA / Application.Run, and Evaluate.
    try { $r = $excel.Run("JuliaHandle", "mkones", 13); "Run -> $r" | Write-Host } catch { "Run threw: $($_.Exception.Message)" | Write-Host }
    try { $r = $excel.Evaluate('JuliaHandle("mkones",14)'); "Evaluate -> $r" | Write-Host } catch { "Evaluate threw: $($_.Exception.Message)" | Write-Host }
    Ev "10 Run and Evaluate"
    Start-Sleep -Seconds 2
    Ev "10b 2 s later"

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
