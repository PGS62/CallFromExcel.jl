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

# Excel rejects COM calls while it is busy (e.g. RTD updates after creating 2000 handles).
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
    $cnt = Retry { $excel.Evaluate('JuliaHandleInfo("count")') }
    Write-Host ""
    Write-Host "=== $title   [$cnt]"
    if ($ev -is [array]) { foreach ($e in $ev) { Write-Host "    $e" } } else { Write-Host "    $ev" }
    [void](Retry { $excel.Evaluate("JuliaHandleEventsClear()") })
}
function Show($ws, $addr) { Retry { $c = $ws.Range($addr); "{0} {1} -> {2}" -f $addr, $c.Formula2, $c.Text | Write-Host } }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    # 1. Before Julia is launched: error, then retry after launch.
    Retry { $ws.Range("A1").Formula2 = '=JuliaCallHandle("mkrand",3)' }
    Show $ws "A1"
    Ev "1a before launch"
    Retry { $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")' }
    Show $ws "Z1"
    Retry { $ws.Range("Z2").Formula2 = '=JuliaEval("mkrand(n)=rand(Int(n)); mkones(n)=ones(Int(n)); mkzeros(n)=zeros(Int(n)); mkfill(x,n)=fill(x,Int(n)); mkkw(n; v=0)=fill(v,Int(n)); ""helpers defined""")' }
    Show $ws "Z2"
    Settle; Show $ws "A1"
    Retry { $excel.CalculateFull(); Show $ws "A1" }
    Ev "1b after launch + recalc"

    # 2. Use it; same args in another cell; recalc with same args.
    Retry { $ws.Range("B1").Formula2 = '=JuliaCall("length",A1)' }
    Retry { $ws.Range("B2").Formula2 = '=JuliaCall("sum",A1)' }
    Retry { $ws.Range("A2").Formula2 = '=JuliaCallHandle("mkrand",3)' }
    Show $ws "B1"; Show $ws "B2"; Show $ws "A2"
    $before = $ws.Range("B2").Text
    Retry { $excel.CalculateFull() }
    "sum before/after CalculateFull: $before / $($ws.Range('B2').Text)" | Write-Host
    Ev "2 use, same args, CalculateFull"

    # 3. New inputs, formula edit, delete.
    Retry { $ws.Range("C3").Value2 = 4 }
    Retry { $ws.Range("A3").Formula2 = '=JuliaCallHandle("mkrand",C3)' }
    Retry { $ws.Range("B3").Formula2 = '=JuliaCall("length",A3)' }
    Show $ws "A3"; Show $ws "B3"
    Ev "3a create from C3=4"
    Retry { $ws.Range("C3").Value2 = 5 }
    Show $ws "A3"; Show $ws "B3"
    Ev "3b C3 changed to 5"
    Retry { $ws.Range("A3").Formula2 = '=JuliaCallHandle("mkones",2)' }
    Show $ws "A3"; Show $ws "B3"
    Ev "3c formula edited"
    Retry { $ws.Range("A3").ClearContents() }
    Show $ws "B3"
    Ev "3d cell cleared"

    # 4. Nested, concatenated, range of handles.
    Retry { $ws.Range("D1").Formula2 = '=JuliaCall("length",JuliaCallHandle("mkones",7))' }
    Show $ws "D1"
    Ev "4a nested"
    Retry { $excel.CalculateFull() }
    Ev "4b nested after CalculateFull"
    Retry { $ws.Range("D2").Formula2 = '=JuliaCallHandle("mkones",8)&"x"' }
    Show $ws "D2"
    Ev "4c concatenated"
    Retry { $ws.Range("D3").Formula2 = '=JuliaCall("length",A1:A2)' }
    Retry { $ws.Range("D4").Formula2 = '=JuliaCall("typeof",A1:A2)' }
    Retry { $ws.Range("D5").Formula2 = '=JuliaEval("Dict(:a=>1, :b=>2)",TRUE)' }
    Retry { $ws.Range("D6").Formula2 = '=JuliaCall("length",D5)' }
    Show $ws "D3"; Show $ws "D4"; Show $ws "D5"; Show $ws "D6"
    Ev "4d range of handles, eval handle"

    # 4e. Named arguments in JuliaCallHandle; ";" as a positional argument (not supported);
    # JuliaEval with ReturnHandle omitted/FALSE; JuliaHandleInfo with an unknown type.
    Retry { $ws.Range("E1").Formula2 = '=JuliaCallHandle("mkkw",3,";","v",7)' }
    Retry { $ws.Range("E2").Formula2 = '=JuliaCall("sum",E1)' }
    Retry { $ws.Range("E3").Formula2 = '=JuliaCallHandle("split","a;b",";")' }
    Retry { $ws.Range("E4").Formula2 = '=JuliaEval("1+1")' }
    Retry { $ws.Range("E5").Formula2 = '=JuliaEval("1+2",FALSE)' }
    Retry { $ws.Range("E6").Formula2 = '=JuliaHandleInfo("foo")' }
    Retry { $ws.Range("E7").Formula2 = '=JuliaEval({"x=10";"x+1"},TRUE)' }
    Retry { $ws.Range("E8").Formula2 = '=JuliaCall("identity",E7)' }
    foreach ($a in "E1","E2","E3","E4","E5","E6","E7","E8") { Show $ws $a }
    Ev "4e named args, ReturnHandle, JuliaHandleInfo"

    # 4f. Review fixes 2026-11-25: broadcasting; ReturnHandle as an error, text or number; a
    # handle with a trailing line feed is plain text; handle text given to JuliaEval is code.
    Retry { $ws.Range("P1").Formula2 = '=JuliaCallHandle("sqrt.",{1;4;9})' }
    Retry { $ws.Range("P2").Formula2 = '=JuliaCall("sum",P1)' }
    Retry { $ws.Range("P3").Formula2 = '=JuliaEval("1+1",NA())' }
    Retry { $ws.Range("P4").Formula2 = '=JuliaEval("1+1","yes")' }
    Retry { $ws.Range("P5").Formula2 = '=JuliaEval("1+1",1)' }
    Retry { $ws.Range("P6").Formula2 = '=JuliaCall("typeof",D5&CHAR(10))' }
    Retry { $ws.Range("P7").Formula2 = '=JuliaEval(D5,TRUE)' }
    foreach ($a in "P1","P2","P3","P4","P5","P6","P7") { Show $ws $a }
    "expect: P1 handle, P2 6, P3 #N/A, P4 #ReturnHandle must be TRUE or FALSE!, P5 handle, P6 String, P7 UndefVarError jl" | Write-Host
    Ev "4f review fixes"
    Retry { $ws.Range("P1:P7").ClearContents() }

    # 5. Spilled handles: one handle per element of an array argument?
    Retry { $ws.Range("G1").Formula2 = '=JuliaCallHandle("mkrand",{2;3;4})' }
    Show $ws "G1"; Show $ws "G2"
    Ev "5 array argument"

    # 6. Moves, inserted rows, copies.
    Retry { $ws.Range("A1").Cut($ws.Range("H10")) | Out-Null }
    Show $ws "H10"; Show $ws "B1"
    Ev "6a cut A1 -> H10"
    Retry { $ws.Rows.Item(1).Insert() | Out-Null }
    Show $ws "H11"; Show $ws "B2"
    Ev "6b row inserted"
    Retry { $ws.Range("A4").Formula2 = '=JuliaCallHandle("mkrand",9)' }
    Retry { $ws.Range("A4").Copy($ws.Range("A5")) | Out-Null }
    Show $ws "A4"; Show $ws "A5"
    Ev "6c copy of a handle cell"
    Retry { $ws.Range("A4").ClearContents() }
    Ev "6d one of the two cleared"
    Retry { $ws.Range("A5").ClearContents() }
    Ev "6e the other cleared"

    # 7. Close a workbook holding handles.
    $wb2 = $excel.Workbooks.Add()
    Retry { $wb2.Worksheets.Item(1).Range("A1").Formula2 = '=JuliaCallHandle("mkzeros",11)' }
    Retry { $wb2.Worksheets.Item(1).Range("A2").Formula2 = '=JuliaCallHandle("mkzeros",12)' }
    Ev "7a second workbook"
    $wb2.Close($false)
    Ev "7b second workbook closed"
    Start-Sleep -Seconds 3
    Ev "7c 3 s later, after Calculate"
    Retry { $ws.Range("C3").Value2 = 6 }
    Ev "7d after an edit in the remaining workbook"
    Retry { $excel.CalculateFull() }
    Ev "7e after CalculateFull"

    # 8. Performance: many handle cells, then use them, then clear.
    $n = $PerfCount
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Retry { $ws.Range("K1:K$n").Formula2 = '=JuliaCallHandle("mkfill",ROW(),3)' }
    Retry { $excel.Calculate() }
    $t1 = $sw.ElapsedMilliseconds
    Retry { "{0}: K1 {1}, K{2} {3}; {4}" -f "8a", $ws.Range("K1").Text, $n, $ws.Range("K$n").Text, $excel.Evaluate('JuliaHandleInfo("count")') | Write-Host }
    Retry { "  #N/A cells in K: {0}" -f $excel.Evaluate("SUMPRODUCT(--ISNA(K1:K$n))") | Write-Host }
    $sw.Restart()
    Retry { $ws.Range("L1:L$n").Formula2 = '=JuliaCall("sum",K1)' }
    Retry { $excel.Calculate() }
    $t2 = $sw.ElapsedMilliseconds
    Retry { "L$n = $($ws.Range("L$n").Text)" | Write-Host }
    $sw.Restart()
    Retry { $excel.CalculateFull() }
    $t3 = $sw.ElapsedMilliseconds
    $sw.Restart()
    Retry { $ws.Range("K1:L$n").ClearContents() }
    Retry { $excel.Calculate() }
    $t4 = $sw.ElapsedMilliseconds
    $sw.Restart()
    Retry { $ws.Range("M1:M$n").Formula2 = '=JuliaCall("mkfill",ROW(),3)' }
    Retry { $excel.Calculate() }
    $t5 = $sw.ElapsedMilliseconds
    Retry { $ws.Range("M1:M$n").ClearContents() }
    "8 perf, n=$n : create {0} ms, use {1} ms, CalculateFull {2} ms, clear {3} ms; baseline JuliaCall fill {4} ms" -f $t1, $t2, $t3, $t4, $t5 | Write-Host
    Retry { [void]$excel.Evaluate("JuliaHandleEventsClear()") }
    $cnt = Retry { $excel.Evaluate('JuliaHandleInfo("count")') }
    "after clearing: [$cnt]" | Write-Host

    # 9. Julia restarted while Excel stays open: JuliaLaunch sees the new session and rebuilds
    # every live handle (more in HandleRestart.ps1). Y2 refers to the launch cell, so it is
    # recalculated - redefining the helpers - before the rebuild runs.
    Show $ws "H11"
    Retry { $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")' }
    Start-Sleep -Seconds 8
    Retry { $ws.Range("Y1").ClearContents() }
    Retry { $ws.Range("Y2").Formula2 = '=IF(LEN(Z1)>0,JuliaEval("mkrand(n)=rand(Int(n)); mkones(n)=ones(Int(n)); mkzeros(n)=zeros(Int(n)); mkfill(x,n)=fill(x,Int(n)); mkkw(n; v=0)=fill(v,Int(n)); ""helpers redefined"""))' }
    Retry { $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")' }
    Settle
    Show $ws "Z1"; Show $ws "Y2"; Show $ws "H11"; Show $ws "B2"
    Ev "9 after Julia restart: expect invalidate, then every live handle rebuilt in the new session"

    # 10. From VBA / Application.Run, and Evaluate.
    try { $r = $excel.Run("JuliaCallHandle", "mkones", 13); "Run -> $r" | Write-Host } catch { "Run threw: $($_.Exception.Message)" | Write-Host }
    try { $r = $excel.Evaluate('JuliaCallHandle("mkones",14)'); "Evaluate -> $r" | Write-Host } catch { "Evaluate threw: $($_.Exception.Message)" | Write-Host }
    Ev "10 Run and Evaluate"
    Start-Sleep -Seconds 2
    Ev "10b 2 s later"

    Retry { $ws.Range("Y1").Formula2 = '=JuliaEval("exit()")' }
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
