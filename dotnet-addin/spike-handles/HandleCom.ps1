# Object-handles spike: handles from VBA, through the COM server.
#  (1) VBA passing a worksheet-created handle (alone and inside a VBA array) to Julia;
#  (2) VBA creating handles (JuliaCallHandle, JuliaEvalHandle) and releasing them explicitly
#      (JuliaHandleRelease, JuliaHandleReleaseAll); a VBA-created handle used by a formula.
# Starts its own invisible Excel, registers the Debug XLL for the worksheet functions, and adds
# a VBA module to a new workbook (needs "Trust access to the VBA project object model").
# CreateObject("CallFromExcel.ComServer") uses the INSTALLED add-in's DLL, so install the
# build under test first.
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"

$vba = @'
Function Jl() As Object
    Static s As Object
    If s Is Nothing Then Set s = CreateObject("CallFromExcel.ComServer")
    Set Jl = s
End Function

Function Try_(ByVal what As String, Optional a1, Optional a2, Optional a3, Optional a4, Optional a5) As Variant
    On Error GoTo Failed
    Select Case what
        Case "CallHandle": Try_ = Jl().JuliaCallHandle(a1, a2)
        Case "CallHandleKw": Try_ = Jl().JuliaCallHandle(a1, a2, a3, a4, a5)
        Case "EvalHandle": Try_ = Jl().JuliaEvalHandle(a1)
        Case "Call": Try_ = Jl().JuliaCall(a1, a2)
        Case "CallVBA": Try_ = Jl().JuliaCallVBA(a1, a2)
        Case "Release": Try_ = Jl().JuliaHandleRelease(a1)
        Case "ReleaseAll": Try_ = Jl().JuliaHandleReleaseAll()
        Case "JuliaCount": Try_ = Jl().JuliaEval("CallFromExcel.handle_count()")
        Case "InArray"
            Dim arr(1 To 2) As Variant
            arr(1) = a1: arr(2) = a1
            Try_ = Jl().JuliaCallVBA("typeof", arr)
        Case "Range": Try_ = Jl().JuliaCall("length", Range(a1))
    End Select
    Exit Function
Failed:
    Try_ = "raised: " & Err.Description
End Function
'@

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
function T($label, $what, $a1, $a2, $a3, $a4, $a5) {
    $r = Retry { $excel.Run("Try_", $what, $a1, $a2, $a3, $a4, $a5) }
    "{0,-58} -> {1}" -f $label, $r | Write-Host
    $r
}
function Show($addr) { Retry { $c = $ws.Range($addr); "{0} {1} -> {2}" -f $addr, $c.Formula2, $c.Text | Write-Host } }
function JuliaCount { [void](T "  Julia stored" "JuliaCount") }

try {
    if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
    $excel.RTD.ThrottleInterval = 0
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    [void]$wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString($vba)

    Write-Host "=== 0 before JuliaLaunch"
    [void](T "JuliaCallHandle before launch (expect raised)" "CallHandle" "ones" 2)

    Retry { $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")' }
    Show "Z1"
    # JuliaLaunch stops waiting after 30 s; precompiling after a package change takes longer.
    for ($i = 0; $i -lt 120 -and -not (Retry { $excel.Evaluate("JuliaIsRunning()") }); $i++) { Start-Sleep -Seconds 2 }
    Write-Host "  JuliaIsRunning: $(Retry { $excel.Evaluate('JuliaIsRunning()') })"
    Retry { $ws.Range("Z2").Formula2 = '=JuliaEval("mkones(n)=ones(Int(n)); mkkw(n; v=0)=fill(v,Int(n)); ""helpers defined""")' }
    Settle; Show "Z2"

    Write-Host ""
    Write-Host "=== 1 VBA using a worksheet handle"
    Retry { $ws.Range("A1").Formula2 = '=JuliaCallHandle("mkones",5)' }
    Settle; Show "A1"
    $h = $ws.Range("A1").Text
    [void](T "JuliaCall length(worksheet handle)" "Call" "length" $h)
    [void](T "JuliaCallVBA typeof([h, h]) (1-based VBA array)" "InArray" $h)
    [void](T "JuliaCall length(Range A1) (a Range holding a handle)" "Range" "A1")
    [void](T "JuliaHandleRelease(worksheet handle) (expect False)" "Release" $h)
    [void](T "worksheet handle still alive: length" "Call" "length" $h)
    JuliaCount

    Write-Host ""
    Write-Host "=== 2 VBA creating and releasing handles"
    $v1 = T "JuliaCallHandle mkones(3)" "CallHandle" "mkones" 3
    $v2 = T "JuliaCallHandle mkkw(3; v=7) (named args)" "CallHandleKw" "mkkw" 3 ";" "v" 7
    $v3 = T "JuliaEvalHandle Dict(:a=>1, :b=>2)" "EvalHandle" "Dict(:a=>1, :b=>2)"
    [void](T "JuliaEvalHandle(handle text) is code (expect raised: UndefVarError jl)" "EvalHandle" $v3)
    [void](T "JuliaCall length(v1)" "Call" "length" $v1)
    [void](T "JuliaCall sum(v2)" "Call" "sum" $v2)
    [void](T "JuliaCall length(v3)" "Call" "length" $v3)
    [void](T "JuliaCallHandle with an error (expect raised)" "CallHandle" "nosuchfunction" 1)
    JuliaCount

    Retry { $ws.Range("B1").Value2 = $v1 }
    Retry { $ws.Range("B2").Formula2 = '=JuliaCall("length",B1)' }
    Settle; Show "B2"
    Write-Host "  (B2: a formula using a VBA-created handle)"

    [void](T "JuliaHandleRelease(v1) (expect True)" "Release" $v1)
    [void](T "JuliaHandleRelease(v1) again (expect False)" "Release" $v1)
    [void](T "JuliaCall length(v1) after release (expect not found)" "Call" "length" $v1)
    Retry { $excel.CalculateFull() }; Show "B2"
    [void](T "JuliaHandleRelease('hello') (expect False)" "Release" "hello")
    [void](T "JuliaHandleRelease('') (expect False)" "Release" "")
    JuliaCount
    [void](T "JuliaHandleReleaseAll (expect 2)" "ReleaseAll")
    [void](T "JuliaCall sum(v2) after ReleaseAll (expect not found)" "Call" "sum" $v2)
    JuliaCount
    Write-Host "  (Julia should still hold 1: the worksheet handle in A1)"

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
