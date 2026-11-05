Attribute VB_Name = "SpeedTestSort"
Option Explicit

' Toolchain-validation prototype only: not wired to Julia for real use -- but this one
' actually talks to the live Julia HTTP server, unlike SpeedTestCSharp.bas's benchmarks.
'
' The "advertisement" scenario: with A1 containing =RANDARRAY(1000000), is
' =JuliaCall("sort",A1#) -- calling out to a separate process over HTTP -- faster than
' Excel's own native =SORT(A1#)? Measured natively: SORT() ~667ms, JuliaCall("sort",...)
' ~1273ms today (VBA-side V-encode/decode dominates that, see SpeedTestCSharp.bas's
' CSharpVFormatSpeedTest: ~580ms encode + ~550ms decode). ComServer.SortViaJulia
' (csharp-prototype\src\ComServer.cs) replaces just the VBA-side encode/decode/HTTP with
' C# (measured ~50-70ms encode, ~8-40ms decode), leaving the HTTP + Julia-side
' decode_from_xl/encode_for_xl cost as the floor (measured ~250-380ms for a 1,000,000-
' element array -- confirmed by timing "identity", which does no real work, at the same
' cost as "sort" -- so this floor isn't improvable by this client-side port; see
' ComServer.SortViaJuliaDiag for the encode/http/decode breakdown). Measured end to end,
' C#'s SortViaJulia (~330-450ms) beats both VBA's JuliaCall("sort",...) and Excel's own
' native SORT().
'
' REQUIRES: SpeedTestCSharp.bas imported into this SAME workbook -- it defines
' RunXlam/GetXlamWorkbook, which this file calls unqualified (VBA resolves an unqualified
' call to any public procedure in the same project, but not across separate workbooks).
' Importing only this file gives a compile error "Sub or Function not defined" on RunXlam.
' Also requires: Julia already launched in this Excel session (JuliaLaunch already called,
' same as for any other JuliaCall use) and JuliaExcelPrototype64.xll loaded.

Sub RunCSharpSortSpeedTest()
    MsgBox CSharpSortSpeedTest()
End Sub

Function CSharpSortSpeedTest() As String

    Const VectorLength As Long = 1000000
    Const NumCalls As Long = 5

    Dim InputData As Variant
    Dim comServer As Object
    Dim Port As Long
    Dim i As Long
    Dim t1 As Double, t2 As Double
    Dim TotalVBA As Double, TotalCS As Double
    Dim VBAResult As Variant, CSResult As Variant
    Dim Report As String

    On Error GoTo ErrHandler

    Set comServer = CreateObject("JuliaExcelPrototype.ComServer")
    Port = CLng(RunXlam("modHttp.GetJuliaPort"))
    InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

    ' --- Correctness: both paths must sort identically ---
    VBAResult = RunXlam("JuliaCall", "sort", InputData)
    CSResult = comServer.SortViaJulia(InputData, "127.0.0.1", Port)
    For i = 1 To VectorLength
        If VBAResult(i, 1) <> CSResult(i - 1, 0) Then Err.Raise vbObjectError + 1, , _
            "VBA JuliaCall(""sort"",...) and C# SortViaJulia results differ at element " & i
    Next i
    If VBAResult(1, 1) > VBAResult(2, 1) Then Err.Raise vbObjectError + 1, , "Result is not actually sorted!"

    ' --- Timing: interleaved, so warmup/drift (and Julia-side GC, etc.) affects both sides equally ---
    For i = 1 To NumCalls
        t1 = Timer
        VBAResult = RunXlam("JuliaCall", "sort", InputData)
        t2 = Timer
        TotalVBA = TotalVBA + (t2 - t1)

        t1 = Timer
        CSResult = comServer.SortViaJulia(InputData, "127.0.0.1", Port)
        t2 = Timer
        TotalCS = TotalCS + (t2 - t1)
    Next i

    Report = "Correctness: VBA JuliaCall(""sort"",...) and C# SortViaJulia give identical, correctly-sorted results." & vbCrLf & vbCrLf & _
        "Vector length = " & Format(VectorLength, "###,###") & ", averaged over " & NumCalls & " calls (interleaved)" & vbCrLf & _
        "JuliaCall(""sort"", RANDARRAY(" & Format(VectorLength, "###,###") & ")) via VBA = " & Format(TotalVBA / NumCalls * 1000, "0.0") & " ms" & vbCrLf & _
        "Same call, C# HTTP client + V-encode/decode (SortViaJulia) = " & Format(TotalCS / NumCalls * 1000, "0.0") & " ms" & vbCrLf & _
        "speedup = " & Format(TotalVBA / TotalCS, "0.0") & "x" & vbCrLf & vbCrLf & _
        "For reference, Excel's native SORT() on the same data was separately measured at ~667 ms."

    Debug.Print Report
    CSharpSortSpeedTest = Report

    Exit Function
ErrHandler:
    CSharpSortSpeedTest = "CSharpSortSpeedTest failed: " & Err.Description
End Function
