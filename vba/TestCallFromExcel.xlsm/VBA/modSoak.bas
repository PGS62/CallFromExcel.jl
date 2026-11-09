Attribute VB_Name = "modSoak"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/CallFromExcel.jl#readme
'
' Soak test for memory leaks. Run by hand (Alt+F8 -> RunSoakTest), not part of RunTests:
' it takes minutes and its result is a trend to look at, not a pass/fail.
'
' Each iteration makes a mix of realistic calls through both call surfaces:
'   - the COM server (VBA): JuliaCallVBA, JuliaEvalVBA, JuliaCall, including a Julia error;
'   - the XLL (worksheet): forced recalculation of a block of JuliaCall/JuliaEval formulas.
' After each batch of iterations it records, in a new workbook:
'   - Excel's private bytes (includes the add-in's .NET heap);
'   - Julia's live heap after a full GC (Base.gc_live_bytes);
'   - Julia's private bytes (catches leaks outside Julia's GC heap, e.g. sockets).
' Memory that rises then levels off is normal (heaps grow to a working size). Steady growth
' batch after batch is a leak: 100,000 calls leaking 100 bytes each is ~10 MB.
'
' Julia's private bytes are blank if Julia runs under WSL (its pid isn't a Windows pid).

Option Explicit

Public Sub RunSoakTest(Optional Batches As Long = 20, Optional IterationsPerBatch As Long = 1000)
    Dim CFE As Object
    Dim Arr() As Double
    Dim i As Long, b As Long, k As Long
    Dim wb As Workbook, wsOut As Worksheet, wsCalc As Worksheet
    Dim Formulas As Range
    Dim ExcelPID As Long, JuliaPID As Long
    Dim t As Double
    Dim Res As Variant
    Dim OldCalc As XlCalculation

    Set CFE = CreateObject("CallFromExcel.ComServer")
    Res = CFE.JuliaLaunch
    ExcelPID = CFE.JuliaExcelPID
    JuliaPID = CFE.JuliaEvalVBA("Int(getpid())")

    ReDim Arr(1 To 1000)
    For k = 1 To 1000
        Arr(k) = Rnd()
    Next k

    Set wb = Application.Workbooks.Add
    Set wsOut = wb.Worksheets(1)
    wsOut.Name = "Soak"
    Set wsCalc = wb.Worksheets.Add(After:=wsOut)
    wsCalc.Name = "Formulas"

    ' Inputs for the worksheet formulas: 1,000 random numbers in column A.
    wsCalc.Range("A1:A1000").Value = Application.WorksheetFunction.Transpose(Arr)
    ' One column each, so the spilling results (sort: C1:C1000, rand: H1:Q100) don't collide.
    wsCalc.Range("C1").Formula2 = "=JuliaCall(""sort"",A1:A1000)"
    wsCalc.Range("E1").Formula2 = "=JuliaCall(""+"",1,2)"
    wsCalc.Range("F1").Formula2 = "=JuliaEval(""repeat(""""abc"""",100)"")"
    wsCalc.Range("G1").Formula2 = "=JuliaEval(""error(""""soak"""")"")"
    wsCalc.Range("H1").Formula2 = "=JuliaEval(""rand(100,10)"")"
    Set Formulas = wsCalc.Range("C1,E1:H1")

    wsOut.Range("A1:G1").Value = Array("Batch", "Iterations", "Seconds", "Excel MB", _
        "Julia live MB", "Julia private MB", "Excel PID / Julia PID: " & ExcelPID & " / " & JuliaPID)

    OldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' Batch 0 is the warm-up: JIT, Julia method compilation and the connection pool settle.
    For b = 0 To Batches
        t = Timer
        For i = 1 To IterationsPerBatch
            Res = CFE.JuliaCallVBA("+", 1, 2)
            Res = CFE.JuliaCallVBA("identity", Arr)
            Res = CFE.JuliaEvalVBA("repeat(""abc"", 100)")
            Res = CFE.JuliaCall("sum", Arr)
            Res = CFE.JuliaEvalVBA("error(""soak"")")
            Formulas.Calculate
        Next i
        t = Timer - t

        wsOut.Cells(b + 2, 1).Value = b
        wsOut.Cells(b + 2, 2).Value = CDbl(b) * IterationsPerBatch
        wsOut.Cells(b + 2, 3).Value = t
        wsOut.Cells(b + 2, 4).Value = PrivateMB(ExcelPID)
        wsOut.Cells(b + 2, 5).Value = CFE.JuliaEvalVBA("GC.gc(); Base.gc_live_bytes() / 2^20")
        wsOut.Cells(b + 2, 6).Value = PrivateMB(JuliaPID)
        Application.StatusBar = "Soak test: batch " & b & " of " & Batches
        DoEvents
    Next b

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = OldCalc

    ' Chart the three memory series against the number of iterations, skipping the warm-up row.
    With wsOut.ChartObjects.Add(wsOut.Range("I2").Left, wsOut.Range("I2").Top, 600, 350).Chart
        .ChartType = xlXYScatterLines
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For k = 4 To 6
            With .SeriesCollection.NewSeries
                .Name = "='Soak'!" & wsOut.Cells(1, k).Address
                .XValues = wsOut.Range(wsOut.Cells(3, 2), wsOut.Cells(Batches + 2, 2))
                .Values = wsOut.Range(wsOut.Cells(3, k), wsOut.Cells(Batches + 2, k))
            End With
        Next k
        .HasTitle = True
        .ChartTitle.Text = "Memory (MB) vs iterations"
    End With
    wsOut.Activate
End Sub

' Private bytes of a process in MB, or Empty if there's no such Windows process.
Private Function PrivateMB(PID As Long) As Variant
    Dim Proc As Object
    On Error GoTo ErrHandler
    Set Proc = GetObject("winmgmts:").Get("Win32_Process.Handle=""" & PID & """")
    PrivateMB = CDbl(Proc.PrivatePageCount) / 2 ^ 20
    Exit Function
ErrHandler:
    PrivateMB = Empty
End Function
