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
          Dim Arr() As Double
          Dim B As Long
          Dim CFE As Object
          Dim ExcelPID As Long
          Dim Formulas As Range
          Dim i As Long
          Dim JuliaPID As Long
          Dim k As Long
          Dim OldCalc As XlCalculation
          Dim Res As Variant
          Dim t As Double
          Dim wb As Workbook
          Dim wsCalc As Worksheet
          Dim wsOut As Worksheet

1         Set CFE = CreateObject("CallFromExcel.ComServer")
2         Res = CFE.JuliaLaunch
3         ExcelPID = CFE.JuliaExcelPID
4         JuliaPID = CFE.JuliaEvalVBA("Int(getpid())")

5         ReDim Arr(1 To 1000)
6         For k = 1 To 1000
7             Arr(k) = Rnd()
8         Next k

9         Set wb = Application.Workbooks.Add
10        Set wsOut = wb.Worksheets(1)
11        wsOut.Name = "Soak"
12        Set wsCalc = wb.Worksheets.Add(After:=wsOut)
13        wsCalc.Name = "Formulas"

          ' Inputs for the worksheet formulas: 1,000 random numbers in column A.
14        wsCalc.Range("A1:A1000").Value = Application.WorksheetFunction.Transpose(Arr)
          ' One column each, so the spilling results (sort: C1:C1000, rand: H1:Q100) don't collide.
15        wsCalc.Range("C1").Formula2 = "=JuliaCall(""sort"",A1:A1000)"
16        wsCalc.Range("E1").Formula2 = "=JuliaCall(""+"",1,2)"
17        wsCalc.Range("F1").Formula2 = "=JuliaEval(""repeat(""""abc"""",100)"")"
18        wsCalc.Range("G1").Formula2 = "=JuliaEval(""error(""""soak"""")"")"
19        wsCalc.Range("H1").Formula2 = "=JuliaEval(""rand(100,10)"")"
20        Set Formulas = wsCalc.Range("C1,E1:H1")

21        wsOut.Range("A1:G1").Value = Array("Batch", "Iterations", "Seconds", "Excel MB", _
              "Julia live MB", "Julia private MB", "Excel PID / Julia PID: " & ExcelPID & " / " & JuliaPID)

22        OldCalc = Application.Calculation
23        Application.Calculation = xlCalculationManual
24        Application.ScreenUpdating = False

          ' Batch 0 is the warm-up: JIT, Julia method compilation and the connection pool settle.
25        For B = 0 To Batches
26            t = Timer
27            For i = 1 To IterationsPerBatch
28                Res = CFE.JuliaCallVBA("+", 1, 2)
29                Res = CFE.JuliaCallVBA("identity", Arr)
30                Res = CFE.JuliaEvalVBA("repeat(""abc"", 100)")
31                Res = CFE.JuliaCall("sum", Arr)
32                Res = CFE.JuliaEvalVBA("error(""soak"")")
33                Formulas.Calculate
34            Next i
35            t = Timer - t

36            wsOut.Cells(B + 2, 1).Value = B
37            wsOut.Cells(B + 2, 2).Value = CDbl(B) * IterationsPerBatch
38            wsOut.Cells(B + 2, 3).Value = t
39            wsOut.Cells(B + 2, 4).Value = PrivateMB(ExcelPID)
40            wsOut.Cells(B + 2, 5).Value = CFE.JuliaEvalVBA("GC.gc(); Base.gc_live_bytes() / 2^20")
41            wsOut.Cells(B + 2, 6).Value = PrivateMB(JuliaPID)
42            Application.StatusBar = "Soak test: batch " & B & " of " & Batches
43            DoEvents
44        Next B

45        Application.StatusBar = False
46        Application.ScreenUpdating = True
47        Application.Calculation = OldCalc

          ' Chart the three memory series against the number of iterations, skipping the warm-up row.
48        With wsOut.ChartObjects.Add(wsOut.Range("I2").Left, wsOut.Range("I2").Top, 600, 350).Chart
49            .ChartType = xlXYScatterLines
50            Do While .SeriesCollection.Count > 0
51                .SeriesCollection(1).Delete
52            Loop
53            For k = 4 To 6
54                With .SeriesCollection.NewSeries
55                    .Name = "='Soak'!" & wsOut.Cells(1, k).Address
56                    .XValues = wsOut.Range(wsOut.Cells(3, 2), wsOut.Cells(Batches + 2, 2))
57                    .Values = wsOut.Range(wsOut.Cells(3, k), wsOut.Cells(Batches + 2, k))
58                End With
59            Next k
60            .HasTitle = True
61            .ChartTitle.Text = "Memory (MB) vs iterations"
62        End With
63        wsOut.Activate
End Sub

' Private bytes of a process in MB, or Empty if there's no such Windows process.
Private Function PrivateMB(PID As Long) As Variant
          Dim Proc As Object
1         On Error GoTo ErrHandler
2         Set Proc = GetObject("winmgmts:").Get("Win32_Process.Handle=""" & PID & """")
3         PrivateMB = CDbl(Proc.PrivatePageCount) / 2 ^ 20
4         Exit Function
ErrHandler:
5         PrivateMB = Empty
End Function

