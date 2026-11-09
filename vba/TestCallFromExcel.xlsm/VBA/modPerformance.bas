Attribute VB_Name = "modPerformance"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

'--------------------------------------------------
'05-Nov-2021 16:18:37        DESKTOP-0VD2AF0
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.47189380999916
'--------------------------------------------------
'06-Nov-2021 12:28:58        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.9295860900078
'--------------------------------------------------
'30-Nov-2021 10:13:30        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    2.82354638000252  <--- Mmm, why the slowdown since 6-Nov version? Use of Assign?
'--------------------------------------------------
'01-Dec-2021 10:30:10       DESKTOP-0VD2AF0
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.25666286000051   <-- also seeing slowdown on PC in the office
'--------------------------------------------------
'20-Sep-2023 16:34:52       DESKTOP-HSGAM5S
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.42395350000006  <-- higher spec PC
'--------------------------------------------------
'29-Oct-2025 18:40:16       PHILIP-HPZ1
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.66512269999985           Averaged over 10 calls
'--------------------------------------------------
'22-Dec-2025 15:57:14       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.7418744300001            Averaged over 20 calls
'--------------------------------------------------
'--------------------------------------------------
'18-Aug-2026 20:06:31       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   0.920728114999656          Averaged over 20 calls
'--------------------------------------------------
'08-Nov-2026 15:16:16       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   0.224108579999302          Averaged over 20 calls '1st run using C# code of CallFromExcel.
'--------------------------------------------------
Private Sub SimpleSpeedTest()

          Const Expression As String = "fill(""xxx"",1000,1000)"
          Const NumCalls = 20
          ' Const Expression As String = "1+1"
          ' Const NumCalls = 1000
          Const UseLinux As Boolean = False
          Dim i As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double
          
1         LoadAddIn
2         CallFromExcel.JuliaEval ("exit()")
3         CallFromExcel.JuliaLaunch UseLinux, , gTestCommandOptions
4         Res = CallFromExcel.JuliaEval(Expression)
5         t1 = ElapsedTime
6         For i = 1 To NumCalls
7             Res = CallFromExcel.JuliaEval(Expression)
8         Next i
9         t2 = ElapsedTime

10        Debug.Print "'" & Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
11        Debug.Print "'Expression = " & Expression
12        Debug.Print "'Average time in JuliaEval", (t2 - t1) / NumCalls, "Averaged over " & CStr(NumCalls) & " calls"
13        Debug.Print "'--------------------------------------------------"
14    End Sub

      '--------------------------------------------------
      '29-Oct-2025 18:37:22       PHILIP-HPZ1
      'Expression = 1+1
      'Average time in JuliaEval   6.16188229999898E-03       Averaged over 1000 calls
      '--------------------------------------------------
      '22-Dec-2025 15:58:22       MSI
      'Expression = 1+1
      'Average time in JuliaEval   0.015039338300001          Averaged over 1000 calls
      '--------------------------------------------------
      '18-Aug-2026 20:09:02       MSI
      'Expression = 1+1
      'Average time in JuliaEval   1.32635509999818E-03       Averaged over 1000 calls
      '--------------------------------------------------
      '08-Nov-2026 11:03:46       MSI
      'Expression = 1+1
      'Average time in JuliaEval   1.60134200006723E-04       Averaged over 1000 calls' '1st run using C# code of CallFromExcel.
      '--------------------------------------------------

      '========================================================================================================================
      'Running method PerformanceTest
      'Time now = 2026-08-03 18:01:11
      'JuliaExcel Version = 107
      'Computer = MSI
      'Latency test
      'Average time for JuliaEval("1+1") = 7.96316399984062 miliseconds (averaged over 500 calls)
      'Two-way data transport test
      'Average time for JuliaCall("identity", vector of length 100,000) = 1.48447485999204 seconds (averaged over 10 calls)
      'One-way data transport test, Excel to Julia
      'Average time for JuliaCall("sum", vector of length 100,000) = 1.27661278001033 seconds (averaged over 10 calls)
      'One-way data transport test, Julia to Excel
      'Average time for JuliaEval("collect((1:100000).*pi)") = 0.294781249994412 seconds (averaged over 10 calls)

      '========================================================================================================================
      'Running method PerformanceTest
      'Time now = 2026-08-04 15:53:30
      'JuliaExcel Version = 109 <- Version 109 had experimental code to use UTF-8 for Excel-> Julia and Julia->Excel _
       communication.Julia To Excel gets faster Excel To Julia gets slower.
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 8.14005120001093 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 1.82576665000015 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 1.57798423999993 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.187884929999564 seconds (averaged over 10 calls)

'Running method PerformanceTest
'Time now = 2026-08-04 16:25:47
'JuliaExcel Version = 110 < using UTF-8 for Julia -> Excel, UTF-16 for Excel -> Julia
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 14.7025480000011 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 2.14857505000036 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 2.03623397000047 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.222088039999653 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-04 17:07:48
'JuliaExcel Version = 111
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 10.9454097999987 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 2.14352065000057 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 1.94212056999968 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.31337268999996 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-06 17:40:36
'JuliaExcel Version = 119
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 11.8458608000074 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.555459060001885 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.154651710001053 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.479368480000994 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-07 13:53:55
'JuliaExcel Version = 122
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 2.09547800000291 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.278066150000086 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.104752339998959 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.154756969999289 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-10 13:17:53
'JuliaExcel Version = 128
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.16016960004345 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.557932109991089 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.276260319992434 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.370191970001906 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-10 14:39:57
'JuliaExcel Version = 128
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.09734679991379 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.274450039991643 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.112631449999753 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.167135900002904 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-13 07:35:29
'JuliaExcel Version = 131
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.31439220000175 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.257647519999591 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.102828399999999 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.148632560000988 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-16 15:04:10
'JuliaExcel Version = 138
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.20061220001662 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.255838729999959 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.100987849998637 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.146759319998091 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-17 11:56:58
'JuliaExcel Version = 139
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.64528539997991 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.199695459997747 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.112783389998367 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 9.42291199986357E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - NOW WITH THE VBA -> JULIA "V" ENCODER ALSO WIRED IN
'(TrySerialiseArrayAsV in modSerialise.bas, decode_xl_array_v in src/decode.jl) - previously only
'Julia -> Excel had the fast "V" path. Compare "sum" (Excel -> Julia) and "identity" (two-way)
'against the runs above and against the pre-V baseline (v138, 2026-08-16 15:04:10: identity =
'0.2558s, sum = 0.1010s, collect = 0.1468s).
'Time now = 2026-08-17 14:25:31
'JuliaExcel Version = 139
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.81287340004928 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.150152799999341 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 6.95522399968468E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.78807900007814E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-18 11:31:05. Slowdown since 2026-08-17 14:25:31 - ROOT CAUSE FOUND AND FIXED (see
'run below): TrySerialiseArrayAsV's finiteness check (modSerialise.bas) called a separate Function
'(IsFiniteHex, ByVal String parameter) once per element - measured at ~80ms of pure VBA function-
'call/BSTR-copy overhead alone for a 100,000-element array, on top of the encoding itself. Fixed by
'inlining the check at each of the three call sites; IsFiniteHex deleted. Not thermal throttling -
'the PC had been idle overnight and was cool when this run was taken.
'JuliaExcel Version = 142
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.56602779999957 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.254055260000314 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.157926959999895 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 9.43088400003035E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.29171199994744E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER inlining TrySerialiseArrayAsV's finiteness check (removing
'the per-element IsFiniteHex Function call). Confirms recovery to, and slightly better than, the
'2026-08-17 14:25:31 numbers (identity = 0.1502s, sum = 0.0696s).
'Time now = 2026-08-18 11:52:28
'JuliaExcel Version = 142
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.20299299999897 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.158094169999822 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.071729609999602 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.29558800003724E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.24318100002711E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER removing TrySerialiseArrayAsV's NaN/Inf finiteness check
'entirely (not just inlining it): shown unnecessary for the Excel -> Julia direction, since the
'general per-scalar Double encode does no NaN/Inf translation either, and Julia's own decode
'reconstructs the identical Float64 bit pattern via either path - unlike the Julia -> Excel
'direction, where NaN -> #N/A / Inf -> #NUM! translation is genuine and stays in place (see
'TrySerialiseArrayAsV's docstring, modSerialise.bas). Best numbers yet.
'Time now = 2026-08-18 12:12:01
'JuliaExcel Version = 143
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.2095915999962 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.14551833000005 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 6.27044100001513E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.13102000000072E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.25317000005452E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER wiring the bulk-CopyMemory hex trick into production V-format
'encode (TrySerialiseArrayAsV, modSerialise.bas) and decode (Unserialise Case 86, modUnserialise.bas)
'- one bulk memory copy instead of N per-element LSet + function-call operations on each side. See
'BulkHexOfDoubleArray/BulkDoublesFromHex's own docstrings, and the now-deleted
'modHexBulkPrototype.bas (see git history) for the prototype that measured this first, in
'isolation. "sum" and "identity" (encode) and "collect" (decode) all improve substantially;
'"collect" roughly halves.
'Time now = 2026-08-19 09:02:52
'JuliaExcel Version = 145
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.24485660000937 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.10265148000035 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 5.39971099991817E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 4.13354000018444E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.41443200001959E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-23 18:51:59
'JuliaExcel Version = 148
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.10643340006936 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 9.43825400026981E-02 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 5.68173000006936E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 4.60276499972679E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.36371799977496E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Test below is against version that has the "Content-Length" change (commit 39445a284cd72b81c4b9bd76ff7f2a4eda2ce217)
'which may be why the latency has fallen from 1.11 to 0.38 milliseconds
'Running method PerformanceTest
'Time now = 2026-09-16 09:43:03
'JuliaExcel Version = 155
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 0.37751240003854 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 9.26379599957727E-02 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 5.37256399984471E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 3.93314900051337E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.14263000024948E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-11-08 11:10:17
'CallFromExcel Version = TODO - write code to get this
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 0.136910200002603 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.040583520004293 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 2.31290099967737E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 1.90326599986292E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.42919200006872E-02 seconds (averaged over 10 calls)

' Two report lines saying which CallFromExcel code is running, as reported by Julia: the
' package version, and the commit from "git describe --always --dirty", which appends
' "-dirty" if there are uncommitted changes. That is "not a git checkout" for a package
' installed with Pkg.add rather than run from the repo via --project=. Julia must be running.
Private Function VersionReport(Prefix As String) As String
          Dim Version As Variant, Commit As Variant
1         Version = CallFromExcel.JuliaEval("string(pkgversion(CallFromExcel))")
2         Commit = CallFromExcel.JuliaEval("try readchomp(`git -C $(pkgdir(CallFromExcel)) describe --always --dirty`) catch; ""not a git checkout"" end")
3         VersionReport = Prefix & "CallFromExcel Version = " & Version & vbLf & Prefix & "Commit = " & Commit
End Function

Function PerformanceTest() As String
          Const NumCallsOnePlusOne As Long = 500
          Const NumCallsVectors = 10
          Const VectorLength As Long = 100000
          Dim i As Long
          Dim InputData As Variant
          Dim j As Long
          Dim JuliaFunction As String
          Dim Report As String
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim WhatWasExecuted As String

1         On Error GoTo ErrHandler

2         LoadAddIn

3         Report = Report & String(120, "=") & vbLf
4         Report = Report & "Running method PerformanceTest" & vbLf
5         Report = Report & "Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss") & vbLf

          'Warm up
6         CallFromExcel.JuliaEval "exit()" 'shuts down Julia if it's running
7         PreciseSleep 1000
8         CallFromExcel.JuliaLaunch , , gTestCommandOptions
9         ThrowIfError CallFromExcel.JuliaEval("1+1")
10        InputData = Application.Evaluate("=RANDARRAY(100)")
11        ThrowIfError CallFromExcel.JuliaCall("identity", InputData)
12        ThrowIfError CallFromExcel.JuliaCall("sum", InputData)
13        ThrowIfError CallFromExcel.JuliaEval("collect((1:" & VectorLength & ").*pi)")

          'Latency test
14        t1 = ElapsedTime
15        For i = 1 To NumCallsOnePlusOne
16            Res = CallFromExcel.JuliaEval("1+1")
17        Next i
18        t2 = ElapsedTime

19        Report = Report & VersionReport("") & vbLf
20        Report = Report & "Computer = " & Environ$("ComputerName") & vbLf
21        Report = Report & "Latency test" & vbLf
22        Report = Report & "Average time for JuliaEval(""1+1"") = " & CStr(1000 * (t2 - t1) / NumCallsOnePlusOne) & _
              " miliseconds (averaged over " & CStr(NumCallsOnePlusOne) & " calls)" & vbLf

23        InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

          'Data transport tests. One Select Case per iteration - covering how to time the call, the
          'report text, and (for "identity" only) the correctness check - rather than testing
          'JuliaFunction against the same string literals in several separate If/ElseIf blocks.
24        For j = 1 To 4
25            JuliaFunction = Choose(j, "identity", "sum", "collect", "range")

26            Select Case JuliaFunction
                  Case "identity"
27                    t1 = ElapsedTime()
28                    For i = 1 To NumCallsVectors
29                        Res = CallFromExcel.JuliaCall(JuliaFunction, InputData)
30                    Next i
31                    t2 = ElapsedTime()
32                    If Not ArraysIdentical(Res, InputData) Then Throw "Ohoh, return from Julia function identity is not equal to its input"
33                    WhatWasExecuted = "JuliaCall(""identity"", vector of " & Format(VectorLength, "###,###") & " doubles)"
34                    Report = Report & "Two-way data transport test" & vbLf

35                Case "sum"
36                    t1 = ElapsedTime()
37                    For i = 1 To NumCallsVectors
38                        Res = CallFromExcel.JuliaCall(JuliaFunction, InputData)
39                    Next i
40                    t2 = ElapsedTime()
41                    WhatWasExecuted = "JuliaCall(""sum"", vector of " & Format(VectorLength, "###,###") & " doubles)"
42                    Report = Report & "One-way data transport test, Excel to Julia" & vbLf

43                Case "collect"
44                    t1 = ElapsedTime()
45                    For i = 1 To NumCallsVectors
46                        Res = CallFromExcel.JuliaEval("collect((1:" & VectorLength & ").*pi)")
47                    Next i
48                    t2 = ElapsedTime()
49                    WhatWasExecuted = "JuliaEval(""collect((1:" & CStr(VectorLength) & ").*pi)"")"
50                    Report = Report & "One-way data transport test, Julia to Excel" & vbLf

51                Case "range"
52                    t1 = ElapsedTime()
53                    For i = 1 To NumCallsVectors
54                        Res = CallFromExcel.JuliaEval("(1:" & VectorLength & ").*pi")
55                    Next i
56                    t2 = ElapsedTime()
57                    WhatWasExecuted = "JuliaEval(""(1:" & CStr(VectorLength) & ").*pi"")"
58                    Report = Report & "One-way data transport (AbstractRange), Julia to Excel" & vbLf
59            End Select

60            Report = Report & "Average time for " & WhatWasExecuted & " = " & _
                  CStr((t2 - t1) / NumCallsVectors) & " seconds (averaged over " & CStr(NumCallsVectors) & " calls)" & vbLf
61        Next j

62        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
63        AppActivate Application.Caption
64        PerformanceTest = Report
65        Exit Function
ErrHandler:
66        PerformanceTest = ReThrow("PerformanceTest", Err, True)
End Function

'--------------------------------------------------
'Running method SerialisationPerformanceTest
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:00:28
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 0.072152280001319 seconds (averaged over 10 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.102856119998614 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:00:55
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 7.19601280003553E-02 seconds (averaged over 50 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.222831700000097 seconds (averaged over 50 calls)
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:01:22
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 7.22701920004329E-02 seconds (averaged over 50 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.22391951800033 seconds (averaged over 50 calls)
'========================================================================================================================

'========================================================================================================================
'Running method DataframePerformanceTest
'Time now = 2026-09-16 11:26:39
'JuliaExcel Version = 155 WITH JULIA CODE EDITED TO OMIT "Content-Length" from HTTP header equivalent to JuliaExcel 2.1.1
'Computer = MSI
'Time to evaluate 'DataFrame(rand(Float64,(5000,1)),:auto)':  8.54120001895353E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,2)),:auto)':  0.017211300029885
'Time to evaluate 'DataFrame(rand(Float64,(5000,4)),:auto)':  4.00946000008844E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,8)),:auto)':  6.74322000122629E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,16)),:auto)':  0.160341700015124
'Time to evaluate 'DataFrame(rand(Float64,(5000,32)),:auto)':  0.287520400015637
'Time to evaluate 'DataFrame(rand(Float64,(5000,64)),:auto)':  0.591385699983221
'Time to evaluate 'DataFrame(rand(Float64,(5000,128)),:auto)':  2.99924229999306
'Time to evaluate 'DataFrame(rand(Float64,(5000,256)),:auto)':  13.1958732000203
'Time to evaluate 'DataFrame(rand(Float64,(5000,512)),:auto)':  50.8415707999957
'Time to evaluate 'DataFrame(rand(Float64,(5000,1024)),:auto)':  241.41591749998

'========================================================================================================================
'Running method DataframePerformanceTest
'Time now = 2026-09-16 11:16:56
'JuliaExcel Version = 155 - To be released as JuliaExcel 2.1.2
'Computer = MSI
'Time to evaluate 'DataFrame(rand(Float64,(5000,1)),:auto)':  1.11267000320368E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,2)),:auto)':  2.03008999815211E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,4)),:auto)':  3.92062999890186E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,8)),:auto)':  0.109932799998205
'Time to evaluate 'DataFrame(rand(Float64,(5000,16)),:auto)':  0.1594414000283
'Time to evaluate 'DataFrame(rand(Float64,(5000,32)),:auto)':  0.317960899963509
'Time to evaluate 'DataFrame(rand(Float64,(5000,64)),:auto)':  0.609124400012661
'Time to evaluate 'DataFrame(rand(Float64,(5000,128)),:auto)':  1.19918080000207
'Time to evaluate 'DataFrame(rand(Float64,(5000,256)),:auto)':  2.37117389997002
'Time to evaluate 'DataFrame(rand(Float64,(5000,512)),:auto)':  4.98593829997117
'Time to evaluate 'DataFrame(rand(Float64,(5000,1024)),:auto)':  10.8832294999738
'========================================================================================================================
'Running method DataframePerformanceTest
'Time now = 2026-11-08 11:12:31
'CallFromExcel Version = TODO write code to get this
'Computer = MSI
'Time to evaluate 'DataFrame(rand(Float64,(5000,1)),:auto)':  0.002693000016734
'Time to evaluate 'DataFrame(rand(Float64,(5000,2)),:auto)':  5.59209997300059E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,4)),:auto)':  0.020365700009279
'Time to evaluate 'DataFrame(rand(Float64,(5000,8)),:auto)':  2.68672000383958E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,16)),:auto)':  7.68823999678716E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,32)),:auto)':  0.12226529995678
'Time to evaluate 'DataFrame(rand(Float64,(5000,64)),:auto)':  0.208693299966399
'Time to evaluate 'DataFrame(rand(Float64,(5000,128)),:auto)':  0.40437090001069
'Time to evaluate 'DataFrame(rand(Float64,(5000,256)),:auto)':  1.0759058999829
'Time to evaluate 'DataFrame(rand(Float64,(5000,512)),:auto)':  2.15701540000737
'Time to evaluate 'DataFrame(rand(Float64,(5000,1024)),:auto)':  6.45206529996358
'========================================================================================================================
'========================================================================================================================
'Running method DataframePerformanceTest
'Time now = 2026-11-08 14:48:53
'CallFromExcel Version = TODO write code to get this
'Computer = MSI
'Time to evaluate 'DataFrame(rand(Float64,(5000,1)),:auto)':  1.13799999235198E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,2)),:auto)':  1.68889999622479E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,4)),:auto)':  2.87070003105327E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,8)),:auto)':  6.71430001966655E-03
'Time to evaluate 'DataFrame(rand(Float64,(5000,16)),:auto)':  1.32147999829613E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,32)),:auto)':  6.27897999947891E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,64)),:auto)':  5.76301999972202E-02
'Time to evaluate 'DataFrame(rand(Float64,(5000,128)),:auto)':  0.144619000027888
'Time to evaluate 'DataFrame(rand(Float64,(5000,256)),:auto)':  0.328970999980811
'Time to evaluate 'DataFrame(rand(Float64,(5000,512)),:auto)':  0.658971500000916
'Time to evaluate 'DataFrame(rand(Float64,(5000,1024)),:auto)':  1.33720409998205

Sub DataframePerformanceTest()
          Dim Report As String

          Const JuliaExpressionTemplate = "DataFrame(rand(Float64,(5000,Ncol)),:auto)"
          Dim i As Long
          Dim JuliaExpression As String
          Dim Ncol As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double

1         On Error GoTo ErrHandler

2         LoadAddIn

3         Report = "'" & String(120, "=") & vbLf
4         Report = Report & "'Running method DataframePerformanceTest" & vbLf
5         Report = Report & "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss") & vbLf
7         Report = Report & "'Computer = " & Environ$("ComputerName")

          'Warm up
8         CallFromExcel.JuliaEval "exit()" 'shuts down Julia if it's running
9         PreciseSleep 1000
10        CallFromExcel.JuliaLaunch , , gTestCommandOptions
11        ThrowIfError CallFromExcel.JuliaEval("1+1")

12        If VarType(CallFromExcel.JuliaEval("using DataFrames")) = vbString Then
13            ThrowIfError CallFromExcel.JuliaEval("using Pkg;Pkg.add(""DataFrames"")")
14        End If

15        ThrowIfError CallFromExcel.JuliaEval("using DataFrames")

16        JuliaExpression = Replace(JuliaExpressionTemplate, "Ncol", 1)

17        ThrowIfError (CallFromExcel.JuliaEval(JuliaExpression))

6         Report = Report & vbLf & VersionReport("'")
18        Debug.Print Report

19        For i = 0 To 10
20            Ncol = 2 ^ i
21            JuliaExpression = Replace(JuliaExpressionTemplate, "Ncol", CStr(Ncol))
22            t1 = ElapsedTime
23            Res = CallFromExcel.JuliaEval(JuliaExpression)
24            t2 = ElapsedTime
25            Debug.Print "'Time to evaluate '" & JuliaExpression & "':  " & CStr(t2 - t1)
26        Next i

27        Exit Sub
ErrHandler:
28        MsgBox CStr(Erl) & " " & Err.Description
End Sub

