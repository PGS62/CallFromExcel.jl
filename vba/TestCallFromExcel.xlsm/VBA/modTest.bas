Attribute VB_Name = "modTest"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/CallFromExcel.jl#readme
'
' Port of vba\JuliaExcel.xlam\VBA\modTest.bas's RunTests for the C#/Excel-DNA add-in
' (dotnet-addin). There is no xlam in the new world, so every old call to a bare
' JuliaCall/JuliaEval/JuliaCallVBA/... Public Function becomes a call to the matching method
' on CallFromExcel, a CallFromExcel.ComServer COM object declared below - see
' dotnet-addin\README.md's migration notes. Two call surfaces now exist where the old VBA-
' only world had one:
'   - CallFromExcel.JuliaCall/JuliaEval/JuliaCallNamedArgs/JuliaLaunch/JuliaIsRunning/
'     JuliaExcelPID/JuliaSetVar/JuliaInclude: the migration surface for code that used to call
'     these as bare Public Functions. Keeps the original worksheet-style restrictions
'     (string-length limit, no nested arrays/Dictionaries, 1-D arrays become a column).
'   - CallFromExcel.JuliaCallVBA/JuliaEvalVBA/JuliaCallNamedArgsVBA: unchanged from before,
'     lifting those restrictions.
' Tests below that used the first group port with a "CallFromExcel." prefix and nothing else
' changed. The WsTest_* functions at the end are new: they exercise the real worksheet-
' formula/XLL surface directly (writing real formulas into real cells), which the old suite
' never needed as a separate thing, since a bare JuliaCall behaved identically whether called
' from a cell or a macro - now the XLL and the COM migration methods are separate
' implementations, so the XLL path needs its own coverage.
'
' Dictionaries, both passed in and returned, are Scripting.Dictionary objects, as in the old
' VBA world - this project requires a reference to "Microsoft Scripting Runtime".
'
' Not ported (no VBA-level equivalent exists any more - the functions they tested are now
' C#-only and already covered by dotnet-addin\CallFromExcel.AddIn.TestHarness\WireFormatTests.cs):
' TestSerialiseArrayAsV, TestMakeJuliaLiteral, TestUnserialiseStringTooLong,
' TestUnserialiseMalformedRank, TestConcatenateExpressions, TestTrim255.
' TestVFormatBoundarySizes is kept but adapted to round-trip through JuliaCallVBA rather than
' calling the (no longer existing) VBA-internal TrySerialiseArrayAsV/UnserialiseFromString
' directly.

Option Explicit

Public CallFromExcel As Object
Private TestLog As String
'The CommandLineOptions RunTests launched Julia with, so tests that relaunch Julia use the same
'environment rather than the default one.
Private TestCommandLineOptions As String
Public Const gPackageName As String = "CallFromExcel"
Public Const gTestCommandOptions = "--project=C:\Projects\CallFromExcel" 'Forces the Julia process
'launched for RunTests/PerformanceTest/SimpleSpeedTest to activate this repo's own environment,
'so "using CallFromExcel" always loads the local working copy under test, regardless of what the
'default (non-project-specific) Julia environment happens to have installed/dev'd.

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Lets a non-interactive caller (Run-VbaTests.ps1) retrieve what PrintTwice wrote, since
' Debug.Print's output (the Immediate window) isn't visible over COM automation.
Function GetTestLog() As String
1         GetTestLog = TestLog
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTests
' Purpose    : Test CallFromExcel's C#/Excel-DNA add-in for a variety of data types and both call
'              surfaces (worksheet/XLL and VBA/COM). Prints results to the Immediate window and
'              to a MsgBox.
' Parameters :
'  CommandLineOptions: Passed straight through to JuliaLaunch. Defaults to "", i.e. JuliaLaunch's
'                       own default environment.
' -----------------------------------------------------------------------------------------------------------------------
Function RunTests(Optional SilentMode As Boolean = False, Optional CommandLineOptions As String = gTestCommandOptions) As Boolean

          Const Title = "CallFromExcel RunTests"
          Dim NumFailed As Long
          Dim NumPassed As Long
          Dim Prompt As String

1         On Error GoTo ErrHandler

2         LoadAddIn

3         CallFromExcel.JuliaEval "exit()"
4         Sleep 1000 'Give time to shut down properly, otherwise launch can fail thinking Julia still running but unresponsive
5         ThrowIfError CallFromExcel.JuliaLaunch(, , CommandLineOptions)
          TestCommandLineOptions = CommandLineOptions

6         PrintTwice vbLf & String(80, "=")
7         PrintTwice "CallFromExcel RunTests (C#/Excel-DNA add-in)"
8         PrintTwice Format(Now, "yyyy-mm-dd hh:mm:ss")
9         PrintTwice "ComputerName = " & Environ("ComputerName")

          'Tests ported from vba\JuliaExcel.xlam\VBA\modTest.bas - the CallFromExcel.<Name>/COM surface.
10        AccResult "TestEmpty", TestEmpty, NumPassed, NumFailed
11        AccResult "TestBoolean", TestBoolean, NumPassed, NumFailed
12        AccResult "TestDouble", TestDouble, NumPassed, NumFailed
13        AccResult "TestString", TestString, NumPassed, NumFailed
14        AccResult "TestWideString", TestWideString, NumPassed, NumFailed
15        AccResult "TestLong", TestLong, NumPassed, NumFailed
16        AccResult "TestLongLong", TestLongLong, NumPassed, NumFailed
17        AccResult "TestSingle", TestSingle, NumPassed, NumFailed
18        AccResult "TestDate", TestDate, NumPassed, NumFailed
19        AccResult "TestDateTime", TestDateTime, NumPassed, NumFailed
20        AccResult "TestByte", TestByte, NumPassed, NumFailed
21        AccResult "Test1DArrayOfDoubles", Test1DArrayOfDoubles, NumPassed, NumFailed
22        AccResult "Test2DArrayOfMixedType", Test2DArrayOfMixedType, NumPassed, NumFailed
23        AccResult "Test3DArray", Test3DArray, NumPassed, NumFailed
24        AccResult "Test4DArray", Test4DArray, NumPassed, NumFailed
25        AccResult "TestDictionary", TestDictionary, NumPassed, NumFailed
26        AccResult "TestExactRoundTripping", TestExactRoundTripping, NumPassed, NumFailed
27        AccResult "TestArrayOfDictionaries", TestArrayOfDictionaries, NumPassed, NumFailed
28        AccResult "TestDictionaryOfArrays", TestDictionaryOfArrays, NumPassed, NumFailed
29        AccResult "TestDictionaryOfTypes", TestDictionaryOfTypes, NumPassed, NumFailed
30        AccResult "TestDictionaryIsScriptingDictionary", TestDictionaryIsScriptingDictionary, NumPassed, NumFailed
31        AccResult "TestIntegerArrayIsLongLong", TestIntegerArrayIsLongLong, NumPassed, NumFailed
32        AccResult "TestOneDArraysDisplayAsOneColumnOnSheet", TestOneDArraysDisplayAsOneColumnOnSheet, NumPassed, NumFailed
33        AccResult "TestElType", TestElType, NumPassed, NumFailed
34        AccResult "TestBroadcasting", TestBroadcasting, NumPassed, NumFailed
35        AccResult "TestNaNFallback", TestNaNFallback, NumPassed, NumFailed
36        AccResult "TestInfFallback", TestInfFallback, NumPassed, NumFailed
37        AccResult "TestVFormatVectorColumn", TestVFormatVectorColumn, NumPassed, NumFailed
38        AccResult "TestVFormatMatrixNonSquare", TestVFormatMatrixNonSquare, NumPassed, NumFailed
39        AccResult "TestMatrixNaNFallback", TestMatrixNaNFallback, NumPassed, NumFailed
40        AccResult "TestVFormatEmptyArrayFallback", TestVFormatEmptyArrayFallback, NumPassed, NumFailed
41        AccResult "TestVEncodeFromVariantArray", TestVEncodeFromVariantArray, NumPassed, NumFailed
42        AccResult "TestVFormat3DArray", TestVFormat3DArray, NumPassed, NumFailed
43        AccResult "TestRangeIntegerVsCollect", TestRangeIntegerVsCollect, NumPassed, NumFailed
44        AccResult "TestRangeStepVsCollect", TestRangeStepVsCollect, NumPassed, NumFailed
45        AccResult "TestRangeFloatVsCollect", TestRangeFloatVsCollect, NumPassed, NumFailed
46        AccResult "TestRangeJuliaEvalVBAIsOneD", TestRangeJuliaEvalVBAIsOneD, NumPassed, NumFailed
47        AccResult "TestRangeJuliaEvalIsColumn", TestRangeJuliaEvalIsColumn, NumPassed, NumFailed
48        AccResult "TestRangeWireFormat", TestRangeWireFormat, NumPassed, NumFailed
49        AccResult "TestNaNInfRoundTripViaV", TestNaNInfRoundTripViaV, NumPassed, NumFailed
50        AccResult "TestExcelErrorRoundTrip", TestExcelErrorRoundTrip, NumPassed, NumFailed
51        AccResult "TestArrayOfSingleCharacters", TestArrayOfSingleCharacters, NumPassed, NumFailed
52        AccResult "TestNamedArgs", TestNamedArgs, NumPassed, NumFailed
53        AccResult "TestNamedArgsVBA", TestNamedArgsVBA, NumPassed, NumFailed
54        AccResult "TestVFormatBoundarySizes", TestVFormatBoundarySizes, NumPassed, NumFailed
55        AccResult "TestJuliaEvalMultiLineWithComment", TestJuliaEvalMultiLineWithComment, NumPassed, NumFailed

          'New: tests of the real worksheet-formula / XLL surface. No equivalent existed in the
          'old suite - a bare JuliaCall behaved identically whether called from a cell or a
          'macro, but the XLL and the COM migration methods above are now separate
          'implementations that can diverge, so this surface needs its own coverage.
56        AccResult "WsTest_Scalar", WsTest_Scalar, NumPassed, NumFailed
57        AccResult "WsTest_String", WsTest_String, NumPassed, NumFailed
58        AccResult "WsTest_OneDArrayBecomesColumn", WsTest_OneDArrayBecomesColumn, NumPassed, NumFailed
59        AccResult "WsTest_TwoDArray", WsTest_TwoDArray, NumPassed, NumFailed
60        AccResult "WsTest_IntegerArrays", WsTest_IntegerArrays, NumPassed, NumFailed
61        AccResult "WsTest_JuliaEvalScalar", WsTest_JuliaEvalScalar, NumPassed, NumFailed
62        AccResult "WsTest_EmptyResult", WsTest_EmptyResult, NumPassed, NumFailed
63        AccResult "WsTest_JuliaLaunchArgOrder", WsTest_JuliaLaunchArgOrder, NumPassed, NumFailed
64        AccResult "WsTest_RelaunchAfterExit", WsTest_RelaunchAfterExit, NumPassed, NumFailed
65        AccResult "WsTest_JuliaIsRunningAndPID", WsTest_JuliaIsRunningAndPID, NumPassed, NumFailed
66        AccResult "WsTest_ErrorString", WsTest_ErrorString, NumPassed, NumFailed
67        AccResult "WsTest_SortRoundTrip", WsTest_SortRoundTrip, NumPassed, NumFailed
          AccResult "WsTest_ErrorArguments", WsTest_ErrorArguments, NumPassed, NumFailed
          AccResult "WsTest_RangeWithBlanksAndErrors", WsTest_RangeWithBlanksAndErrors, NumPassed, NumFailed
          AccResult "WsTest_JuliaSetVar", WsTest_JuliaSetVar, NumPassed, NumFailed
          AccResult "WsTest_JuliaEvalRangeWithBlankLine", WsTest_JuliaEvalRangeWithBlankLine, NumPassed, NumFailed
          AccResult "TestRangeArgumentsVBA", TestRangeArgumentsVBA, NumPassed, NumFailed
          AccResult "WsTest_JuliaLaunchBlankArgs", WsTest_JuliaLaunchBlankArgs, NumPassed, NumFailed
          AccResult "WsTest_JuliaEvalBlankOrErrorCell", WsTest_JuliaEvalBlankOrErrorCell, NumPassed, NumFailed

68        Prompt = NumPassed & " test(s) passed" & vbLf & _
              NumFailed & " test(s) failed"

69        If NumFailed > 0 Then
70            Prompt = Prompt & vbLf & vbLf & "See VBA Immediate window for details"
71        End If

72        PrintTwice NumPassed & " test(s) passed"
73        PrintTwice NumFailed & " test(s) failed"
74        PrintTwice String(80, "=")

75        If Not SilentMode Then
76            AppActivate Application.Caption
77            MsgBox Prompt, IIf(NumFailed = 0, vbInformation, vbCritical), Title
78        End If

79        RunTests = NumFailed = 0

80        Exit Function
ErrHandler:
          Dim ErrText As String
81        ErrText = ReThrow("RunTests", Err, True)
82        TestLog = TestLog & ErrText & vbCrLf
83        If Not SilentMode Then
84            MsgBox ErrText, vbCritical, Title
85        End If
86        RunTests = False
End Function

Sub LoadAddIn()
          Dim XllPath As String
1         Set CallFromExcel = CreateObject("CallFromExcel.ComServer")
2         XllPath = RepoRoot() & "\dotnet-addin\dist\CallFromExcel.AddIn64-packed.xll"
3         If Dir(XllPath) = vbNullString Then
4             Throw "Cannot find the add-in at " & XllPath & _
                  " - build it first (dotnet build -c Release in dotnet-addin\CallFromExcel.AddIn)."
5         End If
6         Application.RegisterXLL XllPath
End Sub

' Echoing into the Julia REPL is best-effort: a logging helper must never abort the test run.
Sub PrintTwice(Text As String)
1         Debug.Print Text
2         TestLog = TestLog & Text & vbCrLf
3         On Error Resume Next
4         CallFromExcel.JuliaEval "println(" & SimpleJuliaStringLiteral(Text) & ")"
5         On Error GoTo 0
End Sub

' Minimal ASCII-safe Julia string literal, enough for the plain diagnostic text PrintTwice is
' ever called with (test names, dates, computer names) - not a port of the old MakeJuliaLiteral
' (which also escaped Unicode bidi-control characters), since that is a C#-only implementation
' detail now (CallFromExcel.AddIn\src\JuliaLiteral.cs), not a public VBA-callable function.
Private Function SimpleJuliaStringLiteral(Text As String) As String
          Dim Escaped As String
1         Escaped = Replace(Text, "\", "\\")
2         Escaped = Replace(Escaped, Chr(34), "\" & Chr(34))
3         Escaped = Replace(Escaped, vbCr, "\r")
4         Escaped = Replace(Escaped, vbLf, "\n")
5         Escaped = Replace(Escaped, "$", "\$")
6         SimpleJuliaStringLiteral = Chr(34) & Escaped & Chr(34)
End Function

Sub AccResult(TestName As String, Result As Boolean, ByRef NumPassed, ByRef NumFailed)
1         On Error GoTo ErrHandler
2         PrintTwice "Test " & TestName & " completed"
3         If Result Then
4             NumPassed = NumPassed + 1
5         Else
6             PrintTwice "Test " & TestName & " Failed!"
7             NumFailed = NumFailed + 1
8         End If
9         Exit Sub
ErrHandler:
10        ReThrow "AccResult", Err
End Sub

Function TestEmpty()
1         On Error GoTo ErrHandler
2         TestEmpty = IsEmpty(CallFromExcel.JuliaCall("identity", Empty))
3         Exit Function
ErrHandler:
4         PrintTwice ReThrow("TestEmpty", Err, True)
5         TestEmpty = False
End Function

Function TestBoolean()
1         On Error GoTo ErrHandler
2         TestBoolean = (CallFromExcel.JuliaCall("identity", True) = True) And (CallFromExcel.JuliaCall("identity", False) = False)
3         Exit Function
ErrHandler:
4         PrintTwice ReThrow("TestBoolean", Err, True)
5         TestBoolean = False
End Function

Function TestDouble()
          Dim x As Double
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = Application.WorksheetFunction.Pi()
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestDouble = (x = y) And (VarType(y) = vbDouble)
5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDouble", Err, True)
7         TestDouble = False
End Function

Function TestString()
          Dim x As String
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = "FooBar"
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestString = x = y
5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestString", Err, True)
7         TestString = False
End Function

Function TestWideString()
          Dim i As Long
          Dim x As String
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = String(10000, " ")
3         For i = 1 To 10000
4             Mid$(x, i, 1) = ChrW(i)
5         Next i
6         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
7         TestWideString = (x = y) And VarType(y) = vbString

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestWideString", Err, True)
10        TestWideString = False
End Function

Function TestLong()
          Dim x As Long
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 123456789
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestLong = (x = y) And VarType(y) = vbLong

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestLong", Err, True)
7         TestLong = False
End Function

Function TestLongLong()
          Dim x As LongLong
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 123456789^
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestLongLong = (x = y) And VarType(y) = vbLongLong

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestLongLong", Err, True)
7         TestLongLong = False
End Function

Function TestSingle()
          Dim x As Single
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = CSng(1 / 3)
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestSingle = (x = y) And (VarType(y) = vbSingle)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestSingle", Err, True)
7         TestSingle = False
End Function

Function TestDate()
          Dim x As Date
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = DateSerial(2025, 12, 22)
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestDate = (x = y) And VarType(y) = vbDate

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDate", Err, True)
7         TestDate = False
End Function

Function TestDateTime()
          Dim x As Date
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = DateSerial(2025, 12, 22) + TimeValue("03:40:33")
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestDateTime = x = y

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDateTime", Err, True)
7         TestDateTime = False
End Function

Function TestByte()
          Dim x As Byte
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 200
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestByte = (x = y) And VarType(y) = vbByte

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestByte", Err, True)
7         TestByte = False
End Function

Function Test1DArrayOfDoubles()
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 3)
3         x(1) = 1 / 3
4         x(2) = 1E+100
5         x(3) = 0
6         y = CallFromExcel.JuliaCallVBA("identity", x)
7         Test1DArrayOfDoubles = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test1DArrayOfDoubles", Err, True)
10        Test1DArrayOfDoubles = False
End Function

Function Test2DArrayOfMixedType()
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 3, 1 To 3)
3         x(1, 1) = 1 / 3:  x(1, 2) = 1E+100:    x(1, 3) = CLng(100)
4         x(2, 1) = "Foo":  x(2, 2) = CSng(3):   x(2, 3) = CLngLng(100)
5         x(3, 1) = "Foo":  x(3, 2) = CSng(3):   x(3, 3) = CInt(100)

6         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
7         Test2DArrayOfMixedType = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test2DArrayOfMixedType", Err, True)
10        Test2DArrayOfMixedType = False
End Function

Function Test3DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 2, 1 To 3, 1 To 4)
3         For i = 1 To 24
4             SetAtLinear x, i, ChrW(i)
5         Next

6         y = CallFromExcel.JuliaCallVBA("identity", x)
7         Test3DArray = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test3DArray", Err, True)
10        Test3DArray = False
End Function

Function Test4DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 2, 1 To 3, 1 To 4, 1 To 5)

3         For i = 1 To 120
4             SetAtLinear x, i, i
5         Next

6         y = CallFromExcel.JuliaCallVBA("identity", x)
7         Test4DArray = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test4DArray", Err, True)
10        Test4DArray = False
End Function

Function TestDictionary()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary
          Dim z As New Scripting.Dictionary

1         On Error GoTo ErrHandler
2         z.Add "alpha", 100
3         z.Add "beta", 200

4         x.Add "a", 1
5         x.Add "b", 2
6         x.Add "c", "d"
7         x.Add "d", Array(1, 2, 3)
8         x.Add "e", z

9         Set y = CallFromExcel.JuliaCallVBA("identity", x)

10        ThrowIfError CallFromExcel.JuliaSetVar("first_dictionary", x)
11        ThrowIfError CallFromExcel.JuliaSetVar("second_dictionary", y)

12        TestDictionary = CallFromExcel.JuliaEval("first_dictionary == second_dictionary")

13        Exit Function
ErrHandler:
14        PrintTwice ReThrow("TestDictionary", Err, True)
15        TestDictionary = False
End Function

Function TestExactRoundTripping()
          Dim i As Long
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 1000)
3         For i = 1 To 1000
4             x(i) = Sqr(i)
5         Next i
6         y = CallFromExcel.JuliaCallVBA("identity", x)
7         TestExactRoundTripping = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestExactRoundTripping", Err, True)
10        TestExactRoundTripping = False
End Function

Function TestArrayOfDictionaries()
          Dim x() As Variant
          Dim y As Variant
          Dim z As New Scripting.Dictionary

1         On Error GoTo ErrHandler
2         z("a") = 1
3         z("b") = 2

4         ReDim x(1 To 2, 1 To 2)

5         Set x(1, 1) = z: Set x(1, 2) = z
6         Set x(2, 1) = z: Set x(2, 2) = z
7         y = CallFromExcel.JuliaCallVBA("identity", x)
8         ThrowIfError CallFromExcel.JuliaSetVar("first_array_of_dictionaries", x)
9         ThrowIfError CallFromExcel.JuliaSetVar("second_array_of_dictionaries", y)

10        TestArrayOfDictionaries = CallFromExcel.JuliaEval("first_array_of_dictionaries == second_array_of_dictionaries")

11        Exit Function
ErrHandler:
12        PrintTwice ReThrow("TestArrayOfDictionaries", Err, True)
13        TestArrayOfDictionaries = False
End Function

Function TestDictionaryOfArrays()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary

1         On Error GoTo ErrHandler

2         x("doubles") = Array(1#, 2#, 3#)
3         x("strings") = Array("foo", "bar")

4         Set y = CallFromExcel.JuliaCallVBA("identity", x)

5         ThrowIfError CallFromExcel.JuliaSetVar("first_dict_of_arrays", x)
6         ThrowIfError CallFromExcel.JuliaSetVar("second_dict_of_arrays", y)

7         TestDictionaryOfArrays = CallFromExcel.JuliaEval("first_dict_of_arrays == second_dict_of_arrays")

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestDictionaryOfArrays", Err, True)
10        TestDictionaryOfArrays = False
End Function

Function TestDictionaryOfTypes()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary

1         On Error GoTo ErrHandler

2         x.Add "Integer", 1
3         x.Add "Long", CLng(1)
4         x.Add "LongLong", CLngLng(1)
5         x.Add "Single", CSng(1)
6         x.Add "Double", CDbl(1)

7         Set y = CallFromExcel.JuliaCallVBA("identity", x)

8         TestDictionaryOfTypes = True
          Dim k As Variant
9         For Each k In y.Keys
10            If TypeName(y(k)) <> k Then
11                TestDictionaryOfTypes = False
12            End If
13        Next

14        Exit Function
ErrHandler:
15        PrintTwice ReThrow("TestDictionaryOfTypes", Err, True)
16        TestDictionaryOfTypes = False
End Function

' Int64 arrays travel in the compact "L" format, but VBA must still receive LongLong elements, as
' it did with JuliaExcel.
Function TestIntegerArrayIsLongLong() As Boolean
          Dim y As Variant
1         On Error GoTo ErrHandler
2         y = CallFromExcel.JuliaEvalVBA("[1, -2, typemax(Int64)]")
3         TestIntegerArrayIsLongLong = LBound(y) = 1 And UBound(y) = 3 And VarType(y(1)) = vbLongLong And _
              y(2) = -2 And y(3) = 9223372036854775807^
4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestIntegerArrayIsLongLong", Err, True)
6         TestIntegerArrayIsLongLong = False
End Function

' Dictionaries returned to VBA must be real Scripting.Dictionary objects, as they were with
' JuliaExcel, so that For Each, CompareMode etc. work, including for nested dictionaries.
Function TestDictionaryIsScriptingDictionary() As Boolean
          Dim d As Scripting.Dictionary
          Dim k As Variant
          Dim NumKeys As Long

1         On Error GoTo ErrHandler
2         Set d = CallFromExcel.JuliaEvalVBA("Dict(""a"" => 1.0, ""b"" => Dict(""c"" => 2.0))")
3         For Each k In d
4             NumKeys = NumKeys + 1
5         Next k
6         TestDictionaryIsScriptingDictionary = (TypeName(d) = "Dictionary") And (NumKeys = 2) And _
              (d("a") = 1) And (TypeName(d("b")) = "Dictionary") And (d("b")("c") = 2) And _
              (d.CompareMode = BinaryCompare)
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestDictionaryIsScriptingDictionary", Err, True)
9         TestDictionaryIsScriptingDictionary = False
End Function

Function TestOneDArraysDisplayAsOneColumnOnSheet()

          Dim OneDArray() As Variant

1         On Error GoTo ErrHandler
2         ReDim OneDArray(1 To 3)
3         OneDArray(1) = 1
4         OneDArray(2) = 2
5         OneDArray(3) = 3

6         TestOneDArraysDisplayAsOneColumnOnSheet = _
              NumDimensions(CallFromExcel.JuliaCall("identity", OneDArray)) = 2 And _
              NumDimensions(CallFromExcel.JuliaCallVBA("identity", OneDArray)) = 1

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestOneDArraysDisplayAsOneColumnOnSheet", Err, True)
9         TestOneDArraysDisplayAsOneColumnOnSheet = False
End Function

' Builds a 1-based (ReDim(1 To n)) Variant array from a ParamArray of values. Deliberately
' not VBA's own Array() function (which returns a 0-based array): when a 0-based array is
' passed as a single argument to a COM method whose trailing parameter is a "params object[]"
' (as JuliaCall/JuliaCallVBA/etc. all are), VBA's COM dispatch auto-spreads it into separate
' positional arguments instead of passing it as one array value - confirmed by isolating this
' exact case (TestElType originally used Array(...) here and failed with Julia reporting
' eltype() called with 3 separate arguments, not 1 array argument). A non-zero-lower-bound
' array (any ReDim(1 To n), the convention this whole codebase otherwise already uses) does
' not trigger that spreading, so this is the correct way to build a single array argument
' from VBA for any of the COM migration/VBA methods - not just a workaround for this test.
Private Function OneBasedArray(ParamArray Values()) As Variant
          Dim i As Long
          Dim Result() As Variant
1         ReDim Result(1 To UBound(Values) - LBound(Values) + 1)
2         For i = LBound(Values) To UBound(Values)
3             Assign Result(i - LBound(Values) + 1), Values(i)
4         Next i
5         OneBasedArray = Result
End Function

' Each array is built by its own statement into a local variable (Arr) before being passed to
' the late-bound JuliaCall - never as a nested OneBasedArray(...) call-expression argument.
' Nesting a function call directly as an argument to a late-bound COM Dispatch call (e.g.
' "CallFromExcel.JuliaCall(\"eltype\", OneBasedArray(1, 2, 3))") was found (while debugging this
' exact test) to unpredictably raise "Invalid procedure call or argument", even though the same
' call succeeds once the array is assigned to a variable first. Not fully explained, but
' reproducible and avoided this way - a real quirk of VBA's late-bound COM dispatch worth
' knowing about for any future code passing a function's return value directly into a
' ComServer method call.
Function TestElType()
          Dim Arr As Variant
          Dim OK As Boolean
          Dim R As Variant

1         On Error GoTo ErrHandler

2         OK = True

3         Arr = OneBasedArray(1, 2, 3)
4         R = CallFromExcel.JuliaCall("eltype", Arr)
5         If R <> "Int16" Then OK = False

6         Arr = OneBasedArray(1!, 2!, 3!)
7         R = CallFromExcel.JuliaCall("eltype", Arr)
8         If R <> "Float32" Then OK = False

9         Arr = OneBasedArray(1#, 2#, 3#)
10        R = CallFromExcel.JuliaCall("eltype", Arr)
11        If R <> "Float64" Then OK = False

12        Arr = OneBasedArray(1, 2#, 3#)
13        R = CallFromExcel.JuliaCall("eltype", Arr)
14        If R <> "Any" Then OK = False

15        Arr = OneBasedArray("a", "b", "c")
16        R = CallFromExcel.JuliaCall("eltype", Arr)
17        If R <> "String" Then OK = False

18        Arr = OneBasedArray("a", 1, True)
19        R = CallFromExcel.JuliaCall("eltype", Arr)
20        If R <> "Any" Then OK = False

21        Arr = OneBasedArray(True, False, True)
22        R = CallFromExcel.JuliaCall("eltype", Arr)
23        If R <> "Bool" Then OK = False

24        Arr = OneBasedArray(Empty, Empty, Empty)
25        R = CallFromExcel.JuliaCall("eltype", Arr)
26        If R <> "Missing" Then OK = False

27        TestElType = OK

28        Exit Function
ErrHandler:
29        PrintTwice ReThrow("TestElType", Err, True)
30        TestElType = False
End Function

Function TestBroadcasting()

          Dim ExpRes(1 To 2, 1 To 2) As Variant
          Dim ObsRes As Variant
          Dim Xs(1 To 1, 1 To 2) As Variant
          Dim Ys(1 To 2, 1 To 1) As Variant

1         On Error GoTo ErrHandler

2         Xs(1, 1) = 7#: Xs(1, 2) = 5#
3         Ys(1, 1) = 2#: Ys(2, 1) = 3#

4         ExpRes(1, 1) = 20#: ExpRes(1, 2) = 16#
5         ExpRes(2, 1) = 23#: ExpRes(2, 2) = 19#

6         ThrowIfError CallFromExcel.JuliaEval("fn(x,y) = 2x + 3y")
          'Note the dot character below, makes this a broadcast use of fn
7         ObsRes = CallFromExcel.JuliaCall("fn.", Xs, Ys)

8         TestBroadcasting = ArraysIdentical(ExpRes, ObsRes)

9         Exit Function
ErrHandler:
10        PrintTwice ReThrow("TestBroadcasting", Err, True)
11        TestBroadcasting = False
End Function

' Checks that a Vector{Float64} containing a NaN falls back from the compact "V" wire format
' to the general "*" format, so that the NaN element arrives as the Excel error #N/A rather
' than corrupting the fast binary decode.
Function TestNaNFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#
3         Expected(2, 1) = 2#
4         Expected(3, 1) = CVErr(2042) '#N/A
5         y = ThrowIfError(CallFromExcel.JuliaEval("[1.0,2.0,NaN]"))
6         TestNaNFallback = ArraysIdentical(Expected, y)

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestNaNFallback", Err, True)
9         TestNaNFallback = False
End Function

' As TestNaNFallback, but for Inf.
Function TestInfFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#
3         Expected(2, 1) = 2#
4         Expected(3, 1) = CVErr(2036) '#NUM!
5         y = ThrowIfError(CallFromExcel.JuliaEval("[1.0,2.0,Inf]"))
6         TestInfFallback = ArraysIdentical(Expected, y)

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestInfFallback", Err, True)
9         TestInfFallback = False
End Function

' Exercises the "V" branch's vector-to-column path, used by JuliaCall/JuliaEval but not by
' JuliaCallVBA.
Function TestVFormatVectorColumn()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 4)
3         x(1) = 1.5: x(2) = -2.25: x(3) = 0#: x(4) = 100000.125
4         Expected(1, 1) = x(1): Expected(2, 1) = x(2)
5         Expected(3, 1) = x(3): Expected(4, 1) = x(4)
6         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
7         TestVFormatVectorColumn = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestVFormatVectorColumn", Err, True)
10        TestVFormatVectorColumn = False
End Function

' Exercises the "V" branch's rank-2 (matrix) path with a non-square shape and distinct values
' in every cell, to catch a row/column transposition bug a square test could mask.
Function TestVFormatMatrixNonSquare()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim x(1 To 2, 1 To 3) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         x(1, 1) = 1.1: x(1, 2) = 2.2: x(1, 3) = 3.3
3         x(2, 1) = 4.4: x(2, 2) = 5.5: x(2, 3) = 6.6
4         Expected(1, 1) = x(1, 1): Expected(1, 2) = x(1, 2): Expected(1, 3) = x(1, 3)
5         Expected(2, 1) = x(2, 1): Expected(2, 2) = x(2, 2): Expected(2, 3) = x(2, 3)
6         y = CallFromExcel.JuliaCallVBA("identity", x)
7         TestVFormatMatrixNonSquare = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestVFormatMatrixNonSquare", Err, True)
10        TestVFormatMatrixNonSquare = False
End Function

' As TestNaNFallback, but for the matrix encoder, which has its own NaN/Inf fallback guard.
Function TestMatrixNaNFallback()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#: Expected(1, 2) = 2#: Expected(1, 3) = 3#
3         Expected(2, 1) = 4#
4         Expected(2, 2) = CVErr(2042) '#N/A
5         Expected(2, 3) = 6#
6         y = ThrowIfError(CallFromExcel.JuliaEval("[1.0 2.0 3.0; 4.0 NaN 6.0]"))
7         TestMatrixNaNFallback = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestMatrixNaNFallback", Err, True)
10        TestMatrixNaNFallback = False
End Function

' Checks an empty Float64 vector correctly routes to the general array format rather than
' attempting a zero-length "V" payload.
Function TestVFormatEmptyArrayFallback()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(CallFromExcel.JuliaEvalVBA("Float64[]"))
3         TestVFormatEmptyArrayFallback = IsArray(y) And (UBound(y) < LBound(y))

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestVFormatEmptyArrayFallback", Err, True)
6         TestVFormatEmptyArrayFallback = False
End Function

' Live round trip through a Variant() array (not a genuinely-typed Double() array) - the
' realistic "worst case" the "V" encoder has to handle in practice, since Range.Value2
' always arrives this way, even when every element holds a number. Built with ReDim
' (1-based), not the VBA.Array() function - Array() returns a 0-based array, which would
' make ArraysIdentical report a spurious mismatch against the decoder's always-1-based
' result, even when every value round-tripped correctly.
Function TestVEncodeFromVariantArray()
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 5)
3         x(1) = 1#: x(2) = -2.5: x(3) = 3.14159265358979: x(4) = 0#: x(5) = 1000000#
4         y = CallFromExcel.JuliaCallVBA("identity", x)
5         TestVEncodeFromVariantArray = ArraysIdentical(x, y)

6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("TestVEncodeFromVariantArray", Err, True)
8         TestVEncodeFromVariantArray = False
End Function

' Exercises the "V" branch's rank 3-9 handling. A genuinely-typed Double() array, so both
' directions of this round trip go via "V". Distinct values at every position (i + 10*j +
' 100*k) so a bug in the column-major index walk (either side) would show up as a value in
' the wrong place, not just a wrong-shaped result.
Function TestVFormat3DArray()
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim x(1 To 2, 1 To 3, 1 To 4) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         For k = 1 To 4
3             For j = 1 To 3
4                 For i = 1 To 2
5                     x(i, j, k) = i + 10 * j + 100 * k
6                 Next i
7             Next j
8         Next k

9         y = CallFromExcel.JuliaCallVBA("identity", x)
10        TestVFormat3DArray = ArraysIdentical(x, y)

11        Exit Function
ErrHandler:
12        PrintTwice ReThrow("TestVFormat3DArray", Err, True)
13        TestVFormat3DArray = False
End Function

' Exercises the "R" (range) branch's integer sub-format - a UnitRange is encoded as just
' first/step/length and reconstructed via arithmetic, rather than transmitting every element.
Function TestRangeIntegerVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(CallFromExcel.JuliaEval("1:1000"))
3         y2 = ThrowIfError(CallFromExcel.JuliaEval("collect(1:1000)"))
4         TestRangeIntegerVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeIntegerVsCollect", Err, True)
7         TestRangeIntegerVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for a StepRange with a non-1 step.
Function TestRangeStepVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(CallFromExcel.JuliaEval("5:3:47"))
3         y2 = ThrowIfError(CallFromExcel.JuliaEval("collect(5:3:47)"))
4         TestRangeStepVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeStepVsCollect", Err, True)
7         TestRangeStepVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for the float sub-format.
Function TestRangeFloatVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(CallFromExcel.JuliaEval("(1:1000).*pi"))
3         y2 = ThrowIfError(CallFromExcel.JuliaEval("collect((1:1000).*pi)"))
4         TestRangeFloatVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeFloatVsCollect", Err, True)
7         TestRangeFloatVsCollect = False
End Function

' Confirms JuliaEvalVBA gives a genuine 1-D array for a range result.
Function TestRangeJuliaEvalVBAIsOneD()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(CallFromExcel.JuliaEvalVBA("1:1000"))
3         TestRangeJuliaEvalVBAIsOneD = (NumDimensions(y) = 1) And (UBound(y) - LBound(y) + 1 = 1000) And _
              (y(LBound(y)) = 1) And (y(UBound(y)) = 1000)

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestRangeJuliaEvalVBAIsOneD", Err, True)
6         TestRangeJuliaEvalVBAIsOneD = False
End Function

' As TestRangeJuliaEvalVBAIsOneD, but via JuliaEval - must give a 2-D, single-column array.
Function TestRangeJuliaEvalIsColumn()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(CallFromExcel.JuliaEval("1:1000"))
3         TestRangeJuliaEvalIsColumn = (NumDimensions(y) = 2) And (UBound(y, 1) = 1000) And (UBound(y, 2) = 1) And _
              (y(1, 1) = 1) And (y(1000, 1) = 1000)

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestRangeJuliaEvalIsColumn", Err, True)
6         TestRangeJuliaEvalIsColumn = False
End Function

' Directly confirms the compact "R" wire format is actually used, for both sub-formats.
Function TestRangeWireFormat()
          Dim s As Variant

1         On Error GoTo ErrHandler
2         s = ThrowIfError(CallFromExcel.JuliaEvalVBA("CallFromExcel.encode_for_xl(1:1000)"))
3         If Left$(s, 3) <> "RI," Then Throw "Expected integer range to use 'RI' wire format, got: " & Left$(s, 10)
4         s = ThrowIfError(CallFromExcel.JuliaEvalVBA("CallFromExcel.encode_for_xl((1:1000).*pi)"))
5         If Left$(s, 3) <> "RF," Then Throw "Expected float range to use 'RF' wire format, got: " & Left$(s, 10)
6         TestRangeWireFormat = True

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestRangeWireFormat", Err, True)
9         TestRangeWireFormat = False
End Function

' Confirms an array containing NaN/Infinity travels Excel -> Julia via the fast "V" path, and
' Julia's own outbound encoder still correctly maps NaN -> #N/A and Infinity -> #NUM! on the
' way back.
Function TestNaNInfRoundTripViaV()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x(1 To 4) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         x(1) = 1#
3         x(2) = HexToDouble("7FF8000000000000") 'quiet NaN
4         x(3) = 2#
5         x(4) = HexToDouble("7FF0000000000000") 'positive infinity

6         Expected(1, 1) = 1#
7         Expected(2, 1) = CVErr(2042) '#N/A
8         Expected(3, 1) = 2#
9         Expected(4, 1) = CVErr(2036) '#NUM!

10        y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
11        TestNaNInfRoundTripViaV = ArraysIdentical(Expected, y)

12        Exit Function
ErrHandler:
13        PrintTwice ReThrow("TestNaNInfRoundTripViaV", Err, True)
14        TestNaNInfRoundTripViaV = False
End Function

' Confirms Excel errors round-trip correctly through Julia via JuliaCall("identity", ...),
' for all 14 Excel error codes, not just the two (2036, 2042) Julia can generate from Inf/NaN.
' Built with ReDim (1-based), not VBA.Array() (0-based).
Function TestExcelErrorRoundTrip()
          Dim Codes(1 To 14) As Long
          Dim Expected(1 To 14, 1 To 1) As Variant
          Dim i As Long
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Codes(1) = 2000: Codes(2) = 2007: Codes(3) = 2015: Codes(4) = 2023: Codes(5) = 2029
3         Codes(6) = 2036: Codes(7) = 2042: Codes(8) = 2043: Codes(9) = 2045: Codes(10) = 2046
4         Codes(11) = 2047: Codes(12) = 2048: Codes(13) = 2049: Codes(14) = 2050

5         For i = 1 To 14
6             Expected(i, 1) = CVErr(Codes(i))
7         Next i

8         y = ThrowIfError(CallFromExcel.JuliaCall("identity", Expected))
9         TestExcelErrorRoundTrip = ArraysIdentical(Expected, y)

10        Exit Function
ErrHandler:
11        PrintTwice ReThrow("TestExcelErrorRoundTrip", Err, True)
12        TestExcelErrorRoundTrip = False
End Function

Function TestArrayOfSingleCharacters()
          Dim x As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
          'All characters with code point <= 100000
2         x = Application.Evaluate("=UNICHAR(VSTACK(SEQUENCE(55295,,1),SEQUENCE(7632,,57344),SEQUENCE(526,,65008),SEQUENCE(34465,,65536)))")
3         y = ThrowIfError(CallFromExcel.JuliaCall("identity", x))
4         TestArrayOfSingleCharacters = ArraysIdentical(x, y)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestArrayOfSingleCharacters", Err, True)
7         TestArrayOfSingleCharacters = False
End Function

' Checks JuliaCallNamedArgs passes the arguments after ";" as keyword arguments, including
' with ";" also used as a positional argument and as the value of the last named argument,
' and that a missing ";", an odd number of arguments after it, and a repeated name are each
' reported as an error.
Function TestNamedArgs()
          Dim Expected1(1 To 4, 1 To 1) As Variant
          Dim Expected3(1 To 2, 1 To 1) As Variant
          Dim OK As Boolean

1         On Error GoTo ErrHandler
2         ThrowIfError CallFromExcel.JuliaEval("named_args_test(x, y=0; a=1, b=2) = [x, y, a, b]; named_args_join(xs...; delim="","") = join(xs, delim)")

3         Expected1(1, 1) = 10#: Expected1(2, 1) = 0#: Expected1(3, 1) = 1#: Expected1(4, 1) = 5#
4         OK = ArraysIdentical(Expected1, ThrowIfError(CallFromExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "b", 5#)))

5         OK = OK And Left$(CallFromExcel.JuliaCallNamedArgs("named_args_test", 10#), 15) = "#ArgumentError:"
6         OK = OK And Left$(CallFromExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "b", 5#, "a"), 15) = "#ArgumentError:"

7         Expected3(1, 1) = "a": Expected3(2, 1) = "b"
8         OK = OK And ArraysIdentical(Expected3, ThrowIfError(CallFromExcel.JuliaCallNamedArgs("split", "a;;b", ";", ";", "keepempty", False)))
9         OK = OK And CallFromExcel.JuliaCallNamedArgs("named_args_join", "a", ";", "b", ";", "delim", ";") = "a;;;b"

10        OK = OK And Left$(CallFromExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "a", 1, "a", 2), 15) = "#ArgumentError:"

11        TestNamedArgs = OK
12        Exit Function
ErrHandler:
13        PrintTwice ReThrow("TestNamedArgs", Err, True)
14        TestNamedArgs = False
End Function

' Checks JuliaCallNamedArgsVBA returns a vector as a 1-dimensional array, and that both
' argument and return types that can't live on a worksheet (here Dictionaries) work as named
' argument values.
Function TestNamedArgsVBA()
          Dim Expected(1 To 4) As Variant
          Dim OK As Boolean
          Dim x As New Scripting.Dictionary
          Dim y As Object

1         On Error GoTo ErrHandler
2         ThrowIfError CallFromExcel.JuliaEval("named_args_test(x, y=0; a=1, b=2) = [x, y, a, b]")
3         Expected(1) = 10#: Expected(2) = 0#: Expected(3) = 1#: Expected(4) = 5#
4         OK = ArraysIdentical(Expected, CallFromExcel.JuliaCallNamedArgsVBA("named_args_test", 10#, ";", "b", 5#))

5         ThrowIfError CallFromExcel.JuliaEval("named_args_dict(; kw...) = Dict(string(k) => v for (k, v) in kw)")
6         x.Add "alpha", 100
7         Set y = CallFromExcel.JuliaCallNamedArgsVBA("named_args_dict", ";", "d", x, "n", 1#)
8         OK = OK And y.Count = 2
9         OK = OK And y("n") = 1#
10        OK = OK And y("d")("alpha") = 100

11        TestNamedArgsVBA = OK
12        Exit Function
ErrHandler:
13        PrintTwice ReThrow("TestNamedArgsVBA", Err, True)
14        TestNamedArgsVBA = False
End Function

' Adapted from vba\JuliaExcel.xlam\VBA\modTest.bas's TestVFormatBoundarySizes: that version
' called the VBA-internal TrySerialiseArrayAsV/UnserialiseFromString directly, which no
' longer exist (now C#-only, covered by WireFormatTests.cs). This instead round-trips
' through the real add-in (JuliaCallVBA) at the same range of sizes - particularly small
' boundaries (1, 2, 3) and well beyond the usual benchmark size.
Function TestVFormatBoundarySizes() As Boolean
          Dim i As Long
          Dim Idx As Long
          Dim OK As Boolean
          Dim Sizes As Variant
          Dim Sz As Long
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Sizes = Array(1, 2, 3, 7, 8, 9, 17, 100, 1000, 1000000)

3         OK = True
4         For Idx = LBound(Sizes) To UBound(Sizes)
5             Sz = Sizes(Idx)
6             ReDim x(1 To Sz)
7             For i = 1 To Sz
8                 x(i) = (i - Sz / 2) * 1.5 + 0.25 'varied: negative, positive, fractional values
9             Next i
10            y = CallFromExcel.JuliaCallVBA("identity", x)
11            If Not ArraysIdentical(x, y) Then
12                Throw "Round trip mismatch for size " & Sz
13            End If
14        Next Idx

15        TestVFormatBoundarySizes = OK
16        Exit Function
ErrHandler:
17        PrintTwice ReThrow("TestVFormatBoundarySizes", Err, True)
18        TestVFormatBoundarySizes = False
End Function

' Regression test: ConcatenateExpressions used to join a multi-line JuliaExpression with ";"
' rather than a newline, so a "#" comment on one line could swallow the rest of the
' expression - including the "end" that closes a function - turning valid Julia code into a
' syntax error.
Function TestJuliaEvalMultiLineWithComment() As Boolean
          Dim Lines(1 To 3) As Variant
          Dim OK As Boolean
1         On Error GoTo ErrHandler
2         Lines(1) = "function RegressionTestForComment(a, b)"
3         Lines(2) = "a + b #this comment messes things up!"
4         Lines(3) = "end"
5         ThrowIfError CallFromExcel.JuliaEvalVBA(Lines)
6         OK = CallFromExcel.JuliaCall("RegressionTestForComment", 2, 3) = 5
7         TestJuliaEvalMultiLineWithComment = OK
8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestJuliaEvalMultiLineWithComment", Err, True)
10        TestJuliaEvalMultiLineWithComment = False
End Function

' =========================================================================================
' Tests of the real worksheet-formula / XLL surface. New for the C#/Excel-DNA add-in: the
' old VBA-only world never needed these as a separate thing from the JuliaCall/JuliaEval
' tests above, since a bare JuliaCall behaved identically whether called from a cell or a
' macro. Now the XLL and the COM migration methods are two independent implementations that
' could silently diverge, so the worksheet-formula path (the thing most users actually type)
' needs its own, direct coverage - not just an inference from the COM-path tests above.
' =========================================================================================

Private Function ScratchSheet() As Worksheet
1         On Error Resume Next
2         Set ScratchSheet = ThisWorkbook.Worksheets("Scratch")
3         On Error GoTo 0
4         If ScratchSheet Is Nothing Then
5             Set ScratchSheet = ThisWorkbook.Worksheets.Add
6             ScratchSheet.Name = "Scratch"
7         End If
8         ScratchSheet.Unprotect
9         ScratchSheet.Cells.Clear
End Function

Function WsTest_Scalar() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Value2 = Application.WorksheetFunction.Pi()
4         ws.Range("B1").Formula2 = "=JuliaCall(""identity"",A1)"
5         WsTest_Scalar = ws.Range("B1").Value2 = Application.WorksheetFunction.Pi()
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_Scalar", Err, True)
8         WsTest_Scalar = False
End Function

Function WsTest_String() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Value2 = "FooBar"
4         ws.Range("B1").Formula2 = "=JuliaCall(""identity"",A1)"
5         WsTest_String = ws.Range("B1").Value2 = "FooBar"
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_String", Err, True)
8         WsTest_String = False
End Function

' Confirms a 1-D range argument to a real worksheet formula comes back as a single column,
' exactly as TestOneDArraysDisplayAsOneColumnOnSheet already confirms for the COM migration
' path - this checks the real XLL path does the same.
' Uses a single-COLUMN range (A1:A3), not a single-row one: EncodeArray only collapses an
' Nx1 (N rows, 1 column) range to the fast 1-D wire representation - a 1xN (1 row, N
' columns) range stays a genuine 2-D array throughout and comes back the same shape it went
' in (a row, not a column) - confirmed by first writing this test with a 1-row range, which
' failed because the result correctly spilled as a row, not a column, matching the old VBA
' convention ("Nx1 2-D arrays collapse", not "1xN").
Function WsTest_OneDArrayBecomesColumn() As Boolean
          Dim Expected As Variant
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1:A3").Value2 = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
4         ws.Range("E1").Formula2 = "=JuliaCall(""identity"",A1:A3)"
5         Expected = ws.Range("E1:E3").Value2
6         WsTest_OneDArrayBecomesColumn = NumDimensions(Expected) = 2 And _
              Expected(1, 1) = 1 And Expected(2, 1) = 2 And Expected(3, 1) = 3
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("WsTest_OneDArrayBecomesColumn", Err, True)
9         WsTest_OneDArrayBecomesColumn = False
End Function

Function WsTest_TwoDArray() As Boolean
          Dim Observed As Variant
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1:B2").Value2 = Array(1, 2) 'row 1 only - fill row 2 explicitly below
4         ws.Range("A1").Value2 = 1: ws.Range("B1").Value2 = 2
5         ws.Range("A2").Value2 = 3: ws.Range("B2").Value2 = 4
6         ws.Range("D1").Formula2 = "=JuliaCall(""identity"",A1:B2)"
7         Observed = ws.Range("D1:E2").Value2
8         WsTest_TwoDArray = Observed(1, 1) = 1 And Observed(1, 2) = 2 And _
              Observed(2, 1) = 3 And Observed(2, 2) = 4
9         Exit Function
ErrHandler:
10        PrintTwice ReThrow("WsTest_TwoDArray", Err, True)
11        WsTest_TwoDArray = False
End Function

' Int64 arrays travel in the compact "L" format, decoded to Doubles for a worksheet.
Function WsTest_IntegerArrays() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaEval(""[10, -20, 30]"")"
4         ws.Range("C1").Formula2 = "=JuliaEval(""[1 2; 3 4]"")"
5         WsTest_IntegerArrays = ws.Range("A1").Value2 = 10 And ws.Range("A2").Value2 = -20 And _
              ws.Range("A3").Value2 = 30 And IsEmpty(ws.Range("A4").Value2) And IsEmpty(ws.Range("B1").Value2) And _
              ws.Range("C1").Value2 = 1 And ws.Range("D1").Value2 = 2 And ws.Range("C2").Value2 = 3 And ws.Range("D2").Value2 = 4
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_IntegerArrays", Err, True)
8         WsTest_IntegerArrays = False
End Function

Function WsTest_JuliaEvalScalar() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaEval(""1+1"")"
4         WsTest_JuliaEvalScalar = ws.Range("A1").Value2 = 2
5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("WsTest_JuliaEvalScalar", Err, True)
7         WsTest_JuliaEvalScalar = False
End Function

' Regression test: a bare .NET null (decoded from Julia's "nothing"/"missing") returned
' unconverted from an XLL worksheet function used to display as #NUM! - Excel-DNA's own
' XLOPER marshaling apparently falls back to that error for a plain null reference it
' doesn't recognise, rather than blank/0. Fixed by converting to ExcelEmpty.Value before
' returning (JuliaFunctions.cs). Matches old VBA's own behaviour: a UDF returning Empty
' displays as 0 on the worksheet (not a literal blank - ISBLANK on it is FALSE, since the
' cell holds a computed formula, not nothing at all).
Function WsTest_EmptyResult() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaEval(""nothing"")"
4         ws.Range("A2").Formula2 = "=JuliaEval(""missing"")"
5         ws.Range("A3").Formula2 = "=ISBLANK(A1)"
6         WsTest_EmptyResult = ws.Range("A1").Value2 = 0 And ws.Range("A2").Value2 = 0 And ws.Range("A3").Value2 = False
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("WsTest_EmptyResult", Err, True)
9         WsTest_EmptyResult = False
End Function

' Regression test: the worksheet JuliaLaunch function's parameter order must exactly match
' the old VBA worksheet function (modMain.bas): UseLinux, MinimiseWindow, CommandLineOptions,
' Packages, BashStatements, TimeOut. An earlier version of the C# port's XLL wrapper dropped
' UseLinux/MinimiseWindow/BashStatements from the signature entirely, silently shifting a
' formula like =JuliaLaunch(,,"-q") (skipping the first two positional args to set
' CommandLineOptions third, exactly as old formulas do) onto the wrong parameter (TimeOut),
' raising "Input string was not in a correct format." That TimeOut-parsing happens in
' JuliaFunctions.cs's wrapper itself, before JuliaLaunch.Launch() even checks whether Julia
' is already running - so this is tested with Julia already running (relying on an earlier
' test having launched it), deliberately avoiding any exit()-then-relaunch timing race
' (Julia's HTTP server can stop responding slightly before its console window actually
' closes, which can make a genuine relaunch attempt immediately afterwards flaky - a real
' but separate concern from the argument-order bug this test targets).
Function WsTest_JuliaLaunchArgOrder() As Boolean
          Dim Result As String
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaLaunch(,,""-q"")"
4         Result = ws.Range("A1").Text
5         WsTest_JuliaLaunchArgOrder = Left$(Result, 1) <> "#"
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_JuliaLaunchArgOrder", Err, True)
8         WsTest_JuliaLaunchArgOrder = False
End Function

' Regression test: JuliaLaunch must succeed when called immediately after JuliaEval("exit()"),
' with no delay at all. Julia's own shutdown tears things down in stages - its HTTP server
' stops responding and its console window's title goes blank before the OS process has
' actually exited - so a relaunch attempt in that window used to be indistinguishable from a
' genuinely stuck/crashed session (JuliaLaunch.cs's own "already running but not responding"
' error), even though it was really just a session finishing a normal, requested exit. Fixed
' by a short retry loop in JuliaLaunch.cs that only engages in that specific ambiguous state.
Function WsTest_RelaunchAfterExit() As Boolean
          Dim Result As String
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         CallFromExcel.JuliaEval "exit()"
4         ws.Range("A1").Formula2 = "=JuliaLaunch(,,""" & TestCommandLineOptions & """)"
5         Result = ws.Range("A1").Text
6         WsTest_RelaunchAfterExit = Left$(Result, 1) <> "#" And CallFromExcel.JuliaIsRunning()
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("WsTest_RelaunchAfterExit", Err, True)
9         WsTest_RelaunchAfterExit = False
End Function

' Smoke test for Julia under WSL. Not part of RunTests (which needs a native session): run it by
' hand from the VBA editor, in a visible Excel with no Julia session running. (It hangs when run
' headless via Run-VbaTests.ps1 - not investigated.) One-time setup - instantiate this repo's
' dependencies in WSL's Julia:
' wsl bash -lc "julia --project=/mnt/c/Projects/CallFromExcel -e 'using Pkg; Pkg.instantiate()'"
Function RunWSLTest(Optional SilentMode As Boolean = False) As Boolean
          Const Title = "CallFromExcel RunWSLTest"
          Dim Prompt As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         LoadAddIn
3         If CallFromExcel.JuliaIsRunning() Then
4             Throw "Julia is already running - close it, then run RunWSLTest again"
5         End If

6         Set ws = ScratchSheet()
7         ws.Range("A1").Formula2 = "=JuliaLaunch(TRUE,,""--project=" & WindowsPathToWSL(RepoRoot()) & """)"
8         ThrowIfError ws.Range("A1").Value2
9         ws.Range("A2").Formula2 = "=JuliaEval(""Sys.islinux()"")"
10        RunWSLTest = (ws.Range("A2").Value2 = True) And (CallFromExcel.JuliaCall("identity", 1) = 1)

11        Prompt = IIf(RunWSLTest, "RunWSLTest passed", "RunWSLTest FAILED: Julia did not report running under Linux")
12        PrintTwice Prompt
13        CallFromExcel.JuliaEval "exit()"
14        If Not SilentMode Then MsgBox Prompt, IIf(RunWSLTest, vbInformation, vbCritical), Title
15        Exit Function
ErrHandler:
16        Prompt = ReThrow("RunWSLTest", Err, True)
17        TestLog = TestLog & Prompt & vbCrLf
18        If Not SilentMode Then MsgBox Prompt, vbCritical, Title
19        RunWSLTest = False
End Function

' Cross-path consistency check: JuliaIsRunning/JuliaExcelPID as real worksheet formulas
' should agree with the same functions called via COM - both run in-process in this same
' Excel.exe, so the process id must be identical either way, and Julia was already launched
' (via COM) earlier in RunTests.
Function WsTest_JuliaIsRunningAndPID() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaIsRunning()"
4         ws.Range("A2").Formula2 = "=JuliaExcelPID()"
5         WsTest_JuliaIsRunningAndPID = (ws.Range("A1").Value2 = True) And _
              (ws.Range("A2").Value2 = CallFromExcel.JuliaExcelPID())
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_JuliaIsRunningAndPID", Err, True)
8         WsTest_JuliaIsRunningAndPID = False
End Function

Function WsTest_ErrorString() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaCall(""thisFunctionDoesNotExist12345"")"
4         WsTest_ErrorString = Left$(ws.Range("A1").Value2, 1) = "#"
5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("WsTest_ErrorString", Err, True)
7         WsTest_ErrorString = False
End Function

' The scenario from dotnet-addin\README.md's performance write-up: JuliaCall("sort", ...)
' against a real worksheet range, kept here as a small (not 1,000,000-element) permanent
' regression test, compared against Excel's own SORT().
Function WsTest_SortRoundTrip() As Boolean
          Dim Expected As Variant
          Dim Observed As Variant
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1:A5").Value2 = Application.WorksheetFunction.Transpose(Array(5, 3, 1, 4, 2))
4         ws.Range("C1").Formula2 = "=SORT(A1:A5)"
5         ws.Range("E1").Formula2 = "=JuliaCall(""sort"",A1:A5)"
6         Expected = ws.Range("C1:C5").Value2
7         Observed = ws.Range("E1:E5").Value2
8         WsTest_SortRoundTrip = ArraysIdentical(Expected, Observed)
9         Exit Function
ErrHandler:
10        PrintTwice ReThrow("WsTest_SortRoundTrip", Err, True)
11        WsTest_SortRoundTrip = False
End Function

' Regression test: Excel-DNA passes an error argument (e.g. 1/0) as its own ExcelError type,
' which the encoder did not recognise, so =JuliaCall("identity",1/0) returned the string
' "#Cannot serialise value of type ExcelError" instead of #DIV/0!.
Function WsTest_ErrorArguments() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaCall(""identity"",1/0)"
4         ws.Range("A2").Formula2 = "=JuliaCall(""identity"",NA())"
5         ws.Range("A3").Formula2 = "=JuliaCall(""string"",1/0)"
6         WsTest_ErrorArguments = IsErrorCode(ws.Range("A1").Value2, xlErrDiv0) And _
              IsErrorCode(ws.Range("A2").Value2, xlErrNA) And _
              ws.Range("A3").Value2 = "#DIV/0!"
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("WsTest_ErrorArguments", Err, True)
9         WsTest_ErrorArguments = False
End Function

' Regression test: Excel-DNA passes blank cells within a range as its own ExcelEmpty type,
' which the encoder did not recognise, so any range containing a blank cell failed with
' "#Cannot serialise value of type ExcelEmpty". Blank cells should arrive in Julia as missing,
' and error cells as ExcelError.
Function WsTest_RangeWithBlanksAndErrors() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Value2 = 1
4         ws.Range("A3").Formula2 = "=NA()"
5         ws.Range("C1").Formula2 = "=JuliaCall(""string"",A1:A3)"
6         ws.Range("E1").Formula2 = "=JuliaCall(""identity"",A1:A3)"
7         WsTest_RangeWithBlanksAndErrors = ws.Range("C1").Value2 = "Any[1.0, missing, #N/A]" And _
              ws.Range("E1").Value2 = 1 And ws.Range("E2").Value2 = 0 And _
              IsErrorCode(ws.Range("E3").Value2, xlErrNA)
8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("WsTest_RangeWithBlanksAndErrors", Err, True)
10        WsTest_RangeWithBlanksAndErrors = False
End Function

' Regression test: JuliaSetVar used to assign the printed form of the value, re-parsed as Julia
' code, which failed for strings ("hello" became the variable name hello) and for error values.
Function WsTest_JuliaSetVar() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Value2 = 1
4         ws.Range("A3").Formula2 = "=NA()"
5         ws.Range("C1").Formula2 = "=JuliaSetVar(""ws_setvar_range"",A1:A3)"
6         ws.Range("C2").Formula2 = "=JuliaSetVar(""ws_setvar_string"",""hello"")"
7         ws.Range("D1").Formula2 = "=JuliaEval(""string(ws_setvar_range)"")"
8         ws.Range("D2").Formula2 = "=JuliaEval(""ws_setvar_string"")"
9         WsTest_JuliaSetVar = ws.Range("D1").Value2 = "Any[1.0, missing, #N/A]" And _
              ws.Range("D2").Value2 = "hello"
10        Exit Function
ErrHandler:
11        PrintTwice ReThrow("WsTest_JuliaSetVar", Err, True)
12        WsTest_JuliaSetVar = False
End Function

' Regression test: a blank cell in a range of expressions passed to JuliaEval was sent to
' Julia as the text "ExcelDna.Integration.ExcelEmpty" instead of as a blank line.
Function WsTest_JuliaEvalRangeWithBlankLine() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("B1").Value2 = "ws_blank_line_test = 41"
4         ws.Range("B3").Value2 = "ws_blank_line_test + 1"
5         ws.Range("A1").Formula2 = "=JuliaEval(B1:B3)"
6         WsTest_JuliaEvalRangeWithBlankLine = ws.Range("A1").Value2 = 42
7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("WsTest_JuliaEvalRangeWithBlankLine", Err, True)
9         WsTest_JuliaEvalRangeWithBlankLine = False
End Function

' Regression test: a Range object passed from VBA to the COM methods arrived as a COM object
' the encoder tried to read as a dictionary, so JuliaCall("sum", Range(...)) failed. JuliaExcel
' (and the README) use the Range's Value2.
Function TestRangeArgumentsVBA() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1:A3").Value2 = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
4         ws.Range("B1").Value2 = "vba_range_test = 41"
5         ws.Range("B3").Value2 = "vba_range_test + 1"
6         ThrowIfError CallFromExcel.JuliaSetVar("vba_range_var", ws.Range("A1:A3"))
7         TestRangeArgumentsVBA = CallFromExcel.JuliaCall("sum", ws.Range("A1:A3")) = 6 And _
              CallFromExcel.JuliaCallVBA("sum", ws.Range("A1:A3")) = 6 And _
              CallFromExcel.JuliaEval("sum(vba_range_var)") = 6 And _
              CallFromExcel.JuliaEvalVBA(ws.Range("B1:B3")) = 42
8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestRangeArgumentsVBA", Err, True)
10        TestRangeArgumentsVBA = False
End Function

' Regression test: JuliaLaunch's arguments given as references to blank cells used to fail
' (Convert.ToBoolean of a blank) or pass the text "ExcelDna.Integration.ExcelEmpty" on as an
' option. A blank cell now means the same as an omitted argument. Julia is already running
' here, so this checks the argument handling, which happens before that check.
Function WsTest_JuliaLaunchBlankArgs() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaLaunch(Z1,Z2,Z3,Z4,Z5,Z6)"
4         ws.Range("A2").Formula2 = "=JuliaLaunch(NA())"
5         WsTest_JuliaLaunchBlankArgs = Left$(ws.Range("A1").Text, 1) <> "#" And _
              IsErrorCode(ws.Range("A2").Value2, xlErrNA)
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_JuliaLaunchBlankArgs", Err, True)
8         WsTest_JuliaLaunchBlankArgs = False
End Function

' Regression test: =JuliaEval(A1) with A1 blank evaluated the text
' "ExcelDna.Integration.ExcelEmpty". A blank is now an empty expression (giving nothing,
' shown as 0), and an error value is returned as itself.
Function WsTest_JuliaEvalBlankOrErrorCell() As Boolean
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ScratchSheet()
3         ws.Range("A1").Formula2 = "=JuliaEval(Z1)"
4         ws.Range("A2").Formula2 = "=JuliaEval(NA())"
5         WsTest_JuliaEvalBlankOrErrorCell = ws.Range("A1").Value2 = 0 And _
              IsErrorCode(ws.Range("A2").Value2, xlErrNA)
6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("WsTest_JuliaEvalBlankOrErrorCell", Err, True)
8         WsTest_JuliaEvalBlankOrErrorCell = False
End Function

Private Function IsErrorCode(Value As Variant, ErrorNumber As Long) As Boolean
1         If IsError(Value) Then IsErrorCode = (CLng(Value) = ErrorNumber)
End Function

