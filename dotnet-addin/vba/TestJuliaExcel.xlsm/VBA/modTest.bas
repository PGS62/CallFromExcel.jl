Attribute VB_Name = "modTest"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme
'
' Port of vba\JuliaExcel.xlam\VBA\modTest.bas's RunTests for the C#/Excel-DNA add-in
' (dotnet-addin). There is no xlam in the new world, so every old call to a bare
' JuliaCall/JuliaEval/JuliaCallVBA/... Public Function becomes a call to the matching method
' on JuliaExcel, a JuliaExcel.AddIn.ComServer COM object declared below - see
' dotnet-addin\README.md's migration notes. Two call surfaces now exist where the old VBA-
' only world had one:
'   - JuliaExcel.JuliaCall/JuliaEval/JuliaCallNamedArgs/JuliaLaunch/JuliaIsRunning/
'     JuliaExcelPID/JuliaSetVar/JuliaInclude: the migration surface for code that used to call
'     these as bare Public Functions. Keeps the original worksheet-style restrictions
'     (string-length limit, no nested arrays/Dictionaries, 1-D arrays become a column).
'   - JuliaExcel.JuliaCallVBA/JuliaEvalVBA/JuliaCallNamedArgsVBA: unchanged from before,
'     lifting those restrictions.
' Tests below that used the first group port with a "JuliaExcel." prefix and nothing else
' changed. The WsTest_* functions at the end are new: they exercise the real worksheet-
' formula/XLL surface directly (writing real formulas into real cells), which the old suite
' never needed as a separate thing, since a bare JuliaCall behaved identically whether called
' from a cell or a macro - now the XLL and the COM migration methods are separate
' implementations, so the XLL path needs its own coverage.
'
' A returned Dictionary is no longer a real Scripting.Dictionary (it was, in the old VBA
' world) - it is a JuliaExcel.AddIn.ComServer JuliaDictionary COM object, exposing the same
' Count/Keys/Items/Item/Exists members (Item is also the default member, so y(k) still
' works) but not assignable to a variable declared As Scripting.Dictionary. Tests that
' receive a dictionary result now declare that variable As Object instead. A Dictionary
' passed as an INPUT argument (never received as a result) is unaffected and stays a real
' Scripting.Dictionary - this project requires a reference to "Microsoft Scripting Runtime".
'
' Not ported (no VBA-level equivalent exists any more - the functions they tested are now
' C#-only and already covered by dotnet-addin\JuliaExcel.AddIn.TestHarness\WireFormatTests.cs):
' TestSerialiseArrayAsV, TestMakeJuliaLiteral, TestUnserialiseStringTooLong,
' TestUnserialiseMalformedRank, TestConcatenateExpressions, TestTrim255.
' TestVFormatBoundarySizes is kept but adapted to round-trip through JuliaCallVBA rather than
' calling the (no longer existing) VBA-internal TrySerialiseArrayAsV/UnserialiseFromString
' directly.

Option Explicit

Private JuliaExcel As Object
Private TestLog As String

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Lets a non-interactive caller (Run-VbaTests.ps1) retrieve what PrintTwice wrote, since
' Debug.Print's output (the Immediate window) isn't visible over COM automation.
Function GetTestLog() As String
          GetTestLog = TestLog
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTests
' Purpose    : Test JuliaExcel's C#/Excel-DNA add-in for a variety of data types and both call
'              surfaces (worksheet/XLL and VBA/COM). Prints results to the Immediate window and
'              to a MsgBox.
' Parameters :
'  CommandLineOptions: Passed straight through to JuliaLaunch. Defaults to "", i.e. JuliaLaunch's
'                       own default environment.
' -----------------------------------------------------------------------------------------------------------------------
Function RunTests(Optional SilentMode As Boolean = False, Optional CommandLineOptions As String = "") As Boolean

          Const Title = "JuliaExcel RunTests"
          Dim NumFailed As Long
          Dim NumPassed As Long
          Dim Prompt As String
          Dim XllPath As String

          On Error GoTo ErrHandler

          Set JuliaExcel = CreateObject("JuliaExcel.AddIn.ComServer")

          XllPath = ThisWorkbook.Path & "\..\..\JuliaExcel.AddIn\bin\Release\net48\publish\JuliaExcel.AddIn64-packed.xll"
          If Dir(XllPath) = vbNullString Then
              Throw "Cannot find the add-in at " & XllPath & _
                  " - build it first (dotnet build -c Release in dotnet-addin\JuliaExcel.AddIn)."
          End If
          Application.RegisterXLL XllPath

          JuliaExcel.JuliaEval "exit()"
          Sleep 1000 'Give time to shut down properly, otherwise launch can fail thinking Julia still running but unresponsive
          ThrowIfError JuliaExcel.JuliaLaunch(, , CommandLineOptions)

          PrintTwice vbLf & String(80, "=")
          PrintTwice "JuliaExcel RunTests (C#/Excel-DNA add-in)"
          PrintTwice Format(Now, "yyyy-mm-dd hh:mm:ss")
          PrintTwice "ComputerName = " & Environ("ComputerName")

          'Tests ported from vba\JuliaExcel.xlam\VBA\modTest.bas - the JuliaExcel.<Name>/COM surface.
          AccResult "TestEmpty", TestEmpty, NumPassed, NumFailed
          AccResult "TestBoolean", TestBoolean, NumPassed, NumFailed
          AccResult "TestDouble", TestDouble, NumPassed, NumFailed
          AccResult "TestString", TestString, NumPassed, NumFailed
          AccResult "TestWideString", TestWideString, NumPassed, NumFailed
          AccResult "TestLong", TestLong, NumPassed, NumFailed
          AccResult "TestLongLong", TestLongLong, NumPassed, NumFailed
          AccResult "TestSingle", TestSingle, NumPassed, NumFailed
          AccResult "TestDate", TestDate, NumPassed, NumFailed
          AccResult "TestDateTime", TestDateTime, NumPassed, NumFailed
          AccResult "TestByte", TestByte, NumPassed, NumFailed
          AccResult "Test1DArrayOfDoubles", Test1DArrayOfDoubles, NumPassed, NumFailed
          AccResult "Test2DArrayOfMixedType", Test2DArrayOfMixedType, NumPassed, NumFailed
          AccResult "Test3DArray", Test3DArray, NumPassed, NumFailed
          AccResult "Test4DArray", Test4DArray, NumPassed, NumFailed
          AccResult "TestDictionary", TestDictionary, NumPassed, NumFailed
          AccResult "TestExactRoundTripping", TestExactRoundTripping, NumPassed, NumFailed
          AccResult "TestArrayOfDictionaries", TestArrayOfDictionaries, NumPassed, NumFailed
          AccResult "TestDictionaryOfArrays", TestDictionaryOfArrays, NumPassed, NumFailed
          AccResult "TestDictionaryOfTypes", TestDictionaryOfTypes, NumPassed, NumFailed
          AccResult "TestOneDArraysDisplayAsOneColumnOnSheet", TestOneDArraysDisplayAsOneColumnOnSheet, NumPassed, NumFailed
          AccResult "TestElType", TestElType, NumPassed, NumFailed
          AccResult "TestBroadcasting", TestBroadcasting, NumPassed, NumFailed
          AccResult "TestNaNFallback", TestNaNFallback, NumPassed, NumFailed
          AccResult "TestInfFallback", TestInfFallback, NumPassed, NumFailed
          AccResult "TestVFormatVectorColumn", TestVFormatVectorColumn, NumPassed, NumFailed
          AccResult "TestVFormatMatrixNonSquare", TestVFormatMatrixNonSquare, NumPassed, NumFailed
          AccResult "TestMatrixNaNFallback", TestMatrixNaNFallback, NumPassed, NumFailed
          AccResult "TestVFormatEmptyArrayFallback", TestVFormatEmptyArrayFallback, NumPassed, NumFailed
          AccResult "TestVEncodeFromVariantArray", TestVEncodeFromVariantArray, NumPassed, NumFailed
          AccResult "TestVFormat3DArray", TestVFormat3DArray, NumPassed, NumFailed
          AccResult "TestRangeIntegerVsCollect", TestRangeIntegerVsCollect, NumPassed, NumFailed
          AccResult "TestRangeStepVsCollect", TestRangeStepVsCollect, NumPassed, NumFailed
          AccResult "TestRangeFloatVsCollect", TestRangeFloatVsCollect, NumPassed, NumFailed
          AccResult "TestRangeJuliaEvalVBAIsOneD", TestRangeJuliaEvalVBAIsOneD, NumPassed, NumFailed
          AccResult "TestRangeJuliaEvalIsColumn", TestRangeJuliaEvalIsColumn, NumPassed, NumFailed
          AccResult "TestRangeWireFormat", TestRangeWireFormat, NumPassed, NumFailed
          AccResult "TestNaNInfRoundTripViaV", TestNaNInfRoundTripViaV, NumPassed, NumFailed
          AccResult "TestExcelErrorRoundTrip", TestExcelErrorRoundTrip, NumPassed, NumFailed
          AccResult "TestArrayOfSingleCharacters", TestArrayOfSingleCharacters, NumPassed, NumFailed
          AccResult "TestNamedArgs", TestNamedArgs, NumPassed, NumFailed
          AccResult "TestNamedArgsVBA", TestNamedArgsVBA, NumPassed, NumFailed
          AccResult "TestVFormatBoundarySizes", TestVFormatBoundarySizes, NumPassed, NumFailed
          AccResult "TestJuliaEvalMultiLineWithComment", TestJuliaEvalMultiLineWithComment, NumPassed, NumFailed

          'New: tests of the real worksheet-formula / XLL surface. No equivalent existed in the
          'old suite - a bare JuliaCall behaved identically whether called from a cell or a
          'macro, but the XLL and the COM migration methods above are now separate
          'implementations that can diverge, so this surface needs its own coverage.
          AccResult "WsTest_Scalar", WsTest_Scalar, NumPassed, NumFailed
          AccResult "WsTest_String", WsTest_String, NumPassed, NumFailed
          AccResult "WsTest_OneDArrayBecomesColumn", WsTest_OneDArrayBecomesColumn, NumPassed, NumFailed
          AccResult "WsTest_TwoDArray", WsTest_TwoDArray, NumPassed, NumFailed
          AccResult "WsTest_JuliaEvalScalar", WsTest_JuliaEvalScalar, NumPassed, NumFailed
          AccResult "WsTest_EmptyResult", WsTest_EmptyResult, NumPassed, NumFailed
          AccResult "WsTest_JuliaLaunchArgOrder", WsTest_JuliaLaunchArgOrder, NumPassed, NumFailed
          AccResult "WsTest_JuliaIsRunningAndPID", WsTest_JuliaIsRunningAndPID, NumPassed, NumFailed
          AccResult "WsTest_ErrorString", WsTest_ErrorString, NumPassed, NumFailed
          AccResult "WsTest_SortRoundTrip", WsTest_SortRoundTrip, NumPassed, NumFailed

          Prompt = NumPassed & " test(s) passed" & vbLf & _
              NumFailed & " test(s) failed"

          If NumFailed > 0 Then
              Prompt = Prompt & vbLf & vbLf & "See VBA Immediate window for details"
          End If

          PrintTwice NumPassed & " test(s) passed"
          PrintTwice NumFailed & " test(s) failed"
          PrintTwice String(80, "=")

          If Not SilentMode Then
              AppActivate Application.Caption
              MsgBox Prompt, IIf(NumFailed = 0, vbInformation, vbCritical), Title
          End If

          RunTests = NumFailed = 0

          Exit Function
ErrHandler:
          Dim ErrText As String
          ErrText = ReThrow("RunTests", Err, True)
          TestLog = TestLog & ErrText & vbCrLf
          If Not SilentMode Then
              MsgBox ErrText, vbCritical, Title
          End If
          RunTests = False
End Function

Sub PrintTwice(Text As String)
          On Error GoTo ErrHandler
          ThrowIfError JuliaExcel.JuliaEval("println(" & SimpleJuliaStringLiteral(Text) & ")")
          Debug.Print Text
          TestLog = TestLog & Text & vbCrLf

          Exit Sub
ErrHandler:
          ReThrow "PrintTwice", Err
End Sub

' Minimal ASCII-safe Julia string literal, enough for the plain diagnostic text PrintTwice is
' ever called with (test names, dates, computer names) - not a port of the old MakeJuliaLiteral
' (which also escaped Unicode bidi-control characters), since that is a C#-only implementation
' detail now (JuliaExcel.AddIn\src\JuliaLiteral.cs), not a public VBA-callable function.
Private Function SimpleJuliaStringLiteral(Text As String) As String
          Dim Escaped As String
          Escaped = Replace(Text, "\", "\\")
          Escaped = Replace(Escaped, Chr(34), "\" & Chr(34))
          Escaped = Replace(Escaped, vbCr, "\r")
          Escaped = Replace(Escaped, vbLf, "\n")
          Escaped = Replace(Escaped, "$", "\$")
          SimpleJuliaStringLiteral = Chr(34) & Escaped & Chr(34)
End Function

Sub AccResult(TestName As String, Result As Boolean, ByRef NumPassed, ByRef NumFailed)
          On Error GoTo ErrHandler
          PrintTwice "Test " & TestName & " completed"
          If Result Then
              NumPassed = NumPassed + 1
          Else
              PrintTwice "Test " & TestName & " Failed!"
              NumFailed = NumFailed + 1
          End If
          Exit Sub
ErrHandler:
          ReThrow "AccResult", Err
End Sub

Function TestEmpty()
          On Error GoTo ErrHandler
          TestEmpty = IsEmpty(JuliaExcel.JuliaCall("identity", Empty))
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestEmpty", Err, True)
          TestEmpty = False
End Function

Function TestBoolean()
          On Error GoTo ErrHandler
          TestBoolean = (JuliaExcel.JuliaCall("identity", True) = True) And (JuliaExcel.JuliaCall("identity", False) = False)
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestBoolean", Err, True)
          TestBoolean = False
End Function

Function TestDouble()
          Dim x As Double
          Dim y As Variant
          On Error GoTo ErrHandler
          x = Application.WorksheetFunction.Pi()
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestDouble = (x = y) And (VarType(y) = vbDouble)
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDouble", Err, True)
          TestDouble = False
End Function

Function TestString()
          Dim x As String
          Dim y As Variant
          On Error GoTo ErrHandler
          x = "FooBar"
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestString = x = y
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestString", Err, True)
          TestString = False
End Function

Function TestWideString()
          Dim i As Long
          Dim x As String
          Dim y As Variant
          On Error GoTo ErrHandler
          x = String(10000, " ")
          For i = 1 To 10000
              Mid$(x, i, 1) = ChrW(i)
          Next i
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestWideString = (x = y) And VarType(y) = vbString

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestWideString", Err, True)
          TestWideString = False
End Function

Function TestLong()
          Dim x As Long
          Dim y As Variant
          On Error GoTo ErrHandler
          x = 123456789
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestLong = (x = y) And VarType(y) = vbLong

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestLong", Err, True)
          TestLong = False
End Function

Function TestLongLong()
          Dim x As LongLong
          Dim y As Variant
          On Error GoTo ErrHandler
          x = 123456789^
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestLongLong = (x = y) And VarType(y) = vbLongLong

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestLongLong", Err, True)
          TestLongLong = False
End Function

Function TestSingle()
          Dim x As Single
          Dim y As Variant
          On Error GoTo ErrHandler
          x = CSng(1 / 3)
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestSingle = (x = y) And (VarType(y) = vbSingle)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestSingle", Err, True)
          TestSingle = False
End Function

Function TestDate()
          Dim x As Date
          Dim y As Variant
          On Error GoTo ErrHandler
          x = DateSerial(2025, 12, 22)
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestDate = (x = y) And VarType(y) = vbDate

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDate", Err, True)
          TestDate = False
End Function

Function TestDateTime()
          Dim x As Date
          Dim y As Variant
          On Error GoTo ErrHandler
          x = DateSerial(2025, 12, 22) + TimeValue("03:40:33")
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestDateTime = x = y

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDateTime", Err, True)
          TestDateTime = False
End Function

Function TestByte()
          Dim x As Byte
          Dim y As Variant
          On Error GoTo ErrHandler
          x = 200
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestByte = (x = y) And VarType(y) = vbByte

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestByte", Err, True)
          TestByte = False
End Function

Function Test1DArrayOfDoubles()
          Dim x() As Double
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 3)
          x(1) = 1 / 3
          x(2) = 1E+100
          x(3) = 0
          y = JuliaExcel.JuliaCallVBA("identity", x)
          Test1DArrayOfDoubles = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("Test1DArrayOfDoubles", Err, True)
          Test1DArrayOfDoubles = False
End Function

Function Test2DArrayOfMixedType()
          Dim x() As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 3, 1 To 3)
          x(1, 1) = 1 / 3:  x(1, 2) = 1E+100:    x(1, 3) = CLng(100)
          x(2, 1) = "Foo":  x(2, 2) = CSng(3):   x(2, 3) = CLngLng(100)
          x(3, 1) = "Foo":  x(3, 2) = CSng(3):   x(3, 3) = CInt(100)

          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          Test2DArrayOfMixedType = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("Test2DArrayOfMixedType", Err, True)
          Test2DArrayOfMixedType = False
End Function

Function Test3DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 2, 1 To 3, 1 To 4)
          For i = 1 To 24
              SetAtLinear x, i, ChrW(i)
          Next

          y = JuliaExcel.JuliaCallVBA("identity", x)
          Test3DArray = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("Test3DArray", Err, True)
          Test3DArray = False
End Function

Function Test4DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 2, 1 To 3, 1 To 4, 1 To 5)

          For i = 1 To 120
              SetAtLinear x, i, i
          Next

          y = JuliaExcel.JuliaCallVBA("identity", x)
          Test4DArray = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("Test4DArray", Err, True)
          Test4DArray = False
End Function

Function TestDictionary()

          Dim x As New Scripting.Dictionary
          Dim y As Object
          Dim z As New Scripting.Dictionary

          On Error GoTo ErrHandler
          z.Add "alpha", 100
          z.Add "beta", 200

          x.Add "a", 1
          x.Add "b", 2
          x.Add "c", "d"
          x.Add "d", Array(1, 2, 3)
          x.Add "e", z

          Set y = JuliaExcel.JuliaCallVBA("identity", x)

          ThrowIfError JuliaExcel.JuliaSetVar("first_dictionary", x)
          ThrowIfError JuliaExcel.JuliaSetVar("second_dictionary", y)

          TestDictionary = JuliaExcel.JuliaEval("first_dictionary == second_dictionary")

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDictionary", Err, True)
          TestDictionary = False
End Function

Function TestExactRoundTripping()
          Dim i As Long
          Dim x() As Double
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 1000)
          For i = 1 To 1000
              x(i) = Sqr(i)
          Next i
          y = JuliaExcel.JuliaCallVBA("identity", x)
          TestExactRoundTripping = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestExactRoundTripping", Err, True)
          TestExactRoundTripping = False
End Function

Function TestArrayOfDictionaries()
          Dim x() As Variant
          Dim y As Variant
          Dim z As New Scripting.Dictionary

          On Error GoTo ErrHandler
          z("a") = 1
          z("b") = 2

          ReDim x(1 To 2, 1 To 2)

          Set x(1, 1) = z: Set x(1, 2) = z
          Set x(2, 1) = z: Set x(2, 2) = z
          y = JuliaExcel.JuliaCallVBA("identity", x)
          ThrowIfError JuliaExcel.JuliaSetVar("first_array_of_dictionaries", x)
          ThrowIfError JuliaExcel.JuliaSetVar("second_array_of_dictionaries", y)

          TestArrayOfDictionaries = JuliaExcel.JuliaEval("first_array_of_dictionaries == second_array_of_dictionaries")

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestArrayOfDictionaries", Err, True)
          TestArrayOfDictionaries = False
End Function

Function TestDictionaryOfArrays()

          Dim x As New Scripting.Dictionary
          Dim y As Object

          On Error GoTo ErrHandler

          x("doubles") = Array(1#, 2#, 3#)
          x("strings") = Array("foo", "bar")

          Set y = JuliaExcel.JuliaCallVBA("identity", x)

          ThrowIfError JuliaExcel.JuliaSetVar("first_dict_of_arrays", x)
          ThrowIfError JuliaExcel.JuliaSetVar("second_dict_of_arrays", y)

          TestDictionaryOfArrays = JuliaExcel.JuliaEval("first_dict_of_arrays == second_dict_of_arrays")

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDictionaryOfArrays", Err, True)
          TestDictionaryOfArrays = False
End Function

Function TestDictionaryOfTypes()

          Dim x As New Scripting.Dictionary
          Dim y As Object

          On Error GoTo ErrHandler

          x.Add "Integer", 1
          x.Add "Long", CLng(1)
          x.Add "LongLong", CLngLng(1)
          x.Add "Single", CSng(1)
          x.Add "Double", CDbl(1)

          Set y = JuliaExcel.JuliaCallVBA("identity", x)

          TestDictionaryOfTypes = True
          Dim k As Variant
          For Each k In y.Keys
              If TypeName(y(k)) <> k Then
                  TestDictionaryOfTypes = False
              End If
          Next

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestDictionaryOfTypes", Err, True)
          TestDictionaryOfTypes = False
End Function

Function TestOneDArraysDisplayAsOneColumnOnSheet()

          Dim OneDArray() As Variant

          On Error GoTo ErrHandler
          ReDim OneDArray(1 To 3)
          OneDArray(1) = 1
          OneDArray(2) = 2
          OneDArray(3) = 3

          TestOneDArraysDisplayAsOneColumnOnSheet = _
              NumDimensions(JuliaExcel.JuliaCall("identity", OneDArray)) = 2 And _
              NumDimensions(JuliaExcel.JuliaCallVBA("identity", OneDArray)) = 1

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestOneDArraysDisplayAsOneColumnOnSheet", Err, True)
          TestOneDArraysDisplayAsOneColumnOnSheet = False
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
          ReDim Result(1 To UBound(Values) - LBound(Values) + 1)
          For i = LBound(Values) To UBound(Values)
              Assign Result(i - LBound(Values) + 1), Values(i)
          Next i
          OneBasedArray = Result
End Function

' Each array is built by its own statement into a local variable (Arr) before being passed to
' the late-bound JuliaCall - never as a nested OneBasedArray(...) call-expression argument.
' Nesting a function call directly as an argument to a late-bound COM Dispatch call (e.g.
' "JuliaExcel.JuliaCall(\"eltype\", OneBasedArray(1, 2, 3))") was found (while debugging this
' exact test) to unpredictably raise "Invalid procedure call or argument", even though the same
' call succeeds once the array is assigned to a variable first. Not fully explained, but
' reproducible and avoided this way - a real quirk of VBA's late-bound COM dispatch worth
' knowing about for any future code passing a function's return value directly into a
' ComServer method call.
Function TestElType()
          Dim OK As Boolean
          Dim Arr As Variant
          Dim R As Variant

          On Error GoTo ErrHandler

          OK = True

          Arr = OneBasedArray(1, 2, 3)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Int16" Then OK = False

          Arr = OneBasedArray(1!, 2!, 3!)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Float32" Then OK = False

          Arr = OneBasedArray(1#, 2#, 3#)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Float64" Then OK = False

          Arr = OneBasedArray(1, 2#, 3#)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Any" Then OK = False

          Arr = OneBasedArray("a", "b", "c")
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "String" Then OK = False

          Arr = OneBasedArray("a", 1, True)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Any" Then OK = False

          Arr = OneBasedArray(True, False, True)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Bool" Then OK = False

          Arr = OneBasedArray(Empty, Empty, Empty)
          R = JuliaExcel.JuliaCall("eltype", Arr)
          If R <> "Missing" Then OK = False

          TestElType = OK

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestElType", Err, True)
          TestElType = False
End Function

Function TestBroadcasting()

          Dim ExpRes(1 To 2, 1 To 2) As Variant
          Dim ObsRes As Variant
          Dim Xs(1 To 1, 1 To 2) As Variant
          Dim Ys(1 To 2, 1 To 1) As Variant

          On Error GoTo ErrHandler

          Xs(1, 1) = 7#: Xs(1, 2) = 5#
          Ys(1, 1) = 2#: Ys(2, 1) = 3#

          ExpRes(1, 1) = 20#: ExpRes(1, 2) = 16#
          ExpRes(2, 1) = 23#: ExpRes(2, 2) = 19#

          ThrowIfError JuliaExcel.JuliaEval("fn(x,y) = 2x + 3y")
          'Note the dot character below, makes this a broadcast use of fn
          ObsRes = JuliaExcel.JuliaCall("fn.", Xs, Ys)

          TestBroadcasting = ArraysIdentical(ExpRes, ObsRes)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestBroadcasting", Err, True)
          TestBroadcasting = False
End Function

' Checks that a Vector{Float64} containing a NaN falls back from the compact "V" wire format
' to the general "*" format, so that the NaN element arrives as the Excel error #N/A rather
' than corrupting the fast binary decode.
Function TestNaNFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          Expected(1, 1) = 1#
          Expected(2, 1) = 2#
          Expected(3, 1) = CVErr(2042) '#N/A
          y = ThrowIfError(JuliaExcel.JuliaEval("[1.0,2.0,NaN]"))
          TestNaNFallback = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestNaNFallback", Err, True)
          TestNaNFallback = False
End Function

' As TestNaNFallback, but for Inf.
Function TestInfFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          Expected(1, 1) = 1#
          Expected(2, 1) = 2#
          Expected(3, 1) = CVErr(2036) '#NUM!
          y = ThrowIfError(JuliaExcel.JuliaEval("[1.0,2.0,Inf]"))
          TestInfFallback = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestInfFallback", Err, True)
          TestInfFallback = False
End Function

' Exercises the "V" branch's vector-to-column path, used by JuliaCall/JuliaEval but not by
' JuliaCallVBA.
Function TestVFormatVectorColumn()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x() As Double
          Dim y As Variant

          On Error GoTo ErrHandler
          ReDim x(1 To 4)
          x(1) = 1.5: x(2) = -2.25: x(3) = 0#: x(4) = 100000.125
          Expected(1, 1) = x(1): Expected(2, 1) = x(2)
          Expected(3, 1) = x(3): Expected(4, 1) = x(4)
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestVFormatVectorColumn = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVFormatVectorColumn", Err, True)
          TestVFormatVectorColumn = False
End Function

' Exercises the "V" branch's rank-2 (matrix) path with a non-square shape and distinct values
' in every cell, to catch a row/column transposition bug a square test could mask.
Function TestVFormatMatrixNonSquare()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim x(1 To 2, 1 To 3) As Double
          Dim y As Variant

          On Error GoTo ErrHandler
          x(1, 1) = 1.1: x(1, 2) = 2.2: x(1, 3) = 3.3
          x(2, 1) = 4.4: x(2, 2) = 5.5: x(2, 3) = 6.6
          Expected(1, 1) = x(1, 1): Expected(1, 2) = x(1, 2): Expected(1, 3) = x(1, 3)
          Expected(2, 1) = x(2, 1): Expected(2, 2) = x(2, 2): Expected(2, 3) = x(2, 3)
          y = JuliaExcel.JuliaCallVBA("identity", x)
          TestVFormatMatrixNonSquare = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVFormatMatrixNonSquare", Err, True)
          TestVFormatMatrixNonSquare = False
End Function

' As TestNaNFallback, but for the matrix encoder, which has its own NaN/Inf fallback guard.
Function TestMatrixNaNFallback()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          Expected(1, 1) = 1#: Expected(1, 2) = 2#: Expected(1, 3) = 3#
          Expected(2, 1) = 4#
          Expected(2, 2) = CVErr(2042) '#N/A
          Expected(2, 3) = 6#
          y = ThrowIfError(JuliaExcel.JuliaEval("[1.0 2.0 3.0; 4.0 NaN 6.0]"))
          TestMatrixNaNFallback = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestMatrixNaNFallback", Err, True)
          TestMatrixNaNFallback = False
End Function

' Checks an empty Float64 vector correctly routes to the general array format rather than
' attempting a zero-length "V" payload.
Function TestVFormatEmptyArrayFallback()
          Dim y As Variant

          On Error GoTo ErrHandler
          y = ThrowIfError(JuliaExcel.JuliaEvalVBA("Float64[]"))
          TestVFormatEmptyArrayFallback = IsArray(y) And (UBound(y) < LBound(y))

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVFormatEmptyArrayFallback", Err, True)
          TestVFormatEmptyArrayFallback = False
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

          On Error GoTo ErrHandler
          ReDim x(1 To 5)
          x(1) = 1#: x(2) = -2.5: x(3) = 3.14159265358979: x(4) = 0#: x(5) = 1000000#
          y = JuliaExcel.JuliaCallVBA("identity", x)
          TestVEncodeFromVariantArray = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVEncodeFromVariantArray", Err, True)
          TestVEncodeFromVariantArray = False
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

          On Error GoTo ErrHandler
          For k = 1 To 4
              For j = 1 To 3
                  For i = 1 To 2
                      x(i, j, k) = i + 10 * j + 100 * k
                  Next i
              Next j
          Next k

          y = JuliaExcel.JuliaCallVBA("identity", x)
          TestVFormat3DArray = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVFormat3DArray", Err, True)
          TestVFormat3DArray = False
End Function

' Exercises the "R" (range) branch's integer sub-format - a UnitRange is encoded as just
' first/step/length and reconstructed via arithmetic, rather than transmitting every element.
Function TestRangeIntegerVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

          On Error GoTo ErrHandler
          y1 = ThrowIfError(JuliaExcel.JuliaEval("1:1000"))
          y2 = ThrowIfError(JuliaExcel.JuliaEval("collect(1:1000)"))
          TestRangeIntegerVsCollect = ArraysIdentical(y1, y2)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeIntegerVsCollect", Err, True)
          TestRangeIntegerVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for a StepRange with a non-1 step.
Function TestRangeStepVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

          On Error GoTo ErrHandler
          y1 = ThrowIfError(JuliaExcel.JuliaEval("5:3:47"))
          y2 = ThrowIfError(JuliaExcel.JuliaEval("collect(5:3:47)"))
          TestRangeStepVsCollect = ArraysIdentical(y1, y2)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeStepVsCollect", Err, True)
          TestRangeStepVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for the float sub-format.
Function TestRangeFloatVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

          On Error GoTo ErrHandler
          y1 = ThrowIfError(JuliaExcel.JuliaEval("(1:1000).*pi"))
          y2 = ThrowIfError(JuliaExcel.JuliaEval("collect((1:1000).*pi)"))
          TestRangeFloatVsCollect = ArraysIdentical(y1, y2)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeFloatVsCollect", Err, True)
          TestRangeFloatVsCollect = False
End Function

' Confirms JuliaEvalVBA gives a genuine 1-D array for a range result.
Function TestRangeJuliaEvalVBAIsOneD()
          Dim y As Variant

          On Error GoTo ErrHandler
          y = ThrowIfError(JuliaExcel.JuliaEvalVBA("1:1000"))
          TestRangeJuliaEvalVBAIsOneD = (NumDimensions(y) = 1) And (UBound(y) - LBound(y) + 1 = 1000) And _
              (y(LBound(y)) = 1) And (y(UBound(y)) = 1000)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeJuliaEvalVBAIsOneD", Err, True)
          TestRangeJuliaEvalVBAIsOneD = False
End Function

' As TestRangeJuliaEvalVBAIsOneD, but via JuliaEval - must give a 2-D, single-column array.
Function TestRangeJuliaEvalIsColumn()
          Dim y As Variant

          On Error GoTo ErrHandler
          y = ThrowIfError(JuliaExcel.JuliaEval("1:1000"))
          TestRangeJuliaEvalIsColumn = (NumDimensions(y) = 2) And (UBound(y, 1) = 1000) And (UBound(y, 2) = 1) And _
              (y(1, 1) = 1) And (y(1000, 1) = 1000)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeJuliaEvalIsColumn", Err, True)
          TestRangeJuliaEvalIsColumn = False
End Function

' Directly confirms the compact "R" wire format is actually used, for both sub-formats.
Function TestRangeWireFormat()
          Dim s As Variant

          On Error GoTo ErrHandler
          s = ThrowIfError(JuliaExcel.JuliaEvalVBA("JuliaExcel.encode_for_xl(1:1000)"))
          If Left$(s, 3) <> "RI," Then Throw "Expected integer range to use 'RI' wire format, got: " & Left$(s, 10)
          s = ThrowIfError(JuliaExcel.JuliaEvalVBA("JuliaExcel.encode_for_xl((1:1000).*pi)"))
          If Left$(s, 3) <> "RF," Then Throw "Expected float range to use 'RF' wire format, got: " & Left$(s, 10)
          TestRangeWireFormat = True

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestRangeWireFormat", Err, True)
          TestRangeWireFormat = False
End Function

' Confirms an array containing NaN/Infinity travels Excel -> Julia via the fast "V" path, and
' Julia's own outbound encoder still correctly maps NaN -> #N/A and Infinity -> #NUM! on the
' way back.
Function TestNaNInfRoundTripViaV()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x(1 To 4) As Double
          Dim y As Variant

          On Error GoTo ErrHandler
          x(1) = 1#
          x(2) = HexToDouble("7FF8000000000000") 'quiet NaN
          x(3) = 2#
          x(4) = HexToDouble("7FF0000000000000") 'positive infinity

          Expected(1, 1) = 1#
          Expected(2, 1) = CVErr(2042) '#N/A
          Expected(3, 1) = 2#
          Expected(4, 1) = CVErr(2036) '#NUM!

          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestNaNInfRoundTripViaV = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestNaNInfRoundTripViaV", Err, True)
          TestNaNInfRoundTripViaV = False
End Function

' Confirms Excel errors round-trip correctly through Julia via JuliaCall("identity", ...),
' for all 14 Excel error codes, not just the two (2036, 2042) Julia can generate from Inf/NaN.
' Built with ReDim (1-based), not VBA.Array() (0-based).
Function TestExcelErrorRoundTrip()
          Dim Codes(1 To 14) As Long
          Dim Expected(1 To 14, 1 To 1) As Variant
          Dim i As Long
          Dim y As Variant

          On Error GoTo ErrHandler
          Codes(1) = 2000: Codes(2) = 2007: Codes(3) = 2015: Codes(4) = 2023: Codes(5) = 2029
          Codes(6) = 2036: Codes(7) = 2042: Codes(8) = 2043: Codes(9) = 2045: Codes(10) = 2046
          Codes(11) = 2047: Codes(12) = 2048: Codes(13) = 2049: Codes(14) = 2050

          For i = 1 To 14
              Expected(i, 1) = CVErr(Codes(i))
          Next i

          y = ThrowIfError(JuliaExcel.JuliaCall("identity", Expected))
          TestExcelErrorRoundTrip = ArraysIdentical(Expected, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestExcelErrorRoundTrip", Err, True)
          TestExcelErrorRoundTrip = False
End Function

Function TestArrayOfSingleCharacters()
          Dim x As Variant
          Dim y As Variant

          On Error GoTo ErrHandler
          'All characters with code point <= 100000
          x = Application.Evaluate("=UNICHAR(VSTACK(SEQUENCE(55295,,1),SEQUENCE(7632,,57344),SEQUENCE(526,,65008),SEQUENCE(34465,,65536)))")
          y = ThrowIfError(JuliaExcel.JuliaCall("identity", x))
          TestArrayOfSingleCharacters = ArraysIdentical(x, y)

          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestArrayOfSingleCharacters", Err, True)
          TestArrayOfSingleCharacters = False
End Function

' Checks JuliaCallNamedArgs passes the arguments after ";" as keyword arguments, including
' with ";" also used as a positional argument and as the value of the last named argument,
' and that a missing ";", an odd number of arguments after it, and a repeated name are each
' reported as an error.
Function TestNamedArgs()
          Dim Expected1(1 To 4, 1 To 1) As Variant
          Dim Expected3(1 To 2, 1 To 1) As Variant
          Dim OK As Boolean

          On Error GoTo ErrHandler
          ThrowIfError JuliaExcel.JuliaEval("named_args_test(x, y=0; a=1, b=2) = [x, y, a, b]; named_args_join(xs...; delim="","") = join(xs, delim)")

          Expected1(1, 1) = 10#: Expected1(2, 1) = 0#: Expected1(3, 1) = 1#: Expected1(4, 1) = 5#
          OK = ArraysIdentical(Expected1, ThrowIfError(JuliaExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "b", 5#)))

          OK = OK And Left$(JuliaExcel.JuliaCallNamedArgs("named_args_test", 10#), 15) = "#ArgumentError:"
          OK = OK And Left$(JuliaExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "b", 5#, "a"), 15) = "#ArgumentError:"

          Expected3(1, 1) = "a": Expected3(2, 1) = "b"
          OK = OK And ArraysIdentical(Expected3, ThrowIfError(JuliaExcel.JuliaCallNamedArgs("split", "a;;b", ";", ";", "keepempty", False)))
          OK = OK And JuliaExcel.JuliaCallNamedArgs("named_args_join", "a", ";", "b", ";", "delim", ";") = "a;;;b"

          OK = OK And Left$(JuliaExcel.JuliaCallNamedArgs("named_args_test", 10#, ";", "a", 1, "a", 2), 15) = "#ArgumentError:"

          TestNamedArgs = OK
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestNamedArgs", Err, True)
          TestNamedArgs = False
End Function

' Checks JuliaCallNamedArgsVBA returns a vector as a 1-dimensional array, and that both
' argument and return types that can't live on a worksheet (here Dictionaries) work as named
' argument values.
Function TestNamedArgsVBA()
          Dim Expected(1 To 4) As Variant
          Dim OK As Boolean
          Dim x As New Scripting.Dictionary
          Dim y As Object

          On Error GoTo ErrHandler
          ThrowIfError JuliaExcel.JuliaEval("named_args_test(x, y=0; a=1, b=2) = [x, y, a, b]")
          Expected(1) = 10#: Expected(2) = 0#: Expected(3) = 1#: Expected(4) = 5#
          OK = ArraysIdentical(Expected, JuliaExcel.JuliaCallNamedArgsVBA("named_args_test", 10#, ";", "b", 5#))

          ThrowIfError JuliaExcel.JuliaEval("named_args_dict(; kw...) = Dict(string(k) => v for (k, v) in kw)")
          x.Add "alpha", 100
          Set y = JuliaExcel.JuliaCallNamedArgsVBA("named_args_dict", ";", "d", x, "n", 1#)
          OK = OK And y.Count = 2
          OK = OK And y("n") = 1#
          OK = OK And y("d")("alpha") = 100

          TestNamedArgsVBA = OK
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestNamedArgsVBA", Err, True)
          TestNamedArgsVBA = False
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

          On Error GoTo ErrHandler
          Sizes = Array(1, 2, 3, 7, 8, 9, 17, 100, 1000, 1000000)

          OK = True
          For Idx = LBound(Sizes) To UBound(Sizes)
              Sz = Sizes(Idx)
              ReDim x(1 To Sz)
              For i = 1 To Sz
                  x(i) = (i - Sz / 2) * 1.5 + 0.25 'varied: negative, positive, fractional values
              Next i
              y = JuliaExcel.JuliaCallVBA("identity", x)
              If Not ArraysIdentical(x, y) Then
                  Throw "Round trip mismatch for size " & Sz
              End If
          Next Idx

          TestVFormatBoundarySizes = OK
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestVFormatBoundarySizes", Err, True)
          TestVFormatBoundarySizes = False
End Function

' Regression test: ConcatenateExpressions used to join a multi-line JuliaExpression with ";"
' rather than a newline, so a "#" comment on one line could swallow the rest of the
' expression - including the "end" that closes a function - turning valid Julia code into a
' syntax error.
Function TestJuliaEvalMultiLineWithComment() As Boolean
          Dim Lines(1 To 3) As Variant
          Dim OK As Boolean
          On Error GoTo ErrHandler
          Lines(1) = "function RegressionTestForComment(a, b)"
          Lines(2) = "a + b #this comment messes things up!"
          Lines(3) = "end"
          ThrowIfError JuliaExcel.JuliaEvalVBA(Lines)
          OK = JuliaExcel.JuliaCall("RegressionTestForComment", 2, 3) = 5
          TestJuliaEvalMultiLineWithComment = OK
          Exit Function
ErrHandler:
          PrintTwice ReThrow("TestJuliaEvalMultiLineWithComment", Err, True)
          TestJuliaEvalMultiLineWithComment = False
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
          On Error Resume Next
          Set ScratchSheet = ThisWorkbook.Worksheets("Scratch")
          On Error GoTo 0
          If ScratchSheet Is Nothing Then
              Set ScratchSheet = ThisWorkbook.Worksheets.Add
              ScratchSheet.Name = "Scratch"
          End If
          ScratchSheet.Cells.Clear
End Function

Function WsTest_Scalar() As Boolean
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Value2 = Application.WorksheetFunction.Pi()
          ws.Range("B1").Formula2 = "=JuliaCall(""identity"",A1)"
          Application.CalculateFullRebuild
          WsTest_Scalar = ws.Range("B1").Value2 = Application.WorksheetFunction.Pi()
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_Scalar", Err, True)
          WsTest_Scalar = False
End Function

Function WsTest_String() As Boolean
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Value2 = "FooBar"
          ws.Range("B1").Formula2 = "=JuliaCall(""identity"",A1)"
          Application.CalculateFullRebuild
          WsTest_String = ws.Range("B1").Value2 = "FooBar"
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_String", Err, True)
          WsTest_String = False
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
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1:A3").Value2 = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
          ws.Range("E1").Formula2 = "=JuliaCall(""identity"",A1:A3)"
          Application.CalculateFullRebuild
          Expected = ws.Range("E1:E3").Value2
          WsTest_OneDArrayBecomesColumn = NumDimensions(Expected) = 2 And _
              Expected(1, 1) = 1 And Expected(2, 1) = 2 And Expected(3, 1) = 3
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_OneDArrayBecomesColumn", Err, True)
          WsTest_OneDArrayBecomesColumn = False
End Function

Function WsTest_TwoDArray() As Boolean
          Dim Observed As Variant
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1:B2").Value2 = Array(1, 2) 'row 1 only - fill row 2 explicitly below
          ws.Range("A1").Value2 = 1: ws.Range("B1").Value2 = 2
          ws.Range("A2").Value2 = 3: ws.Range("B2").Value2 = 4
          ws.Range("D1").Formula2 = "=JuliaCall(""identity"",A1:B2)"
          Application.CalculateFullRebuild
          Observed = ws.Range("D1:E2").Value2
          WsTest_TwoDArray = Observed(1, 1) = 1 And Observed(1, 2) = 2 And _
              Observed(2, 1) = 3 And Observed(2, 2) = 4
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_TwoDArray", Err, True)
          WsTest_TwoDArray = False
End Function

Function WsTest_JuliaEvalScalar() As Boolean
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Formula2 = "=JuliaEval(""1+1"")"
          Application.CalculateFullRebuild
          WsTest_JuliaEvalScalar = ws.Range("A1").Value2 = 2
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_JuliaEvalScalar", Err, True)
          WsTest_JuliaEvalScalar = False
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
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Formula2 = "=JuliaEval(""nothing"")"
          ws.Range("A2").Formula2 = "=JuliaEval(""missing"")"
          ws.Range("A3").Formula2 = "=ISBLANK(A1)"
          Application.CalculateFullRebuild
          WsTest_EmptyResult = ws.Range("A1").Value2 = 0 And ws.Range("A2").Value2 = 0 And ws.Range("A3").Value2 = False
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_EmptyResult", Err, True)
          WsTest_EmptyResult = False
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
          Dim ws As Worksheet
          Dim Result As String
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Formula2 = "=JuliaLaunch(,,""-q"")"
          Application.CalculateFullRebuild
          Result = ws.Range("A1").Text
          WsTest_JuliaLaunchArgOrder = Left$(Result, 1) <> "#"
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_JuliaLaunchArgOrder", Err, True)
          WsTest_JuliaLaunchArgOrder = False
End Function

' Cross-path consistency check: JuliaIsRunning/JuliaExcelPID as real worksheet formulas
' should agree with the same functions called via COM - both run in-process in this same
' Excel.exe, so the process id must be identical either way, and Julia was already launched
' (via COM) earlier in RunTests.
Function WsTest_JuliaIsRunningAndPID() As Boolean
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Formula2 = "=JuliaIsRunning()"
          ws.Range("A2").Formula2 = "=JuliaExcelPID()"
          Application.CalculateFullRebuild
          WsTest_JuliaIsRunningAndPID = (ws.Range("A1").Value2 = True) And _
              (ws.Range("A2").Value2 = JuliaExcel.JuliaExcelPID())
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_JuliaIsRunningAndPID", Err, True)
          WsTest_JuliaIsRunningAndPID = False
End Function

Function WsTest_ErrorString() As Boolean
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1").Formula2 = "=JuliaCall(""thisFunctionDoesNotExist12345"")"
          Application.CalculateFullRebuild
          WsTest_ErrorString = Left$(ws.Range("A1").Value2, 1) = "#"
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_ErrorString", Err, True)
          WsTest_ErrorString = False
End Function

' The scenario from dotnet-addin\README.md's performance write-up: JuliaCall("sort", ...)
' against a real worksheet range, kept here as a small (not 1,000,000-element) permanent
' regression test, compared against Excel's own SORT().
Function WsTest_SortRoundTrip() As Boolean
          Dim Expected As Variant
          Dim Observed As Variant
          Dim ws As Worksheet
          On Error GoTo ErrHandler
          Set ws = ScratchSheet()
          ws.Range("A1:A5").Value2 = Application.WorksheetFunction.Transpose(Array(5, 3, 1, 4, 2))
          ws.Range("C1").Formula2 = "=SORT(A1:A5)"
          ws.Range("E1").Formula2 = "=JuliaCall(""sort"",A1:A5)"
          Application.CalculateFullRebuild
          Expected = ws.Range("C1:C5").Value2
          Observed = ws.Range("E1:E5").Value2
          WsTest_SortRoundTrip = ArraysIdentical(Expected, Observed)
          Exit Function
ErrHandler:
          PrintTwice ReThrow("WsTest_SortRoundTrip", Err, True)
          WsTest_SortRoundTrip = False
End Function
