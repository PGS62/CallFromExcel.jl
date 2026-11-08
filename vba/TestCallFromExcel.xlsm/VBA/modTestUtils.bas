Attribute VB_Name = "modTestUtils"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/CallFromExcel.jl#readme
'
' Generic, JuliaExcel-independent helpers used by modTest.bas, ported from
' vba\JuliaExcel.xlam\VBA (modTestDependents.bas, modUtils.bas, modSetGet.bas,
' modPrivateMethods.bas, modUnserialise.bas). These are plain VBA array/error-handling/
' bit-pattern utilities with no dependency on CallFromExcel itself, so they port to the new
' (no-xlam) world unchanged, except ThrowIfError (see note below).

Option Explicit
Option Private Module

Private Type TDouble
    d As Double
End Type

Private Type TLongs
    Lo As Long
    Hi As Long
End Type

Private Type TBytes8
    B(0 To 7) As Byte
End Type

' Differs from the original (vba\JuliaExcel.xlam\VBA\modPrivateMethods.bas): the old
' convention was that an internal error returns as a string starting with "#" and ending
' with "!" (see ReThrow below) - the trailing "!" let ThrowIfError tell an internal error
' apart from a genuine Excel error value that happens to start with "#" (such as "#N/A",
' which does not end "!"). The new C# add-in's error returns ("#" & exception message) do
' not follow that convention, so this checks only the "#" prefix. This is safe here because
' every function this is used against (JuliaCall/JuliaEval/.../JuliaCallVBA/JuliaEvalVBA/...
' via ComServer) never legitimately returns a plain string starting with "#" as a successful
' result - genuine Excel error values (CVErr) are a different VBA type (vbError), not a
' string, so they are never affected by this check either way.
Function ThrowIfError(Data As Variant)
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 Throw CStr(Data)
5             End If
6         End If
End Function

Function ReThrow(FunctionName As String, Error As ErrObject, Optional ReturnString As Boolean = False)
          Dim ErrorDescription As String
          Dim ErrorNumber As Long
          Dim LineDescription As String

1         ErrorDescription = Error.Description
2         ErrorNumber = Err.Number

3         If Erl = 0 Then
4             LineDescription = " (line unknown): "
5         Else
6             LineDescription = " (line " & CStr(Erl) & "): "
7         End If
8         ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"

9         If ReturnString Then
10            ReThrow = ErrorDescription
11        Else
12            Err.Raise ErrorNumber, , ErrorDescription
13        End If
End Function

Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim Lbnd As Long
1         On Error GoTo ErrHandler
2         Do
3             i = i + 1
4             Lbnd = LBound(x, i)
5         Loop
6         Exit Function
ErrHandler:
7         NumDimensions = i - 1
End Function

Sub Throw(ByVal ErrorString As String)
          'Out of stack space errors can lead to enormous error strings, but Excel cannot
          'handle strings longer than 32767, so just take the right part.
1         If Len(ErrorString) > 32000 Then
2             Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
3         Else
4             Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
5         End If
End Sub

Sub Assign(ByRef A, B)
1         If IsObject(B) Then
2             Set A = B
3         Else
4             Let A = B
5         End If
End Sub

Function GetAt(ByRef Arr As Variant, ByRef Idx() As Long) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(Idx)
              Case 1: GetAt = Arr(Idx(1))
3             Case 2: GetAt = Arr(Idx(1), Idx(2))
4             Case 3: GetAt = Arr(Idx(1), Idx(2), Idx(3))
5             Case 4: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4))
6             Case 5: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5))
7             Case 6: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6))
8             Case 7: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7))
9             Case 8: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8))
10            Case 9: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8), Idx(9))
11            Case Else
12                Throw "Rank > 9 not supported"
13        End Select
14        Exit Function
ErrHandler:
15        ReThrow "GetAt", Err
End Function

Function SetAt(ByRef Arr As Variant, ByRef Idx() As Long, Value As Variant) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(Idx)
              Case 1: Arr(Idx(1)) = Value
3             Case 2: Arr(Idx(1), Idx(2)) = Value
4             Case 3: Arr(Idx(1), Idx(2), Idx(3)) = Value
5             Case 4: Arr(Idx(1), Idx(2), Idx(3), Idx(4)) = Value
6             Case 5: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5)) = Value
7             Case 6: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6)) = Value
8             Case 7: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7)) = Value
9             Case 8: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8)) = Value
10            Case 9: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8), Idx(9)) = Value
11            Case Else
12                Throw "Rank > 9 not supported"
13        End Select
14        Exit Function
ErrHandler:
15        ReThrow "SetAt", Err
End Function

Sub SetAtLinear(Arr As Variant, Idx As Long, Value As Variant)
          Dim Dims As Long
          Dim i As Long
          Dim Indices() As Long
          Dim Lb() As Long
          Dim Offset As Long
          Dim Sizes() As Long
          Dim t As Long

1         On Error GoTo ErrHandler
2         Dims = NumDimensions(Arr)
3         ReDim Sizes(1 To Dims)
4         ReDim Lb(1 To Dims)
5         ReDim Indices(1 To Dims)

6         For i = 1 To Dims
7             Lb(i) = LBound(Arr, i)
8             Sizes(i) = UBound(Arr, i) - Lb(i) + 1
9         Next i

10        t = Idx - 1
11        For i = 1 To Dims
12            Offset = t Mod Sizes(i)
13            Indices(i) = Lb(i) + Offset
14            t = t \ Sizes(i)
15        Next i

16        SetAt Arr, Indices, Value

17        Exit Sub
ErrHandler:
18        ReThrow "SetAtLinear", Err
End Sub

' Module: ArrayCompare, written by Copilot 22 Dec 2025, with amendments by Philip Swannell
' (vba\JuliaExcel.xlam\VBA\modTestDependents.bas). Returns True iff both arrays are arrays
' (both initialised or both uninitialised in the same way), have the same number of
' dimensions, the same bounds per dimension, and every corresponding element compares equal
' (scalars only).
Function ArraysIdentical(ByVal A As Variant, ByVal B As Variant) As Boolean
1         If Not IsArray(A) Or Not IsArray(B) Then
2             ArraysIdentical = False
3             Exit Function
4         End If

          Dim InitA As Boolean
          Dim InitB As Boolean
5         InitA = IsArrayInitialized(A)
6         InitB = IsArrayInitialized(B)

7         If InitA Xor InitB Then
8             ArraysIdentical = False
9             Exit Function
10        End If

11        If Not InitA And Not InitB Then
12            ArraysIdentical = True
13            Exit Function
14        End If

          Dim nA As Long
          Dim nB As Long
15        nA = NumDimensions(A)
16        nB = NumDimensions(B)
17        If nA <> nB Then
18            ArraysIdentical = False
19            Exit Function
20        End If

          Dim d As Long
21        For d = 1 To nA
22            If LBound(A, d) <> LBound(B, d) Then
23                ArraysIdentical = False
24                Exit Function
25            End If
26            If UBound(A, d) <> UBound(B, d) Then
27                ArraysIdentical = False
28                Exit Function
29            End If
30        Next d

31        If nA = 0 Then
32            ArraysIdentical = True
33            Exit Function
34        End If

35        ArraysIdentical = WalkAndCompare(A, B, nA)
End Function

Function ScalarsEqual(ByVal x As Variant, ByVal y As Variant) As Boolean
1         If IsEmpty(x) Then
2             ScalarsEqual = IsEmpty(y)
3             Exit Function
4         ElseIf IsEmpty(y) Then
5             ScalarsEqual = False
6             Exit Function
7         End If

8         If IsNull(x) Then
9             ScalarsEqual = IsNull(y)
10            Exit Function
11        ElseIf IsNull(y) Then
12            ScalarsEqual = False
13            Exit Function
14        End If

15        If VarType(x) = vbString And VarType(y) = vbString Then
16            ScalarsEqual = (StrComp(CStr(x), CStr(y), vbBinaryCompare) = 0)
17            Exit Function
18        End If

19        On Error GoTo NotComparable
20        ScalarsEqual = (x = y) And (VarType(x) = VarType(y))
21        Exit Function

NotComparable:
22        ScalarsEqual = False
End Function

Function WalkAndCompare(ByRef A As Variant, ByRef B As Variant, ByVal n As Long) As Boolean
          Dim d As Long
          Dim Idx() As Long
          Dim Lb() As Long
          Dim Ub() As Long

1         ReDim Idx(1 To n)
2         ReDim Lb(1 To n)
3         ReDim Ub(1 To n)

4         For d = 1 To n
5             Lb(d) = LBound(A, d)
6             Ub(d) = UBound(A, d)
7             Idx(d) = Lb(d)
8         Next d

9         Do
10            If Not ScalarsEqual(GetAt(A, Idx), GetAt(B, Idx)) Then
11                WalkAndCompare = False
12                Exit Function
13            End If

14            d = n
15            Do While d >= 1
16                Idx(d) = Idx(d) + 1
17                If Idx(d) <= Ub(d) Then Exit Do
18                Idx(d) = Lb(d)
19                d = d - 1
20            Loop
21            If d = 0 Then Exit Do
22        Loop

23        WalkAndCompare = True
End Function

Function IsArrayInitialized(ByVal A As Variant) As Boolean
1         On Error GoTo NotInit
2         If Not IsArray(A) Then Exit Function
          Dim Lb As Long
3         Lb = LBound(A, 1)
4         IsArrayInitialized = True
5         Exit Function
NotInit:
6         IsArrayInitialized = False
End Function

' Ported from vba\JuliaExcel.xlam\VBA\modUnserialise.bas - constructs a Double from its raw
' IEEE-754 bit pattern (16 hex characters), used by TestNaNInfRoundTripViaV to build NaN/
' Infinity values VBA's own arithmetic cannot produce (it raises a runtime error instead).
Function HexToDouble(ByVal Hex As String) As Double
          Dim Hi As Long
          Dim Lo As Long
          Dim TD As TDouble
          Dim Tl As TLongs

1         On Error GoTo ErrHandler
2         If Len(Hex) <> 16 Then Throw "Hex must be 16 hex characters, but got " & Len(Hex)
3         Hi = CLng("&H" & Left$(Hex, 8))
4         Lo = CLng("&H" & Right$(Hex, 8))
5         Tl.Hi = Hi
6         Tl.Lo = Lo
7         LSet TD = Tl
8         HexToDouble = TD.d

9         Exit Function
ErrHandler:
10        ReThrow "HexToDouble", Err
End Function

Function RepoRoot() As String
1         RepoRoot = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(ThisWorkbook.Path & "\..")
End Function

' "C:\foo\bar" -> "/mnt/c/foo/bar"
Function WindowsPathToWSL(WindowsAddress As String) As String
1         On Error GoTo ErrHandler
2         If Mid(WindowsAddress, 2, 1) <> ":" Then Throw "WindowsAddress must start with characters ""x:\"" for some drive-letter x"
3         WindowsPathToWSL = "/mnt/" & LCase(Left(WindowsAddress, 1)) & Replace(Mid(WindowsAddress, 3), "\", "/")
4         Exit Function
ErrHandler:
5         ReThrow "WindowsPathToWSL", Err
End Function
