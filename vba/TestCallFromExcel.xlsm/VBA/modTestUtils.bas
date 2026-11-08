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
          ThrowIfError = Data
          If VarType(Data) = vbString Then
              If Left$(Data, 1) = "#" Then
                  Throw CStr(Data)
              End If
          End If
End Function

Function ReThrow(FunctionName As String, Error As ErrObject, Optional ReturnString As Boolean = False)
          Dim ErrorDescription As String
          Dim ErrorNumber As Long
          Dim LineDescription As String

          ErrorDescription = Error.Description
          ErrorNumber = Err.Number

          If Erl = 0 Then
              LineDescription = " (line unknown): "
          Else
              LineDescription = " (line " & CStr(Erl) & "): "
          End If
          ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"

          If ReturnString Then
              ReThrow = ErrorDescription
          Else
              Err.Raise ErrorNumber, , ErrorDescription
          End If
End Function

Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim Lbnd As Long
          On Error GoTo ErrHandler
          Do
              i = i + 1
              Lbnd = LBound(x, i)
          Loop
          Exit Function
ErrHandler:
          NumDimensions = i - 1
End Function

Sub Throw(ByVal ErrorString As String)
          'Out of stack space errors can lead to enormous error strings, but Excel cannot
          'handle strings longer than 32767, so just take the right part.
          If Len(ErrorString) > 32000 Then
              Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
          Else
              Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
          End If
End Sub

Sub Assign(ByRef A, B)
          If IsObject(B) Then
              Set A = B
          Else
              Let A = B
          End If
End Sub

Function GetAt(ByRef Arr As Variant, ByRef Idx() As Long) As Variant
          On Error GoTo ErrHandler
          Select Case UBound(Idx)
              Case 1: GetAt = Arr(Idx(1))
              Case 2: GetAt = Arr(Idx(1), Idx(2))
              Case 3: GetAt = Arr(Idx(1), Idx(2), Idx(3))
              Case 4: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4))
              Case 5: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5))
              Case 6: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6))
              Case 7: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7))
              Case 8: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8))
              Case 9: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8), Idx(9))
              Case Else
                  Throw "Rank > 9 not supported"
          End Select
          Exit Function
ErrHandler:
          ReThrow "GetAt", Err
End Function

Function SetAt(ByRef Arr As Variant, ByRef Idx() As Long, Value As Variant) As Variant
          On Error GoTo ErrHandler
          Select Case UBound(Idx)
              Case 1: Arr(Idx(1)) = Value
              Case 2: Arr(Idx(1), Idx(2)) = Value
              Case 3: Arr(Idx(1), Idx(2), Idx(3)) = Value
              Case 4: Arr(Idx(1), Idx(2), Idx(3), Idx(4)) = Value
              Case 5: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5)) = Value
              Case 6: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6)) = Value
              Case 7: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7)) = Value
              Case 8: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8)) = Value
              Case 9: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8), Idx(9)) = Value
              Case Else
                  Throw "Rank > 9 not supported"
          End Select
          Exit Function
ErrHandler:
          ReThrow "SetAt", Err
End Function

Sub SetAtLinear(Arr As Variant, Idx As Long, Value As Variant)
          Dim Dims As Long
          Dim i As Long
          Dim Indices() As Long
          Dim Lb() As Long
          Dim Offset As Long
          Dim Sizes() As Long
          Dim t As Long

          On Error GoTo ErrHandler
          Dims = NumDimensions(Arr)
          ReDim Sizes(1 To Dims)
          ReDim Lb(1 To Dims)
          ReDim Indices(1 To Dims)

          For i = 1 To Dims
              Lb(i) = LBound(Arr, i)
              Sizes(i) = UBound(Arr, i) - Lb(i) + 1
          Next i

          t = Idx - 1
          For i = 1 To Dims
              Offset = t Mod Sizes(i)
              Indices(i) = Lb(i) + Offset
              t = t \ Sizes(i)
          Next i

          SetAt Arr, Indices, Value

          Exit Sub
ErrHandler:
          ReThrow "SetAtLinear", Err
End Sub

' Module: ArrayCompare, written by Copilot 22 Dec 2025, with amendments by Philip Swannell
' (vba\JuliaExcel.xlam\VBA\modTestDependents.bas). Returns True iff both arrays are arrays
' (both initialised or both uninitialised in the same way), have the same number of
' dimensions, the same bounds per dimension, and every corresponding element compares equal
' (scalars only).
Function ArraysIdentical(ByVal A As Variant, ByVal B As Variant) As Boolean
          If Not IsArray(A) Or Not IsArray(B) Then
              ArraysIdentical = False
              Exit Function
          End If

          Dim InitA As Boolean
          Dim InitB As Boolean
          InitA = IsArrayInitialized(A)
          InitB = IsArrayInitialized(B)

          If InitA Xor InitB Then
              ArraysIdentical = False
              Exit Function
          End If

          If Not InitA And Not InitB Then
              ArraysIdentical = True
              Exit Function
          End If

          Dim nA As Long
          Dim nB As Long
          nA = NumDimensions(A)
          nB = NumDimensions(B)
          If nA <> nB Then
              ArraysIdentical = False
              Exit Function
          End If

          Dim d As Long
          For d = 1 To nA
              If LBound(A, d) <> LBound(B, d) Then
                  ArraysIdentical = False
                  Exit Function
              End If
              If UBound(A, d) <> UBound(B, d) Then
                  ArraysIdentical = False
                  Exit Function
              End If
          Next d

          If nA = 0 Then
              ArraysIdentical = True
              Exit Function
          End If

          ArraysIdentical = WalkAndCompare(A, B, nA)
End Function

Function ScalarsEqual(ByVal x As Variant, ByVal y As Variant) As Boolean
          If IsEmpty(x) Then
              ScalarsEqual = IsEmpty(y)
              Exit Function
          ElseIf IsEmpty(y) Then
              ScalarsEqual = False
              Exit Function
          End If

          If IsNull(x) Then
              ScalarsEqual = IsNull(y)
              Exit Function
          ElseIf IsNull(y) Then
              ScalarsEqual = False
              Exit Function
          End If

          If VarType(x) = vbString And VarType(y) = vbString Then
              ScalarsEqual = (StrComp(CStr(x), CStr(y), vbBinaryCompare) = 0)
              Exit Function
          End If

          On Error GoTo NotComparable
          ScalarsEqual = (x = y) And (VarType(x) = VarType(y))
          Exit Function

NotComparable:
          ScalarsEqual = False
End Function

Function WalkAndCompare(ByRef A As Variant, ByRef B As Variant, ByVal n As Long) As Boolean
          Dim d As Long
          Dim Idx() As Long
          Dim Lb() As Long
          Dim Ub() As Long

          ReDim Idx(1 To n)
          ReDim Lb(1 To n)
          ReDim Ub(1 To n)

          For d = 1 To n
              Lb(d) = LBound(A, d)
              Ub(d) = UBound(A, d)
              Idx(d) = Lb(d)
          Next d

          Do
              If Not ScalarsEqual(GetAt(A, Idx), GetAt(B, Idx)) Then
                  WalkAndCompare = False
                  Exit Function
              End If

              d = n
              Do While d >= 1
                  Idx(d) = Idx(d) + 1
                  If Idx(d) <= Ub(d) Then Exit Do
                  Idx(d) = Lb(d)
                  d = d - 1
              Loop
              If d = 0 Then Exit Do
          Loop

          WalkAndCompare = True
End Function

Function IsArrayInitialized(ByVal A As Variant) As Boolean
          On Error GoTo NotInit
          If Not IsArray(A) Then Exit Function
          Dim Lb As Long
          Lb = LBound(A, 1)
          IsArrayInitialized = True
          Exit Function
NotInit:
          IsArrayInitialized = False
End Function

' Ported from vba\JuliaExcel.xlam\VBA\modUnserialise.bas - constructs a Double from its raw
' IEEE-754 bit pattern (16 hex characters), used by TestNaNInfRoundTripViaV to build NaN/
' Infinity values VBA's own arithmetic cannot produce (it raises a runtime error instead).
Function HexToDouble(ByVal Hex As String) As Double
          Dim Hi As Long
          Dim Lo As Long
          Dim TD As TDouble
          Dim Tl As TLongs

          On Error GoTo ErrHandler
          If Len(Hex) <> 16 Then Throw "Hex must be 16 hex characters, but got " & Len(Hex)
          Hi = CLng("&H" & Left$(Hex, 8))
          Lo = CLng("&H" & Right$(Hex, 8))
          Tl.Hi = Hi
          Tl.Lo = Lo
          LSet TD = Tl
          HexToDouble = TD.d

          Exit Function
ErrHandler:
          ReThrow "HexToDouble", Err
End Function

Function RepoRoot() As String
          RepoRoot = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(ThisWorkbook.Path & "\..")
End Function

' "C:\foo\bar" -> "/mnt/c/foo/bar"
Function WindowsPathToWSL(WindowsAddress As String) As String
          On Error GoTo ErrHandler
          If Mid(WindowsAddress, 2, 1) <> ":" Then Throw "WindowsAddress must start with characters ""x:\"" for some drive-letter x"
          WindowsPathToWSL = "/mnt/" & LCase(Left(WindowsAddress, 1)) & Replace(Mid(WindowsAddress, 3), "\", "/")
          Exit Function
ErrHandler:
          ReThrow "WindowsPathToWSL", Err
End Function
