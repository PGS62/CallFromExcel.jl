Attribute VB_Name = "SpeedTestCSharp"
Option Explicit

' Toolchain-validation prototype only: not wired to Julia.
' Compares the C# port of the "V" wire-format encode/decode (ComServer.EncodeV/DecodeV,
' csharp-prototype\src\VFormat.cs) against the real, unmodified production VBA
' implementation (TrySerialiseArrayAsV / Unserialise's 'V' branch, in modSerialise.bas /
' modUnserialise.bas of the actual JuliaExcel.xlam add-in) on identical data, to decide
' whether porting the wire format to C#/Excel-DNA is worth the effort.
'
' Calls into JuliaExcel.xlam via Application.Run with an explicit workbook-qualified name
' (the same pattern scripts\Run-VbaTests.ps1 uses for modTest.RunTests), which works even
' though those modules use "Option Private Module" -- no changes to JuliaExcel.xlam itself
' are needed or made.
'
' Requires: the real, installed JuliaExcel.xlam add-in loaded in this Excel session (as
' normal), and this module imported into any workbook (does not need to be the xlam
' itself). GetXlamWorkbook pins to the installed add-in's exact path rather than trusting
' a bare "JuliaExcel.xlam" name match, and fails loudly instead of silently running
' against the wrong copy if, e.g., a second JuliaExcel.xlam (such as this repo's
' workbooks\JuliaExcel.xlam) has been opened in the same session -- Excel's Workbooks
' collection is keyed by name, so a bare-name lookup can silently resolve to whichever
' copy happened to load first, see this file's own git history for how that nearly
' produced a benchmark against the wrong VBA code.

Const XlamFullPath As String = "C:\Users\Public\JuliaExcel\JuliaExcel.xlam"

' Convenience wrapper for interactive use: Functions don't appear in the Alt+F8 macro
' list, only Subs do. Run this one manually; CSharpVFormatSpeedTest itself returns a
' String so it can also be driven programmatically (e.g. via Application.Run from
' outside VBA) without a blocking MsgBox.
Sub RunCSharpVFormatSpeedTest()
    MsgBox CSharpVFormatSpeedTest()
End Sub

Function CSharpVFormatSpeedTest() As String

    Const VectorLength As Long = 1000000
    Const NumCalls As Long = 10

    Dim InputData As Variant
    Dim comServer As Object
    Dim i As Long
    Dim t1 As Double, t2 As Double
    Dim TotalVBAEncode As Double, TotalCSEncode As Double
    Dim TotalVBADecode As Double, TotalCSDecode As Double
    Dim VBAEncoded As String, CSEncoded As String
    Dim VBADecoded As Variant, CSDecoded As Variant
    Dim Report As String

    On Error GoTo ErrHandler

    Set comServer = CreateObject("JuliaExcelPrototype.ComServer")
    InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

    ' --- Correctness: wire-format compatibility, both directions ---
    VBAEncoded = RunXlam("modSerialise.SerialiseElement", InputData)
    If Left$(VBAEncoded, 1) <> "V" Then Err.Raise vbObjectError + 1, , _
        "Expected VBA's SerialiseElement to take the 'V' fast path for an all-Double RANDARRAY, got '" & Left$(VBAEncoded, 1) & "'"

    CSEncoded = comServer.EncodeV(InputData)
    If VBAEncoded <> CSEncoded Then Err.Raise vbObjectError + 1, , _
        "VBA and C# 'V'-format encodings differ for identical input"

    VBADecoded = RunXlam("modUnserialise.UnserialiseFromString", CSEncoded, False, 0, True)   ' VBA decoding C#'s output
    CSDecoded = comServer.DecodeV(VBAEncoded)                                                 ' C# decoding VBA's output
    For i = 1 To VectorLength
        If VBADecoded(i, 1) <> InputData(i, 1) Then Err.Raise vbObjectError + 1, , "VBA decode of C#-encoded string mismatch at element " & i
        If CSDecoded(i - 1, 0) <> InputData(i, 1) Then Err.Raise vbObjectError + 1, , "C# decode of VBA-encoded string mismatch at element " & i
    Next i

    ' --- Timing: interleaved, so warmup/drift affects both sides equally ---
    For i = 1 To NumCalls
        t1 = Timer
        VBAEncoded = RunXlam("modSerialise.SerialiseElement", InputData)
        t2 = Timer
        TotalVBAEncode = TotalVBAEncode + (t2 - t1)

        t1 = Timer
        CSEncoded = comServer.EncodeV(InputData)
        t2 = Timer
        TotalCSEncode = TotalCSEncode + (t2 - t1)

        t1 = Timer
        VBADecoded = RunXlam("modUnserialise.UnserialiseFromString", VBAEncoded, False, 0, True)
        t2 = Timer
        TotalVBADecode = TotalVBADecode + (t2 - t1)

        t1 = Timer
        CSDecoded = comServer.DecodeV(VBAEncoded)
        t2 = Timer
        TotalCSDecode = TotalCSDecode + (t2 - t1)
    Next i

    Report = "Correctness: VBA and C# 'V'-format encodings are byte-identical; cross-decodes match." & vbCrLf & vbCrLf & _
        "Vector length = " & Format(VectorLength, "###,###") & ", averaged over " & NumCalls & " calls (interleaved)" & vbCrLf & _
        "Encode: VBA = " & Format(TotalVBAEncode / NumCalls * 1000, "0.0") & " ms   C# (via COM) = " & Format(TotalCSEncode / NumCalls * 1000, "0.0") & " ms   speedup = " & Format(TotalVBAEncode / TotalCSEncode, "0.0") & "x" & vbCrLf & _
        "Decode: VBA = " & Format(TotalVBADecode / NumCalls * 1000, "0.0") & " ms   C# (via COM) = " & Format(TotalCSDecode / NumCalls * 1000, "0.0") & " ms   speedup = " & Format(TotalVBADecode / TotalCSDecode, "0.0") & "x"

    Debug.Print Report
    CSharpVFormatSpeedTest = Report

    Exit Function
ErrHandler:
    CSharpVFormatSpeedTest = "CSharpVFormatSpeedTest failed: " & Err.Description
End Function

' Compares the C# port of the general "*" per-element wire format (ComServer.EncodeGeneral/
' DecodeGeneral, csharp-prototype\src\GeneralFormat.cs) against VBA's SerialiseElement/
' Unserialise general-path fallback -- the path taken for Strings/mixed types, which has
' no bulk-CopyMemory equivalent in VBA (unlike the all-Double "V" path tested by
' CSharpVFormatSpeedTest above). Uses an all-String array ("1x", "2x", ..., "<n>x") so
' every element takes the general per-element-length-prefixed encoding in both VBA and C#.
Sub RunCSharpGeneralFormatSpeedTest()
    MsgBox CSharpGeneralFormatSpeedTest()
End Sub

Function CSharpGeneralFormatSpeedTest() As String

    Const VectorLength As Long = 1000000
    Const NumCalls As Long = 5

    Dim InputData As Variant
    Dim comServer As Object
    Dim i As Long
    Dim t1 As Double, t2 As Double
    Dim TotalVBAEncode As Double, TotalCSEncode As Double
    Dim TotalVBADecode As Double, TotalCSDecode As Double
    Dim VBAEncoded As String, CSEncoded As String
    Dim VBADecoded As Variant, CSDecoded As Variant
    Dim Report As String

    On Error GoTo ErrHandler

    Set comServer = CreateObject("JuliaExcelPrototype.ComServer")
    InputData = Application.Evaluate("=SEQUENCE(" & VectorLength & ")&""x""")

    ' --- Correctness: wire-format compatibility, both directions ---
    VBAEncoded = RunXlam("modSerialise.SerialiseElement", InputData)
    If Left$(VBAEncoded, 1) <> "*" Then Err.Raise vbObjectError + 1, , _
        "Expected VBA's SerialiseElement to take the general '*' path for an all-String array, got '" & Left$(VBAEncoded, 1) & "'"

    CSEncoded = comServer.EncodeGeneral(InputData)
    If VBAEncoded <> CSEncoded Then Err.Raise vbObjectError + 1, , _
        "VBA and C# general-format encodings differ for identical input"

    VBADecoded = RunXlam("modUnserialise.UnserialiseFromString", CSEncoded, False, 0, True)   ' VBA decoding C#'s output
    CSDecoded = comServer.DecodeGeneral(VBAEncoded)                                           ' C# decoding VBA's output
    For i = 1 To VectorLength
        If VBADecoded(i, 1) <> InputData(i, 1) Then Err.Raise vbObjectError + 1, , "VBA decode of C#-encoded string mismatch at element " & i
        If CSDecoded(i - 1, 0) <> InputData(i, 1) Then Err.Raise vbObjectError + 1, , "C# decode of VBA-encoded string mismatch at element " & i
    Next i

    ' --- Timing: interleaved, so warmup/drift affects both sides equally ---
    For i = 1 To NumCalls
        t1 = Timer
        VBAEncoded = RunXlam("modSerialise.SerialiseElement", InputData)
        t2 = Timer
        TotalVBAEncode = TotalVBAEncode + (t2 - t1)

        t1 = Timer
        CSEncoded = comServer.EncodeGeneral(InputData)
        t2 = Timer
        TotalCSEncode = TotalCSEncode + (t2 - t1)

        t1 = Timer
        VBADecoded = RunXlam("modUnserialise.UnserialiseFromString", VBAEncoded, False, 0, True)
        t2 = Timer
        TotalVBADecode = TotalVBADecode + (t2 - t1)

        t1 = Timer
        CSDecoded = comServer.DecodeGeneral(VBAEncoded)
        t2 = Timer
        TotalCSDecode = TotalCSDecode + (t2 - t1)
    Next i

    Report = "Correctness: VBA and C# general-format encodings are byte-identical; cross-decodes match." & vbCrLf & vbCrLf & _
        "Vector length = " & Format(VectorLength, "###,###") & " Strings, averaged over " & NumCalls & " calls (interleaved)" & vbCrLf & _
        "Encode: VBA = " & Format(TotalVBAEncode / NumCalls * 1000, "0.0") & " ms   C# (via COM) = " & Format(TotalCSEncode / NumCalls * 1000, "0.0") & " ms   speedup = " & Format(TotalVBAEncode / TotalCSEncode, "0.0") & "x" & vbCrLf & _
        "Decode: VBA = " & Format(TotalVBADecode / NumCalls * 1000, "0.0") & " ms   C# (via COM) = " & Format(TotalCSDecode / NumCalls * 1000, "0.0") & " ms   speedup = " & Format(TotalVBADecode / TotalCSDecode, "0.0") & "x"

    Debug.Print Report
    CSharpGeneralFormatSpeedTest = Report

    Exit Function
ErrHandler:
    CSharpGeneralFormatSpeedTest = "CSharpGeneralFormatSpeedTest failed: " & Err.Description
End Function

' Calls a Public (but possibly Option-Private-Module) procedure inside JuliaExcel.xlam via
' Application.Run, workbook-qualified so it resolves regardless of Option Private Module
' and regardless of whether the xlam is loaded as an ordinary workbook or an installed
' add-in. Supports up to 5 call-time arguments, which is all this file needs.
Function RunXlam(ProcName As String, ParamArray Args())
    Dim FullName As String
    FullName = "'" & GetXlamWorkbook().FullName & "'!" & ProcName

    Select Case UBound(Args) - LBound(Args) + 1
        Case 0: RunXlam = Application.Run(FullName)
        Case 1: RunXlam = Application.Run(FullName, Args(0))
        Case 2: RunXlam = Application.Run(FullName, Args(0), Args(1))
        Case 3: RunXlam = Application.Run(FullName, Args(0), Args(1), Args(2))
        Case 4: RunXlam = Application.Run(FullName, Args(0), Args(1), Args(2), Args(3))
        Case Else: Err.Raise vbObjectError + 1, , "RunXlam: unsupported argument count"
    End Select
End Function

' Pins to the installed add-in's exact path (XlamFullPath) rather than trusting a bare
' "JuliaExcel.xlam" name match -- see this module's header comment for why that matters.
Function GetXlamWorkbook() As Object
    Dim wb As Object
    On Error Resume Next
    Set wb = Application.Workbooks.Item("JuliaExcel.xlam")
    On Error GoTo 0
    If wb Is Nothing Then
        Err.Raise vbObjectError + 1, , "JuliaExcel.xlam is not loaded in this Excel session."
    End If
    If LCase$(wb.FullName) <> LCase$(XlamFullPath) Then
        Err.Raise vbObjectError + 1, , "A workbook/add-in named 'JuliaExcel.xlam' is loaded from '" & _
            wb.FullName & "', not the expected installed add-in at '" & XlamFullPath & _
            "'. Close the other copy first so this test runs against the real production code, not a different copy."
    End If
    Set GetXlamWorkbook = wb
End Function
