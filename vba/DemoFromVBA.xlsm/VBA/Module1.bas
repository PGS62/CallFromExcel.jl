Attribute VB_Name = "Module1"
Option Explicit

Sub Demo()

    Dim CallFromExcel As Object, ResultFromJulia As Variant, PasteHere As Range
    
    Set CallFromExcel = CreateObject("CallFromExcel.ComServer")
    CallFromExcel.JuliaLaunch
    
    ResultFromJulia = CallFromExcel.JuliaEvalVBA("(1:10).^(1:10)'")

    Set PasteHere = Application.Workbooks.Add.Worksheets(1) _
        .Cells(1, 1).Resize(UBound(ResultFromJulia, 1), _
        UBound(ResultFromJulia, 2))
    
    PasteHere.Value = ResultFromJulia

End Sub
