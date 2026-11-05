Attribute VB_Name = "TestPrototype"
Option Explicit

' Toolchain-validation prototype only: not wired to Julia.
' Exercises both halves of the C#/Excel-DNA prototype from VBA:
'   1) the XLL worksheet function, called via Application.Run
'   2) the COM-visible class, called via CreateObject
' Load JuliaExcelPrototype64.xll as an add-in before running this.

Sub TestPrototype()

    Dim xllResult As Double
    xllResult = Application.Run("JepAdd", 3, 4)

    Dim comServer As Object
    Set comServer = CreateObject("JuliaExcelPrototype.ComServer")
    Dim comResult As String
    comResult = comServer.HelloFromCSharp("Philip")

    MsgBox "XLL function JepAdd(3, 4) = " & xllResult & vbCrLf & _
           "COM method HelloFromCSharp(""Philip"") = " & comResult

End Sub
