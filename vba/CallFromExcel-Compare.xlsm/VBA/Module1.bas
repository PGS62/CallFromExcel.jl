Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=""String""&InputData"
End Sub
