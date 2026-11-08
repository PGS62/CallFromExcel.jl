Attribute VB_Name = "modUtils"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/CallFromExcel.jl#readme

Option Explicit
Option Private Module

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
' -----------------------------------------------------------------------------------------------------------------------
Function ElapsedTime() As Double
          Dim A As Currency
          Dim B As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter A
3         QueryPerformanceFrequency B
4         ElapsedTime = A / B

5         Exit Function
ErrHandler:
6         ReThrow "ElapsedTime", Err
End Function

'Called from "Menu..." button on sheet Audit.
Sub MenuButton()
1         On Error GoTo ErrHandler
2         Application.Run "SolumAddin.xlam!AuditMenu"
3         Exit Sub
ErrHandler:
4         MsgBox ReThrow("MenuButton", Err, True), vbCritical
End Sub

Sub PreciseSleep(Milliseconds As Double)
          Dim StartTime As Double
1         StartTime = ElapsedTime()
2         Do Until ((ElapsedTime() - StartTime) > Milliseconds / 1000) Or (ElapsedTime() < StartTime)
3         Loop
End Sub
