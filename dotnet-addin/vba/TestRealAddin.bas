Attribute VB_Name = "TestRealAddin"
Option Explicit

' Manual test module for the real C# add-in (dotnet-addin\JuliaExcel.AddIn), covering
' both call paths: JuliaLaunch/JuliaCall/JuliaEval as worksheet (XLL) functions, and
' JuliaCallVBA/JuliaEvalVBA via the COM-visible class (for Dictionary/nested-array
' results that Application.Run/XLL functions can't carry).
'
' Setup (one-time):
' 1. Build: cd dotnet-addin\JuliaExcel.AddIn && dotnet build -c Release
' 2. Load the add-in: Excel > File > Options > Add-ins > Manage Excel Add-ins > Go... >
'    Browse... > select bin\Release\net48\publish\JuliaExcel.AddIn-AddIn64-packed.xll
' 3. COM registration for JuliaCallVBA/JuliaEvalVBA (via ComServer): this session
'    registered JuliaExcel.AddIn.ComServer under HKCU:\Software\Classes (no admin
'    needed) pointing at bin\Release\net48\JuliaExcel.AddIn.dll. If you rebuild and the
'    DLL path changes, or on a different machine/user account, re-register with:
'      $dllPath = "<repo>\dotnet-addin\JuliaExcel.AddIn\bin\Release\net48\JuliaExcel.AddIn.dll"
'      (see this session's own commands for the exact HKCU registration script, or use
'      an elevated RegAsm /codebase against that DLL for a normal HKLM registration)
' 4. Import this module into any workbook.
' 5. Run TestWorksheetFunctions, then TestComPath.

Sub TestWorksheetFunctions()
    Dim result As Variant

    result = JuliaLaunch()
    Debug.Print "JuliaLaunch: " & result

    result = JuliaEval("1+1")
    Debug.Print "JuliaEval(""1+1""): " & result

    result = JuliaCall("sqrt", 16#)
    Debug.Print "JuliaCall(""sqrt"", 16): " & result

    result = JuliaCall("uppercase", "hello")
    Debug.Print "JuliaCall(""uppercase"", ""hello""): " & result

    result = JuliaEval("[1.0,2.0,3.0]")
    Debug.Print "JuliaEval array result, element 1: " & result(1, 1) & ", element 3: " & result(3, 1)

    MsgBox "Worksheet-function tests complete - see Immediate window (Ctrl+G) for results."
End Sub

Sub TestComPath()
    Dim comServer As Object
    Set comServer = CreateObject("JuliaExcel.AddIn.ComServer")

    Dim result As Variant
    result = comServer.JuliaCallVBA("+", 3, 4)
    Debug.Print "JuliaCallVBA(""+"", 3, 4): " & result

    result = comServer.JuliaCallVBA("sqrt", 16#)
    Debug.Print "JuliaCallVBA(""sqrt"", 16): " & result

    ' The whole point of the COM path: Dictionary and nested-array results, which
    ' JuliaCall/JuliaEval (XLL functions) cannot return.
    Dim dict As Object
    Set dict = comServer.JuliaEvalVBA("Dict(""name""=>""Julia"",""version""=>1.13)")
    Debug.Print "Dictionary Count: " & dict.Count
    Debug.Print "Dictionary Item(""name""): " & dict.Item("name")
    Debug.Print "Dictionary Exists(""version""): " & dict.Exists("version")

    ' Note: arrays returned via the COM path are 0-based (LBound=0), unlike the XLL/
    ' worksheet-function path (JuliaCall/JuliaEval), which Excel-DNA automatically
    ' rebases to 1-based for Excel's own worksheet display convention. COM Automation
    ' itself does no such rebasing.
    Dim nested As Variant
    nested = comServer.JuliaEvalVBA("[[1.0,2.0],[3.0,4.0,5.0]]")
    Debug.Print "Nested array: nested(0) is an array of length " & (UBound(nested(0)) - LBound(nested(0)) + 1)
    Debug.Print "Nested array: nested(1) is an array of length " & (UBound(nested(1)) - LBound(nested(1)) + 1)

    MsgBox "COM-path tests complete - see Immediate window (Ctrl+G) for results."
End Sub
