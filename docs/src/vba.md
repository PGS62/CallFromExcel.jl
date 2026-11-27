# Calling from VBA

The VBA code below makes a call to [`JuliaLaunch`](@ref julialaunch) and [`JuliaEvalVBA`](@ref juliaevalvba) and then pastes the result to range A1:J10 in a new worksheet. VBA reaches CallFromExcel's functions through an object created with `CreateObject("CallFromExcel.ComServer")`, which needs no reference. Alternatively, add a reference to CallFromExcel (VBA editor, Tools menu -> References) and declare `Dim jl As New CallFromExcel.ComServer`: the VBA editor then shows each function's arguments as you type. But a workbook with that reference won't compile on a PC without CallFromExcel installed.

```vba
Sub Demo()

    Dim jl As Object, ResultFromJulia As Variant, PasteHere As Range

    Set jl = CreateObject("CallFromExcel.ComServer")
    jl.JuliaLaunch

    ResultFromJulia = jl.JuliaEvalVBA("(1:10).^(1:10)'")

    Set PasteHere = Application.Workbooks.Add.Worksheets(1).Range("A1:J10")

    PasteHere.Value = ResultFromJulia

End Sub
```

The functions ending in `VBA` - [`JuliaEvalVBA`](@ref juliaevalvba), [`JuliaCallVBA`](@ref juliacallvba) and [`JuliaCallNamedArgsVBA`](@ref juliacallnamedargsvba) - and [`JuliaEvalHandle`](@ref juliaevalhandle), [`JuliaHandleRelease`](@ref juliahandlerelease) and [`JuliaHandleReleaseAll`](@ref juliahandlereleaseall) can only be called from VBA. From VBA, the `JuliaCall` family takes at most 30 arguments after the function name.

## Errors

A call from VBA that fails raises an error, for `On Error` to handle, rather than returning a value. (On a worksheet, where a formula can't raise an error, the same failure shows as a string such as `#MethodError: ... Julia REPL has more details and stacktrace!`.) `Err.Source` is `"CallFromExcel"`, `Err.Description` the message, and `Err.Number` one of:

|`Err.Number`|Meaning|
|:-----------|:------|
|`vbObjectError + 1001`|The Julia code threw an exception. `Err.Description` starts with its type, such as `MethodError:`, and the Julia REPL shows the full error and stack trace.|
|`vbObjectError + 1002`|Julia isn't running: not yet launched, exited, or still starting.|
|`vbObjectError + 1003`|Any other failure: a result that can't be returned, such as a dictionary from `JuliaCall` rather than `JuliaCallVBA`; an argument that can't be sent to Julia; or `JuliaLaunch` failing to start Julia.|

For example:

```vba
Sub FitModel()
    Dim jl As Object, Model As Variant
    Set jl = CreateObject("CallFromExcel.ComServer")

    On Error GoTo Failed
    Model = jl.JuliaCallVBA("fit_model", Range("A1:A100").Value, Range("B1:B100").Value)
    ' ...use Model...
    Exit Sub

Failed:
    If Err.Number = vbObjectError + 1002 Then
        MsgBox "Please launch Julia first."
    Else
        MsgBox "fit_model failed: " & Err.Description
    End If
End Sub
```

Excel error values are data, not failures: a result that is, or contains, `#N/A` or another error value is returned as such (`IsError` is `True` for it), and a string that merely looks like an error, such as `"#N/A"` or `"#Done!"`, is returned as a string.

Until version 0.2.0, these functions returned failures as strings starting with `#`, which VBA code had to check for after each call.
