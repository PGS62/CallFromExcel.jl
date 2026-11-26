# Calling from VBA

The VBA code below makes a call to [`JuliaLaunch`](@ref julialaunch) and [`JuliaEvalVBA`](@ref juliaevalvba) and then pastes the result to range A1:J10 in a new worksheet. VBA reaches CallFromExcel's functions through an object created with `CreateObject("CallFromExcel.ComServer")`, which needs no reference. Alternatively, add a reference to CallFromExcel (VBA editor, Tools menu -> References) and declare `Dim jl As New CallFromExcel.ComServer`: the VBA editor then shows each function's arguments as you type. But a workbook with that reference won't compile on a PC without CallFromExcel installed.

```vba
Sub Demo()

    Dim jl As Object, ResultFromJulia As Variant, PasteHere As Range

    Set jl = CreateObject("CallFromExcel.ComServer")
    jl.JuliaLaunch

    ResultFromJulia = jl.JuliaEvalVBA("(1:10).^(1:10)'")

    Set PasteHere = Application.Workbooks.Add.Worksheets(1) _
        .Cells(1, 1).Resize(UBound(ResultFromJulia, 1), _
        UBound(ResultFromJulia, 2))

    PasteHere.Value = ResultFromJulia

End Sub
```

The functions ending in `VBA` - [`JuliaEvalVBA`](@ref juliaevalvba), [`JuliaCallVBA`](@ref juliacallvba) and [`JuliaCallNamedArgsVBA`](@ref juliacallnamedargsvba) - and [`JuliaEvalHandle`](@ref juliaevalhandle), [`JuliaHandleRelease`](@ref juliahandlerelease) and [`JuliaHandleReleaseAll`](@ref juliahandlereleaseall) can only be called from VBA. From VBA, the `JuliaCall` family takes at most 30 arguments after the function name.
