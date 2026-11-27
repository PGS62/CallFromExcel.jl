# Object handles

Some Julia objects are large, or have no sensible representation on a worksheet: a big matrix, a fitted model, a DataFrame, a dictionary. Rather than returning such an object to Excel, [`JuliaCallHandle`](@ref juliacallhandle) (and [`JuliaEval`](@ref juliaeval) with `ReturnHandle` TRUE) keeps it in Julia and returns a _handle_, a short string naming it, such as `jl:Array#17@k3x9`. Pass the handle to any function in place of the object, and Julia receives the object itself:

|Cell|Formula|Result|
|:---|:------|:-----|
|A1|`=JuliaEval("rand(1000, 1000)", TRUE)`|`jl:Array#1@k3x9`|
|A2|`=JuliaCall("sum", A1)`|the sum of the matrix's million elements|
|A3|`=JuliaCallHandle("inv", A1)`|`jl:Array#2@k3x9`, a handle to its inverse|
|A4|`=JuliaCall("maximum", A3)`|the inverse's largest element|

The matrix is built once and stays in Julia, however many formulas use it. A handle can be passed to `JuliaCall`, `JuliaCallNamedArgs`, `JuliaCallHandle` and their VBA counterparts, including inside a range or array: a range of handles arrives in Julia as a vector of the objects. `=JuliaCall("identity", A1)` returns a handle's object to the worksheet as data. A handle's text is `jl:<label>#<id>@<session>`: the label is the object's type (or a label you choose, see below), for people reading the sheet, the id numbers each object, and the session identifies the Julia process holding it.

The functions for handles are [`JuliaCallHandle`](@ref juliacallhandle), [`JuliaEvalHandle`](@ref juliaevalhandle), [`JuliaHandleRelease`](@ref juliahandlerelease), [`JuliaHandleReleaseAll`](@ref juliahandlereleaseall) and [`JuliaHandleInfo`](@ref juliahandleinfo), and `JuliaEval` with `ReturnHandle` TRUE.

## Points to note

### Lifetime
A handle's object stays in Julia while a formula that creates it exists. Edit or delete that formula, or change its inputs, and the object is freed (and a new one created for new inputs). Excel sometimes frees an object only at a later recalculation, for example after a workbook is closed.

### Same formula, same object
Formulas with the same function and the same arguments share one object, as do copies of a cell, and recalculating - even with Ctrl+Alt+F9 - doesn't rebuild it. So `=JuliaEval("rand(3)", TRUE)` keeps its first random numbers. To get a separate, or new, object, change an input, for example an extra argument that the Julia function ignores. Running the macro `JuliaHandleInvalidateAll` (Alt+F8, type the name, Run) rebuilds every handle's object.

### Volatile arguments
A handle formula whose arguments change at every recalculation - because it uses a volatile function such as `RAND`, `RANDARRAY`, `NOW`, `OFFSET` or `INDIRECT`, directly or through the cells it refers to - can't keep a handle. In automatic calculation, each new handle makes Excel recalculate its cell, which then has new arguments, so another handle is created, and so on. CallFromExcel notices after a few rounds and returns an error saying the arguments change at every recalculation. Instead:

- generate the data in Julia, e.g. `=JuliaEval("rand(1000)", TRUE)` rather than `=JuliaCallHandle("identity", RANDARRAY(1000))`;
- if the result needn't stay in Julia, use `JuliaCall`, e.g. `=JuliaCall("sum", RANDARRAY(1000))`;
- copy the volatile values and paste them as values;
- or, to get a new object at every recalculation (F9), switch to manual calculation (Formulas > Calculation Options > Manual), where Excel doesn't recalculate after each new handle. Note that this setting applies to all open workbooks.

### Don't change a handle's object
A function that changes its argument, such as `push!` or `sort!`, changes the stored object without changing its handle, so cells using it aren't recalculated, and results depend on the order Excel calculates in. Use functions that return a new object, such as `sort`, and make a new handle of the result.

### When Julia restarts
`JuliaLaunch`, and the ribbon's Launch Julia button, notice a new Julia session and rebuild every handle's object in it, in the order Excel's dependencies require. A cell that defines functions the handles need should refer to the launch cell, so that it is recalculated first, e.g. `=IF(LEN(Z1)>0, JuliaInclude("C:\MyCode\models.jl"))` where `Z1` holds `=JuliaLaunch()`. Until then - for example after Julia is restarted from VBA - a handle from the old session gives an error saying it was not found.

### Strings that look like handles are handles
Any string matching the pattern `jl:<label>#<id>@<session>` is taken to be a handle, whether a formula returned it or it was typed. If its object has been freed, or it comes from an earlier Julia session (for example, saved in a workbook), the call gives a "not found" error, rather than passing the string to Julia.

### Errors pass through
If an argument to `JuliaCallHandle`, or to a function exposed with `handle=true`, is an error value, such as `#N/A`, it returns that error without calling Julia. (`JuliaCall` instead passes the error to Julia, as an `ExcelError`.) This is because a handle cell shows `#N/A` while its object is being built, and after Julia restarts until it is rebuilt: in a chain of handles, each cell waits for the one before it, rather than running its function on `#N/A` and creating an object that is replaced moments later. An error inside an array argument, such as a range containing `#N/A`, is passed to Julia as usual.

### From VBA
VBA can pass handles from cells to `jl.JuliaCall` etc. Handles that VBA creates, with `jl.JuliaCallHandle` or `jl.JuliaEvalHandle`, belong to no formula, so they last until `jl.JuliaHandleRelease` or `jl.JuliaHandleReleaseAll` frees them, or Julia exits. (Calling the worksheet function through `Application.Run` or `Evaluate` gives a handle that is freed straight away.)

### Labels for your own types
The label is the name of the object's type, without its module or parameters (`Array` for a `Matrix{Float64}`). To choose your own, add a method to `CallFromExcel.handle_label`, for example `CallFromExcel.handle_label(x::YieldCurve) = "Curve_$(x.currency)"`, so that a sheet shows `jl:Curve_USD#3@k3x9`. The label is worked out once, when the handle is created, and any `#`, `@`, whitespace or control character in it is replaced by `_`.
