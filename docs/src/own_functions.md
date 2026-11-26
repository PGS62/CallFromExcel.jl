# Your own worksheet functions

Instead of `=JuliaCall("birthday", A1)`, a Julia function can be a worksheet function of its own, `=jl.birthday(A1)`, with its arguments and description shown by Excel's Insert Function (fx) dialog and in the tooltips as you type. Mark it with `@excel`:

```julia
"""
    birthday(people, days=365)

The probability that, in a group of people, at least two share a birthday.

# Arguments
- `people`: the number of people in the group.
- `days`: the number of days in the year.
"""
@excel broadcast=true function birthday(people, days=365)
    1 - prod((days - k) / days for k in 0:people-1)
end
```

Include the file - with [`JuliaInclude`](@ref juliainclude), or the ribbon's Include button - and `=jl.birthday(23)` works in any cell of any open workbook, returning 0.507: in a group of just 23 people, it's more likely than not that two share a birthday. `broadcast=true` makes it work on ranges too: with the numbers 1 to 60 in A1:A60, `=jl.birthday(A1:A60)` returns a column of 60 probabilities, ready to chart. Typing `=jl.` lists all such functions. A Julia function that's already defined, including one from a package, can be made available without changing its definition with `expose_to_excel`, for example `expose_to_excel("hypot")`, or `expose_to_excel(LinearAlgebra.det)` for `=jl.det(A1:C3)`.

`@excel` and `expose_to_excel` take the same options:

|Option|Description|
|:-----|:----------|
|`name`|The Excel name. Default `"jl."` followed by the Julia name. The prefix keeps your functions apart from Excel's own (when names clash, Excel's built-in function wins) and from cell references: `abc123` is a cell, `jl.abc123` isn't. A name without it, such as `name="BIRTHDAY"`, is allowed. A name Excel can't use is refused with an error saying why: one with characters Excel doesn't allow (`push!`), one Excel would read as a cell reference, one beginning with `Julia` (reserved for the add-in's functions), or one differing only in case from a name already in use, since Excel ignores case.|
|`description`, `argdescriptions`|Descriptions of the function and of each argument, for the Insert Function dialog. By default they come from the docstring: the first paragraph after the signature, and the `# Arguments` list. Each is cut to 253 characters, Excel's limit.|
|`broadcast`|Which arguments, given a range, are broadcast over - see below. Default `false`, none.|

Points to note:
 * **Optional arguments.** Arguments with default values (`days=365`) are optional: leave them out at the end of a call, as in `=jl.birthday(23)`, and Julia uses the defaults; give them to override the defaults, as in `=jl.birthday(23, 366)`. Leaving out a required argument (`=jl.birthday()`), or one before an argument that is given (`=jl.birthday(, 366)`), is an error naming it. Keyword arguments aren't available as worksheet arguments; use [`JuliaCallNamedArgs`](@ref juliacallnamedargs) for those.
 * **Which method.** `@excel` follows the method it's attached to. `expose_to_excel` follows the method with the most arguments, not counting methods with a variable number of arguments (`f(x...)`).
 * **Ranges are passed whole.** By default a range arrives in Julia as a matrix (a single cell as a single value), as with `JuliaCall`. An argument declared `::AbstractVector` receives a vector, from a single cell, a row or a column.
 * **Broadcasting.** For a function written for single values, `broadcast` makes it work on ranges: a range given to a broadcasting argument is broadcast over, as Julia's `f.(x, y)` does, in one call from Excel, and the result spills. `broadcast=true` broadcasts every argument, except those declared as arrays (`w::AbstractVector`), which are passed whole; `broadcast=(:x, :y)`, or `broadcast=:x` for one, names the arguments to broadcast. For example, `birthday` above is marked `broadcast=true`, so `=jl.birthday(A1:A60)` returns a column of 60 probabilities, and with 365 and 366 in B1:C1, `=jl.birthday(A1:A60, B1:C1)` - a column against a row - a 60 by 2 table. Without `broadcast=true`, `birthday` would receive A1:A60 as a matrix, and fail. Ranges broadcast together must have the same number of rows, or one of them a single row, and likewise for columns; otherwise the error names the two arguments (where Excel would pad with `#N/A`). A broadcast function that returns an array for each element gives an error, as the result can't be shown in one range.
 * **When functions are registered.** After `JuliaLaunch`, `JuliaInclude` and `JuliaEval`, and the ribbon's Launch Julia and Include buttons, the add-in registers any functions that have been exposed since it last did, once Excel is next idle. Functions exposed from the Julia REPL, or from VBA's `jl.JuliaEval` or `jl.JuliaInclude`, are registered by the ribbon's Register Functions button, or from VBA with `Application.Run "JuliaRegisterFunctions"`, which registers them straight away. Registered functions stay registered until Excel closes: a function no longer exposed in a new Julia session gives an error saying so.
 * **Formulas typed too soon show `#NAME?`.** A formula that uses a function before it is registered shows `#NAME?`, and goes on doing so after registration, until it is re-entered (select it, press F2, then Enter). This doesn't affect saved workbooks: reopened before Julia is running, their cells keep their saved values, and recalculate correctly once the functions are registered (in manual calculation mode, at the next recalculation). VBA code that writes such formulas should run `Application.Run "JuliaRegisterFunctions"` first.
 * **Changing a function.** Change its code and re-include the file: the worksheet function calls whatever the Julia function now does, with no need to register again. (Cells using it aren't recalculated by the change; press Ctrl+Alt+F9.) Changing its arguments needs it to be exposed again, which re-including a file with `@excel` does.
 * **From VBA.** `Application.Run("jl.birthday", 23)` calls a registered function.

See also [`@excel`](@ref) and [`expose_to_excel`](@ref) in the [Julia API](api.md).
