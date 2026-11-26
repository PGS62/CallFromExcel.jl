# Marshalling

Two questions arose during implementation. First, what Julia type should data from a worksheet or VBA arrive as? Mostly that's easy to decide, but should a single-column or single-row range, or a one-dimensional VBA array, have one dimension or two in Julia? Second, in the other direction, how should a Julia vector come back to Excel?

The design had three objectives:
 1) Round-tripping: `=JuliaCall("identity",x)` should return an identical copy of `x`, whatever its shape.
 2) Natural matrix arithmetic: `=JuliaCall("*",Range1,Range2)` should give the same result as `=MMULT(Range1,Range2)`, since `*` in Julia is matrix multiplication.
 3) Julia's dot syntax for broadcasting should work from `JuliaCall`.

This scheme achieves them:
 * Scalars marshal back and forth as scalars.
 * Ranges or arrays with more than one row and more than one column marshal back and forth as two-dimensional arrays.
 * A single-column range arrives in Julia as a vector, and a Julia vector returns to the worksheet as a single column.
 * A single-row range arrives in Julia as a two-dimensional array with one row.

From VBA, one-dimensional arrays arrive in Julia as vectors, and `JuliaCallVBA` and `JuliaEvalVBA` return Julia vectors as one-dimensional arrays, so VBA variables round-trip too.

!!! details "Round-tripping of vectors and matrices"
    ![Round-tripping vectors and matrices between an Excel worksheet and Julia](assets/images/roundtripping.gif)

!!! details "Matrix arithmetic"
    ![Matrix arithmetic in Julia on Excel ranges, results spilling as dynamic arrays](assets/images/matrixarithmetic.gif)

!!! details "Function broadcasting"
    ![Broadcasting a Julia function over Excel ranges with JuliaCall](assets/images/functionbroadcasting.gif)

## [Numbers from a worksheet arrive as `Float64`](@id numbers-from-a-worksheet-arrive-as-float64)
Excel stores every number on a worksheet as a double-precision float, so numbers passed from a worksheet to `JuliaCall`, `JuliaCallNamedArgs` or `JuliaSetVar` always arrive in Julia as `Float64`, even whole numbers such as `2`. This includes the results of other functions in the same formula, since Excel converts a function's result to a worksheet value before passing it on. So `=JuliaCall("typeof", JuliaEval("2"))` returns `Float64`, even though `JuliaEval("2")` returns an `Int64` to VBA (but see the third workaround below).

This matters when calling a Julia function that requires an integer, whether as a positional argument (e.g. `rand(3)`) or a named one (e.g. `round(x; digits=2)`, `split(s, ","; limit=2)`). Passing `Float64` gives a `MethodError` or `TypeError`. There are three workarounds:
 * Write the whole call in `JuliaEval`, e.g. `=JuliaEval("rand(3)")`, if the integer is a constant.
 * Write a small Julia function that converts, e.g. `myrand(n) = rand(Int(n))`, loaded with `JuliaInclude`, and call that instead.
 * Pass the integer as a [handle](handles.md). `JuliaEval("3", TRUE)` evaluates `3` in Julia and returns a handle to the resulting `Int64`, which Julia receives as the integer itself, not as a worksheet number. So `=JuliaCall("rand", JuliaEval("3", TRUE))` calls `rand(3)`, and `=JuliaCallNamedArgs("round", A1, ";", "digits", JuliaEval("2", TRUE))` calls `round(A1; digits=2)`. Unlike the first workaround, the integer can come from a cell: `JuliaEval(B1, TRUE)` evaluates B1's number as Julia code.

From VBA, numbers keep their VBA type: `Integer` arrives as `Int16`, `Long` as `Int32`, `LongLong` as `Int64`, `Single` as `Float32` and `Double` as `Float64`. So pass a `LongLong` (e.g. `2^` or `CLngLng(n)`) where Julia needs an `Int64`. Note that a literal such as `2` in VBA code is an `Integer`, which arrives as `Int16` - some Julia functions, such as `split`, require an `Int64` and won't accept that either.

Julia's integer arithmetic wraps around on overflow rather than raising an error, and the narrower the type, the sooner that happens. So from VBA, `jl.JuliaCall("*", 200, 200)` returns -25536, since both arguments arrive as `Int16`, where VBA itself would raise an Overflow error. Pass `LongLong` or `Double` values when results may be large.

## Complex numbers return as text
Excel has no complex number type. Its functions such as `IMREAL`, `IMAGINARY` and `IMSUM` work on text such as `"3+4i"` instead. So complex numbers in Julia return to Excel as text in exactly the format of Excel's own `COMPLEX` function, e.g. `"3+4i"`, `"-i"` or `"1.5+i"`, with 15 significant digits. As with `COMPLEX`, a number with a zero imaginary part returns as text without one, e.g. `"3"`. Text in this form passed from Excel to Julia arrives as a `String`, not as a complex number.

## Dates use the 1900 date system
Dates and times returned from Julia to a worksheet become Excel date serial numbers in the 1900 date system, Excel's default on Windows. In a workbook set to the 1904 date system (File > Options > Advanced > "Use 1904 date system"), they display 4 years and 1 day late. Dates passed from a worksheet to Julia arrive as plain numbers, as Excel stores them, whichever date system the workbook uses. Calls from VBA are unaffected, since VBA dates aren't serial numbers.
