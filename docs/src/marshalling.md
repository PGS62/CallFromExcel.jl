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

## Types

From Excel or VBA to Julia:

|Excel or VBA|Julia|
|:-----------|:----|
|A number on a worksheet|`Float64` - see [below](@ref numbers-from-a-worksheet-arrive-as-float64)|
|VBA `Double`, `Single`|`Float64`, `Float32`|
|VBA `Integer`, `Long`, `LongLong`, `Byte`|`Int16`, `Int32`, `Int64`, `UInt8`|
|Text, VBA `String`|`String`|
|`TRUE`, `FALSE`, VBA `Boolean`|`Bool`|
|An empty cell, VBA `Empty`|`missing`|
|VBA `Null`|`nothing`|
|An error value, such as `#N/A`|[`ExcelError`](@ref)|
|A date on a worksheet|`Float64`, its serial number - see [Dates](@ref dates-use-the-1900-date-system)|
|VBA `Date`|`Date`, or `DateTime` if it has a time of day|
|A range, or VBA array|An array, shaped as described above. If its elements all have the same type, the array has that element type (a range of numbers is a `Vector{Float64}` or `Matrix{Float64}`); otherwise `Any`.|
|`Scripting.Dictionary` (from VBA)|`Dict`|
|A [handle](handles.md)|The object it stands for|
|VBA `Currency`, `Decimal`|Not supported: the call fails|

From Julia to a worksheet, or to VBA:

|Julia|Worksheet|VBA|
|:----|:--------|:--|
|`Float64`|Number|`Double`|
|`Float32`, `Float16`|Number|`Single`|
|`Int8`, `Int16`|Number|`Integer`|
|`Int32`, `UInt16`|Number|`Long`|
|`Int64`, `UInt32`|Number, exact up to 2^53|`LongLong`|
|`UInt8`|Number|`Byte`|
|`Int128`, `UInt64`, `BigInt`, `BigFloat`, `Rational`, `Irrational`|Number, exact up to 2^53|`Double`, exact up to 2^53|
|`NaN`, `Inf`|`#N/A`, `#NUM!`|The same error values|
|`Bool`|`TRUE`, `FALSE`|`Boolean`|
|`String`, `Char`|Text, of at most 32,767 characters|`String`. With the functions ending in `VBA`, of any length.|
|`missing`, `nothing`|An empty value, which a cell shows as 0|`Empty`|
|`Date`, `DateTime`|Its serial number: format the cell as a date to show it as one|`Date`|
|`Complex`|Text, as Excel's `COMPLEX` gives - see [below](@ref complex-numbers-return-as-text)|`String`|
|[`ExcelError`](@ref)|The error value|The error value (as `CVErr` gives)|
|A vector, `Tuple` or range such as `1:10`|A column|A column (a two-dimensional array); a one-dimensional array from the functions ending in `VBA`|
|A matrix|A range of the same shape|A two-dimensional array|
|An array of three or more dimensions|Refused: the call fails|An array of up to 9 dimensions, from the functions ending in `VBA`; otherwise refused|
|An array of arrays|Refused|An array of arrays, from the functions ending in `VBA`; otherwise refused|
|`Dict`|Refused|A `Scripting.Dictionary`, from the functions ending in `VBA`; otherwise refused|
|`DataFrame`|A range: a header row of column names above the columns|The same, as a two-dimensional array|
|Anything else - a `struct`, `Symbol`, function or type|Text, as the Julia REPL shows it|`String`, likewise|

A call that is refused fails as any other: with a string starting with `#` on a worksheet, and an [error](@ref vba-errors) in VBA. To keep an object that can't go to the worksheet in Julia, return a [handle](handles.md) to it.

## [Numbers from a worksheet arrive as `Float64`](@id numbers-from-a-worksheet-arrive-as-float64)
Excel stores every number on a worksheet as a double-precision float, so numbers passed from a worksheet to `JuliaCall`, `JuliaCallNamedArgs` or `JuliaSetVar` always arrive in Julia as `Float64`, even whole numbers such as `2`. This includes the results of other functions in the same formula, since Excel converts a function's result to a worksheet value before passing it on. So `=JuliaCall("typeof", JuliaEval("2"))` returns `Float64`, even though `JuliaEval("2")` returns an `Int64` to VBA (but see the third workaround below).

This matters when calling a Julia function that requires an integer, whether as a positional argument (e.g. `rand(3)`) or a named one (e.g. `round(x; digits=2)`, `split(s, ","; limit=2)`). Passing `Float64` gives a `MethodError` or `TypeError`. There are three workarounds:
 * Write the whole call in `JuliaEval`, e.g. `=JuliaEval("rand(3)")`, if the integer is a constant.
 * Write a small Julia function that converts, e.g. `myrand(n) = rand(Int(n))`, loaded with `JuliaInclude`, and call that instead.
 * Pass the integer as a [handle](handles.md). `JuliaEval("3", TRUE)` evaluates `3` in Julia and returns a handle to the resulting `Int64`, which Julia receives as the integer itself, not as a worksheet number. So `=JuliaCall("rand", JuliaEval("3", TRUE))` calls `rand(3)`, and `=JuliaCallNamedArgs("round", A1, ";", "digits", JuliaEval("2", TRUE))` calls `round(A1; digits=2)`. Unlike the first workaround, the integer can come from a cell: `JuliaEval(B1, TRUE)` evaluates B1's number as Julia code.

From VBA, numbers keep their VBA type: `Integer` arrives as `Int16`, `Long` as `Int32`, `LongLong` as `Int64`, `Single` as `Float32` and `Double` as `Float64`. So pass a `LongLong` (e.g. `2^` or `CLngLng(n)`) where Julia needs an `Int64`. Note that a literal such as `2` in VBA code is an `Integer`, which arrives as `Int16` - some Julia functions, such as `split`, require an `Int64` and won't accept that either.

Julia's integer arithmetic wraps around on overflow rather than raising an error, and the narrower the type, the sooner that happens. So from VBA, `jl.JuliaCall("*", 200, 200)` returns -25536, since both arguments arrive as `Int16`, where VBA itself would raise an Overflow error. Pass `LongLong` or `Double` values when results may be large.

## [Complex numbers return as text](@id complex-numbers-return-as-text)
Excel has no complex number type. Its functions such as `IMREAL`, `IMAGINARY` and `IMSUM` work on text such as `"3+4i"` instead. So complex numbers in Julia return to Excel as text in exactly the format of Excel's own `COMPLEX` function, e.g. `"3+4i"`, `"-i"` or `"1.5+i"`, with 15 significant digits. As with `COMPLEX`, a number with a zero imaginary part returns as text without one, e.g. `"3"`. Text in this form passed from Excel to Julia arrives as a `String`, not as a complex number.

## [Dates use the 1900 date system](@id dates-use-the-1900-date-system)
Dates and times returned from Julia to a worksheet become Excel date serial numbers in the 1900 date system, Excel's default on Windows. In a workbook set to the 1904 date system (File > Options > Advanced > "Use 1904 date system"), they display 4 years and 1 day late. Dates passed from a worksheet to Julia arrive as plain numbers, as Excel stores them, whichever date system the workbook uses. Calls from VBA are unaffected, since VBA dates aren't serial numbers.
