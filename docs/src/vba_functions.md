# VBA functions

These functions can only be called from VBA, through the object described in [Calling from VBA](vba.md), as for example `jl.JuliaCallVBA(...)`. VBA can also call the [worksheet functions](functions.md), all except `JuliaHandleInfo`. A call that fails raises an [error](@ref vba-errors).

|Name|Description|
|:---|:----------|
|[JuliaEvalVBA](@ref juliaevalvba)|Evaluate a Julia expression from VBA. Differs from `JuliaEval` in handling of 1-d arrays and strings longer than 32,767 characters. May return data of types that cannot be displayed on a worksheet, such as a dictionary or an array of arrays.|
|[JuliaCallVBA](@ref juliacallvba)|Call a named Julia function from VBA. Differs from `JuliaCall` in handling of 1-d arrays and strings longer than 32,767 characters. May return data of types that cannot be displayed on a worksheet, such as a dictionary, an array of arrays, or arrays of dimension up to 9.|
|[JuliaCallNamedArgsVBA](@ref juliacallnamedargsvba)|Call a named Julia function from VBA, passing positional and named arguments, with `";"` marking where the named arguments start. Differs from `JuliaCallNamedArgs` in the same ways that `JuliaCallVBA` differs from `JuliaCall`.|
|[JuliaEvalHandle](@ref juliaevalhandle)|Evaluate a Julia expression from VBA and return a handle to the result. The VBA counterpart of `JuliaEval(JuliaExpression, TRUE)`.|
|[JuliaHandleRelease](@ref juliahandlerelease)|Free the Julia object of a handle created from VBA.|
|[JuliaHandleReleaseAll](@ref juliahandlereleaseall)|Free the Julia objects of all handles created from VBA.|

## [`JuliaEvalVBA`](@id juliaevalvba)
Evaluate a Julia expression from VBA. Differs from `JuliaEval` in handling of 1-d arrays and strings longer than 32,767 characters. May return data of types that cannot be displayed on a worksheet, such as a dictionary or an array of arrays.
```
jl.JuliaEvalVBA(JuliaExpression)
```

|Argument|Description|
|:-------|:----------|
|`JuliaExpression`|Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.|

**Example.** `Set d = jl.JuliaEvalVBA("Dict(w => length(w) for w in split(""the quick brown fox""))")` returns a `Scripting.Dictionary` mapping each word to its length, so `d("quick")` is 5. A dictionary can't be returned to a worksheet.

## [`JuliaCallVBA`](@id juliacallvba)
Call a named Julia function from VBA. Differs from `JuliaCall` in handling of 1-d arrays and strings longer than 32,767 characters. May return data of types that cannot be displayed on a worksheet, such as a dictionary, an array of arrays, or arrays of dimension up to 9.
```
jl.JuliaCallVBA(JuliaFunction, [Arg1], [Arg2], ...)
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function visible from the Julia REPL.|
|`Args...`|Zero or more arguments. Each may be a number, string, Boolean, empty cell, array or range. A range's values are passed as Excel stores them (its `Value2`), so cells formatted as dates arrive as serial numbers, not as Julia `Date`s.|

**Example.** `words = jl.JuliaCallVBA("split", "the quick brown fox")` returns a one-dimensional VBA array, `words(1)` to `words(4)`.

## [`JuliaCallNamedArgsVBA`](@id juliacallnamedargsvba)
Call a named Julia function from VBA, passing positional and named arguments, with `";"` marking where the named arguments start. Differs from `JuliaCallNamedArgs` in the same ways that `JuliaCallVBA` differs from `JuliaCall`.
```
jl.JuliaCallNamedArgsVBA(JuliaFunction, [Arg1], ..., ";", Name1, Value1, ...)
```

|Argument|Description|
|:-------|:----------|
|`JuliaFunction`|The name of a Julia function visible from the Julia REPL.|
|`Args...`|Zero or more positional arguments, then `";"`, then pairs of name and value, e.g. `"a",1,"b",2`. Names must be valid Julia identifiers. Each argument may be a number, string, Boolean, empty cell, array or range.|

**Example.** `parts = jl.JuliaCallNamedArgsVBA("split", "a,b,,c", ",", ";", "keepempty", False)` returns `a`, `b` and `c`, without the empty string between the two commas.

## [`JuliaEvalHandle`](@id juliaevalhandle)
Evaluate a Julia expression from VBA and return a handle to the result. The VBA counterpart of `JuliaEval(JuliaExpression, TRUE)`. The handle lasts until `JuliaHandleRelease` or `JuliaHandleReleaseAll` frees it.
```
jl.JuliaEvalHandle(JuliaExpression)
```

|Argument|Description|
|:-------|:----------|
|`JuliaExpression`|Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.|

**Example.**
```vba
odds = jl.JuliaEvalHandle("Set(1:2:999)")
Debug.Print jl.JuliaCall("in", 7, odds)   ' True
jl.JuliaHandleRelease odds
```

## [`JuliaHandleRelease`](@id juliahandlerelease)
Free the Julia object of a handle created from VBA. Returns TRUE if it freed one, and FALSE otherwise - when the handle was already freed, was created by a worksheet formula (which frees its own), or isn't a handle at all, such as an empty string. So it is always safe to call, for example from a class's `Class_Terminate` event.
```
jl.JuliaHandleRelease(Handle)
```

|Argument|Description|
|:-------|:----------|
|`Handle`|A handle returned by `jl.JuliaCallHandle` or `jl.JuliaEvalHandle`.|

## [`JuliaHandleReleaseAll`](@id juliahandlereleaseall)
Free the Julia objects of all handles created from VBA, and return how many. Handles created by worksheet formulas are unaffected.
```
jl.JuliaHandleReleaseAll()
```
