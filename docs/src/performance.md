# Performance and alternatives

## Performance
Calls are fast: a simple call takes under a millisecond, and `=JuliaCall("sort", A1:A1000000)` sorts a million numbers in about 0.4 seconds, including sending them to Julia and back. That's faster than Excel's own `=SORT(A1:A1000000)`, which takes about 0.6 seconds, and about a third of the time taken by CallFromExcel's predecessor, JuliaExcel. See the [changelog](https://github.com/PGS62/CallFromExcel.jl/blob/main/CHANGELOG.md) for more timings.

Sending data is usually the larger cost, and the one you can avoid. If many formulas use the same large range, or a result that's expensive to build, keep it in Julia as a [handle](handles.md): it is sent or built once, and each formula then sends only the handle's short text. Handles are cheap: 2,000 cells each creating a handle calculate in under a second (0.5 to 0.9 seconds in tests), and 2,000 `JuliaCall`s using them in about 0.3 seconds.

## Alternatives
Apart from CallFromExcel's own predecessor, [JuliaExcel](https://github.com/PGS62/JuliaExcel.jl), there is one other method of calling Julia from Excel of which I am aware: [JuliaInXL.jl](https://github.com/JuliaComputing/JuliaInXL.jl).

JuliaComputing (now JuliaHub) made JuliaInXL open source in October 2021; it previously required a licence for commercial use. When last checked (November 2026), JuliaInXL did not support dynamic array formulas or calling Julia from VBA. Filling those gaps was why I wrote JuliaExcel, this project's predecessor, in 2021.

Early versions of JuliaExcel were slower than JuliaInXL, both in latency and in moving data between Excel and Julia. That has since reversed: in tests (see `workbooks/CallFromExcel-Compare.xlsm` in the repository), CallFromExcel was between 1.6 and 16 times faster. For example, `=jlcall("sort", A1:A1000000)` takes about 4.5 seconds, and `=JuliaCall("sort", A1:A1000000)` about 0.4 seconds.
