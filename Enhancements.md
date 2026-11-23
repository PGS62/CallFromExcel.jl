# Potential enhancements

Ideas considered but not yet scheduled, each with a sketch of how it might be implemented.

## Named worksheet functions

**Problem.** Every Julia function is called through `=JuliaCall("price_swap", A1, B1)`. A demanding user wants `=PRICE_SWAP(A1, B1)`: a real worksheet function with its own name, argument names in the Function Wizard and IntelliSense, and help text. In-house bank add-ins are typically built this way. Probably the feature such users would notice most.

**Sketch.** The calling path is unchanged: a named function forwards to the same code as `JuliaCall`, with the same wire format. What's new is registration.

- **Julia side (small).** A macro or function marks what to expose, e.g. `@excel function price_swap(notional, fixed_rate, curve) ... end` with a docstring, or `CallFromExcel.expose(price_swap; name=..., description=..., args=...)`. It records the Excel name, argument names (from the method, `Base.method_argnames`), descriptions (from the docstring, kept within Excel-DNA's 253/255-character limits) and flags such as volatile in a registry. A new HTTP endpoint returns that list.
- **Add-in side (moderate).** After `JuliaLaunch`, and on request (a ribbon button, or after `JuliaInclude`/`JuliaEval` change the list), the add-in fetches the list and calls `ExcelIntegration.RegisterDelegates`. Each delegate takes the right number of `object` parameters (built with expression trees) and in effect calls `JuliaCall("price_swap", args...)`. Registration must run in a macro context (`ExcelAsyncUtil.QueueAsMacro`). Registering a name again replaces it, so redefining a function in Julia and refreshing works.
- Free once registered: Function Wizard and IntelliSense help, and `Application.Run "PRICE_SWAP"` from VBA.

**Broadcasting over broadcastable arguments.** `JuliaCall("f.", ...)` broadcasts over every argument; in Julia, a function `f(a, b, c)` whose `c` is an array is broadcast over `a` and `b` alone as `f.(a, b, Ref(c))`. Named functions can do that automatically, so the Julia author writes `f` for scalars and ranges are handled correctly and consistently, with no `.` for the user to remember.

- **Which arguments broadcast.** Declared at registration, e.g. `@excel broadcast=(:a, :b) function f(a, b, c) ... end`: the listed arguments broadcast, the rest are passed whole. This suits the common style of leaving arguments unannotated unless needed for dispatch. Where there is no list, declared types decide: an argument declared as an array type (`::AbstractVector`, `::AbstractMatrix`) is passed whole, any other argument broadcasts. So without a list or annotations everything broadcasts, as `JuliaCall("f.", ...)` does today. `broadcast=false` opts out entirely.
- **The rule per call.** An array given to a broadcasting argument is broadcast over; every other value is wrapped in `Ref` (harmless for scalars, and it protects structs that don't support broadcasting). If nothing is broadcast, it's a plain call, so the result isn't a zero-dimensional array. A column of handles given to a broadcasting argument (see Object handles) broadcasts over the objects.
- **In Julia, not by Excel's lifting.** Excel can lift an add-in function over a range itself if its parameters are registered as scalar types, but it calls the function once per cell: one HTTP round trip each. Registering every parameter as a generic value and broadcasting in Julia makes one round trip and one compiled broadcast.
- **Consistent where Excel isn't.** Excel lifts a built-in function only when its parameter is registered as a scalar value, and many aren't: `=SQRT(A1:A3)` spills, but `=COMPLEX(A1:A3, B1:B3)` and `=EDATE(A1:A3, 1)` (Analysis ToolPak functions, whose parameters take references) give `#VALUE!` unless coerced with `+0`. Named Julia functions follow one rule; a surprised user can read the answer off the registration.
- **Conversions for arguments passed whole.** A single cell arrives as a scalar and an `n x 1` range as a `Matrix`; for an argument declared `::AbstractVector`, wrap or `vec` it, so the Julia author never sees the difference. With no declared type, pass what arrives.
- **Shapes.** Julia's rules match Excel's for compatible shapes: a column against a row gives a table, as `=A1:A10*B1:F1` does. For incompatible shapes (10 rows against 7) Excel pads with `#N/A`; don't copy that - give a clear error naming the two arguments.
- **Functions returning arrays.** Broadcasting them gives an array of arrays, which can't spill: a clear error, suggesting `broadcast=false`.
- **Several methods.** The method marked `@excel` defines the registered signature.
- **Optional arguments.** Trailing Julia defaults become optional Excel arguments, arriving as missing - separate from broadcasting, but part of the same registration.
- **Possible bonus for `JuliaCall`.** `JuliaCall("f.", ...)` could follow the same rule for a function with a single method whose declared types say which arguments are arrays. A change of behaviour, so separate and optional.

**Decisions to make.**

- **Workbooks opened before Julia is running.** Cells with `=PRICE_SWAP(...)` show `#NAME?`, and may not recover when the function is registered later. Add-ins that start their language runtime when Excel starts mostly avoid this; we launch on demand. Options: cache the last exposed list (per project folder) and register **stubs** when the add-in loads, which return "Julia not launched" until Julia is running and then forward normally; and/or a "launch Julia when Excel starts" setting.
- **Removing functions.** Excel-DNA can't cleanly unregister, so a removed function stays as a stub returning an error until Excel restarts. Document it.
- **Name clashes** with Excel's built-in functions or other add-ins (Excel's win). Warn, or offer an optional prefix.
- **Several projects or sessions.** Which set applies after relaunching into another environment - probably whatever the current session exposes, plus cached stubs.
- **IntelliSense** should pick up functions registered after load; confirm.

**Effort.** Roughly 1-2 weeks: the Julia macro and endpoint about a day, dynamic registration a few days, the rest on stubs and caching, `#NAME?` behaviour, tests and docs. Purely additive: `JuliaCall` users are unaffected.

**First step: a spike** (about a day). Register one function from Julia at run time and check the Function Wizard, IntelliSense, and what a saved workbook using it does when reopened before and after launch.

**Test.** Julia tests for the registry and endpoint (argument names, descriptions, length limits). VBA tests that call a registered function on a sheet and via `Application.Run`, redefine it and check the new definition is used. Julia tests for broadcasting: the `broadcast=` list, classification by declared type, `Ref` wrapping, plain call when nothing broadcasts, a column against a row, the shape-mismatch and array-of-arrays errors, the scalar-to-vector and `n x 1` conversions.

## Object handles

**Problem.** Every result comes back to Excel as data. A large or structured Julia object (a curve, a model, a DataFrame, a portfolio) must be rebuilt on every call that needs it, or be kept in a global by hand (`JuliaEval("global curve = ...")`), which doesn't recalculate with the sheet. The handle pattern common in C++ and other Excel add-ins solves this: a function returns a short string naming an object kept in the add-in's language, and other functions accept that string in place of the object. Combines well with named worksheet functions: `=BUILD_CURVE(...)` returns a handle, `=PRICE(curve, trade)` uses it.

The examples below come from finance, where Excel is the working environment for building models, not just a place to view results, and where the handle pattern is long established. The same need arises wherever models are built in workbooks (insurance, energy, engineering, operations): any object that is large or expensive to build once and then used from many cells, such as a fitted model, a loaded dataset, a simulation's results or an interpolator.

**Sketch.**

- **Creating.** Two functions, mirroring `JuliaCall` and `JuliaEval`: `=JuliaHandle("build_curve", A1:B20)` calls a function as `JuliaCall` does, and `=JuliaEvalHandle("(1:100).*(1:100)'")` evaluates code as `JuliaEval` does. Each stores the result Julia-side in a `Dict` (not as globals) and returns a handle such as `jl:Curve#17@k3x9`. Rejected: a blank first argument, `=JuliaHandle(,"code")`, meaning "evaluate code" - hard to discover, and the Function Arguments dialog would label the code as an argument to `f`.
- **Using.** Every function accepts handles, with nothing to declare. In `decode_from_xl`'s string branch, a string that exactly matches the handle pattern is replaced by the cached object. So:
  - The add-in doesn't change: to Excel and C# a handle is just a string.
  - Handles work wherever values do: single cells, ranges, nested arrays, `JuliaCallKw`'s named arguments, VBA arrays.
  - `_maybe_typed` narrows a range of `Trade` handles to a `Vector{Trade}`, not `Vector{Any}`.
  - Ordinary strings cost one `startswith` check; numeric arrays ("V" format) never reach that branch.
  - Any Julia function works, and dispatch picks the method: `=JuliaCall("length", A1)`, `=JuliaCall("price", A1, B1)` calling `price(curve::Curve, ...)`.
- **Not found.** A string matching the pattern whose object isn't cached (freed, or from an earlier Julia session) raises an error such as `Handle jl:Curve#17@k3x9 not found - recalculate the cell that created it`, rather than passing the string on, which would end in a baffling `MethodError`. The cost: a genuine string of exactly that form can't be passed as a value.
- **Same object, not a copy.** A 2 GB DataFrame used by 50 cells is stored once. But a function that mutates its argument (`sort!`, `push!`) changes the cached object without changing its handle, so dependent cells aren't recalculated and results depend on calculation order. Document "don't mutate" (Julia's `!` convention helps); copying on every use would defeat the purpose.
- **Code can't take handles.** `JuliaEvalHandle` receives only a string, so to build an object from other handles use the call form, e.g. `=JuliaHandle("build_portfolio", A1:A10)` with trade handles in A1:A10. A `handle("jl:...")` lookup for use inside code could be added if asked for.
- **Reading.** `=JuliaCall("identity", D1)` already returns the object as data; a `JuliaObject(handle)` function would be a friendlier name for it.

**Handle format: `jl:<label>#<id>@<session>`.**

- **Label.** Only for people reading the sheet - lookup ignores it, so it can't cause a wrong match. Default `string(nameof(typeof(x)))`, which is always short: `Curve` rather than `MyPkg.Internal.Curve`, `Dict` rather than `Dict{String, Vector{Float64}}` (though `Array` rather than `Matrix{Float64}`). Closure types contain `#`, so parse from the right (session after the last `@`, id after the last `#` before it), or strip labels to letters, digits and `_`. Users could customise labels through dispatch rather than a setting, e.g. `CallFromExcel.handle_label(c::Curve) = "Curve_" * c.currency` - add only if asked for.
- **Id.** Increases with every handle created and is never reused, so a recalculated cell returns a new string (making Excel recalculate its dependents), and a freed handle can never come to name a different object.
- **Session.** A few random characters chosen when the server starts. Without them, ids restart at 1 when Julia restarts, so a stale handle in a reopened workbook could silently pick up a new, unrelated object #17 - a wrong answer with no error. With them, a stale handle always gives "not found".
- **Rejected:** a non-printing marker (e.g. a zero width space) to make accidental matches less likely. A handle typed by hand (in a formula, in VBA, while debugging) would silently not be one while looking identical; some paste paths and tools strip such characters; `LEN` and `=A1=B1` would puzzle. And an accidental exact match of this pattern is already implausible.

**Lifetime (most of the work).**

**Preferred: Excel-DNA's built-in object handles.** Excel-DNA (which the add-in uses, with automatic registration) has built-in handles since 1.9: `[return: ExcelHandle]` on a function's result, `[ExcelHandle]` on a parameter. Lifetime rests on RTD: each handle is an RTD topic, and Excel itself tells the add-in when the last formula using a topic has gone, at which point Excel-DNA calls `Dispose` on the stored object (`HandleInfo.Dispose` -> `ObjectHandler.Remove`, seen in its source). Excel knows its own formulas better than any scan, so:

- Moved cells, inserted rows, renamed sheets, Save As, closed workbooks, copied cells: all tracked by Excel.
- Nested `=JuliaCall("price", JuliaHandle(...))` and concatenation: the object lives while the formula exists - harmless.
- No scan, so the scan-cost and `SpecialCells` questions below disappear.

How it would fit: the C# object stored is a small proxy holding the Julia id (and session tag), whose `Dispose` tells Julia to drop the object - queued and batched, as `Dispose` comes from Excel's RTD machinery rather than a calculation. `JuliaCall` takes untyped arguments, so `[ExcelHandle]` parameter conversion doesn't apply; instead the add-in looks each string argument up (`ObjectHandler.TryGetObject`) and sends a found handle to Julia as a new wire-format type (e.g. `O<id>@<session>`), which the decoder replaces with the object. Julia then needs no string-pattern matching at all, and look-alike strings can't be mistaken for handles - only strings Excel-DNA actually issued are. The handle text is Excel-DNA's `<tag>:<n>`; the tag can be set (a custom `DisplayName` was added in early 2026, probably needing 1.10), so the type label could still appear. A Julia restarted while Excel stays open leaves proxies naming objects that no longer exist: the session tag turns those into the "not found" error.

Status: live and maintained, not an abandoned experiment. In Excel-DNA's core (`ExcelDna.Integration/ObjectHandles`), with runtime tests and NativeAOT support; commits through 2026 (custom display name Feb, more than 16 arguments Apr, AOT fix Jul); issue #824 (handles from user code) closed Feb 2026; issue #860 (nested handle stops an RTD-based async function restarting) open, Oct 2026 - relevant to the nested case, though `JuliaCall` isn't async. RTD has its own quirks to check, e.g. issue #209 (RTD functions recalculated on cut and paste anywhere in a sheet).

**Fallback: keep what's reachable.** If the Excel-DNA route fails the spike. The rule, as for Julia's own garbage collector: an object stays alive while some cell in an open workbook **displays its handle exactly**; everything else is freed. After each calculation (Excel's `AfterCalculate` event, which Excel-DNA can hook - nothing is half-calculated then, and no calculation runs during the scan) the add-in collects the handle strings visible in open workbooks and sends them to Julia (a new endpoint), which drops every other entry. It doesn't matter how formulas are written, only which handles are on show:

- **Nested**, `=JuliaCall("price", JuliaHandle(...))`: the object has done its job once the price is calculated; nothing can refer to it later, so freeing it is correct.
- **Concatenated**, `=JuliaHandle(...)&JuliaHandle(...)`: the cell's string isn't a handle (exact match), so nothing can use either object; freeing them is correct.
- **Moved cells, inserted rows, renamed sheets, Save As**: the scan reads values, not addresses, so none of these matter.
- **Copied cells**: both show the handle, so it's kept; each gets its own object when it next recalculates.
- **Closed workbooks**: their cells drop out of the next scan. No special case.
- **Manual calculation**: a cell awaiting recalculation still shows its old handle, which keeps the object its dependents still use.
- **Spilled arrays of handles**: each spilled cell shows a handle and is scanned like any other (but see the `SpecialCells` caveat below).

Rejected approaches:

- **Keying entries by calling cell** (`xlfCaller`, workbook/sheet/address), replacing a cell's entry when it recalculates. Addresses change under routine edits without any recalculation: inserting a row moves every cell below it, as do cut and paste, renaming a sheet and Save As. A new handle formula at a vacated address would then free another cell's object, and entries for moved cells would leak.
- **Asking the cell** (`xlCoerce` on the calling cell to read the handle it is about to replace). Assumes one visible handle per cell: `=JuliaCall("price", JuliaHandle(...))` would leak an object on every recalculation, and concatenation, `IF`, `LET`, `LAMBDA` and named formulas all break it.

Open points:

- **Scan cost.** Skip the sweep when no handle has been created since the last one. Two ways to scan, to measure on a big workbook:
  - `SpecialCells(xlCellTypeFormulas, xlTextValues)` per sheet: formula cells **and** text result (the second argument filters the first; within it, `xlNumbers + xlTextValues` would be an OR). Raises error 1004 "No cells were found" rather than returning nothing, so each call needs a catch. Has historically failed on some protected sheets. **Unknown whether spilled cells count** - only the top-left cell holds the formula; if they don't, a spilled column of handles would be freed while in use. Test this first.
  - Fallback: each sheet's `UsedRange.Value2` read as one array, checking strings with `startswith("jl:")`. Slower on a huge sheet, but misses nothing.
- **Handles created from VBA** are invisible to the scan. Exclude them from sweeping and release them explicitly (`JuliaRelease`), or have VBA store them in a cell.
- **Defined names** holding a handle as a constant: rare, cheap to include in the scan.
- **Julia restarted, workbooks reopened.** Every saved handle gives "not found", and Excel won't recalculate on open unless made to; `JuliaLaunch` could offer to recalculate open workbooks.
- **Safety valves.** `JuliaClearHandles()` or a ribbon button; the handle count (and perhaps memory) in the About dialog; a REPL function listing the cache (handle, label, size) - what a user would send when a complicated workbook misbehaves.

Design aim: since no test suite can cover every workbook, a case we missed should end in a clear error or a leak, never a silent wrong answer.

**Decisions to make.** Whether named worksheet functions can declare that they return a handle; whether a function returning an array of objects gives one handle or a spilled range of handles (one per element); whether handles from one workbook may be used in another (probably yes, and the sweep allows it).

**Test.** Most of the logic can be tested in Julia without Excel: the pattern (look-alikes, labels containing `#`, the session tag), replacement inside arrays and the typed result, the not-found error, the sweep as a pure function (cache plus live set gives what remains), labels. In Excel, a few VBA scenarios, each checking the handle count after calculating: a handle used from another cell, with dependents recalculating when it's recalculated; nested; concatenated; a moved cell and an inserted row; a copied cell; a closed workbook; manual calculation; spilled handles. The complexity of users' workbooks doesn't multiply these cases, because the rule depends only on what is visible. No random or stress tests.

**First step: a spike of the Excel-DNA route.** A `JuliaHandle` returning a disposable proxy via `[return: ExcelHandle]`; check that `Dispose` is called on recalculation with new inputs, cell deletion, formula edits, closing the workbook, and for nested and spilled handles; that moves, inserted rows and copies keep handles alive; whether `DisplayName` needs Excel-DNA 1.10; and what RTD costs with thousands of handle cells. Only if that fails: hook `AfterCalculate` for the sweep, scan a large workbook both ways and time it, and check whether `SpecialCells` reports spilled cells.

**Spike results (2026-11-23, branch `spike/object-handles`).** The RTD route works; no need for the sweep.

- *Mechanism.* Excel-DNA 1.9.0's built-in handles are a dictionary over the public `ExcelAsyncUtil.Observe` (RTD). Their handle text is `<function name>:<n>`; `DisplayName` exists only on master, unreleased. And `ObjectHandler.GetHandle(name, params, obj)` takes an already-built object, so it would build one per recalculation and leak it. The spike calls `ExcelAsyncUtil.Observe` itself with its own observable: `Subscribe` asks Julia to build the object (only for a new topic), disposing the subscription tells Julia to drop it. Same lifetime machinery, our own handle text (`jl:Array#1@k907`). Code: `src/handles.jl`, wire type `O<id>@<session>` in `decode.jl`, `JuliaHandles.cs`; scripts in `dotnet-addin/spike-handles/`.
- *Works as hoped.* New inputs, an edited formula and a cleared cell each create the new object and release the old one at once. Cut/paste and inserted rows: no events, dependents follow. Nested `=JuliaCall("length",JuliaHandle(...))`: the object lives while the formula exists. A range of handles arrives as `Vector{Vector{Float64}}`. A failed create (e.g. before `JuliaLaunch`) is retried on the next recalculation. The handle appears on the first calculation (no `#N/A` first).
- *Cost.* 2,000 handle cells: create 457 ms, 2,000 `JuliaCall`s using them 291 ms, `CalculateFull` 395 ms (no Julia calls), clear 231 ms.
- *Topics are keyed by function and arguments, not by cell.* Two cells with the same formula, or a copied cell, share one object, released when the last goes. `CalculateFull` doesn't rebuild: `=JuliaHandle("rand",3)` keeps its numbers. Arguably right ("same inputs, same object"); document it, and offer a rebuild command (below) for when it isn't.
- *Closing a workbook* doesn't release its handles at once, nor on `Calculate` or an edit elsewhere; only at the next `CalculateFull`. A bounded leak; probably Excel disconnecting RTD topics lazily.
- *Julia restarted while Excel stays open* - the one real problem found, and fixed in the spike. The topic keeps its dead handle while its arguments are unchanged, so re-entering or recalculating the formula doesn't help; nor does completing the topic (Excel-DNA then returns its final value for as long as the arguments are unchanged). Fix: a generation number in every topic key, and a command (`JuliaHandleInvalidateAll` in the spike) that bumps it and pushes `#N/A` into each live topic. The RTD update makes Excel recalculate those cells, which ask for a topic under the new generation and so rebuild their objects in the new Julia, in Excel's dependency order (a handle built from another handle comes after it). `JuliaHandle` returns an error argument unchanged, so nothing is built from an upstream `#N/A`. To do: call it from `JuliaLaunch` when it starts a new session; the same command could serve as "Rebuild all objects".
- *From VBA* (`Application.Run`, `Evaluate`): returns a handle, but it's released at once, since no cell holds the topic. Handles for VBA need explicit lifetime (`JuliaRelease`), as expected.
- *Not tested.* A function returning a spilled range of handles (the spike makes one handle per call); thousands of handles over hours of editing; Excel-DNA issue #209 (cut and paste recalculating RTD functions).

## Untrusted workbooks can run Julia code

**Problem.** CallFromExcel's worksheet functions run arbitrary Julia code, but a workbook that uses them needs no macros, so Excel opens it without any macro-security warning. A workbook received from someone else containing `=JuliaLaunch()` and `=JuliaEval("...")` runs that code, with the user's permissions, as soon as it calculates (once the user clicks "Enable Editing", if it opened in Protected View). This is inherent to any add-in whose functions execute code, and JuliaExcel has the same exposure.

**Rejected:** asking with a dialog, once per workbook, before running anything - the user doesn't want dialogs.

**Sketches (not yet evaluated):**

- **Trusted folders setting.** A setting (e.g. a registry value or a config file next to the installed add-in) listing folders whose workbooks may call Julia, analogous to Excel's Trusted Locations. Formulas in other workbooks return an error such as `#Workbook not in a CallFromExcel trusted folder` instead of running. The add-in finds the calling workbook via `XlCall.Excel(XlCall.xlfCaller)` and its sheet reference. Needs a default that doesn't break existing users (e.g. no setting = allow everything, so it's opt-in), which limits its value.
- **Use Excel's own Trusted Locations.** Allow calls only from workbooks in a folder Excel already trusts (read from the registry under `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations`). No new setting, but most users' workbooks aren't in a trusted location, so this would need to be opt-in too.
- **Launch only by the user.** `JuliaLaunch` as a worksheet formula is what lets a workbook start Julia by itself; without it, a workbook can only use a Julia session the user started. An option could make the formula version refuse, leaving launching to VBA, a ribbon button, or `serve_xl`. Reduces exposure only while no Julia session is running.

**Test.** Depends on the approach - e.g. a workbook outside the trusted folders gets the error and Julia receives nothing.

## Document running as administrator and Constrained Language Mode

**Excel run as administrator.** Not tested; from Windows' documented behaviour:

- Worksheet functions should work: the add-in is registered in the user's own Excel settings (`HKCU\...\Excel\Options`), which an elevated Excel run by the same user reads. A Julia started by `JuliaLaunch` is elevated too.
- VBA's `CreateObject("CallFromExcel.ComServer")` probably fails ("ActiveX component can't create object"): `Install.ps1` registers the COM class per user (`HKCU\Software\Classes`), and Windows ignores per-user COM registrations in elevated processes.
- Elevating as a different account (an admin account's password typed in) runs Excel as that user, for whom the add-in isn't installed at all.
- Quick check: open Excel as administrator and run `?TypeName(CreateObject("CallFromExcel.ComServer"))` in the Immediate window.

**PowerShell Constrained Language Mode (CLM).** Where AppLocker or WDAC locks PowerShell down, scripts from user-writable folders such as `~\.julia\packages\...` run in CLM, which blocks COM objects and most .NET types (`-ExecutionPolicy Bypass` doesn't get round it).

- `CallFromExcel.installme()` runs `dotnet-addin\installer\Install.ps1`, which uses `New-Object -ComObject Excel.Application` and `[System.IO.File]::OpenRead` with a `BinaryReader` (to find Office's version and Excel's bitness) - these fail with "Cannot create type. Only core types are supported in this language mode", so installation fails. Its registry writes use plain cmdlets, which CLM allows.
- Day-to-day use doesn't involve PowerShell, except `serve_xl` under WSL, which asks `powershell.exe` for `[IO.Path]::GetTempPath()`.

**Sketches:**

- **README section** "Corporate and locked-down machines": worksheet functions work in an elevated Excel but VBA's `CreateObject` doesn't (don't run Excel elevated, or install as that account); and, for when the installer fails under CLM, what it does - copy the add-in files to `%APPDATA%\Microsoft\AddIns\CallFromExcel` and set the listed registry values - so IT can apply it by hand.
- **Make Install.ps1 CLM-safe:** use only cmdlets - read Office's version and Excel's bitness from the registry rather than via COM and the executable's header. A moderate rewrite of the installer; testing needs CLM enabled on a machine (e.g. `__PSLockdownPolicy` or a WDAC policy in a VM).