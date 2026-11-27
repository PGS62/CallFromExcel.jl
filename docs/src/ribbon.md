# The Julia ribbon tab

The add-in adds a Julia tab to Excel's ribbon:

![The Julia tab on Excel's ribbon, with one group, CallFromExcel.jl: Launch Julia, Launch Settings, Show Julia, Close Julia, Display Results, a Julia File box with Browse and Include, Register Functions, Rebuild Objects, Help and About](assets/images/ribbon.png)

|Control|What it does|
|:------|:-----------|
|**Launch Julia**|Starts Julia for this Excel session, using the Launch Settings.|
|**Launch Settings...**|Sets the project folder, other command line options, packages to load, whether to run under WSL, and so on. The settings are saved for you on this PC.|
|**Show Julia**|Brings Julia's window to the front.|
|**Close Julia**|Closes Julia, after asking for confirmation.|
|**Display Results**|Shows each call from Excel, and its result, in Julia's window - as `CallFromExcel.display_results(true)` does.|
|**File**, **Browse**, **Include**|Choose a Julia file - typed, pasted, browsed to or picked from recently used files - and include it, as `JuliaInclude` does. Include doesn't trigger recalculation.|
|**Register Functions**|Makes the Julia functions marked with `@excel` or `expose_to_excel` callable as worksheet functions - see [Your own worksheet functions](own_functions.md). That happens by itself after `JuliaLaunch`, `JuliaInclude` and `JuliaEval`; use the button for functions exposed from the Julia REPL.|
|**Rebuild Objects**|Rebuilds the Julia object of every [handle](handles.md) on the worksheets, so that they are built by your Julia code as it is now - for example after you change and re-include it. Excel then recalculates the cells that use handles, and no others. Its arrow offers **Rebuild Objects and Recalculate All**, which then recalculates every formula in every open workbook, as Ctrl+Alt+F9 does, so that `JuliaCall` and other formulas use the changed code too; that can be slow. In manual calculation mode, press F9 after Rebuild Objects. Handles created from VBA aren't rebuilt.|
|**Help**|Opens this documentation.|
|**About**|Shows the versions of the add-in, and of the CallFromExcel package and Julia in the Julia session it is connected to, and a link to this documentation.|

Launch Julia and a `JuliaLaunch` formula start Julia in the same way, but each with its own settings: a formula uses only its own arguments, so a workbook records how Julia should start and does so the same way on every PC, while the ribbon's Launch Settings are yours.
