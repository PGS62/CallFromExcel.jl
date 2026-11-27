# Installation

Installation does not require admin rights on the PC.
 * Both Julia and Microsoft Office must be installed on your PC.
 * Close Excel - installation fails if Excel is running.
 * Launch Julia, and copy-paste the following commands into the REPL. To paste into the REPL, right-click.
   ```julia
   using Pkg
   Pkg.add("CallFromExcel")
   using CallFromExcel
   CallFromExcel.installme()
   ```
   `Pkg.add` installs the Julia package, and `installme()` installs the Excel add-in.
 * Launch Excel. Check that it shows a Julia tab on the ribbon, and that typing `=Julia` into a worksheet cell makes the auto-complete offer `JuliaCall`, `JuliaEval`, `JuliaInclude` etc.
 * After updating CallFromExcel (e.g. with `Pkg.update()`), close Excel and run `using CallFromExcel; CallFromExcel.installme()` again: the add-in and the Julia package must be the same version, and if they aren't, Excel says so. Do the same if the functions go missing from Excel.
 * To uninstall the Excel add-in, close Excel and find CallFromExcel in Windows' Settings > Apps > Installed apps. This also removes the ribbon's saved settings. The Julia package is separate: remove it with `using Pkg; Pkg.rm("CallFromExcel")`.

## [Installation (Linux)](@id installation-linux)
As an alternative to calling Julia running in a Windows process, CallFromExcel can call Julia running in a Linux process under Windows Subsystem for Linux (WSL). If that's your preference, then in addition to the steps described above, you need to:
 * If necessary, install [WSL](https://docs.microsoft.com/en-us/windows/wsl/install) with the default Linux distribution, Ubuntu.
 * Install Julia under WSL by running `curl -fsSL https://install.julialang.org | sh` at the Ubuntu prompt, per the [official instructions](https://julialang.org/downloads/). This installs `juliaup`, which puts `julia` on the PATH automatically.
 * At the Julia prompt (under WSL) install CallFromExcel by copy-pasting `using Pkg; Pkg.add("CallFromExcel")` into the REPL.

## [Upgrading from JuliaExcel](@id upgrading-from-juliaexcel)
CallFromExcel replaces [JuliaExcel](https://github.com/PGS62/JuliaExcel.jl). Its worksheet functions have the same names, but some changes are needed to workbooks and code written for JuliaExcel.

**Installation.** Installing CallFromExcel (see above) also uninstalls JuliaExcel's Excel add-in, `JuliaExcel.xlam`. The JuliaExcel Julia package is no longer used, and can be removed with `using Pkg; Pkg.rm("JuliaExcel")`.

**Worksheet formulas.** In a workbook saved with JuliaExcel, Excel records the location of `JuliaExcel.xlam` in each formula that uses its functions, so after upgrading a formula such as `=JuliaLaunch()` appears as `='C:\Users\Public\JuliaExcel\JuliaExcel.xlam'!JuliaLaunch()`. Remove that path with Find and Replace (Ctrl+H): set "Within" to Workbook and "Look in" to Formulas, and replace the text `'C:\Users\Public\JuliaExcel\JuliaExcel.xlam'!` with nothing.

**VBA.**
 * Replace the reference to JuliaExcel with one to CallFromExcel (VBA editor, Tools menu -> References).
 * Create an object to call CallFromExcel's functions through, and put it in front of each call, so `JuliaCall("f", x)` becomes `jl.JuliaCall("f", x)` - see [Calling from VBA](vba.md):
   ```vba
   Dim jl As New CallFromExcel.ComServer
   ```
 * From VBA, `JuliaCall`, `JuliaCallVBA`, `JuliaCallNamedArgs` and `JuliaCallNamedArgsVBA` take at most 30 arguments after the function name.
 * A call that fails raises an error, rather than returning a string starting with `#` as JuliaExcel's functions did. Code that tests results for that, such as `If Left(Result, 1) = "#" Then`, no longer catches anything: use `On Error` instead - see [Errors](vba.md#Errors).

**Julia.** Functions and variables formerly accessed as `JuliaExcel.xxx`, such as `JuliaExcel.serve_xl()` and `JuliaExcel.last_answer`, are now `CallFromExcel.xxx`, after `using CallFromExcel`.

## [Compatibility](@id compatibility)
CallFromExcel requires 64-bit Excel for Windows - the default for Microsoft 365 - and has been tested on Excel under Microsoft 365. It is not compatible with 32-bit Excel: the installer checks, and stops with an error if Excel is 32-bit. It requires Excel 2016 or later, and _should_ work on 64-bit Excel 2016, 2019 and 2021, but has not been tested on them.
