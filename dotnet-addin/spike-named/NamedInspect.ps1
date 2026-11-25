# Named-functions spike: opens a VISIBLE Excel with the Debug XLL, launches Julia, exposes and
# registers jl.price_swap and jl.hypot, and leaves Excel open so the Function Wizard,
# IntelliSense and "=jl." autocomplete can be checked by eye. Close Excel when done (Julia then
# stops serving it; close its window too).
$ErrorActionPreference = "Stop"
$xll = "C:\Projects\CallFromExcel\dotnet-addin\CallFromExcel.AddIn\bin\Debug\net48\CallFromExcel.AddIn64.xll"
function Retry($sb) {
    for ($i = 0; $i -lt 100; $i++) {
        try { return & $sb } catch { if ($_.Exception.Message -notmatch "rejected|0x8001010A") { throw }; Start-Sleep -Milliseconds 200 }
    }
    throw "Excel stayed busy"
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
if (-not $excel.RegisterXLL($xll)) { throw "RegisterXLL failed" }
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
Retry { $ws.Range("A1").Value2 = "Spike: named worksheet functions. Julia was launched and jl.price_swap and jl.hypot registered by the cells in column Z." }
Retry { $ws.Range("Z1").Formula2 = '=JuliaLaunch(FALSE,TRUE,"--project=C:\Projects\CallFromExcel")' }
Retry { $ws.Range("Z2").Formula2 = '=JuliaEval({"price_swap(notional, fixed_rate, years) = notional * fixed_rate * years";"CallFromExcel.expose_to_excel(""price_swap""; description=""Toy swap price: notional x fixed rate x years."", argdescriptions=[""The notional amount."", ""The fixed rate, e.g. 0.05 for 5%."", ""The term in years.""])";"CallFromExcel.expose_to_excel(""hypot"")"})' }
Retry { [void]$excel.Run("JuliaRegisterFunctions") }
Retry { $ws.Range("A3").Formula2 = '=jl.price_swap(1000000, 0.05, 10)' }
Retry { $ws.Range("A4").Formula2 = '=jl.hypot(3, 4)' }
Retry { $ws.Range("A6").Value2 = "Try: type =jl. in a cell (autocomplete list); type =jl.price_swap( (IntelliSense tooltip); Insert Function (fx) with category Julia." }
Retry { $ws.Range("A3").Select() }
"Excel is open. A3 = " + (Retry { $ws.Range("A3").Text }) + ", A4 = " + (Retry { $ws.Range("A4").Text })
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
