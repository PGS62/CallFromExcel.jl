using System.Diagnostics;
using ExcelDna.Integration;

namespace JuliaExcel.AddIn
{
    // Diagnostic XLL functions kept for future performance investigation -- not part of
    // the production API. Used to fully explain why the real, end-to-end
    // JuliaCall("sort", RANDARRAY(1000000)) measured ~720-770ms (via both Application.Run
    // and genuine worksheet-formula recalculation, so not a COM-specific artifact) when
    // the wire-format work alone (encode+HTTP+decode, timed internally) only accounts for
    // ~310-340ms of that:
    //
    //   JuliaDiagInputOnly  (receive a 1,000,000-element array, return a tiny result):
    //     ~237ms -- Excel-DNA's own XLOPER->CLR marshaling of the incoming array into a
    //     boxed object[,], which happens before the function body even starts running.
    //   JuliaDiagOutputOnly (receive a tiny argument, return a 1,000,000-element array):
    //     ~200ms -- the same marshaling in the opposite direction (CLR array -> XLOPER),
    //     after the function body returns.
    //
    // 237 + 200 = 437ms, which (plus the ~310-340ms wire-format cost) fully accounts for
    // the observed ~720-770ms total. This is inherent to Excel-DNA's generic `object`
    // array marshaling for a function that must accept/return arbitrary types (JuliaCall's
    // args can be numbers, strings, arrays, etc., so a more specific parameter type like
    // double[,] isn't an option) -- not something this project's own code can optimize
    // away, and not specific to this implementation: any XLL add-in taking/returning a
    // generic-typed array this large would pay a comparable cost. Still a genuine ~1.7-1.9x
    // improvement over today's VBA JuliaCall (~1273-1473ms measured for the same case), but
    // this means the earlier csharp-prototype SortViaJulia demo's "beats Excel's native
    // SORT()" result doesn't hold for a real worksheet-formula JuliaCall -- that demo
    // returned a strongly-typed double[,] via a direct COM method call, a different (and
    // evidently cheaper) marshaling path than a genuine Excel-DNA XLL function return.
    public static class JuliaCallDiag
    {
        [ExcelFunction(IsHidden = true, Description = "Diagnostic: takes a large array argument but returns a trivial result -- isolates Excel-DNA's input-side XLOPER->CLR marshaling cost.")]
        public static object JuliaDiagInputOnly(object arg)
        {
            int count = (arg is object[,] a2) ? a2.Length : 1;
            return "received " + count + " elements (marshaling-in time is NOT included here - it happens before this function body starts)";
        }

        // Tested: double[,] (~210-220ms) vs object (~240ms) for the same 1,000,000-element
        // input -- only a modest ~10-15% improvement, not a path to eliminating the
        // marshaling cost. And not applicable to JuliaCall specifically anyway, since its
        // arguments must support arbitrary mixed types (numbers, strings, arrays) decided
        // at runtime, not a fixed numeric type declared at compile time. Confirms the
        // ~435ms is substantially inherent to Excel-DNA's own XLOPER-array-walking, not
        // mostly a boxing cost that a more specific type could skip.
        [ExcelFunction(IsHidden = true, Description = "Diagnostic: same as JuliaDiagInputOnly but with a double[,] parameter type instead of object -- tests whether a specific numeric type marshals faster than generic object.")]
        public static object JuliaDiagInputOnlyTyped(double[,] arg)
        {
            return "received " + arg.Length + " elements (typed double[,] parameter)";
        }

        [ExcelFunction(IsHidden = true, Description = "Diagnostic: takes a trivial argument but returns a large array -- isolates Excel-DNA's output-side CLR->XLOPER marshaling cost.")]
        public static object JuliaDiagOutputOnly(int n)
        {
            var result = new double[n, 1];
            for (int i = 0; i < n; i++) result[i, 0] = i;
            return result;
        }

        [ExcelFunction(IsHidden = true, Description = "Diagnostic: times JuliaCall's own encode/HTTP/decode phases (excludes Excel-DNA's XLOPER marshaling of the argument/result, which happens outside this function body).")]
        public static object JuliaCallDiagTimed(string juliaFunction, object arg)
        {
            var sw = Stopwatch.StartNew();
            string fnEncoded = WireFormatEncode.Encode(juliaFunction);
            string argEncoded = WireFormatEncode.Encode(arg);
            string payload = "*1,2;" + fnEncoded.Length + "," + argEncoded.Length + ",;" + fnEncoded + argEncoded;
            sw.Stop();
            double encodeMs = sw.Elapsed.TotalMilliseconds;

            sw.Restart();
            string responseWire = JuliaHttpClient.Post(payload, "/call");
            sw.Stop();
            double httpMs = sw.Elapsed.TotalMilliseconds;

            sw.Restart();
            WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: 32768, vectorToColumn: true);
            sw.Stop();
            double decodeMs = sw.Elapsed.TotalMilliseconds;

            return "encode=" + encodeMs.ToString("0.0") + "ms http=" + httpMs.ToString("0.0") + "ms decode=" + decodeMs.ToString("0.0") + "ms";
        }
    }
}
