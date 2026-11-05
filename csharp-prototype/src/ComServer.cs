using System.Diagnostics;
using System.Runtime.InteropServices;

namespace JuliaExcelPrototype
{
    // Toolchain-validation prototype only: not wired to Julia.
    // Proves VBA can obtain a live object reference from this add-in via
    // CreateObject(ProgId) -- the mechanism JuliaCallVBA/JuliaEvalVBA would need,
    // since Application.Run cannot return a live object reference (only
    // worksheet-safe value types).

    [ComVisible(true)]
    [Guid("0dafe843-0c64-4de7-9ed2-e2b9e1f91fff")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJepComServer
    {
        string HelloFromCSharp(string name);
        double JepAdd(double a, double b);
        string EncodeV(object values);
        double[,] DecodeV(string wire);
        string EncodeVDiag(object values, out double methodBodyMs);
        double[,] DecodeVDiag(string wire, out double methodBodyMs);
        string EncodeGeneral(object values);
        object[,] DecodeGeneral(string wire);
        double[,] SortViaJulia(object values, string host, int port);
        double[,] SortViaJuliaDiag(object values, string host, int port, out double encodeMs, out double httpMs, out double decodeMs);
    }

    [ComVisible(true)]
    [Guid("3c9e8f53-0e94-478b-b232-0588bd08080c")]
    [ProgId("JuliaExcelPrototype.ComServer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJepComServer))]
    public class JepComServer : IJepComServer
    {
        public string HelloFromCSharp(string name)
        {
            return "Hello, " + name + "! (from C#/.NET via COM)";
        }

        // Same underlying method as the worksheet/Application.Run-facing
        // JepFunctions.JepAdd -- no argument translation needed, since double
        // marshals identically for both the XLL and COM Automation call paths.
        public double JepAdd(double a, double b)
        {
            return JepFunctions.JepAdd(a, b);
        }

        // Speed-comparison microbenchmark against VBA's TrySerialiseArrayAsV
        // (modSerialise.bas) -- see VFormat.cs for the wire-format details.
        public string EncodeV(object values)
        {
            return VFormat.EncodeV((object[,])values);
        }

        public double[,] DecodeV(string wire)
        {
            int nr, nc;
            double[] flat = VFormat.DecodeV(wire, out nr, out nc);
            var result = new double[nr, nc];
            int k = 0;
            for (int j = 0; j < nc; j++)
            {
                for (int i = 0; i < nr; i++)
                {
                    result[i, j] = flat[k];
                    k++;
                }
            }
            return result;
        }

        // Diagnostic twins of EncodeV/DecodeV: isolate time spent in the method body
        // (cast + VFormat logic) from time spent in the COM marshaling that happens
        // before the method is entered (the incoming VARIANT array -> object[,]
        // conversion) and after it returns (the outgoing string/array -> BSTR/SAFEARRAY
        // conversion) -- added after an initial VFormat benchmark showed only a 4-7x
        // speedup over VBA, far short of the ~15-30x expected from the algorithm change
        // alone, to find out how much of the gap is COM marshaling rather than the
        // hex encode/decode logic itself.
        public string EncodeVDiag(object values, out double methodBodyMs)
        {
            var sw = Stopwatch.StartNew();
            object[,] arr = (object[,])values;
            string result = VFormat.EncodeV(arr);
            sw.Stop();
            methodBodyMs = sw.Elapsed.TotalMilliseconds;
            return result;
        }

        public double[,] DecodeVDiag(string wire, out double methodBodyMs)
        {
            var sw = Stopwatch.StartNew();
            int nr, nc;
            double[] flat = VFormat.DecodeV(wire, out nr, out nc);
            var result = new double[nr, nc];
            int k = 0;
            for (int j = 0; j < nc; j++)
            {
                for (int i = 0; i < nr; i++)
                {
                    result[i, j] = flat[k];
                    k++;
                }
            }
            sw.Stop();
            methodBodyMs = sw.Elapsed.TotalMilliseconds;
            return result;
        }

        // Speed-comparison microbenchmark against VBA's SerialiseElement/Unserialise
        // general "*" per-element path (not the "V" fast path above) -- see
        // GeneralFormat.cs for the wire-format details. Exercises the path VBA uses for
        // Strings/mixed types, which has no bulk-CopyMemory equivalent and pays far more
        // per-element overhead in VBA than the all-Double "V" case does.
        public string EncodeGeneral(object values)
        {
            return GeneralFormat.EncodeGeneral((object[,])values);
        }

        public object[,] DecodeGeneral(string wire)
        {
            int nr, nc;
            return GeneralFormat.DecodeGeneral(wire, out nr, out nc);
        }

        // Full end-to-end equivalent of VBA's JuliaCall("sort", values) (modMain.bas) --
        // same outer "*1,2;<len1>,<len2>,;<fnNameEncoded><argEncoded>" envelope JuliaCall
        // itself builds (a function-name element plus one array argument element, the
        // array V-encoded since it's all-Double), POSTed to the already-running Julia
        // HTTP.jl server at /call exactly as JuliaHttpPost does, decoding the "V"-format
        // response the same way. Point of this: JuliaCall("sort", RANDARRAY(1000000)) in
        // VBA measures ~1273ms, of which VBA-side V-encode/decode account for ~1130ms (see
        // the VFormat speed benchmark) -- leaving ~143ms for HTTP + Julia's own sort. With
        // C#'s V-encode/decode (~134ms combined), the full round trip should plausibly
        // land under Excel's native SORT() (~667ms on the same data) -- this method
        // measures whether that actually holds.
        public double[,] SortViaJulia(object values, string host, int port)
        {
            string arrayEncoded = VFormat.EncodeV((object[,])values);
            string fnEncoded = ((char)163) + "sort"; // Chr(163) in VBA -- the String type-indicator tag
            string payload = "*1,2;" + fnEncoded.Length + "," + arrayEncoded.Length + ",;" + fnEncoded + arrayEncoded;

            string responseBody = JuliaHttpClient.Post(host, port, "/call", payload);

            int nr, nc;
            double[] flat = VFormat.DecodeV(responseBody, out nr, out nc);
            var result = new double[nr, nc];
            int k = 0;
            for (int j = 0; j < nc; j++)
            {
                for (int i = 0; i < nr; i++)
                {
                    result[i, j] = flat[k];
                    k++;
                }
            }
            return result;
        }

        public double[,] SortViaJuliaDiag(object values, string host, int port, out double encodeMs, out double httpMs, out double decodeMs)
        {
            var sw = Stopwatch.StartNew();
            string arrayEncoded = VFormat.EncodeV((object[,])values);
            string fnEncoded = ((char)163) + "sort";
            string payload = "*1,2;" + fnEncoded.Length + "," + arrayEncoded.Length + ",;" + fnEncoded + arrayEncoded;
            sw.Stop();
            encodeMs = sw.Elapsed.TotalMilliseconds;

            sw.Restart();
            string responseBody = JuliaHttpClient.Post(host, port, "/call", payload);
            sw.Stop();
            httpMs = sw.Elapsed.TotalMilliseconds;

            sw.Restart();
            int nr, nc;
            double[] flat = VFormat.DecodeV(responseBody, out nr, out nc);
            var result = new double[nr, nc];
            int k = 0;
            for (int j = 0; j < nc; j++)
            {
                for (int i = 0; i < nr; i++)
                {
                    result[i, j] = flat[k];
                    k++;
                }
            }
            sw.Stop();
            decodeMs = sw.Elapsed.TotalMilliseconds;
            return result;
        }
    }
}
