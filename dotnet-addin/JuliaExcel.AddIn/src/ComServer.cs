using System.Runtime.InteropServices;

namespace JuliaExcel.AddIn
{
    // The VBA-facing COM-visible class for JuliaCallVBA/JuliaEvalVBA (step 5 of the
    // plan) -- needed because Application.Run/XLL function return types can't carry a
    // live object reference, so results JuliaCall/JuliaEval can't display on a worksheet
    // (Dictionaries, nested arrays) go through here instead. Also still carries the
    // JuliaLaunch/JuliaHttpClient test-harness methods used while validating those ports.
    [ComVisible(true)]
    [Guid("6a6e3e0a-6f3a-4b9a-8a2a-1d6f7c9a2b4e")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJuliaExcelComServer
    {
        string Launch(bool minimiseWindow, string commandLineOptions, string packages, int timeoutSeconds);
        string Post(string payload, string path);
        bool IsListening(int port);
        int GetPort();

        // Port of JuliaCallVBA/JuliaEvalVBA (modMain.bas): differ from JuliaCall/JuliaEval
        // in allowing nested arrays and Dictionary results (AllowNesting=true,
        // JuliaVectorToXLColumn=false, no string-length limit), matching how those are
        // only ever called from VBA, never bound to a worksheet cell.
        object JuliaCallVBA(string juliaFunction, params object[] args);
        object JuliaEvalVBA(object juliaExpression);
        object JuliaCallNamedArgsVBA(string juliaFunction, params object[] args);

        // Migration surface for VBA code that used to call these as bare Public Functions
        // in a referenced JuliaExcel.xlam (Tools > References -> JuliaExcel). Since the new
        // world has no xlam at all, old call sites become e.g.
        //   Result = JuliaExcel.JuliaCall("sqrt", 16)
        // with JuliaExcel declared once as CreateObject("JuliaExcel.AddIn.ComServer") (or the
        // early-bound equivalent after a Tools > References to this assembly's type library).
        // These intentionally keep the worksheet-style restrictions (string-length limit,
        // no nested arrays/Dictionaries, 1-D arrays become a column) that the original
        // JuliaCall/JuliaEval/JuliaCallNamedArgs/JuliaSetVar/JuliaInclude had even when called
        // from VBA - JuliaCallVBA/JuliaEvalVBA/JuliaCallNamedArgsVBA above remain the way to
        // lift those restrictions, exactly as before.
        object JuliaCall(string juliaFunction, params object[] args);
        object JuliaEval(object juliaExpression);
        object JuliaCallNamedArgs(string juliaFunction, params object[] args);
        object JuliaLaunch(bool useLinux = false, bool minimiseWindow = false, string commandLineOptions = "", string packages = "", string bashStatements = "", int timeOut = 30);
        bool JuliaIsRunning();
        int JuliaExcelPID();
        object JuliaSetVar(string variableName, object refersTo);
        object JuliaInclude(string fileName);
    }

    [ComVisible(true)]
    [Guid("9d7b5c1a-2e4f-4a8b-9c3d-7e1a5b6f8c2d")]
    [ProgId("JuliaExcel.AddIn.ComServer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJuliaExcelComServer))]
    public class ComServer : IJuliaExcelComServer
    {
        public string Launch(bool minimiseWindow, string commandLineOptions, string packages, int timeoutSeconds)
        {
            return global::JuliaExcel.AddIn.JuliaLaunch.Launch(minimiseWindow, commandLineOptions, packages, timeoutSeconds);
        }

        public string Post(string payload, string path)
        {
            return JuliaHttpClient.Post(payload, path);
        }

        public bool IsListening(int port)
        {
            return JuliaHttpClient.IsListening(port);
        }

        public int GetPort()
        {
            return JuliaPort.GetJuliaPort();
        }

        public object JuliaCallVBA(string juliaFunction, params object[] args)
        {
            return CallVBACore(juliaFunction, args, "/call");
        }

        public object JuliaCallNamedArgsVBA(string juliaFunction, params object[] args)
        {
            return CallVBACore(juliaFunction, args, "/callkw");
        }

        public object JuliaEvalVBA(object juliaExpression)
        {
            string expression = JuliaExpressionToString(juliaExpression);
            string responseWire = JuliaHttpClient.Post(expression, "/eval");
            return WireFormatDecode.Decode(responseWire, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false);
        }

        public object JuliaCall(string juliaFunction, params object[] args)
        {
            return CallCore(juliaFunction, args, "/call");
        }

        public object JuliaEval(object juliaExpression)
        {
            try
            {
                if (JuliaPort.GetJuliaPortOrZero() == 0)
                {
                    return "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!";
                }
                string expression = JuliaExpressionToString(juliaExpression);
                string responseWire = JuliaHttpClient.Post(expression, "/eval");
                return WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true);
            }
            catch (System.Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        public object JuliaCallNamedArgs(string juliaFunction, params object[] args)
        {
            return CallCore(juliaFunction, args, "/callkw");
        }

        public bool JuliaIsRunning()
        {
            int port = JuliaPort.GetJuliaPortOrZero();
            return port != 0 && JuliaHttpClient.IsListening(port);
        }

        public int JuliaExcelPID()
        {
            return System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        public object JuliaSetVar(string variableName, object refersTo)
        {
            return CallCore("JuliaExcel.setvar", new[] { (object)variableName, refersTo }, "/call");
        }

        public object JuliaInclude(string fileName)
        {
            return CallCore("JuliaExcel.include", new object[] { fileName.Replace("\\", "/") }, "/call");
        }

        // Unlike the other migration wrappers above, this can't just delegate to
        // JuliaFunctions.JuliaLaunch (its CommandLineOptions/Packages/TimeOut parameters are
        // typed `object` to detect ExcelMissing, an XLL-only concept). UseLinux/BashStatements
        // (WSL launch) are not yet supported by this port - old VBA code passing UseLinux:=True
        // gets a clear error here rather than silently launching on Windows instead.
        public object JuliaLaunch(bool useLinux = false, bool minimiseWindow = false, string commandLineOptions = "", string packages = "", string bashStatements = "", int timeOut = 30)
        {
            if (useLinux || !string.IsNullOrEmpty(bashStatements))
            {
                throw new System.NotSupportedException("JuliaLaunch's UseLinux/BashStatements (launching Julia under WSL) is not yet supported by the C# port of JuliaExcel - launch Julia natively on Windows instead.");
            }
            try
            {
                return global::JuliaExcel.AddIn.JuliaLaunch.Launch(minimiseWindow: minimiseWindow, commandLineOptions: commandLineOptions, packages: packages, timeoutSeconds: timeOut);
            }
            catch (System.Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        // Fixed at the modern (Excel 2016+) limit rather than querying Excel's own version
        // (as JuliaFunctions.StringLengthLimit does via ExcelDnaUtil.ExcelVersion) -- that
        // would need ExcelDna.Integration.dll loadable, which this COM-only code path must
        // not depend on (see the note on JuliaCall/JuliaEval below). Pre-2016 Excel is not a
        // realistic target in 2026.
        private const int StringLengthLimit = 32768;

        // Worksheet-style counterpart to CallVBACore below, used by the migration methods
        // (JuliaCall/JuliaCallNamedArgs/JuliaSetVar/JuliaInclude) so that old VBA code calling
        // these by their original (non-"VBA"-suffixed) names keeps their original restricted
        // behaviour (string-length cap, no nested arrays/Dictionaries, 1-D arrays become a
        // column) and their original "return an error string" handling, rather than silently
        // switching to JuliaCallVBA's unrestricted, exception-raising behaviour just because
        // the call now goes through COM. Deliberately does not call into JuliaFunctions.cs:
        // its ExcelMissing/ExcelEmpty handling requires ExcelDna.Integration.dll, which is
        // embedded in the packed .xll and only resolvable when Excel-DNA's own loader has
        // already loaded it into the process -- not guaranteed for a bare CreateObject of
        // this COM server, which COM/VBA calls never produce those sentinel types for anyway.
        private static object CallCore(string juliaFunction, object[] args, string endpoint)
        {
            try
            {
                if (JuliaPort.GetJuliaPortOrZero() == 0)
                {
                    return "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!";
                }

                string fnEncoded = WireFormatEncode.Encode(juliaFunction);
                var encoded = new string[1 + args.Length];
                encoded[0] = fnEncoded;
                for (int i = 0; i < args.Length; i++)
                {
                    encoded[i + 1] = WireFormatEncode.Encode(args[i]);
                }
                var lens = new System.Text.StringBuilder();
                foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }
                string payload = "*1," + encoded.Length + ";" + lens + ";" + string.Concat(encoded);

                string responseWire = JuliaHttpClient.Post(payload, endpoint);
                return WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true);
            }
            catch (System.Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        private static object CallVBACore(string juliaFunction, object[] args, string endpoint)
        {
            string fnEncoded = WireFormatEncode.Encode(juliaFunction);
            var encoded = new string[1 + args.Length];
            encoded[0] = fnEncoded;
            for (int i = 0; i < args.Length; i++)
            {
                encoded[i + 1] = WireFormatEncode.Encode(args[i]);
            }
            var lens = new System.Text.StringBuilder();
            foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }
            string payload = "*1," + encoded.Length + ";" + lens + ";" + string.Concat(encoded);

            string responseWire = JuliaHttpClient.Post(payload, endpoint);
            return WireFormatDecode.Decode(responseWire, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false);
        }

        // Port of ConcatenateExpressions (modPrivateMethods.bas) for the VBA-facing path.
        private static string JuliaExpressionToString(object juliaExpression)
        {
            if (juliaExpression is object[,] arr2)
            {
                int nc = arr2.GetUpperBound(1) - arr2.GetLowerBound(1) + 1;
                if (nc > 1) throw new System.ArgumentException("When passed as an array or a Range, JuliaExpression should have only one column, but got " + nc + " columns");
                int lo = arr2.GetLowerBound(0);
                int hi = arr2.GetUpperBound(0);
                var parts = new string[hi - lo + 1];
                for (int i = lo; i <= hi; i++) parts[i - lo] = System.Convert.ToString(arr2[i, arr2.GetLowerBound(1)]);
                return string.Join("\n", parts);
            }
            if (juliaExpression is object[] arr1)
            {
                var parts = new string[arr1.Length];
                for (int i = 0; i < arr1.Length; i++) parts[i] = System.Convert.ToString(arr1[i]);
                return string.Join("\n", parts);
            }
            return System.Convert.ToString(juliaExpression);
        }
    }
}
