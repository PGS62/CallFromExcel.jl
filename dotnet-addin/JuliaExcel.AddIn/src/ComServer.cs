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
            object result = WireFormatDecode.Decode(responseWire, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false);
            result = WireFormatDecode.MapErrorValues(result, MapToComError);
            return Rebase1Based(result);
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
                object result = WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true);
                result = WireFormatDecode.MapErrorValues(result, MapToComError);
                return Rebase1Based(result);
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
                object result = WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true);
                result = WireFormatDecode.MapErrorValues(result, MapToComError);
                return Rebase1Based(result);
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
            object result = WireFormatDecode.Decode(responseWire, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false);
            result = WireFormatDecode.MapErrorValues(result, MapToComError);
            return Rebase1Based(result);
        }

        // Maps a wire-format error code (VBA's own XlCVError UI numbering, e.g. 2042 for
        // #N/A) to a genuine COM VT_ERROR Variant, which VBA then sees exactly as if it had
        // called CVErr(code) itself. ExcelErrorValue (the decoded wire value) is a plain
        // struct with no COM significance of its own, so .NET cannot marshal it to a VBA
        // Variant without this conversion. The well-known Excel/VBA error SCODEs (e.g.
        // 0x800A07FA for #N/A) are exactly 0x800A0000 | code for all 14 codes this project's
        // wire format carries (confirmed against the documented SCODEs for the 7 "classic"
        // errors: 2000/2007/2015/2023/2029/2036/2042 -> 0x800A07D0/D7/DF/E7/ED/F4/FA) - unlike
        // the XLL path (JuliaFunctions.MapToExcelDnaError), there is no enum with a limited
        // set of members to worry about here, so this covers all 14 codes uniformly.
        private static object MapToComError(int code)
        {
            return new ErrorWrapper(unchecked((int)(0x800A0000u | (uint)code)));
        }

        // Port of ConcatenateExpressions (modPrivateMethods.bas) for the VBA-facing path.
        // Uses Array.GetLowerBound/GetValue rather than casting to object[]/object[,]
        // directly: a VBA array arriving over COM is virtually always 1-based in this
        // codebase's own convention (Dim x(1 To n)), and a non-zero-lower-bound 1-D array is
        // not an object[] at all (a different runtime type - see WireFormatEncode.cs's own
        // note on this), so a direct cast would throw InvalidCastException instead of just
        // mis-handling the bounds the way a naive 0-based assumption would.
        private static string JuliaExpressionToString(object juliaExpression)
        {
            var arr = juliaExpression as System.Array;
            if (arr != null)
            {
                if (arr.Rank == 1)
                {
                    int lb = arr.GetLowerBound(0);
                    int ub = arr.GetUpperBound(0);
                    var parts = new string[ub - lb + 1];
                    for (int i = lb; i <= ub; i++) parts[i - lb] = System.Convert.ToString(arr.GetValue(i));
                    return string.Join("\n", parts);
                }
                if (arr.Rank == 2)
                {
                    int lb0 = arr.GetLowerBound(0);
                    int ub0 = arr.GetUpperBound(0);
                    int lb1 = arr.GetLowerBound(1);
                    int nc = arr.GetUpperBound(1) - lb1 + 1;
                    if (nc > 1) throw new System.ArgumentException("When passed as an array or a Range, JuliaExpression should have only one column, but got " + nc + " columns");
                    var parts = new string[ub0 - lb0 + 1];
                    for (int i = lb0; i <= ub0; i++) parts[i - lb0] = System.Convert.ToString(arr.GetValue(i, lb1));
                    return string.Join("\n", parts);
                }
                throw new System.ArgumentException("Too many dimensions");
            }
            return System.Convert.ToString(juliaExpression);
        }

        // Rebases every object[]/object[,]/double[]/double[,]/... in a decoded result tree
        // from 0-based (WireFormatDecode's own convention, matching Excel-DNA's) to 1-based,
        // matching the old, pure-VBA JuliaCallVBA/JuliaEvalVBA's behaviour exactly (its
        // Unserialise always built results via ReDim(1 To n)) - without this, migrating VBA
        // code indexing a result with a hard-coded 1 (not relative to LBound/UBound, which is
        // how most of this codebase's own test suite was written) would silently read the
        // wrong element rather than erroring, since COM doesn't care about VBA's own 1-based
        // convention on its own.
        private static object Rebase1Based(object value)
        {
            var dict = value as JuliaDictionary;
            if (dict != null)
            {
                var mapped = new JuliaDictionary();
                foreach (var kv in dict.Entries) mapped.Add(Rebase1Based(kv.Key), Rebase1Based(kv.Value));
                return mapped;
            }

            var arr = value as System.Array;
            if (arr == null) return value;

            int rank = arr.Rank;
            for (int d = 0; d < rank; d++)
            {
                if (arr.GetLowerBound(d) != 0) return value; // already non-zero-based - leave alone
            }

            var lengths = new int[rank];
            var lowerBounds = new int[rank];
            for (int d = 0; d < rank; d++)
            {
                lengths[d] = arr.GetLength(d);
                lowerBounds[d] = 1;
            }
            System.Array rebased = System.Array.CreateInstance(arr.GetType().GetElementType(), lengths, lowerBounds);

            bool recurseIntoElements = arr.GetType().GetElementType() == typeof(object);
            var srcIdx = new int[rank];
            var dstIdx = new int[rank];
            int total = arr.Length;
            for (int count = 0; count < total; count++)
            {
                object element = arr.GetValue(srcIdx);
                object mappedElement = recurseIntoElements ? Rebase1Based(element) : element;
                for (int d = 0; d < rank; d++) dstIdx[d] = srcIdx[d] + 1;
                rebased.SetValue(mappedElement, dstIdx);

                for (int d = 0; d < rank; d++)
                {
                    srcIdx[d]++;
                    if (srcIdx[d] < arr.GetLength(d)) break;
                    srcIdx[d] = 0;
                }
            }
            return rebased;
        }
    }
}
