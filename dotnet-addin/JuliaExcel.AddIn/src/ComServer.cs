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
            return JuliaLaunch.Launch(minimiseWindow, commandLineOptions, packages, timeoutSeconds);
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
