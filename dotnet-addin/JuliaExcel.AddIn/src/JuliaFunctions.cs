using System;
using System.Text;
using ExcelDna.Integration;

namespace JuliaExcel.AddIn
{
    // C# port of the worksheet-facing JuliaCall/JuliaEval (modMain.bas) as ordinary
    // Excel-DNA XLL functions -- per the August plan: "JuliaCall/JuliaEval (worksheet-
    // facing) -> ordinary Excel-DNA XLL functions, registered thread-safe. No behavior
    // regression: they're already documented as restricted to worksheet-displayable
    // results." JuliaCallVBA/JuliaEvalVBA (which can return Dictionaries/nested arrays,
    // needing a live object reference Application.Run/XLL functions can't carry) are a
    // separate COM-visible class -- step 5 of the plan, not yet done.
    public static class JuliaFunctions
    {
        private static int _cachedStringLengthLimit;

        private static int StringLengthLimit()
        {
            // Port of GetStringLengthLimit (modUnserialise.bas): Excel 2010 and earlier
            // capped array-element strings at 255 chars; 2016+ (all report version 16)
            // raised it to 32,767. Cached after the first call, not just as a micro-
            // optimization: ExcelDnaUtil.ExcelVersion reads an Excel COM property, and
            // JuliaCall/JuliaEval run with IsThreadSafe=true, meaning Excel-DNA may invoke
            // them on a worker thread rather than Excel's main thread -- querying a COM
            // property cross-thread back onto Excel's STA thread on every single call was
            // found (via this project's own JuliaCall("sort", ...) timing investigation) to
            // roughly double the real end-to-end latency of every call, versus computing it
            // once. Excel's version obviously can't change mid-session.
            if (_cachedStringLengthLimit == 0)
            {
                _cachedStringLengthLimit = ExcelDnaUtil.ExcelVersion <= 15 ? 256 : 32768;
            }
            return _cachedStringLengthLimit;
        }

        [ExcelFunction(Description = "Launches a local Julia session which listens to the current Excel session and responds to calls to JuliaEval etc..", IsThreadSafe = false)]
        public static object JuliaLaunch(
            [ExcelArgument(Name = "CommandLineOptions", Description = "Command line options set when launching Julia.")] object commandLineOptions,
            [ExcelArgument(Name = "Packages", Description = "Packages to load, comma-delimited, which must be available in the default Julia environment.")] object packages,
            [ExcelArgument(Name = "TimeOut", Description = "Seconds to wait for Julia to fully start before giving up waiting. Defaults to 30.")] object timeOut)
        {
            try
            {
                string cmdOpts = (commandLineOptions is ExcelMissing) ? "" : Convert.ToString(commandLineOptions);
                string pkgs = (packages is ExcelMissing) ? "" : Convert.ToString(packages);
                int timeout = (timeOut is ExcelMissing) ? 30 : Convert.ToInt32(Convert.ToDouble(timeOut));
                return JuliaExcel.AddIn.JuliaLaunch.Launch(commandLineOptions: cmdOpts, packages: pkgs, timeoutSeconds: timeout);
            }
            catch (Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        [ExcelFunction(Description = "Call a named Julia function, passing in data from the worksheet. Returns an error string for results that cannot be displayed on a worksheet (nested arrays, dictionaries etc).", IsThreadSafe = true)]
        public static object JuliaCall(
            [ExcelArgument(Name = "JuliaFunction", Description = "The name of a Julia function visible from the Julia REPL.")] string juliaFunction,
            [ExcelArgument(Name = "Args", Description = "Zero or more arguments: a number, string, Boolean, empty cell, or array/range.")] params object[] args)
        {
            return CallCore(juliaFunction, args, "/call");
        }

        [ExcelFunction(Description = "Call a named Julia function, passing positional and named arguments from the worksheet. The argument \";\" marks where the named arguments start, given as name-value pairs. A named argument whose value is \";\" must be the last pair.", IsThreadSafe = true)]
        public static object JuliaCallNamedArgs(
            [ExcelArgument(Name = "JuliaFunction", Description = "The name of a Julia function visible from the Julia REPL.")] string juliaFunction,
            [ExcelArgument(Name = "Args", Description = "Positional arguments, then \";\", then name/value pairs for keyword arguments.")] params object[] args)
        {
            return CallCore(juliaFunction, args, "/callkw");
        }

        [ExcelFunction(Description = "Returns TRUE if a Julia session is running and listening to this Excel session.", IsThreadSafe = false)]
        public static bool JuliaIsRunning()
        {
            // Improves on the VBA original's approach (matching a console window titled
            // "serving Excel PID <n>"): this project's own investigation found that
            // window-title matching never works when Windows Terminal is the default
            // terminal host (see JuliaLaunch.cs's own notes), so this checks the thing
            // that actually matters -- is something listening on the cached port -- same
            // signal JuliaLaunch itself already relies on.
            int port = JuliaPort.GetJuliaPortOrZero();
            return port != 0 && JuliaHttpClient.IsListening(port);
        }

        [ExcelFunction(Description = "Returns the process id of the current Excel session. Only needed if more than one Excel instance is running - pass it to serve_xl(pid) to say which one to attach to.", IsThreadSafe = false)]
        public static int JuliaExcelPID()
        {
            return System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        [ExcelFunction(Description = "Set a global variable in the Julia process.", IsThreadSafe = true)]
        public static object JuliaSetVar(
            [ExcelArgument(Name = "VariableName", Description = "The name of the variable to be set. Must follow Julia's rules for allowed variable names.")] string variableName,
            [ExcelArgument(Name = "RefersTo", Description = "An Excel range, or a number, string, Boolean, empty cell or array of such types.")] object refersTo)
        {
            return CallCore("JuliaExcel.setvar", new[] { (object)variableName, refersTo }, "/call");
        }

        [ExcelFunction(Description = "Load a Julia source file into the Julia process, to make additional functions available via JuliaEval and JuliaCall.", IsThreadSafe = true)]
        public static object JuliaInclude(
            [ExcelArgument(Name = "FileName", Description = "The full name of the file to be included.")] string fileName)
        {
            return CallCore("JuliaExcel.include", new object[] { fileName.Replace("\\", "/") }, "/call");
        }

        private static object CallCore(string juliaFunction, object[] args, string endpoint)
        {
            try
            {
                if (JuliaPort.GetJuliaPortOrZero() == 0)
                {
                    return "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!";
                }

                string payload = BuildCallPayload(juliaFunction, args);
                string responseWire = JuliaHttpClient.Post(payload, endpoint);
                return WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit(), vectorToColumn: true);
            }
            catch (Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        [ExcelFunction(Description = "Evaluate a Julia expression and return the result to an Excel worksheet.", IsThreadSafe = true)]
        public static object JuliaEval(
            [ExcelArgument(Name = "JuliaExpression", Description = "Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple Julia statements.")] object juliaExpression)
        {
            try
            {
                if (JuliaPort.GetJuliaPortOrZero() == 0)
                {
                    return "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!";
                }

                string expression = ConcatenateExpressions(juliaExpression);
                string responseWire = JuliaHttpClient.Post(expression, "/eval");
                return WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit(), vectorToColumn: true);
            }
            catch (Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        // Port of JuliaCall's own argument-encoding loop (modMain.bas): wraps the function
        // name and each argument as elements of the general "*1,<n>;..." wire envelope.
        // No separate Range->.Value2 step is needed here (unlike JuliaCallVBA's VBA
        // original): a worksheet-called XLL function with `object`/`object[]` parameters
        // already receives resolved values, not a live Range reference, from Excel's own
        // calling convention -- confirmed for the VBA original earlier in this project's
        // own investigation (a Variant-typed UDF parameter never receives a Range object
        // when called from a worksheet formula).
        private static string BuildCallPayload(string juliaFunction, object[] args)
        {
            string fnEncoded = WireFormatEncode.Encode(juliaFunction);
            var encoded = new string[1 + args.Length];
            encoded[0] = fnEncoded;
            for (int i = 0; i < args.Length; i++)
            {
                object arg = args[i];
                if (arg is ExcelMissing || arg is ExcelEmpty) arg = null;
                encoded[i + 1] = WireFormatEncode.Encode(arg);
            }

            var lens = new StringBuilder();
            foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }

            return "*1," + encoded.Length + ";" + lens + ";" + string.Concat(encoded);
        }

        // Port of ConcatenateExpressions (modPrivateMethods.bas).
        private static string ConcatenateExpressions(object juliaExpression)
        {
            if (juliaExpression is object[,] arr2)
            {
                int nc = arr2.GetUpperBound(1) - arr2.GetLowerBound(1) + 1;
                if (nc > 1) throw new ArgumentException("When passed as an array or a Range, JuliaExpression should have only one column, but got " + nc + " columns");
                int lo = arr2.GetLowerBound(0);
                int hi = arr2.GetUpperBound(0);
                var parts = new string[hi - lo + 1];
                for (int i = lo; i <= hi; i++) parts[i - lo] = Convert.ToString(arr2[i, arr2.GetLowerBound(1)]);
                return string.Join("\n", parts);
            }
            if (juliaExpression is object[] arr1)
            {
                var parts = new string[arr1.Length];
                for (int i = 0; i < arr1.Length; i++) parts[i] = Convert.ToString(arr1[i]);
                return string.Join("\n", parts);
            }
            return Convert.ToString(juliaExpression);
        }
    }
}
