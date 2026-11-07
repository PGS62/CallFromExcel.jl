using System;
using System.Text;
using ExcelDna.Integration;

namespace CallFromExcel.AddIn
{
    // C# port of the worksheet-facing functions (modMain.bas) as Excel-DNA XLL functions.
    // JuliaCallVBA/JuliaEvalVBA, which can return dictionaries and nested arrays, are in
    // ComServer.cs.
    //
    // None is registered thread-safe, so Excel runs them one at a time on its main thread, as
    // it did JuliaExcel's VBA functions. Julia functions can have side effects, and the Julia
    // server keeps each call's arguments in globals (args_from_xl etc.), so concurrent calls
    // could see each other's arguments.
    public static class JuliaFunctions
    {
        // Port of GetStringLengthLimit (modUnserialise.bas): Excel 2010 and earlier capped
        // array-element strings at 255 chars; 2016+ (all report version 16) raised it to 32,767.
        private static int _stringLengthLimit = 32768;

        // Called from AutoOpen, on Excel's main thread: ExcelDnaUtil.ExcelVersion reads an Excel
        // COM property, which deadlocks Excel if read from a calculation thread while the main
        // thread waits for that calculation.
        internal static void CacheStringLengthLimit()
        {
            _stringLengthLimit = ExcelDnaUtil.ExcelVersion <= 15 ? 256 : 32768;
        }

        private static int StringLengthLimit()
        {
            return _stringLengthLimit;
        }

        // Parameter order (UseLinux, MinimiseWindow, CommandLineOptions, Packages,
        // BashStatements, TimeOut) deliberately matches the old VBA worksheet function
        // (modMain.bas) exactly - dropping any of them from the signature (as an earlier
        // version of this wrapper did, before WSL support existed here) broke every old
        // formula that set a later positional argument (e.g.
        // "=JuliaLaunch(,,""--project=...."")" to set CommandLineOptions) by silently
        // shifting it onto the wrong parameter.
        [ExcelFunction(Description = "Launches a local Julia session which listens to the current Excel session and responds to calls to JuliaEval etc..", IsThreadSafe = false)]
        public static object JuliaLaunch(
            [ExcelArgument(Name = "UseLinux", Description = "Launch Julia under WSL instead of natively on Windows.")] object useLinux,
            [ExcelArgument(Name = "MinimiseWindow", Description = "Launch Julia's console window minimised.")] object minimiseWindow,
            [ExcelArgument(Name = "CommandLineOptions", Description = "Command line options set when launching Julia.")] object commandLineOptions,
            [ExcelArgument(Name = "Packages", Description = "Packages to load, comma-delimited, which must be available in the default Julia environment.")] object packages,
            [ExcelArgument(Name = "BashStatements", Description = "Relevant only when UseLinux is TRUE. Bash statements to run before launching Julia under WSL, e.g. to set environment variables. Delimit multiple statements with the line feed character.")] object bashStatements,
            [ExcelArgument(Name = "TimeOut", Description = "Seconds to wait for Julia to fully start before giving up waiting. Defaults to 30.")] object timeOut)
        {
            try
            {
                bool linux = (useLinux is ExcelMissing) ? false : Convert.ToBoolean(useLinux);
                string bash = (bashStatements is ExcelMissing) ? "" : Convert.ToString(bashStatements);
                bool minimise = (minimiseWindow is ExcelMissing) ? false : Convert.ToBoolean(minimiseWindow);
                string cmdOpts = (commandLineOptions is ExcelMissing) ? "" : Convert.ToString(commandLineOptions);
                string pkgs = (packages is ExcelMissing) ? "" : Convert.ToString(packages);
                int timeout = (timeOut is ExcelMissing) ? 30 : Convert.ToInt32(Convert.ToDouble(timeOut));
                return CallFromExcel.AddIn.JuliaLaunch.Launch(useLinux: linux, minimiseWindow: minimise, commandLineOptions: cmdOpts, packages: pkgs, bashStatements: bash, timeoutSeconds: timeout);
            }
            catch (Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        [ExcelFunction(Description = "Call a named Julia function, passing in data from the worksheet. Returns an error string for results that cannot be displayed on a worksheet (nested arrays, dictionaries etc).", IsThreadSafe = false)]
        public static object JuliaCall(
            [ExcelArgument(Name = "JuliaFunction", Description = "The name of a Julia function visible from the Julia REPL.")] string juliaFunction,
            [ExcelArgument(Name = "Args", Description = "Zero or more arguments: a number, string, Boolean, empty cell, or array/range.")] params object[] args)
        {
            return CallCore(juliaFunction, args, "/call");
        }

        [ExcelFunction(Description = "Call a named Julia function, passing positional and named arguments from the worksheet. The argument \";\" marks where the named arguments start, given as name-value pairs. A named argument whose value is \";\" must be the last pair.", IsThreadSafe = false)]
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

        [ExcelFunction(Description = "Set a global variable in the Julia process.", IsThreadSafe = false)]
        public static object JuliaSetVar(
            [ExcelArgument(Name = "VariableName", Description = "The name of the variable to be set. Must follow Julia's rules for allowed variable names.")] string variableName,
            [ExcelArgument(Name = "RefersTo", Description = "An Excel range, or a number, string, Boolean, empty cell or array of such types.")] object refersTo)
        {
            return CallCore("CallFromExcel.setvar", new[] { (object)variableName, refersTo }, "/call");
        }

        [ExcelFunction(Description = "Load a Julia source file into the Julia process, to make additional functions available via JuliaEval and JuliaCall.", IsThreadSafe = false)]
        public static object JuliaInclude(
            [ExcelArgument(Name = "FileName", Description = "The full name of the file to be included.")] string fileName)
        {
            return CallCore("CallFromExcel.include", new object[] { fileName.Replace("\\", "/") }, "/call");
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
                return DecodeForWorksheet(responseWire);
            }
            catch (Exception ex)
            {
                return "#" + ex.Message;
            }
        }

        // One pass over the result: Julia's Excel errors become worksheet errors, and
        // nothing/missing an empty cell.
        private static object DecodeForWorksheet(string responseWire)
        {
            object result = WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit(), vectorToColumn: true, forWorksheet: true);
            return WireFormatDecode.MapLeaves(result, v => v is ExcelErrorValue ev ? MapToExcelDnaError(ev.Code) : v ?? ExcelEmpty.Value);
        }

        [ExcelFunction(Description = "Evaluate a Julia expression and return the result to an Excel worksheet.", IsThreadSafe = false)]
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
                return DecodeForWorksheet(responseWire);
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

        // Maps a wire-format error code (VBA's own XlCVError UI numbering, e.g. 2042 for
        // #N/A) to Excel-DNA's ExcelError enum, which Excel-DNA recognises and converts to a
        // genuine XLOPER error in a worksheet cell. VBA's numbering is exactly 2000 + the
        // underlying xlerror code Excel-DNA itself uses (confirmed empirically: 2007-2000=7=
        // ExcelErrorDiv0's own value, etc.), but this Excel-DNA version (1.9.0) only defines
        // enum members for the 8 "classic" errors - the 6 newer dynamic-array-era codes
        // (#SPILL!, #CALC!, #BLOCKED!, #CONNECT!, #FIELD!, #UNKNOWN!, codes 2045-2050) have
        // no native XLOPER error representation here, so they fall back to a text
        // approximation rather than losing the value entirely.
        private static object MapToExcelDnaError(int code)
        {
            switch (code)
            {
                case 2000: return ExcelError.ExcelErrorNull;
                case 2007: return ExcelError.ExcelErrorDiv0;
                case 2015: return ExcelError.ExcelErrorValue;
                case 2023: return ExcelError.ExcelErrorRef;
                case 2029: return ExcelError.ExcelErrorName;
                case 2036: return ExcelError.ExcelErrorNum;
                case 2042: return ExcelError.ExcelErrorNA;
                case 2043: return ExcelError.ExcelErrorGettingData;
                default: return "#ERR" + code + "!";
            }
        }

        // Port of ConcatenateExpressions (modPrivateMethods.bas). Uses Array.GetLowerBound/
        // GetValue rather than casting to object[]/object[,] directly - see ComServer.cs's
        // own copy of this method for why a direct cast isn't safe in general (a non-zero-
        // lower-bound 1-D array isn't an object[] at all).
        private static string ConcatenateExpressions(object juliaExpression)
        {
            var arr = juliaExpression as Array;
            if (arr != null)
            {
                if (arr.Rank == 1)
                {
                    int lb = arr.GetLowerBound(0);
                    int ub = arr.GetUpperBound(0);
                    var parts = new string[ub - lb + 1];
                    for (int i = lb; i <= ub; i++) parts[i - lb] = Convert.ToString(arr.GetValue(i));
                    return string.Join("\n", parts);
                }
                if (arr.Rank == 2)
                {
                    int lb0 = arr.GetLowerBound(0);
                    int ub0 = arr.GetUpperBound(0);
                    int lb1 = arr.GetLowerBound(1);
                    int nc = arr.GetUpperBound(1) - lb1 + 1;
                    if (nc > 1) throw new ArgumentException("When passed as an array or a Range, JuliaExpression should have only one column, but got " + nc + " columns");
                    var parts = new string[ub0 - lb0 + 1];
                    for (int i = lb0; i <= ub0; i++) parts[i - lb0] = Convert.ToString(arr.GetValue(i, lb1));
                    return string.Join("\n", parts);
                }
                throw new ArgumentException("Too many dimensions");
            }
            return Convert.ToString(juliaExpression);
        }
    }
}
