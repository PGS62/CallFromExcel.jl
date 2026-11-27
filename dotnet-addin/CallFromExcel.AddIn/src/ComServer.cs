using System.Runtime.InteropServices;

namespace CallFromExcel.AddIn
{
    // The VBA-facing COM-visible class for JuliaCallVBA/JuliaEvalVBA (step 5 of the
    // plan) -- needed because Application.Run/XLL function return types can't carry a
    // live object reference, so results JuliaCall/JuliaEval can't display on a worksheet
    // (Dictionaries, nested arrays) go through here instead. Also still carries the
    // JuliaLaunch/JuliaHttpClient test-harness methods used while validating those ports.
    //
    // Callable both late-bound (Dim x As Object: Set x = CreateObject("CallFromExcel.ComServer"))
    // and early-bound (Tools > References -> CallFromExcel, which the installer registers;
    // Dim x As New CallFromExcel.ComServer), the latter giving the VBA editor's IntelliSense
    // tooltips. Parameter names match the old JuliaExcel.xlam's, so the tooltips read as before.
    //
    // The JuliaCall family takes 30 [Optional] arguments, like Application.Run, rather than
    // "params object[]": VBA can't call a vararg (params) method early-bound at all ("Function
    // or interface marked as restricted, or the function uses an Automation type not supported
    // in Visual Basic"), and with IDispatch-only + vararg its IntelliSense crashed outright.
    // See ArgsList below for how omitted arguments are dropped.
    //
    // Every failure is raised as a VBA error (see Raising below), never returned as a string:
    // a Julia exception, Julia not running, an argument that can't be sent or a result that
    // can't be returned. Excel error values in a result (#N/A etc.) are data, returned as
    // CVErr values.
    [ComVisible(true)]
    [Guid("8ba635c2-d745-4d97-8fec-6cbfb398ae6d")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICallFromExcelComServer
    {
        string Launch(bool MinimiseWindow, string CommandLineOptions, string Packages, int TimeoutSeconds);
        string Post(string payload, string path);
        bool IsListening(int port);
        int GetPort();

        // Port of JuliaCallVBA/JuliaEvalVBA (modMain.bas): differ from JuliaCall/JuliaEval
        // in allowing nested arrays and Dictionary results (AllowNesting=true,
        // JuliaVectorToXLColumn=false, no string-length limit), matching how those are
        // only ever called from VBA, never bound to a worksheet cell.
        object JuliaCallVBA(string JuliaFunction,
            [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5,
            [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10,
            [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15,
            [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20,
            [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25,
            [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
        object JuliaEvalVBA(object JuliaExpression);
        object JuliaCallNamedArgsVBA(string JuliaFunction,
            [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5,
            [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10,
            [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15,
            [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20,
            [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25,
            [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);

        // Migration surface for VBA code that used to call these as bare Public Functions
        // in a referenced JuliaExcel.xlam (Tools > References -> JuliaExcel). Since the new
        // world has no xlam at all, old call sites become e.g.
        //   Result = jl.JuliaCall("sqrt", 16)
        // with jl declared once as CreateObject("CallFromExcel.ComServer") (or
        // early-bound, as New CallFromExcel.ComServer).
        // These intentionally keep the worksheet-style restrictions (string-length limit,
        // no nested arrays/Dictionaries, 1-D arrays become a column) that the original
        // JuliaCall/JuliaEval/JuliaCallNamedArgs/JuliaSetVar/JuliaInclude had even when called
        // from VBA - JuliaCallVBA/JuliaEvalVBA/JuliaCallNamedArgsVBA above remain the way to
        // lift those restrictions, exactly as before.
        object JuliaCall(string JuliaFunction,
            [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5,
            [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10,
            [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15,
            [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20,
            [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25,
            [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
        object JuliaEval(object JuliaExpression);
        object JuliaCallNamedArgs(string JuliaFunction,
            [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5,
            [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10,
            [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15,
            [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20,
            [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25,
            [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
        object JuliaLaunch(bool UseLinux = false, bool MinimiseWindow = false, string CommandLineOptions = "", string Packages = "", string BashStatements = "", int TimeOut = 30);
        bool JuliaIsRunning();
        int JuliaExcelPID();
        object JuliaSetVar(string VariableName, object RefersTo);
        object JuliaInclude(string FileName);

        // Object handles created from VBA (documentation, "Object handles"). Added at the
        // end of the interface so code compiled against the earlier type library is unaffected.
        // A worksheet handle's lifetime comes from Excel's RTD machinery (it lives while a
        // formula uses it); a COM call has no calling cell, so a handle created here lives
        // until JuliaHandleRelease or JuliaHandleReleaseAll frees it.
        string JuliaCallHandle(string JuliaFunction,
            [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5,
            [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10,
            [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15,
            [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20,
            [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25,
            [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
        string JuliaEvalHandle(object JuliaExpression);
        bool JuliaHandleRelease(string Handle);
        int JuliaHandleReleaseAll();
    }

    [ComVisible(true)]
    [Guid("0581209a-2052-43b8-8875-83bfe419e8c5")]
    [ProgId("CallFromExcel.ComServer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICallFromExcelComServer))]
    public class ComServer : ICallFromExcelComServer
    {
        public string Launch(bool MinimiseWindow, string CommandLineOptions, string Packages, int TimeoutSeconds)
        {
            return global::CallFromExcel.AddIn.JuliaLaunch.Launch(minimiseWindow: MinimiseWindow, commandLineOptions: CommandLineOptions, packages: Packages, timeoutSeconds: TimeoutSeconds);
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

        public object JuliaCallVBA(string JuliaFunction,
            object Arg1, object Arg2, object Arg3, object Arg4, object Arg5,
            object Arg6, object Arg7, object Arg8, object Arg9, object Arg10,
            object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
            object Arg21, object Arg22, object Arg23, object Arg24, object Arg25,
            object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return CallVBACore(JuliaFunction, ArgsList(
                Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10,
                Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20,
                Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30), "/call");
        }

        // Errors raised to VBA: Err.Number is one of these, Err.Description the message
        // (without the "#" and "!" a worksheet cell shows around it) and Err.Source
        // "CallFromExcel". vbObjectError + n, as VBA asks of a component's own errors; the
        // documentation ("Calling from VBA") lists them, so they must not change.
        private const int VbObjectError = unchecked((int)0x80040000);
        private const int ErrJuliaException = VbObjectError + 1001;   // the Julia code threw
        private const int ErrJuliaNotRunning = VbObjectError + 1002;  // not launched, exited or still starting
        private const int ErrOther = VbObjectError + 1003;            // anything else

        // Runs a method's body, raising any failure to VBA as a COMException with one of the
        // numbers above.
        private static T Raising<T>(System.Func<T> body)
        {
            try
            {
                return body();
            }
            catch (System.Exception ex) when (!(ex is COMException))
            {
                int number = ex is JuliaException ? ErrJuliaException
                    : ex is JuliaNotRunningException ? ErrJuliaNotRunning
                    : ErrOther;
                string message = ex.Message;
                if (message.StartsWith("#", System.StringComparison.Ordinal)) message = message.Substring(1);
                throw new COMException(message, number) { Source = "CallFromExcel" };
            }
        }

        public object JuliaCallNamedArgsVBA(string JuliaFunction,
            object Arg1, object Arg2, object Arg3, object Arg4, object Arg5,
            object Arg6, object Arg7, object Arg8, object Arg9, object Arg10,
            object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
            object Arg21, object Arg22, object Arg23, object Arg24, object Arg25,
            object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return CallVBACore(JuliaFunction, ArgsList(
                Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10,
                Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20,
                Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30), "/callkw");
        }

        // The arguments actually passed to a JuliaCall-family method: trailing omitted ones are
        // dropped. An omitted argument arrives as System.Reflection.Missing when called
        // late-bound (via IDispatch), but early-bound (via the vtable) as the VT_ERROR
        // DISP_E_PARAMNOTFOUND, which .NET boxes as a plain Int32 (see WireFormatEncode's note
        // on VT_ERROR) - a genuine Long argument of exactly that value is not a realistic
        // collision. An omission before a supplied argument (JuliaCall("f", , 2)) becomes
        // Missing either way, which the encoder rejects.
        private const int DISP_E_PARAMNOTFOUND = unchecked((int)0x80020004);
        private static object[] ArgsList(params object[] args)
        {
            int n = args.Length;
            for (int i = 0; i < n; i++)
            {
                if (args[i] is int && (int)args[i] == DISP_E_PARAMNOTFOUND) args[i] = System.Reflection.Missing.Value;
            }
            while (n > 0 && args[n - 1] is System.Reflection.Missing) n--;
            var result = new object[n];
            System.Array.Copy(args, result, n);
            return result;
        }

        public object JuliaEvalVBA(object JuliaExpression)
        {
            return Raising(() => EvalCore(JuliaExpression, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false));
        }

        public object JuliaCall(string JuliaFunction,
            object Arg1, object Arg2, object Arg3, object Arg4, object Arg5,
            object Arg6, object Arg7, object Arg8, object Arg9, object Arg10,
            object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
            object Arg21, object Arg22, object Arg23, object Arg24, object Arg25,
            object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return CallCore(JuliaFunction, ArgsList(
                Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10,
                Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20,
                Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30), "/call");
        }

        public object JuliaEval(object JuliaExpression)
        {
            return Raising(() => EvalCore(JuliaExpression, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true));
        }

        private static object EvalCore(object juliaExpression, bool allowNesting, int stringLengthLimit, bool vectorToColumn)
        {
            string expression = JuliaExpressionToString(juliaExpression);
            string responseWire = JuliaHttpClient.Post(expression, "/eval");
            object result = WireFormatDecode.Decode(responseWire, allowNesting: allowNesting, stringLengthLimit: stringLengthLimit, vectorToColumn: vectorToColumn);
            result = WireFormatDecode.MapErrorValues(result, MapToComError);
            return ForVba(result);
        }

        public object JuliaCallNamedArgs(string JuliaFunction,
            object Arg1, object Arg2, object Arg3, object Arg4, object Arg5,
            object Arg6, object Arg7, object Arg8, object Arg9, object Arg10,
            object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
            object Arg21, object Arg22, object Arg23, object Arg24, object Arg25,
            object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            return CallCore(JuliaFunction, ArgsList(
                Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10,
                Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20,
                Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30), "/callkw");
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

        public object JuliaSetVar(string VariableName, object RefersTo)
        {
            return CallCore("CallFromExcel.setvar", new[] { (object)VariableName, RefersTo }, "/call");
        }

        public object JuliaInclude(string FileName)
        {
            return CallCore("include", new object[] { FileName.Replace("\\", "/") }, "/call");
        }

        // Unlike the other migration wrappers above, this can't just delegate to
        // JuliaFunctions.JuliaLaunch (its CommandLineOptions/Packages/TimeOut parameters are
        // typed `object` to detect ExcelMissing, an XLL-only concept).
        public object JuliaLaunch(bool UseLinux = false, bool MinimiseWindow = false, string CommandLineOptions = "", string Packages = "", string BashStatements = "", int TimeOut = 30)
        {
            return Raising<object>(() => global::CallFromExcel.AddIn.JuliaLaunch.Launch(useLinux: UseLinux, minimiseWindow: MinimiseWindow, commandLineOptions: CommandLineOptions, packages: Packages, bashStatements: BashStatements, timeoutSeconds: TimeOut));
        }

        // Handles created by JuliaCallHandle/JuliaEvalHandle and not yet released. Only
        // these can be released from VBA: a worksheet handle is released by Excel when no
        // formula uses it, and freeing it here would break the cells using it. Shared by every
        // ComServer instance in this Excel (it's static), but not with the worksheet functions,
        // which run in a separate copy of this assembly (see HandlePattern).
        private static readonly System.Collections.Generic.Dictionary<string, string> VbaHandles =
            new System.Collections.Generic.Dictionary<string, string>();

        // JuliaCallHandle covers both JuliaCall and JuliaCallNamedArgs, as the worksheet
        // function does: an argument ";" marks where named arguments start.
        public string JuliaCallHandle(string JuliaFunction,
            object Arg1, object Arg2, object Arg3, object Arg4, object Arg5,
            object Arg6, object Arg7, object Arg8, object Arg9, object Arg10,
            object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
            object Arg21, object Arg22, object Arg23, object Arg24, object Arg25,
            object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            object[] args = ArgsList(
                Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10,
                Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20,
                Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30);
            var createArgs = new object[1 + args.Length];
            createArgs[0] = JuliaFunction;
            System.Array.Copy(args, 0, createArgs, 1, args.Length);
            return Raising(() => CreateHandle("CallFromExcel.handle_create", createArgs));
        }

        public string JuliaEvalHandle(object JuliaExpression)
        {
            // The code goes as text, even if it looks like a handle.
            return Raising(() => CreateHandle("CallFromExcel.handle_eval", new object[] { JuliaExpressionToString(JuliaExpression) }, mapHandles: false));
        }

        // TRUE if Handle was a live handle created from VBA, now freed. FALSE if it isn't one
        // (already released, created by a worksheet formula, from before Julia restarted and
        // already released, or not a handle at all, e.g. "" when creating it failed) - so
        // releasing is always harmless, e.g. twice, or from a class's Terminate event after
        // JuliaHandleReleaseAll or after its create raised an error.
        public bool JuliaHandleRelease(string Handle)
        {
            if (Handle == null || !VbaHandles.TryGetValue(Handle, out string key)) return false;
            VbaHandles.Remove(Handle);
            FreeInJulia(key);
            return true;
        }

        // Frees every live handle created from VBA; returns how many. Worksheet handles are
        // unaffected.
        public int JuliaHandleReleaseAll()
        {
            var keys = new System.Collections.Generic.List<string>(VbaHandles.Values);
            VbaHandles.Clear();
            foreach (string key in keys) FreeInJulia(key);
            return keys.Count;
        }

        private static string CreateHandle(string juliaFunction, object[] args, bool mapHandles = true)
        {
            string responseWire = JuliaHttpClient.Post(BuildCallPayload(juliaFunction, args, mapHandles), "/call");
            object result = WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: 0, vectorToColumn: true);
            string handle = result as string;
            if (HandlePattern.TryGetKey(handle, out string key))
            {
                VbaHandles[handle] = key;
                return handle;
            }
            // A failed call was raised by Decode, so this would be a bug.
            throw new System.InvalidOperationException("Julia returned " + System.Convert.ToString(result) + " rather than a handle");
        }

        // Julia not running, or restarted: its objects are gone anyway (and handle_free ignores
        // a key from an earlier session), so a failure here needs no report.
        private static void FreeInJulia(string key)
        {
            try
            {
                if (JuliaPort.GetJuliaPortOrZero() != 0)
                {
                    JuliaHttpClient.Post(BuildCallPayload("CallFromExcel.handle_free", new object[] { key }), "/call");
                }
            }
            catch (System.Exception)
            {
            }
        }

        // The "*1,<n>;..." wire envelope of a call: the function name, then each argument.
        // A Range argument becomes its values, and each object handle (a string of the handle
        // pattern, here or in an array) a reference to its Julia object.
        // mapHandles false sends string arguments as they are: for code to evaluate.
        private static string BuildCallPayload(string juliaFunction, object[] args, bool mapHandles = true)
        {
            var encoded = new string[1 + args.Length];
            encoded[0] = WireFormatEncode.Encode(juliaFunction);
            for (int i = 0; i < args.Length; i++)
            {
                object value = RangeToValue(args[i]);
                encoded[i + 1] = WireFormatEncode.Encode(!mapHandles && value is string ? value : HandlePattern.MapHandles(value));
            }
            var lens = new System.Text.StringBuilder();
            foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }
            return "*1," + encoded.Length + ";" + lens + ";" + string.Concat(encoded);
        }

        // Excel 2016 and later, as in JuliaFunctions.cs.
        private const int StringLengthLimit = 32768;

        // Worksheet-style counterpart to CallVBACore below, used by the migration methods
        // (JuliaCall/JuliaCallNamedArgs/JuliaSetVar/JuliaInclude) so that old VBA code calling
        // these by their original (non-"VBA"-suffixed) names keeps their original restricted
        // behaviour (string-length cap, no nested arrays/Dictionaries, 1-D arrays become a
        // column), rather than silently switching to JuliaCallVBA's unrestricted behaviour just
        // because the call now goes through COM. Failures are raised, as by every method here
        // (until 0.2.0 these returned "#...!" strings). Deliberately does not call into JuliaFunctions.cs:
        // its ExcelMissing/ExcelEmpty handling requires ExcelDna.Integration.dll, which is
        // embedded in the packed .xll and only resolvable when Excel-DNA's own loader has
        // already loaded it into the process -- not guaranteed for a bare CreateObject of
        // this COM server, which COM/VBA calls never produce those sentinel types for anyway.
        private static object CallCore(string juliaFunction, object[] args, string endpoint)
        {
            return Raising(() =>
            {
                string responseWire = JuliaHttpClient.Post(BuildCallPayload(juliaFunction, args), endpoint);
                object result = WireFormatDecode.Decode(responseWire, allowNesting: false, stringLengthLimit: StringLengthLimit, vectorToColumn: true);
                result = WireFormatDecode.MapErrorValues(result, MapToComError);
                return ForVba(result);
            });
        }

        private static object CallVBACore(string juliaFunction, object[] args, string endpoint)
        {
            return Raising(() =>
            {
                string responseWire = JuliaHttpClient.Post(BuildCallPayload(juliaFunction, args), endpoint);
                object result = WireFormatDecode.Decode(responseWire, allowNesting: true, stringLengthLimit: 0, vectorToColumn: false);
                result = WireFormatDecode.MapErrorValues(result, MapToComError);
                return ForVba(result);
            });
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
            juliaExpression = RangeToValue(juliaExpression);
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

        // VBA code may pass a Range object, e.g. JuliaCall("sum", Range("A1:A10")). It arrives
        // as a COM object; use its Value2, as JuliaExcel's VBA did. Other COM objects (a
        // Scripting.Dictionary) have no Value2 property and are returned unchanged.
        private static object RangeToValue(object value)
        {
            if (value == null || !value.GetType().IsCOMObject) return value;
            try
            {
                return value.GetType().InvokeMember("Value2", System.Reflection.BindingFlags.GetProperty, null, value, null);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return value;
            }
            catch (System.MissingMethodException)
            {
                return value;
            }
        }

        // Converts a decoded result tree to what JuliaExcel's VBA returned, so migrated code
        // behaves identically: arrays rebased from 0-based (WireFormatDecode's convention) to
        // 1-based, since the old Unserialise built them via ReDim(1 To n) and code indexing a
        // result with a hard-coded 1 would otherwise silently read the wrong element; and
        // dictionaries as real Scripting.Dictionary objects, so For Each, CompareMode etc. and
        // variables declared As Scripting.Dictionary all keep working.
        private static object ForVba(object value)
        {
            var dict = value as JuliaDictionary;
            if (dict != null)
            {
                System.Type scriptingDictionary = System.Type.GetTypeFromProgID("Scripting.Dictionary", throwOnError: true);
                object mapped = System.Activator.CreateInstance(scriptingDictionary);
                foreach (var kv in dict.Entries)
                {
                    scriptingDictionary.InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, mapped,
                        new[] { ForVba(kv.Key), ForVba(kv.Value) });
                }
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
            // Copies in the same linear order for both, so element (i, j) lands at (i+1, j+1).
            System.Array.Copy(arr, rebased, arr.Length);
            if (arr.GetType().GetElementType() != typeof(object)) return rebased;

            // Nested arrays and dictionaries need the same conversion. Direct indexing for the
            // shapes the decoder produces; GetValue/SetValue (much slower) otherwise.
            var vector = arr as object[];
            var matrix = arr as object[,];
            if (vector != null)
            {
                for (int i = 0; i < vector.Length; i++)
                {
                    if (NeedsForVba(vector[i])) rebased.SetValue(ForVba(vector[i]), i + 1);
                }
            }
            else if (matrix != null)
            {
                var rebasedMatrix = (object[,])rebased;
                int nr = matrix.GetLength(0);
                int nc = matrix.GetLength(1);
                for (int i = 0; i < nr; i++)
                {
                    for (int j = 0; j < nc; j++)
                    {
                        if (NeedsForVba(matrix[i, j])) rebasedMatrix[i + 1, j + 1] = ForVba(matrix[i, j]);
                    }
                }
            }
            else
            {
                var srcIdx = new int[rank];
                var dstIdx = new int[rank];
                for (int count = 0; count < arr.Length; count++)
                {
                    object element = arr.GetValue(srcIdx);
                    if (NeedsForVba(element))
                    {
                        for (int d = 0; d < rank; d++) dstIdx[d] = srcIdx[d] + 1;
                        rebased.SetValue(ForVba(element), dstIdx);
                    }
                    for (int d = 0; d < rank; d++)
                    {
                        srcIdx[d]++;
                        if (srcIdx[d] < arr.GetLength(d)) break;
                        srcIdx[d] = 0;
                    }
                }
            }
            return rebased;
        }

        private static bool NeedsForVba(object element)
        {
            return element is System.Array || element is JuliaDictionary;
        }
    }
}
