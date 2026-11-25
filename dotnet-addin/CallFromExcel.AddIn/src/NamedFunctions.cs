using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelDna.Integration;

namespace CallFromExcel.AddIn
{
    // SPIKE (branch spike/named-functions): named worksheet functions, see "Named worksheet
    // functions" in Enhancements.md and src/expose.jl. Fetches the functions Julia has exposed
    // and registers each as a worksheet function - by default named "jl.<julia name>" - that
    // forwards its arguments to the same code as JuliaCall.
    public static class NamedFunctions
    {
        private const char ArgSep = '\x1f';

        // Recent registrations, for checking the spike from a sheet.
        private static readonly List<string> Log = new List<string>();

        // Julia's exposed_version() when the add-in last registered: "<session>#<count>".
        private static string registeredVersion = "";

        // A registration is queued and hasn't started: several checks in one calculation (a
        // JuliaLaunch and a JuliaEval cell, say) queue it once. Main thread only.
        private static bool registrationQueued;

        // Called after JuliaLaunch (worksheet and ribbon), JuliaInclude (worksheet and ribbon)
        // and JuliaEval, which may have run code calling expose_to_excel: if Julia's exposed_version()
        // differs from the version last registered, queues the registration, which needs a
        // macro context. One small call to Julia; nothing queued when nothing changed.
        internal static void AfterJuliaChange()
        {
            string version = ExposedVersion();
            if (version == null || version == registeredVersion || registrationQueued) return;
            registrationQueued = true;
            ExcelAsyncUtil.QueueAsMacro(RegisterIfChanged);
        }

        // The queued registration. By the time it runs, the JuliaRegisterFunctions command (as
        // VBA code calls it) may already have registered this version: then nothing to do.
        private static void RegisterIfChanged()
        {
            registrationQueued = false;
            string version = ExposedVersion();
            if (version == null || version == registeredVersion) return;
            Register();
        }

        // null when Julia gives no answer (not running, or an error).
        private static string ExposedVersion()
        {
            object version;
            try { version = JuliaFunctions.CallForHandles("CallFromExcel.exposed_version", new object[0]); }
            catch (Exception) { return null; }
            string s = version as string;
            return string.IsNullOrEmpty(s) || s.StartsWith("#", StringComparison.Ordinal) ? null : s;
        }

        // Registration needs a macro context, so this is a command (Application.Run, or queued
        // with ExcelAsyncUtil.QueueAsMacro).
        [ExcelCommand(Name = "JuliaRegisterFunctions")]
        public static void RegisterAll()
        {
            Register();
        }

        // Registers every exposed function; returns how many. The version is read first, so an
        // expose_to_excel that runs while this does leaves the next check seeing a change.
        internal static int Register()
        {
            string version = ExposedVersion();
            if (version == null) { Log.Add("no answer from Julia"); return 0; }
            object list;
            try { list = JuliaFunctions.CallForHandles("CallFromExcel.exposed_functions", new object[0]); }
            catch (Exception ex) { Log.Add("fetch failed: " + ex.Message); return 0; }
            var rows = list as object[,];
            if (rows == null)
            {
                Log.Add("nothing to register: " + Convert.ToString(list));
                registeredVersion = version;
                return 0;
            }

            var lambdas = new List<LambdaExpression>();
            var functionAttributes = new List<object>();
            var argumentAttributes = new List<List<object>>();
            for (int i = 0; i < rows.GetLength(0); i++)
            {
                string excelName = Convert.ToString(rows[i, 0]);
                string juliaName = Convert.ToString(rows[i, 1]);
                string description = Convert.ToString(rows[i, 2]);
                string[] argNames = SplitList(rows[i, 3]);
                string[] argDescriptions = SplitList(rows[i, 4]);
                int nRequired = Convert.ToInt32(rows[i, 5]);

                lambdas.Add(Forwarder(excelName, juliaName, nRequired, argNames));
                functionAttributes.Add(new ExcelFunctionAttribute
                {
                    Name = excelName,
                    Description = description,
                    Category = "Julia",
                    IsThreadSafe = false,
                    SuppressOverwriteError = true,
                });
                var args = new List<object>();
                for (int k = 0; k < argNames.Length; k++)
                {
                    args.Add(new ExcelArgumentAttribute
                    {
                        Name = argNames[k],
                        Description = k < argDescriptions.Length ? argDescriptions[k] : "",
                    });
                }
                argumentAttributes.Add(args);
                Log.Add("register " + excelName + "(" + string.Join(", ", argNames) + ") -> " + juliaName);
            }
            ExcelIntegration.RegisterLambdaExpressions(lambdas, functionAttributes, argumentAttributes);
            // IntelliSense reads the registered functions when the add-in loads; without this it
            // doesn't show the arguments of functions registered later (the Function Wizard
            // and Ctrl+Shift+A, which ask Excel, do).
            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
            registeredVersion = version;
            Log.Add("registered " + lambdas.Count + " (" + version + ")");
            // A saved workbook's cells that use a function show #NAME? after a full recalculation
            // before it is registered; registering doesn't recalculate them, but this does. Not
            // a formula TYPED while its function was unregistered: Excel keeps that #NAME?
            // through every kind of recalculation until it is re-entered (decided 2026-11-25:
            // document it, don't re-enter cells for the user).
            XlCall.Excel(XlCall.xlcCalculateNow);
            return lambdas.Count;
        }

        [ExcelFunction(Description = "SPIKE. Recent named-function registrations.", IsThreadSafe = false, IsVolatile = true)]
        public static object JuliaRegisterLog()
        {
            if (Log.Count == 0) return "(none)";
            var result = new object[Log.Count, 1];
            for (int i = 0; i < Log.Count; i++) result[i, 0] = Log[i];
            return result;
        }

        private static string[] SplitList(object cell)
        {
            string s = cell as string;
            return string.IsNullOrEmpty(s) ? new string[0] : s.Split(ArgSep);
        }

        // (object a1, ..., object an) =>
        //     Forward("<excel name>", "<julia name>", nRequired, argNames, new object[] { a1, ..., an })
        private static LambdaExpression Forwarder(string excelName, string juliaName, int nRequired, string[] argNames)
        {
            int n = argNames.Length;
            var parameters = new ParameterExpression[n];
            for (int k = 0; k < n; k++) parameters[k] = Expression.Parameter(typeof(object), "a" + k);
            Expression body = Expression.Call(
                typeof(NamedFunctions).GetMethod(nameof(Forward), System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic),
                Expression.Constant(excelName),
                Expression.Constant(juliaName),
                Expression.Constant(nRequired),
                Expression.Constant(argNames),
                Expression.NewArrayInit(typeof(object), parameters));
            return Expression.Lambda(body, parameters);
        }

        // Arguments left out at the end of the call are dropped, so Julia uses the defaults of
        // optional ones. Any other argument left out is an error here: Julia would receive
        // missing, which can give a wrong answer rather than an error (the spike's 0), and
        // can't skip a positional argument to use its default. (An empty cell, as opposed to an
        // argument left out, is passed as usual.)
        private static object Forward(string excelName, string juliaName, int nRequired, string[] argNames, object[] args)
        {
            int n = args.Length;
            while (n > 0 && args[n - 1] is ExcelMissing) n--;
            if (n < nRequired)
            {
                return "#" + excelName + " needs at least " + nRequired + (nRequired == 1 ? " argument" : " arguments") + "!";
            }
            for (int k = 0; k < n; k++)
            {
                if (args[k] is ExcelMissing)
                {
                    return "#" + excelName + ": " + argNames[k] + " can't be left out when a later argument is given!";
                }
            }
            // Julia's _call_exposed finds the function by its Excel name and broadcasts as
            // expose_to_excel says; juliaName is kept for this file's log.
            var callArgs = new object[n + 1];
            callArgs[0] = excelName;
            Array.Copy(args, 0, callArgs, 1, n);
            return JuliaFunctions.CallCore("CallFromExcel._call_exposed", callArgs, "/call");
        }
    }
}
