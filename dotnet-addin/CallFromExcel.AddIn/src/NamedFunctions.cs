using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelDna.Integration;

namespace CallFromExcel.AddIn
{
    // Named worksheet functions (README, "Your own worksheet functions"; Julia side in
    // src/expose.jl). Fetches the functions Julia has exposed with @excel or expose_to_excel
    // and registers each as a worksheet function - by default named "jl.<julia name>" - that
    // forwards its arguments to Julia's CallFromExcel._call_exposed.
    public static class NamedFunctions
    {
        private const char ArgSep = '\x1f';

        // Recent registrations, newest last, for the hidden JuliaRegisterLog (diagnostics and
        // the drivers in dotnet-addin/spike-named). Capped, as it grows on every registration.
        private static readonly List<string> Log = new List<string>();
        private const int LogLimit = 200;

        private static void AddToLog(string entry)
        {
            Log.Add(entry);
            if (Log.Count > LogLimit) Log.RemoveRange(0, Log.Count - LogLimit);
        }

        // Julia's exposed_version() when the add-in last registered: "<session>#<count>".
        private static string registeredVersion = "";

        // What each Excel name was last registered with in this Excel session (descriptions,
        // argument names and descriptions, number required), so Register skips the unchanged.
        private static readonly Dictionary<string, string> registeredSignatures = new Dictionary<string, string>();

        // A registration is queued and hasn't started: several checks in one calculation (a
        // JuliaLaunch and a JuliaEval cell, say) queue it once. Main thread only.
        private static bool registrationQueued;

        // Called after JuliaLaunch (worksheet and ribbon), JuliaInclude (worksheet and ribbon)
        // and JuliaEval (including with ReturnHandle), which may have run code calling
        // expose_to_excel. Doesn't call Julia itself: it queues one check, which needs a macro
        // context to register anyway, so a calculation with many JuliaEval cells checks once,
        // when Excel is next idle, rather than once per cell.
        internal static void AfterJuliaChange()
        {
            if (registrationQueued) return;
            registrationQueued = true;
            ExcelAsyncUtil.QueueAsMacro(RegisterIfChanged);
        }

        // The queued check: registers if Julia's exposed_version() differs from the version last
        // registered. The JuliaRegisterFunctions command (as VBA code calls it) may already have
        // registered it, and a JuliaEval that exposes nothing leaves it unchanged.
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
            if (version == null) { AddToLog("no answer from Julia"); return 0; }
            object list;
            try { list = JuliaFunctions.CallForHandles("CallFromExcel.exposed_functions", new object[0]); }
            catch (Exception ex) { AddToLog("fetch failed: " + ex.Message); return 0; }
            // Julia returns nothing (here ExcelEmpty) when no function is exposed. Anything else
            // that isn't a table is a failure - an error text - so the version isn't recorded as
            // registered, and the next check tries again.
            if (list is ExcelEmpty)
            {
                AddToLog("nothing to register (" + version + ")");
                registeredVersion = version;
                return 0;
            }
            var rows = list as object[,];
            if (rows == null)
            {
                AddToLog("fetch failed: " + Convert.ToString(list));
                return 0;
            }

            var lambdas = new List<LambdaExpression>();
            var functionAttributes = new List<object>();
            var argumentAttributes = new List<List<object>>();
            var newSignatures = new Dictionary<string, string>();
            for (int i = 0; i < rows.GetLength(0); i++)
            {
                string excelName = Convert.ToString(rows[i, 0]);
                string juliaName = Convert.ToString(rows[i, 1]);
                string description = Convert.ToString(rows[i, 2]);
                string[] argNames = SplitList(rows[i, 3]);
                string[] argDescriptions = SplitList(rows[i, 4]);
                int nRequired = Convert.ToInt32(rows[i, 5]);

                // Only what's new or changed. Each registration takes one of Excel-DNA's export
                // slots, which aren't freed when a name is registered again: re-registering 50
                // functions 300 times (15,000 registrations) left none for a new function, which
                // then gave #NAME?, and each round got slower (checked 2026-11-25). The Julia
                // name isn't part of it: calls are forwarded by Excel name.
                string signature = string.Join("\x1e", description, Convert.ToString(rows[i, 3]),
                    Convert.ToString(rows[i, 4]), nRequired.ToString());
                if (registeredSignatures.TryGetValue(excelName, out string old) && old == signature) continue;
                newSignatures[excelName] = signature;

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
                AddToLog("register " + excelName + "(" + string.Join(", ", argNames) + ") -> " + juliaName);
            }
            if (lambdas.Count == 0)
            {
                registeredVersion = version;
                AddToLog("nothing new or changed (" + version + ")");
                return 0;
            }
            ExcelIntegration.RegisterLambdaExpressions(lambdas, functionAttributes, argumentAttributes);
            foreach (var kv in newSignatures) registeredSignatures[kv.Key] = kv.Value;
            // IntelliSense reads the registered functions when the add-in loads; without this it
            // doesn't show the arguments of functions registered later (the Function Wizard
            // and Ctrl+Shift+A, which ask Excel, do).
            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();
            registeredVersion = version;
            AddToLog("registered " + lambdas.Count + " (" + version + ")");
            // A saved workbook's cells that use a function show #NAME? after a full recalculation
            // before it is registered; registering doesn't recalculate them, but this does. Not
            // a formula TYPED while its function was unregistered: Excel keeps that #NAME?
            // through every kind of recalculation until it is re-entered (decided 2026-11-25:
            // document it, don't re-enter cells for the user). Not in manual calculation mode,
            // whose user chooses when to recalculate - such cells then recover at the next F9.
            if (!CalculationIsManual()) XlCall.Excel(XlCall.xlcCalculateNow);
            return lambdas.Count;
        }

        // GET.DOCUMENT(14): 1 automatic, 2 automatic except tables, 3 manual.
        private static bool CalculationIsManual()
        {
            try { return Convert.ToInt32(XlCall.Excel(XlCall.xlfGetDocument, 14)) == 3; }
            catch (Exception) { return false; }
        }

        [ExcelFunction(Description = "Recent registrations of named worksheet functions, for diagnostics.", IsThreadSafe = false, IsVolatile = true, IsHidden = true)]
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
