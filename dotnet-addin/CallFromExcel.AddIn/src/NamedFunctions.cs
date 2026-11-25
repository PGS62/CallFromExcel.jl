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

        // Registration needs a macro context, so this is a command (Alt+F8, or Application.Run).
        [ExcelCommand(Name = "JuliaRegisterFunctions")]
        public static void RegisterAll()
        {
            object list;
            try { list = JuliaFunctions.CallForHandles("CallFromExcel.exposed_functions", new object[0]); }
            catch (Exception ex) { Log.Add("fetch failed: " + ex.Message); return; }
            var rows = list as object[,];
            if (rows == null) { Log.Add("nothing to register: " + Convert.ToString(list)); return; }

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

                lambdas.Add(Forwarder(juliaName, argNames.Length));
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
            Log.Add("registered " + lambdas.Count);
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

        // (object a1, ..., object an) => Forward("<julia name>", new object[] { a1, ..., an })
        private static LambdaExpression Forwarder(string juliaName, int n)
        {
            var parameters = new ParameterExpression[n];
            for (int k = 0; k < n; k++) parameters[k] = Expression.Parameter(typeof(object), "a" + k);
            Expression body = Expression.Call(
                typeof(NamedFunctions).GetMethod(nameof(Forward), System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic),
                Expression.Constant(juliaName),
                Expression.NewArrayInit(typeof(object), parameters));
            return Expression.Lambda(body, parameters);
        }

        private static object Forward(string juliaName, object[] args)
        {
            return JuliaFunctions.CallCore(juliaName, args, "/call");
        }
    }
}
