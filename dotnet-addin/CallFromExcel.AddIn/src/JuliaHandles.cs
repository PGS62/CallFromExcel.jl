using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace CallFromExcel.AddIn
{
    // SPIKE (branch spike/object-handles): object handles, see "Object handles" in
    // Enhancements.md and src/handles.jl.
    //
    // Lifetime rests on Excel-DNA's RTD-based ExcelAsyncUtil.Observe - the same mechanism its
    // built-in [ExcelHandle] uses (Excel-DNA 1.9.0's ObjectHandler.GetHandle is a thin layer
    // over it). Observe keys a topic on the function name and its arguments: the first formula
    // to ask for a topic makes Excel-DNA call our observable's Subscribe, which asks Julia to
    // build the object; later recalculations with the same arguments get the same handle
    // without calling Julia; when the last formula using the topic goes, Excel disconnects it
    // and Excel-DNA disposes the subscription, which tells Julia to drop the object.
    //
    // Using Observe directly rather than [return: ExcelHandle]: JuliaCall-style functions take
    // params object[] and untyped arguments, the handle text should come from Julia (type
    // label, session tag), and ObjectHandler.GetHandle(name, params, obj) builds the object
    // before Observe decides whether the topic already exists, which would leak one Julia
    // object per recalculation.
    public static class JuliaHandles
    {
        // Handle text -> key "<id>@<session>" for every handle this add-in has issued and not
        // yet released. Only strings in here are sent to Julia as handles.
        private static readonly Dictionary<string, string> Live = new Dictionary<string, string>();
        // Handle text -> the observer of its topic, to push a new value into the cells using it.
        private static readonly Dictionary<string, IExcelObserver> Observers = new Dictionary<string, IExcelObserver>();

        // Part of every topic's key. Bumped when Julia restarts, so each handle formula asks for
        // a new topic, which creates its object in the new Julia.
        private static int generation;

        // SPIKE: what to do when Julia restarts. Every live handle names an object that no
        // longer exists, and its topic would otherwise keep that dead handle for as long as its
        // arguments are unchanged (completing the topic doesn't help: Excel-DNA then returns
        // its final value for as long as the arguments are unchanged). So: new generation, and
        // push #N/A into each old topic. The RTD update makes Excel recalculate each such cell,
        // which asks for a topic under the new generation (creating the object again, in
        // Excel's dependency order), and the old topic, no longer used, is disposed.
        [ExcelCommand(Name = "JuliaHandleInvalidateAll")]
        public static void InvalidateAll()
        {
            generation++;
            var handles = new List<string>(Observers.Keys);
            foreach (string handle in handles)
            {
                IExcelObserver observer = Observers[handle];
                Observers.Remove(handle);
                Live.Remove(handle);
                Log("invalidate " + handle);
                observer.OnNext(ExcelError.ExcelErrorNA);
            }
        }

        // Create/release events, for checking the spike's scenarios from a sheet.
        private static readonly List<string> Events = new List<string>();
        private static int eventCount;

        internal static bool TryGetKey(string text, out string key)
        {
            key = null;
            return text.StartsWith("jl:", StringComparison.Ordinal) && Live.TryGetValue(text, out key);
        }

        private static void Log(string what)
        {
            eventCount++;
            Events.Add(eventCount + " " + DateTime.Now.ToString("HH:mm:ss.fff") + " " + what);
            if (Events.Count > 500) Events.RemoveAt(0);
        }

        [ExcelFunction(Description = "SPIKE. Calls a Julia function and returns a handle to the result, which stays in Julia while a formula uses it.", IsThreadSafe = false)]
        public static object JuliaHandle(
            [ExcelArgument(Name = "JuliaFunction", Description = "The name of a Julia function.")] string juliaFunction,
            [ExcelArgument(Name = "Args", Description = "Zero or more arguments; handles are replaced by their objects.")] params object[] args)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return "#Disabled in Function Wizard!";
            // An error argument (e.g. an invalidated handle upstream) propagates, as in most
            // Excel functions, rather than creating an object from it.
            foreach (object arg in args) if (arg is ExcelError) return arg;
            var topic = new object[2 + args.Length];
            topic[0] = generation;
            topic[1] = juliaFunction;
            Array.Copy(args, 0, topic, 2, args.Length);
            // Excel-DNA calls the lambda only when no topic matches, so the copying in
            // HandleCreateArgs happens only when Julia is asked to build an object, not on every
            // recalculation that reuses an existing handle.
            return ExcelAsyncUtil.Observe("JuliaHandle", topic,
                () => new HandleObservable("JuliaHandle(" + juliaFunction + ")", "CallFromExcel.handle_create",
                    HandleCreateArgs(juliaFunction, args)));
        }

        // Arguments for handle_create: the function name, then copies of the arguments. Copies,
        // because the payload builder converts arrays in place and Excel-DNA keeps the original
        // arrays (in the topic) to compare against later calls.
        private static object[] HandleCreateArgs(string juliaFunction, object[] args)
        {
            object[] callArgs = new object[1 + args.Length];
            callArgs[0] = juliaFunction;
            for (int i = 0; i < args.Length; i++) callArgs[i + 1] = args[i] is Array a ? a.Clone() : args[i];
            return callArgs;
        }

        [ExcelFunction(Description = "SPIKE. Evaluates Julia code and returns a handle to the result, which stays in Julia while a formula uses it.", IsThreadSafe = false)]
        public static object JuliaEvalHandle(
            [ExcelArgument(Name = "JuliaExpression", Description = "Julia code, as a string.")] string juliaExpression)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return "#Disabled in Function Wizard!";
            return ExcelAsyncUtil.Observe("JuliaEvalHandle", new object[] { generation, juliaExpression },
                () => new HandleObservable("JuliaEvalHandle", "CallFromExcel.handle_eval", new object[] { juliaExpression }));
        }

        [ExcelFunction(Description = "SPIKE. Handle counts: [handles the add-in holds, objects Julia stores].", IsThreadSafe = false, IsVolatile = true)]
        public static object JuliaHandleCount()
        {
            object julia;
            try { julia = JuliaFunctions.CallForHandles("CallFromExcel.handle_count", new object[0]); }
            catch (Exception ex) { julia = "#" + ex.Message; }
            return "addin live " + Live.Count + ", julia stored " + Convert.ToString(julia);
        }

        [ExcelFunction(Description = "SPIKE. Recent handle create/release events, newest last.", IsThreadSafe = false, IsVolatile = true)]
        public static object JuliaHandleEvents()
        {
            if (Events.Count == 0) return "(none)";
            var result = new object[Events.Count, 1];
            for (int i = 0; i < Events.Count; i++) result[i, 0] = Events[i];
            return result;
        }

        [ExcelFunction(Description = "SPIKE. Clears the event log.", IsThreadSafe = false)]
        public static object JuliaHandleEventsClear()
        {
            Events.Clear();
            return "cleared";
        }

        private sealed class HandleObservable : IExcelObservable
        {
            private readonly string description, juliaFunction;
            private readonly object[] args;

            internal HandleObservable(string description, string juliaFunction, object[] args)
            {
                this.description = description;
                this.juliaFunction = juliaFunction;
                this.args = args;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                object result;
                try { result = JuliaFunctions.CallForHandles(juliaFunction, args); }
                catch (Exception ex) { result = "#" + ex.Message; }

                string handle = result as string;
                int hash = handle == null ? -1 : handle.LastIndexOf('#');
                if (handle != null && handle.StartsWith("jl:", StringComparison.Ordinal) && hash > 0)
                {
                    string key = handle.Substring(hash + 1);
                    Live[handle] = key;
                    Observers[handle] = observer;
                    Log("create " + handle + " for " + description);
                    observer.OnNext(handle);
                    return new Release(handle, key);
                }

                // Failed (Julia not running, error in the function...): report it and complete,
                // so the topic ends and the next recalculation tries again rather than keeping
                // the error for as long as the arguments are unchanged.
                Log("failed " + description + ": " + Convert.ToString(result));
                observer.OnNext(result);
                observer.OnCompleted();
                return new Release(null, null);
            }
        }

        private sealed class Release : IDisposable
        {
            private string handle;
            private readonly string key;

            internal Release(string handle, string key)
            {
                this.handle = handle;
                this.key = key;
            }

            public void Dispose()
            {
                if (handle == null) return;
                Live.Remove(handle);
                Observers.Remove(handle);
                string remaining;
                try { remaining = Convert.ToString(JuliaFunctions.CallForHandles("CallFromExcel.handle_free", new object[] { key })); }
                catch (Exception ex) { remaining = "error: " + ex.Message; }
                Log("release " + handle + " (Julia now holds " + remaining + ")");
                handle = null;
            }
        }
    }

    // A handle argument, sent to Julia as the wire-format type "O<id>@<session>".
    public sealed class JuliaHandleRef
    {
        public readonly string Key;
        public JuliaHandleRef(string key) { Key = key; }
    }
}
