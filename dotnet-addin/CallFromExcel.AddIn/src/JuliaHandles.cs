using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace CallFromExcel.AddIn
{
    // Object handles (documentation, "Object handles"; Julia side in src/handles.jl).
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
        // Handle text -> the observer of its topic, to push a new value into the cells using it.
        // One entry per live worksheet handle. Which strings are handles is decided by their
        // text (HandlePattern), not by this table, so that handles created from VBA - by the
        // COM server, which can't see this table - work in formulas too.
        private static readonly Dictionary<string, IExcelObserver> Observers = new Dictionary<string, IExcelObserver>();

        // Part of every topic's key. Bumped when Julia restarts, so each handle formula asks for
        // a new topic, which creates its object in the new Julia.
        private static int generation;

        // What to do when Julia restarts. Every live handle names an object that no
        // longer exists, and its topic would otherwise keep that dead handle for as long as its
        // arguments are unchanged. Completing the topic now doesn't help: once a topic has
        // delivered a value, Excel-DNA keeps it and returns its final value for as long as the
        // arguments are unchanged (checked 2026-11-23). That differs from a failed create in
        // Subscribe below, which completes before the first value is returned, so Excel-DNA
        // keeps no topic and the next recalculation tries again. So: new generation, and
        // push #N/A into each old topic. The RTD update makes Excel recalculate each such cell,
        // which asks for a topic under the new generation (creating the object again, in
        // Excel's dependency order), and the old topic, no longer used, is disposed.
        // Called after JuliaLaunch (the worksheet function) and the ribbon's Launch Julia button
        // - not from the shared JuliaLaunch.Launch, which the COM server also calls from its own
        // copy of this assembly, without Excel-DNA. Asks Julia for its session tag, so a
        // launch that found the same Julia already running changes nothing, while a Julia
        // restarted any way (closed and relaunched, from VBA, or serve_xl from a REPL) is
        // caught here. Queued as a macro, to run after the calculation JuliaLaunch may be in.
        internal static void AfterLaunch()
        {
            if (Observers.Count == 0) return;
            string session;
            try { session = JuliaFunctions.CallForHandles("CallFromExcel.handle_session", new object[0]) as string; }
            catch (Exception) { return; }
            // No answer (e.g. Julia still starting): the first use of a stale handle gives
            // "not found", and the next JuliaLaunch tries again.
            if (string.IsNullOrEmpty(session) || session.StartsWith("#", StringComparison.Ordinal)) return;
            foreach (string handle in Observers.Keys)
            {
                if (!handle.EndsWith("@" + session, StringComparison.Ordinal))
                {
                    Log("Julia session is now " + session + ": invalidating handles");
                    ExcelAsyncUtil.QueueAsMacro(InvalidateAll);
                    return;
                }
            }
        }

        [ExcelCommand(Name = "JuliaHandleInvalidateAll")]
        public static void InvalidateAll()
        {
            generation++;
            var handles = new List<string>(Observers.Keys);
            foreach (string handle in handles)
            {
                IExcelObserver observer = Observers[handle];
                Observers.Remove(handle);
                Log("invalidate " + handle);
                observer.OnNext(ExcelError.ExcelErrorNA);
            }
        }

        // Create/release events, for diagnostics and the drivers in dotnet-addin/spike-handles.
        private static readonly List<string> Events = new List<string>();
        private static int eventCount;

        private static void Log(string what)
        {
            eventCount++;
            Events.Add(eventCount + " " + DateTime.Now.ToString("HH:mm:ss.fff") + " " + what);
            if (Events.Count > 500) Events.RemoveAt(0);
        }

        // Handle formulas whose arguments change at every calculation - volatile functions such
        // as RANDARRAY, NOW or OFFSET, directly or in cells they depend on. Each new handle's
        // topic posts an RTD update as it connects (Excel-DNA's ExcelRtdServer.ConnectData does
        // so whenever the topic has a value), which in automatic calculation recalculates the
        // cell; its arguments are then new, so another handle is created, and so on for ever.
        // An observable that fails also posts that update, so the loop is stopped by returning an
        // error WITHOUT calling Observe: with no topic, there is no update.
        //
        // A loop is recognised per calling cell: LoopRepeats new handles in a row, each created
        // within one RTD throttle interval (plus a second) of the previous one finishing - the
        // time the update takes to come back. Manual calculation never loops (an RTD update
        // doesn't recalculate), so its recalculations, however quick, aren't counted.
        private sealed class CreateHistory
        {
            internal DateTime LastEnd;      // when this cell's last handle was created
            internal int QuickRepeats;      // new handles in a row, each soon after the last
            internal DateTime LoopSeen;     // when the error was last returned; MinValue if never
        }
        private static readonly Dictionary<string, CreateHistory> History = new Dictionary<string, CreateHistory>();
        private const int LoopRepeats = 3;

        // The calling cell, e.g. "123!4,5" (sheet id, row, column); null if not called from a cell.
        private static string CallerKey()
        {
            try
            {
                return XlCall.Excel(XlCall.xlfCaller) is ExcelReference r
                    ? r.SheetId.ToString() + "!" + r.RowFirst + "," + r.ColumnFirst
                    : null;
            }
            catch (Exception) { return null; }
        }

        // How soon after a handle is created a looping cell asks for the next one, at most.
        private static double QuickGapMs()
        {
            try
            {
                const System.Reflection.BindingFlags get = System.Reflection.BindingFlags.GetProperty;
                object app = ExcelDnaUtil.Application;
                int calculation = Convert.ToInt32(app.GetType().InvokeMember("Calculation", get, null, app, null));
                if (calculation == -4135) return -1;   // xlCalculationManual: never quick
                object rtd = app.GetType().InvokeMember("RTD", get, null, app, null);
                int throttle = Convert.ToInt32(rtd.GetType().InvokeMember("ThrottleInterval", get, null, rtd, null));
                return Math.Max(0, throttle) + 1000;
            }
            catch (Exception) { return 3000; }
        }

        // Before Observe: true while the caller is looping, so the function returns an error.
        // Stays true while the cell keeps being recalculated soon after each error; once that
        // stops (e.g. the formula no longer uses a volatile function, or the next recalculation
        // is a while later), handles are created again.
        private static bool IsLooping(string caller)
        {
            if (caller == null || !History.TryGetValue(caller, out CreateHistory h) || h.LoopSeen == DateTime.MinValue) return false;
            DateTime now = DateTime.UtcNow;
            double quick = QuickGapMs();
            if (quick >= 0 && (now - h.LoopSeen).TotalMilliseconds < quick)
            {
                h.LoopSeen = now;
                return true;
            }
            History.Remove(caller);
            return false;
        }

        // In Subscribe, after a new handle is created for the caller.
        private static void NoteCreate(string caller, DateTime started)
        {
            if (caller == null) return;
            if (!History.TryGetValue(caller, out CreateHistory h))
            {
                History[caller] = h = new CreateHistory { LoopSeen = DateTime.MinValue };
                if (History.Count > 10000) History.Clear();   // cells deleted or moved; keep it bounded
            }
            else
            {
                double quick = QuickGapMs();
                h.QuickRepeats = quick >= 0 && (started - h.LastEnd).TotalMilliseconds < quick ? h.QuickRepeats + 1 : 0;
            }
            h.LastEnd = DateTime.UtcNow;
            if (h.QuickRepeats >= LoopRepeats)
            {
                h.LoopSeen = h.LastEnd;
                Log("loop: " + caller + " creates a new handle at every recalculation");
            }
        }

        private const string CallLoopError = "#JuliaCallHandle: its arguments change at every recalculation (RAND, NOW, OFFSET...?), so it can't keep a handle - see Help";
        private const string EvalLoopError = "#JuliaEval: its expression changes at every recalculation (RAND, NOW, OFFSET...?), so it can't keep a handle - see Help";

        // Covers both JuliaCall and JuliaCallNamedArgs, so there is no JuliaCallNamedArgsHandle:
        // handle_create treats the arguments as JuliaCallNamedArgs does if one of them is ";",
        // else as JuliaCall does. The cost: a positional argument can't be the string ";".
        [ExcelFunction(Description = "Call a named Julia function and return a handle to the result, which stays in Julia for other functions to use. An argument \";\" marks where named arguments start, as in JuliaCallNamedArgs.", HelpTopic = JuliaFunctions.HelpRoot + "juliacallhandle", IsThreadSafe = false)]
        public static object JuliaCallHandle(
            [ExcelArgument(Name = "JuliaFunction", Description = "The name of a Julia function visible from the Julia REPL. A trailing \".\" broadcasts, as for JuliaCall.")] string juliaFunction,
            [ExcelArgument(Name = "Args", Description = "Positional arguments, optionally followed by \";\" and name/value pairs for named arguments. Handles among them are replaced by their objects.")] params object[] args)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return "#Disabled in Function Wizard!";
            // An error argument (e.g. an invalidated handle upstream) propagates, as in most
            // Excel functions, rather than creating an object from it.
            foreach (object arg in args) if (arg is ExcelError) return arg;
            string caller = CallerKey();
            if (IsLooping(caller)) return CallLoopError;
            var topic = new object[2 + args.Length];
            topic[0] = generation;
            topic[1] = juliaFunction;
            Array.Copy(args, 0, topic, 2, args.Length);
            // Excel-DNA calls the lambda only when no topic matches, so the copying in
            // HandleCreateArgs happens only when Julia is asked to build an object, not on every
            // recalculation that reuses an existing handle.
            return ExcelAsyncUtil.Observe("JuliaCallHandle", topic,
                () => new HandleObservable("JuliaCallHandle(" + juliaFunction + ")", "CallFromExcel.handle_create",
                    HandleCreateArgs(juliaFunction, args), caller));
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

        // JuliaEval(expression, TRUE). The topic is keyed on the expression text after
        // JuliaEval has joined a range of cells into one string.
        internal static object EvalHandle(string expression)
        {
            string caller = CallerKey();
            if (IsLooping(caller)) return EvalLoopError;
            return ExcelAsyncUtil.Observe("JuliaEvalHandle", new object[] { generation, expression },
                () => new HandleObservable("JuliaEval(ReturnHandle)", "CallFromExcel.handle_eval", new object[] { expression }, caller, mapHandles: false));
        }

        // Info types JuliaHandleInfo recognises, as listed in its error for any other.
        private static readonly string[] InfoTypes = { "count", "events" };

        // One function for information about handles, in the style of Excel's INFO(type_text).
        [ExcelFunction(Description = "Information about object handles: how many are live, or recent events.", HelpTopic = JuliaFunctions.HelpRoot + "juliahandleinfo", IsThreadSafe = false, IsVolatile = true)]
        public static object JuliaHandleInfo(
            [ExcelArgument(Name = "Type", Description = "\"count\" for the number of live handles created by worksheet formulas, and of objects Julia holds (including those created from VBA). \"events\" for a column of recent create and release events, newest last.")] object type)
        {
            string t = type is string s ? s.Trim().ToLowerInvariant() : "";
            if (t == "count")
            {
                object julia;
                try { julia = JuliaFunctions.CallForHandles("CallFromExcel.handle_count", new object[0]); }
                catch (Exception ex) { julia = "#" + ex.Message; }
                return "worksheet handles " + Observers.Count + ", julia stored " + Convert.ToString(julia);
            }
            if (t == "events")
            {
                if (Events.Count == 0) return "(none)";
                var result = new object[Events.Count, 1];
                for (int i = 0; i < Events.Count; i++) result[i, 0] = Events[i];
                return result;
            }
            return "#Type must be one of: \"" + string.Join("\", \"", InfoTypes) + "\"";
        }

        // Hidden: a worksheet function that changes state is a poor fit for the documented set,
        // but the drivers in dotnet-addin/spike-handles use it to see each step's events on their own.
        [ExcelFunction(Description = "Clears the log of handle events that JuliaHandleInfo(\"events\") shows.", IsThreadSafe = false, IsHidden = true)]
        public static object JuliaHandleEventsClear()
        {
            Events.Clear();
            return "cleared";
        }

        private sealed class HandleObservable : IExcelObservable
        {
            private readonly string description, juliaFunction, caller;
            private readonly object[] args;
            // False for code to evaluate, which must reach Julia as text even if it looks like
            // a handle.
            private readonly bool mapHandles;

            internal HandleObservable(string description, string juliaFunction, object[] args, string caller, bool mapHandles = true)
            {
                this.description = description;
                this.juliaFunction = juliaFunction;
                this.args = args;
                this.caller = caller;
                this.mapHandles = mapHandles;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                DateTime started = DateTime.UtcNow;
                object result;
                try { result = JuliaFunctions.CallForHandles(juliaFunction, args, mapHandles); }
                catch (Exception ex) { result = "#" + ex.Message; }

                string handle = result as string;
                if (HandlePattern.TryGetKey(handle, out string key))
                {
                    Observers[handle] = observer;
                    Log("create " + handle + " for " + description);
                    NoteCreate(caller, started);
                    observer.OnNext(handle);
                    return new Release(handle, key);
                }

                // Failed (Julia not running, error in the function...): report it and complete.
                // Completing here, inside the first Subscribe, means Excel-DNA keeps no topic,
                // so the next recalculation tries again rather than keeping the error for as
                // long as the arguments are unchanged (tested: a create before JuliaLaunch
                // succeeds on the next recalculation after it). Completing a topic that has
                // already delivered a value is different - see InvalidateAll.
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
                Observers.Remove(handle);
                string remaining;
                try { remaining = Convert.ToString(JuliaFunctions.CallForHandles("CallFromExcel.handle_free", new object[] { key })); }
                catch (Exception ex) { remaining = "error: " + ex.Message; }
                Log("release " + handle + " (Julia now holds " + remaining + ")");
                handle = null;
            }
        }
    }
}
