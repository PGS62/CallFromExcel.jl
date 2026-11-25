using System;
using System.Text.RegularExpressions;

namespace CallFromExcel.AddIn
{
    // Recognises object handles by their text, so that a handle can be passed to Julia from a
    // worksheet formula or from VBA, whichever created it. Uses no Excel-DNA types: the COM
    // server runs in a separate copy of this assembly (registered via mscoree, loaded from the
    // DLL file), which can't see the worksheet functions' state (loaded by Excel-DNA from the
    // packed .xll), and may run where ExcelDna.Integration.dll can't be resolved.
    //
    // A handle is "jl:<label>#<id>@<session>" (see src/handles.jl). The label is the type's
    // name, or whatever a method of CallFromExcel.handle_label returns, with "#", "@",
    // whitespace, control and format characters replaced by "_" (_clean_label), so it may
    // contain any other characters, including non-ASCII ones. The session is random letters
    // and digits (4 of them). Any string of exactly that form is sent to Julia as a handle,
    // and Julia, which knows which objects are alive, looks it up or raises "not found". A
    // look-alike string would have to name a live session tag to pick up an object, so the
    // realistic outcome of one is that error, not a wrong answer.
    public static class HandlePattern
    {
        private static readonly Regex Pattern =
            // \z, not $: in .NET, $ also matches before a final line feed, so a handle pasted with
            // a trailing line feed would count as one. Every character excluded from the label
            // here must be one _clean_label replaces, or Julia could make a handle this doesn't
            // recognise. (.NET's \s is Unicode whitespace, including \p{Z}; Julia's _clean_label
            // replaces both.)
            new Regex(@"^jl:[^#@\s\p{Cc}\p{Cf}]*#([0-9]+@[A-Za-z0-9]+)\z", RegexOptions.CultureInvariant);

        // The key "<id>@<session>" of a handle, or false if text isn't one.
        public static bool TryGetKey(string text, out string key)
        {
            key = null;
            // Cheap test first: most strings passed to Julia are not handles.
            if (text == null || !text.StartsWith("jl:", StringComparison.Ordinal)) return false;
            Match m = Pattern.Match(text);
            if (!m.Success) return false;
            key = m.Groups[1].Value;
            return true;
        }

        // value with each handle string replaced by a JuliaHandleRef, which the encoder sends
        // as the wire type "O<id>@<session>". For the COM server's arguments: strings,
        // arrays of Variants (any rank and lower bounds) and arrays nested in them. Arrays
        // are changed in place - COM marshals a VBA array, and Range.Value2 returns, a fresh
        // copy. Typed arrays (Double, String...) are left alone: a handle in a VBA String
        // array is rare enough not to be worth a copy of every String array.
        public static object MapHandles(object value)
        {
            if (value is string s) return TryGetKey(s, out string key) ? new JuliaHandleRef(key) : value;
            var arr = value as Array;
            if (arr == null || arr.GetType().GetElementType() != typeof(object)) return value;

            if (arr is object[,] m)
            {
                int lb0 = m.GetLowerBound(0), ub0 = m.GetUpperBound(0);
                int lb1 = m.GetLowerBound(1), ub1 = m.GetUpperBound(1);
                for (int i = lb0; i <= ub0; i++)
                    for (int j = lb1; j <= ub1; j++)
                        m[i, j] = MapHandles(m[i, j]);
                return m;
            }
            if (arr is object[] v)
            {
                for (int i = 0; i < v.Length; i++) v[i] = MapHandles(v[i]);
                return v;
            }

            // Other shapes: a 1-D array not starting at 0 (VBA's usual Dim x(1 To n)), rank 3+.
            int rank = arr.Rank;
            var idx = new int[rank];
            for (int d = 0; d < rank; d++) idx[d] = arr.GetLowerBound(d);
            for (int count = 0; count < arr.Length; count++)
            {
                arr.SetValue(MapHandles(arr.GetValue(idx)), idx);
                for (int d = rank - 1; d >= 0; d--)
                {
                    if (++idx[d] <= arr.GetUpperBound(d)) break;
                    idx[d] = arr.GetLowerBound(d);
                }
            }
            return arr;
        }
    }

    // A handle argument, sent to Julia as the wire-format type "O<id>@<session>". Here rather
    // than in JuliaHandles.cs so the COM server can use it without Excel-DNA's types.
    public sealed class JuliaHandleRef
    {
        public readonly string Key;
        public JuliaHandleRef(string key) { Key = key; }
    }
}
