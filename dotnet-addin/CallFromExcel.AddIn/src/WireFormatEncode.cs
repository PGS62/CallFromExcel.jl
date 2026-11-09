using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace CallFromExcel.AddIn
{
    // C# port of SerialiseElement/TrySerialiseArrayAsV (modSerialise.bas) -- encodes a
    // single .NET value (scalar or array) into the CallFromExcel wire format. See
    // WireFormatDecode.cs for the full format specification (mirrored from
    // modUnserialise.bas's own wire-format spec comment) and the corresponding decoder.
    //
    // .NET <-> VBA Variant type mapping used throughout:
    //   null       <-> Empty   (standard COM interop convention: a null object reference
    //                           marshals to/from VT_EMPTY, matching how `object` COM
    //                           parameters/returns already behave)
    //   DBNull.Value <-> Null  (standard COM interop convention: DBNull marshals to/from
    //                           VT_NULL)
    //   double, string, bool, short (Integer), int (Long), byte, float (Single),
    //   long (LongLong), DateTime, decimal (Currency/Decimal, decode only -- see below)
    //   ExcelErrorValue <-> a VBA error Variant (CVErr)
    //   JuliaDictionary (decoded internally; ComServer converts it for VBA) <-> Scripting.Dictionary
    //   System.Array (any rank, any of the above element types) <-> a VBA array
    public static class WireFormatEncode
    {
        public static string Encode(object value)
        {
            if (value == null) return "E";
            if (value is Array) return EncodeArray((Array)value);

            if (value is double) return "#" + Hex.DoubleToHex((double)value);
            if (value is string) return (char)163 + (string)value;
            if (value is bool) return ((bool)value) ? "T" : "F";
            if (value is DBNull) return "N";
            if (value is short) return "%" + ((short)value).ToString(CultureInfo.InvariantCulture);
            if (value is int)
            {
                // A VBA CVErr(code) Variant (VT_ERROR) crossing COM into .NET arrives boxed
                // as a plain Int32, not as ExcelErrorValue - .NET's default Variant<->Object
                // marshaling has no distinct managed type for VT_ERROR. The underlying SCODE
                // is exactly 0x800A0000 | code (see ComServer.MapToComError's own note on
                // this, which performs the reverse conversion for results going the other
                // way), so an incoming int with that upper-16-bits pattern can only
                // plausibly be a disguised CVErr value - a genuine Long data value landing on
                // one of these exact facility-code bit patterns is not a realistic
                // collision. Checked ahead of the plain "&" Long encoding below.
                int iv = (int)value;
                if (((uint)iv & 0xFFFF0000u) == 0x800A0000u)
                {
                    return "!" + ((uint)iv & 0xFFFFu).ToString(CultureInfo.InvariantCulture);
                }
                return "&" + iv.ToString(CultureInfo.InvariantCulture);
            }
            if (value is byte) return "B" + ((byte)value).ToString(CultureInfo.InvariantCulture);
            if (value is float) return "S" + Hex.SingleToHex((float)value);
            if (value is long) return "^" + ((long)value).ToString(CultureInfo.InvariantCulture);
            if (value is ExcelErrorValue) return "!" + ((ExcelErrorValue)value).Code.ToString(CultureInfo.InvariantCulture);
            if (value is JuliaDictionary) return EncodeDictionary((JuliaDictionary)value);
            if (value is DateTime)
            {
                double serial = DateSerial.FromDateTime((DateTime)value);
                return serial == Math.Floor(serial)
                    ? "D" + ((long)serial).ToString(CultureInfo.InvariantCulture)
                    : "G" + Hex.DoubleToHex(serial);
            }
            if (value is decimal)
            {
                // Currency/Decimal are reserved, not currently implemented in Julia's
                // encode_for_xl -- and SerialiseElement (modSerialise.bas) itself has no
                // case for them either (they'd hit its own "Case Else: Throw"). Decode
                // support exists (WireFormatDecode.cs) for forward-compatibility/fidelity,
                // but nothing produces one today, matching the VBA original exactly.
                throw new NotSupportedException("Cannot serialise Currency/Decimal values - reserved, not implemented (matches SerialiseElement's own omission)");
            }

            // A Scripting.Dictionary passed in from VBA arrives as a generic __ComObject RCW
            // (this project doesn't reference the Scripting library), so it's read via
            // late-bound reflection of its Count/Keys()/Item() members.
            if (value.GetType().IsCOMObject) return EncodeComDictionary(value);

            throw new NotSupportedException("Cannot serialise value of type " + value.GetType().Name);
        }

        private static string EncodeDictionary(JuliaDictionary dict)
        {
            return EncodeDictionaryEntries(dict.Entries, dict.Count);
        }

        private static string EncodeComDictionary(object comObj)
        {
            Type t = comObj.GetType();
            int n;
            try
            {
                n = Convert.ToInt32(t.InvokeMember("Count", BindingFlags.GetProperty, null, comObj, null));
            }
            catch (Exception)
            {
                throw new NotSupportedException("Cannot serialise value of type " + t.Name);
            }
            if (n == 0) return "H0;;";

            Array keys = (Array)t.InvokeMember("Keys", BindingFlags.InvokeMethod, null, comObj, null);
            var entries = new List<KeyValuePair<object, object>>(n);
            foreach (object key in keys)
            {
                object val = t.InvokeMember("Item", BindingFlags.GetProperty, null, comObj, new[] { key });
                entries.Add(new KeyValuePair<object, object>(key, val));
            }
            return EncodeDictionaryEntries(entries, n);
        }

        private static string EncodeDictionaryEntries(IEnumerable<KeyValuePair<object, object>> entries, int n)
        {
            if (n == 0) return "H0;;";
            var encoded = new List<string>(n * 2);
            foreach (var kv in entries)
            {
                encoded.Add(Encode(kv.Key));
                encoded.Add(Encode(kv.Value));
            }
            var lens = new StringBuilder();
            foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }
            return "H" + n.ToString(CultureInfo.InvariantCulture) + ";" + lens + ";" + string.Concat(encoded);
        }

        private static string EncodeArray(Array arr)
        {
            int rank = arr.Rank;
            if (rank > 9) throw new NotSupportedException("Cannot serialise arrays with more than 9 dimensions");

            var dims = new int[rank];
            var lb = new int[rank];
            int total = 1;
            for (int d = 0; d < rank; d++)
            {
                lb[d] = arr.GetLowerBound(d);
                dims[d] = arr.GetUpperBound(d) - lb[d] + 1;
                total *= dims[d];
            }

            // Nx1 2-D arrays collapse to the 1-D wire representation, matching VBA's own
            // SerialiseElement (its own comment: "matches JuliaCallOld / README single-
            // column ranges arrive as vectors"). 3-D+ arrays are left as-is.
            bool collapseTo1D = rank == 2 && dims[1] == 1;
            int effectiveRank = collapseTo1D ? 1 : rank;
            int[] effectiveDims = collapseTo1D ? new[] { dims[0] } : dims;
            string dimStr = DimsToString(effectiveDims);

            if (total == 0)
            {
                return "*" + effectiveRank.ToString(CultureInfo.InvariantCulture) + "," + dimStr + ";;";
            }

            string vResult;
            if (TryEncodeArrayAsV(arr, rank, dims, lb, total, effectiveRank, dimStr, out vResult))
            {
                return vResult;
            }

            var encoded = new string[total];
            // Fast paths for a genuine (zero-based) object[]/object[,] -- by far the common
            // case: Excel always hands worksheet ranges to Excel-DNA this way -- use direct
            // typed indexer access. Array.GetValue(int[]) -- needed for everything else,
            // including any non-zero-lower-bound array (e.g. a VBA Dim x(1 To n) array
            // arriving over COM, which is a *different* runtime type from object[]/object[,]
            // and cannot be cast to either, whatever its rank -- .NET's own "[*]" suffix on
            // such a type's name denotes exactly this) and any genuinely-typed array (a VBA
            // Double()/Long() array arrives as double[]/int[], never boxed into object[]) --
            // measured as a dominant cost for large arrays (a real regression found via this
            // project's own end-to-end JuliaCall("sort", ...) timing: ~800-1750ms through
            // this path versus ~330-450ms through the csharp-prototype's rank-specific
            // VFormat.cs, for the same 1,000,000-element input) -- generic per-element
            // Array.GetValue/SetValue is measurably slower than direct [,] indexing,
            // consistent with .NET's own well-known multi-dimensional-array performance
            // characteristics. That cost only matters for the XLL path though (Excel-DNA
            // always hands over a zero-based object[]/object[,]), so the slower fallback
            // being used for every COM/VBA-sourced array is an acceptable trade, not a
            // regression.
            //
            // Note: object[] is always zero-based (a CLR SZArray guarantee - a non-zero-
            // lower-bound 1-D array is a genuinely different runtime type, never matching
            // "is object[]"), so lb[0] is always 0 here. object[,] has no such guarantee
            // though - a COM-marshaled VBA 2-D array (almost always 1-based in this
            // codebase's own convention) still matches "is object[,]" while having non-zero
            // bounds, so this still has to index via lb[], not from 0.
            if (arr is object[] a1)
            {
                int k = 0;
                for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++) encoded[k++] = Encode(a1[i]);
            }
            else if (arr is object[,] a2)
            {
                int k = 0;
                for (int j = lb[1]; j <= lb[1] + dims[1] - 1; j++)
                    for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++)
                        encoded[k++] = Encode(a2[i, j]);
            }
            else
            {
                var idx = (int[])lb.Clone();
                for (int k = 0; k < total; k++)
                {
                    encoded[k] = Encode(arr.GetValue(idx));
                    AdvanceIndex(idx, lb, dims, rank);
                }
            }

            var lens = new StringBuilder();
            foreach (string e in encoded) { lens.Append(e.Length); lens.Append(','); }

            return "*" + effectiveRank.ToString(CultureInfo.InvariantCulture) + "," + dimStr + ";" + lens + ";" + string.Concat(encoded);
        }

        // Port of TrySerialiseArrayAsV (modSerialise.bas): optimistic single pass over the
        // array, bailing (returning false) the instant a non-Double element is found, so
        // the caller falls back to the general per-element encoding unchanged. See the
        // fast-path note above EncodeArray's own rank 1/2 handling for why this also
        // special-cases rank 1/2 rather than using Array.GetValue uniformly.
        private static bool TryEncodeArrayAsV(Array arr, int rank, int[] dims, int[] lb, int total, int effectiveRank, string dimStr, out string result)
        {
            result = null;
            var buf = new double[total];

            // See EncodeArray's fast-path note: object[] is always zero-based, but object[,]
            // is not (a COM-marshaled VBA 2-D array still matches "is object[,]" while having
            // non-zero bounds), so both still have to index via lb[]. Anything that isn't one
            // of these two types (non-zero-lower-bound 1-D, or a genuinely-typed
            // double[]/double[,] etc. - either possible from a VBA array arriving over COM)
            // falls back to Array.GetValue, which handles both correctly.
            if (arr is object[] a1)
            {
                int k = 0;
                for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++)
                {
                    object element = a1[i];
                    if (!(element is double)) return false;
                    buf[k++] = (double)element;
                }
            }
            else if (arr is object[,] a2)
            {
                int k = 0;
                for (int j = lb[1]; j <= lb[1] + dims[1] - 1; j++)
                {
                    for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++)
                    {
                        object element = a2[i, j];
                        if (!(element is double)) return false;
                        buf[k++] = (double)element;
                    }
                }
            }
            else
            {
                var idx = (int[])lb.Clone();
                for (int k = 0; k < total; k++)
                {
                    object element = arr.GetValue(idx);
                    if (!(element is double)) return false;
                    buf[k] = (double)element;
                    AdvanceIndex(idx, lb, dims, rank);
                }
            }

            string header = "V" + effectiveRank.ToString(CultureInfo.InvariantCulture) + "," + dimStr;
            var payload = new char[total * 16];
            Hex.WriteBulkHex(buf, payload, 0);
            result = header + ";" + new string(payload);
            return true;
        }

        // Advances idx in column-major order (dimension 0 fastest), matching Julia's array
        // layout and VBA's own index-walking loops (SerialiseElement's Case Else, and
        // TrySerialiseArrayAsV's own copy of the same walk).
        private static void AdvanceIndex(int[] idx, int[] lb, int[] dims, int rank)
        {
            for (int d = 0; d < rank; d++)
            {
                idx[d]++;
                if (idx[d] <= lb[d] + dims[d] - 1) return;
                idx[d] = lb[d];
            }
        }

        private static string DimsToString(int[] dims)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < dims.Length; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(dims[i].ToString(CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }
    }
}
