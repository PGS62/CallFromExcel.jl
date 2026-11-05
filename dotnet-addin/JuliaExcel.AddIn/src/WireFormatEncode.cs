using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace JuliaExcel.AddIn
{
    // C# port of SerialiseElement/TrySerialiseArrayAsV (modSerialise.bas) -- encodes a
    // single .NET value (scalar or array) into the JuliaExcel wire format. See
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
    //   JuliaDictionary <-> Scripting.Dictionary
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
            if (value is int) return "&" + ((int)value).ToString(CultureInfo.InvariantCulture);
            if (value is byte) return "B" + ((byte)value).ToString(CultureInfo.InvariantCulture);
            if (value is float) return "S" + Hex.SingleToHex((float)value);
            if (value is long) return "^" + ((long)value).ToString(CultureInfo.InvariantCulture);
            if (value is ExcelErrorValue) return "!" + ((ExcelErrorValue)value).Code.ToString(CultureInfo.InvariantCulture);
            if (value is JuliaDictionary) return EncodeDictionary((JuliaDictionary)value);
            if (value is DateTime)
            {
                double serial = ((DateTime)value).ToOADate(); // Excel's date serial IS the OLE Automation date
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

            throw new NotSupportedException("Cannot serialise value of type " + value.GetType().Name);
        }

        private static string EncodeDictionary(JuliaDictionary dict)
        {
            int n = dict.Count;
            if (n == 0) return "H0;;";
            var encoded = new List<string>(n * 2);
            foreach (var kv in dict.Entries)
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
            // Fast paths for rank 1/2 (by far the common case: Excel always hands worksheet
            // ranges to Excel-DNA as object[]/object[,]) use direct typed indexer access.
            // Array.GetValue(int[]) -- needed for the generic rank-3+ path -- measured as a
            // dominant cost for large arrays (a real regression found via this project's own
            // end-to-end JuliaCall("sort", ...) timing: ~800-1750ms through this path versus
            // ~330-450ms through the csharp-prototype's rank-specific VFormat.cs, for the
            // same 1,000,000-element input) -- generic per-element Array.GetValue/SetValue is
            // measurably slower than direct [,] indexing, consistent with .NET's own
            // well-known multi-dimensional-array performance characteristics.
            if (rank == 1)
            {
                var a1 = (object[])arr;
                int k = 0;
                for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++) encoded[k++] = Encode(a1[i]);
            }
            else if (rank == 2)
            {
                var a2 = (object[,])arr;
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

            if (rank == 1)
            {
                var a1 = (object[])arr;
                int k = 0;
                for (int i = lb[0]; i <= lb[0] + dims[0] - 1; i++)
                {
                    object element = a1[i];
                    if (!(element is double)) return false;
                    buf[k++] = (double)element;
                }
            }
            else if (rank == 2)
            {
                var a2 = (object[,])arr;
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
