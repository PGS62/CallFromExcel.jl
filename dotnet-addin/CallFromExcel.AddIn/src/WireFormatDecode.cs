using System;
using System.Globalization;

namespace CallFromExcel.AddIn
{
    // C# port of Unserialise (modUnserialise.bas) -- decodes a CallFromExcel wire-format
    // string back into a .NET value. See WireFormatEncode.cs for the .NET<->VBA Variant
    // type mapping used throughout.
    //
    // Wire format (ported from modUnserialise.bas's own spec comment):
    //   # Double, payload is hex e.g. 1.5 encoded as #3FF8000000000000
    //   Chr(163) (pound sterling sign) String
    //   T / F Boolean True / False
    //   D Date, payload is decimal of Excel's date serial e.g. 22-Dec-2025 is D64013
    //   G DateTime, payload is hex (Excel date serial, which is exactly .NET's OADate)
    //   E Empty (-> null)   N Null (-> DBNull.Value)
    //   % Integer (Int16)   & Long (Int32)   B Byte   S Single (hex)   ^ LongLong (Int64)
    //   C Currency (decimal string) -- reserved, not currently produced by Julia
    //   @ Decimal (decimal string) -- reserved, not currently produced by Julia
    //   ! Error, payload is VBA's own XlCVError UI code (e.g. 2007 for #DIV/0!)
    //   * Array: "*<rank>,<dims>;<lengths-comma-terminated>;<concatenated-encodings>"
    //   H Dictionary: "H<n>;<key/value lengths-comma-terminated>;<concatenated encodings>"
    //   V Array of Float64 only, no per-element type indicator/length (Julia -> Excel
    //     only): "V<rank>,<dims>;<hex>", 16 hex chars/element, no delimiters
    //   R Range (Julia -> Excel only), "RI,<n>,<first>,<step>;" (integer) or
    //     "RF,<n>;<hex first><hex step>" (float), reconstructed via arithmetic
    public static class WireFormatDecode
    {
        public static object Decode(string wire, bool allowNesting = true, int stringLengthLimit = 0, bool vectorToColumn = false, bool forWorksheet = false)
        {
            return Decode(wire, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, 0);
        }

        private static object Decode(string wire, bool allowNesting, int stringLengthLimit, bool vectorToColumn, bool forWorksheet, int depth)
        {
            depth++;
            char tag = wire[0];
            switch (tag)
            {
                case '#': return Hex.HexToDouble(wire, 1);
                case (char)163:
                    if (stringLengthLimit > 0)
                    {
                        int limit = depth == 1 ? 32768 : stringLengthLimit;
                        if (wire.Length > limit)
                        {
                            throw new InvalidOperationException(
                                "Data contains a string of length " + (wire.Length - 1).ToString("#,###", CultureInfo.InvariantCulture) +
                                ", too long to be returned to an Excel worksheet" +
                                (stringLengthLimit == 32768 ? ", for which the limit is 32,767" :
                                    ", for which the limit is 32,767 for a string and " + (stringLengthLimit - 1).ToString("#,###", CultureInfo.InvariantCulture) + " for string elements of an array"));
                        }
                    }
                    return wire.Substring(1);
                case 'T': return true;
                case 'F': return false;
                case 'D': return DateTime.FromOADate(long.Parse(wire.Substring(1), CultureInfo.InvariantCulture));
                case 'G': return DateTime.FromOADate(Hex.HexToDouble(wire, 1));
                case 'E': return null;
                case 'N': return DBNull.Value;
                case '%': return short.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case '&': return int.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case 'B': return byte.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case '^': return long.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case 'S': return Hex.HexToSingle(wire, 1);
                case 'C': return decimal.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case '@': return decimal.Parse(wire.Substring(1), CultureInfo.InvariantCulture);
                case '!': return new ExcelErrorValue(int.Parse(wire.Substring(1), CultureInfo.InvariantCulture));
                case '*': return DecodeArray(wire, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                case 'H': return DecodeDictionary(wire, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                case 'V': return DecodeV(wire, vectorToColumn);
                case 'L': return DecodeL(wire, allowNesting, vectorToColumn, forWorksheet);
                case 'R': return DecodeRange(wire, vectorToColumn, forWorksheet);
                default: throw new FormatException("Character '" + tag + "' is not recognised as a type identifier");
            }
        }

        // Decodes the array element at wire[pos, pos + len). The commonest kinds are read where
        // they are, rather than first copied out with Substring; everything else goes to Decode.
        private static object DecodeElement(string wire, int pos, int len, bool allowNesting, int stringLengthLimit, bool vectorToColumn, bool forWorksheet, int depth)
        {
            switch (wire[pos])
            {
                case '#': if (len == 17) return Hex.HexToDouble(wire, pos + 1); break;
                case 'T': if (len == 1) return true; break;
                case 'F': if (len == 1) return false; break;
                case 'E': if (len == 1) return null; break;
                case '^': if (len > 1) return ParseLongFast(wire, pos + 1, pos + len); break;
                case '!': if (len > 1) return new ExcelErrorValue((int)ParseLongFast(wire, pos + 1, pos + len)); break;
                case (char)163:
                    // A string over the limit goes to Decode, which reports it.
                    if (stringLengthLimit > 0 && len > stringLengthLimit) break;
                    return wire.Substring(pos + 1, len - 1);
            }
            return Decode(wire.Substring(pos, len), allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
        }

        private static object DecodeArray(string wire, bool allowNesting, int stringLengthLimit, bool vectorToColumn, bool forWorksheet, int depth)
        {
            if (depth > 1 && !allowNesting) throw new InvalidOperationException("Excel cannot display arrays containing arrays");

            int p1 = wire.IndexOf(';');
            int p2 = wire.IndexOf(';', p1 + 1);
            if (wire[2] != ',') throw new FormatException("Cannot unserialise arrays with more than 9 dimensions");
            int rank = wire[1] - '0';

            if (rank == 1)
            {
                int n = int.Parse(wire.Substring(3, p1 - 3), CultureInfo.InvariantCulture);
                if (n == 0)
                {
                    if (!allowNesting) throw new InvalidOperationException("Excel cannot display arrays with zero elements");
                    return new object[0];
                }

                var flat = new object[n];
                int m = p1 + 1;
                int pos = p2 + 1;
                for (int i = 0; i < n; i++)
                {
                    int comma = wire.IndexOf(',', m);
                    int len = ParseIntFast(wire, m, comma);
                    flat[i] = DecodeElement(wire, pos, len, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                    pos += len;
                    m = comma + 1;
                }

                if (!vectorToColumn) return flat;
                var column = new object[n, 1];
                for (int i = 0; i < n; i++) column[i, 0] = flat[i];
                return column;
            }

            if (rank == 2)
            {
                int commaPos = wire.IndexOf(',', 3);
                int nr = int.Parse(wire.Substring(3, commaPos - 3), CultureInfo.InvariantCulture);
                int nc = int.Parse(wire.Substring(commaPos + 1, p1 - commaPos - 1), CultureInfo.InvariantCulture);
                if (nr == 0 || nc == 0) throw new InvalidOperationException("Cannot create array of size zero");

                var result = new object[nr, nc];
                int m = p1 + 1;
                int pos = p2 + 1;
                for (int j = 0; j < nc; j++)
                {
                    for (int i = 0; i < nr; i++)
                    {
                        int comma = wire.IndexOf(',', m);
                        int len = ParseIntFast(wire, m, comma);
                        result[i, j] = DecodeElement(wire, pos, len, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                        pos += len;
                        m = comma + 1;
                    }
                }
                return result;
            }

            // Rank 3-9.
            if (!allowNesting) throw new InvalidOperationException("Excel cannot display arrays with more than 2 dimensions");
            int[] dims = ParseDims(wire, 3, p1, rank);
            int total = 1;
            foreach (int d in dims)
            {
                if (d <= 0) throw new InvalidOperationException("Cannot create array of size zero");
                total *= d;
            }

            Array nd = Array.CreateInstance(typeof(object), dims);
            var idx = new int[rank];
            int mPos = p1 + 1;
            int payloadPos = p2 + 1;
            for (int count = 0; count < total; count++)
            {
                int comma = wire.IndexOf(',', mPos);
                int len = ParseIntFast(wire, mPos, comma);
                object val = DecodeElement(wire, payloadPos, len, allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                nd.SetValue(val, idx);
                payloadPos += len;
                mPos = comma + 1;
                AdvanceIndex(idx, dims, rank);
            }
            return nd;
        }

        private static object DecodeDictionary(string wire, bool allowNesting, int stringLengthLimit, bool vectorToColumn, bool forWorksheet, int depth)
        {
            if (!allowNesting) throw new InvalidOperationException("Excel cannot display variables of type Dictionary");
            int p1 = wire.IndexOf(';');
            int p2 = wire.IndexOf(';', p1 + 1);
            int n = int.Parse(wire.Substring(1, p1 - 1), CultureInfo.InvariantCulture);

            var dict = new JuliaDictionary();
            int m = p1 + 1;
            int pos = p2 + 1;
            for (int i = 0; i < n; i++)
            {
                int comma1 = wire.IndexOf(',', m);
                int keyLen = ParseIntFast(wire, m, comma1);
                int comma2 = wire.IndexOf(',', comma1 + 1);
                int valueLen = ParseIntFast(wire, comma1 + 1, comma2);

                object key = Decode(wire.Substring(pos, keyLen), allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                pos += keyLen;
                object val = Decode(wire.Substring(pos, valueLen), allowNesting, stringLengthLimit, vectorToColumn, forWorksheet, depth);
                pos += valueLen;
                m = comma2 + 1;

                dict.Add(key, val);
            }
            return dict;
        }

        // Returns a genuinely-typed double[]/double[,]/double[...] rather than a boxed
        // object array -- not just a style choice. Found via this project's own
        // JuliaCall("sort", ...) end-to-end timing investigation: Excel-DNA must marshal
        // whatever JuliaCall/JuliaEval return into Excel's XLOPER format, and doing that
        // for an object[,] of 1,000,000 boxed doubles (each needing its own per-element
        // XLOPER conversion) measured at roughly double the latency of the same data as a
        // genuinely-typed double[,] (which Excel-DNA can convert in bulk). The "V" format
        // is already guaranteed all-Double by construction (that's its entire reason to
        // exist - see this file's own format-spec comment), so there's no correctness
        // reason to box it; only the general "*" per-element path (genuinely
        // mixed-type) needs object[]/object[,].
        private static object DecodeV(string wire, bool vectorToColumn)
        {
            int p1 = wire.IndexOf(';');
            if (wire[2] != ',') throw new FormatException("Cannot unserialise 'V'-format arrays with more than 9 dimensions");
            int rank = wire[1] - '0';
            int payloadStart = p1 + 1;
            const int bytesPerElement = 16;

            if (rank == 1)
            {
                int n = int.Parse(wire.Substring(3, p1 - 3), CultureInfo.InvariantCulture);
                if (wire.Length - payloadStart != n * bytesPerElement)
                    throw new FormatException("'V'-format string has the wrong number of hex characters for a 1-D array of " + n + " element(s)");
                double[] buf = Hex.ReadBulkHex(wire, payloadStart, n);
                if (!vectorToColumn) return buf;
                var column = new double[n, 1];
                for (int i = 0; i < n; i++) column[i, 0] = buf[i];
                return column;
            }

            if (rank == 2)
            {
                int commaPos = wire.IndexOf(',', 3);
                int nr = int.Parse(wire.Substring(3, commaPos - 3), CultureInfo.InvariantCulture);
                int nc = int.Parse(wire.Substring(commaPos + 1, p1 - commaPos - 1), CultureInfo.InvariantCulture);
                if (wire.Length - payloadStart != nr * nc * bytesPerElement)
                    throw new FormatException("'V'-format string has the wrong number of hex characters for a " + nr + "x" + nc + " array");
                double[] buf = Hex.ReadBulkHex(wire, payloadStart, nr * nc);
                var result = new double[nr, nc];
                int k = 0;
                for (int j = 0; j < nc; j++)
                    for (int i = 0; i < nr; i++)
                        result[i, j] = buf[k++];
                return result;
            }

            // Rank 3-9.
            int[] dims = ParseDims(wire, 3, p1, rank);
            int total = 1;
            foreach (int d in dims)
            {
                if (d <= 0) throw new InvalidOperationException("Cannot create array of size zero");
                total *= d;
            }
            if (wire.Length - payloadStart != total * bytesPerElement)
                throw new FormatException("'V'-format string has the wrong number of hex characters for a " + rank + "-D array");
            double[] flatBuf = Hex.ReadBulkHex(wire, payloadStart, total);

            Array nd = Array.CreateInstance(typeof(double), dims);
            var idx = new int[rank];
            for (int count = 0; count < total; count++)
            {
                nd.SetValue(flatBuf[count], idx);
                AdvanceIndex(idx, dims, rank);
            }
            return nd;
        }

        // "L<rank>,<dims>;<hex>": Int64 array, 16 hex chars per element. For a worksheet the
        // values become a double array, since Excel holds every number as a Double anyway and
        // Excel-DNA converts a double array to Excel's format much faster than boxed values.
        // For VBA they stay Int64 (LongLong), boxed in a Variant array, exactly as the general
        // "*" format has always delivered an array of Int64.
        private static object DecodeL(string wire, bool allowNesting, bool vectorToColumn, bool forWorksheet)
        {
            int p1 = wire.IndexOf(';');
            if (wire[2] != ',') throw new FormatException("Cannot unserialise 'L'-format arrays with more than 9 dimensions");
            int rank = wire[1] - '0';
            if (rank > 2 && !allowNesting) throw new InvalidOperationException("Excel cannot display arrays with more than 2 dimensions");
            int[] dims = ParseDims(wire, 3, p1, rank);
            int total = 1;
            foreach (int d in dims)
            {
                if (d <= 0) throw new InvalidOperationException("Cannot create array of size zero");
                total *= d;
            }
            if (wire.Length - (p1 + 1) != total * 16)
                throw new FormatException("'L'-format string has the wrong number of hex characters for its dimensions");
            long[] flat = Hex.ReadBulkHexInt64(wire, p1 + 1, total);

            if (rank == 1 && !vectorToColumn)
            {
                if (forWorksheet)
                {
                    var doubles = new double[total];
                    for (int i = 0; i < total; i++) doubles[i] = flat[i];
                    return doubles;
                }
                var boxed = new object[total];
                for (int i = 0; i < total; i++) boxed[i] = flat[i];
                return boxed;
            }

            if (rank <= 2)
            {
                int nr = dims[0];
                int nc = rank == 1 ? 1 : dims[1];
                if (forWorksheet)
                {
                    var doubles = new double[nr, nc];
                    int k = 0;
                    for (int j = 0; j < nc; j++)
                        for (int i = 0; i < nr; i++)
                            doubles[i, j] = flat[k++];
                    return doubles;
                }
                var boxed = new object[nr, nc];
                int m = 0;
                for (int j = 0; j < nc; j++)
                    for (int i = 0; i < nr; i++)
                        boxed[i, j] = flat[m++];
                return boxed;
            }

            Array nd = Array.CreateInstance(typeof(object), dims);
            var idx = new int[rank];
            for (int count = 0; count < total; count++)
            {
                nd.SetValue(flat[count], idx);
                AdvanceIndex(idx, dims, rank);
            }
            return nd;
        }

        private static object DecodeRange(string wire, bool vectorToColumn, bool forWorksheet)
        {
            int p1 = wire.IndexOf(';');
            int n;
            double first, step;
            bool isInteger = wire[1] == 'I';
            if (isInteger)
            {
                string[] parts = wire.Substring(3, p1 - 3).Split(',');
                n = int.Parse(parts[0], CultureInfo.InvariantCulture);
                first = long.Parse(parts[1], CultureInfo.InvariantCulture);
                step = long.Parse(parts[2], CultureInfo.InvariantCulture);
            }
            else if (wire[1] == 'F')
            {
                n = int.Parse(wire.Substring(3, p1 - 3), CultureInfo.InvariantCulture);
                first = Hex.HexToDouble(wire, p1 + 1);
                step = Hex.HexToDouble(wire, p1 + 17);
            }
            else
            {
                throw new FormatException("Character '" + wire[1] + "' is not a recognised 'R'-format sub-type (expected 'I' or 'F')");
            }

            if (forWorksheet)
            {
                // A worksheet holds every number as a Double, so build a double array directly,
                // which Excel-DNA converts to Excel's format much faster than boxed values.
                if (!vectorToColumn)
                {
                    var doubles = new double[n];
                    for (int i = 0; i < n; i++) doubles[i] = first + i * step;
                    return doubles;
                }
                var doubleColumn = new double[n, 1];
                for (int i = 0; i < n; i++) doubleColumn[i, 0] = first + i * step;
                return doubleColumn;
            }

            var flat = new object[n];
            for (int i = 0; i < n; i++)
            {
                double v = first + i * step;
                flat[i] = isInteger ? (object)(long)v : v;
            }
            if (!vectorToColumn) return flat;
            var column = new object[n, 1];
            for (int i = 0; i < n; i++) column[i, 0] = flat[i];
            return column;
        }

        // Walks a decoded result tree structurally (into object-element arrays of any rank
        // and JuliaDictionary entries), applying `mapper` to every leaf value - a leaf being
        // anything that isn't itself one of those two container shapes. `mapper` should
        // return its argument unchanged for any value it doesn't want to transform. A
        // genuinely-typed double[]/double[,]/... ("V", or "L"/"R" decoded for a worksheet) is
        // never walked into (its element type isn't `object`), which is always correct: the
        // mappers used (error values, and nothing/missing -> empty) never apply to a number.
        public static object MapLeaves(object value, Func<object, object> mapper)
        {
            var dict = value as JuliaDictionary;
            if (dict != null)
            {
                var mapped = new JuliaDictionary();
                foreach (var kv in dict.Entries)
                {
                    mapped.Add(MapLeaves(kv.Key, mapper), MapLeaves(kv.Value, mapper));
                }
                return mapped;
            }

            // Direct indexing for the two shapes the decoder produces, rather than the
            // GetValue/SetValue path below, which is several times slower per element.
            var vector = value as object[];
            if (vector != null)
            {
                for (int i = 0; i < vector.Length; i++)
                {
                    object element = vector[i];
                    object mappedElement = MapLeaves(element, mapper);
                    if (!ReferenceEquals(mappedElement, element)) vector[i] = mappedElement;
                }
                return vector;
            }
            var matrix = value as object[,];
            if (matrix != null && matrix.GetLowerBound(0) == 0 && matrix.GetLowerBound(1) == 0)
            {
                int nr = matrix.GetLength(0);
                int nc = matrix.GetLength(1);
                for (int i = 0; i < nr; i++)
                {
                    for (int j = 0; j < nc; j++)
                    {
                        object element = matrix[i, j];
                        object mappedElement = MapLeaves(element, mapper);
                        if (!ReferenceEquals(mappedElement, element)) matrix[i, j] = mappedElement;
                    }
                }
                return matrix;
            }

            var arr = value as Array;
            if (arr != null && arr.GetType().GetElementType() == typeof(object))
            {
                int rank = arr.Rank;
                var lb = new int[rank];
                var idx = new int[rank];
                for (int d = 0; d < rank; d++)
                {
                    lb[d] = arr.GetLowerBound(d);
                    idx[d] = lb[d];
                }
                int total = arr.Length;
                for (int count = 0; count < total; count++)
                {
                    object element = arr.GetValue(idx);
                    object mappedElement = MapLeaves(element, mapper);
                    if (!ReferenceEquals(mappedElement, element)) arr.SetValue(mappedElement, idx);
                    for (int d = 0; d < rank; d++)
                    {
                        idx[d]++;
                        if (idx[d] <= arr.GetUpperBound(d)) break;
                        idx[d] = lb[d];
                    }
                }
                return arr;
            }

            return mapper(value);
        }

        // Replaces every ExcelErrorValue in a decoded result tree with mapper(Code). Needed
        // because ExcelErrorValue is a plain struct with no COM or Excel-DNA significance of
        // its own - .NET cannot marshal it to a VBA Variant over COM, nor does Excel-DNA
        // recognise it as a worksheet error - each return path supplies its own conversion
        // (JuliaFunctions.cs -> ExcelDna.Integration.ExcelError for a real XLOPER error;
        // ComServer.cs -> System.Runtime.InteropServices.ErrorWrapper for a genuine COM
        // VT_ERROR Variant).
        public static object MapErrorValues(object value, Func<int, object> mapper)
        {
            return MapLeaves(value, v => v is ExcelErrorValue ev ? mapper(ev.Code) : v);
        }

        private static int[] ParseDims(string wire, int start, int end, int rank)
        {
            string[] parts = wire.Substring(start, end - start).Split(',');
            if (parts.Length != rank) throw new FormatException("Malformed array header: expected " + rank + " dimensions, found " + parts.Length);
            var dims = new int[rank];
            for (int i = 0; i < rank; i++) dims[i] = int.Parse(parts[i], CultureInfo.InvariantCulture);
            return dims;
        }

        private static void AdvanceIndex(int[] idx, int[] dims, int rank)
        {
            for (int d = 0; d < rank; d++)
            {
                idx[d]++;
                if (idx[d] < dims[d]) return;
                idx[d] = 0;
            }
        }

        // Parses an optionally negative integer between [start, end) without allocating a
        // substring. Accumulates negatively so that long.MinValue doesn't overflow.
        private static long ParseLongFast(string s, int start, int end)
        {
            bool negative = s[start] == '-';
            long v = 0;
            for (int i = negative ? start + 1 : start; i < end; i++) v = v * 10 - (s[i] - '0');
            return negative ? v : -v;
        }

        // Parses digits between [start, end) without allocating a substring.
        private static int ParseIntFast(string s, int start, int end)
        {
            int v = 0;
            for (int i = start; i < end; i++) v = v * 10 + (s[i] - '0');
            return v;
        }
    }
}
