using System;
using System.Globalization;
using System.Text;

namespace JuliaExcelPrototype
{
    // Wire-format compatible C# port of the general "*" per-element array format only
    // (SerialiseElement's non-"V" fallback path in modSerialise.bas, and Unserialise's
    // Case 42 '*' branch in modUnserialise.bas) -- a speed-comparison microbenchmark, not
    // a general wire-format port. Only handles 1-D (or NRx1 column) arrays of Double or
    // String elements -- the two types needed to compare against VBA's general per-element
    // encode path for an all-String array (Doubles already have their own fast "V" path in
    // VFormat.cs; this file exists because strings/mixed types never take that path).
    // Format: "*1,<n>;<len1>,<len2>,...,<lenN>,;<concatenated per-element encodings>" --
    // see modUnserialise.bas's own wire-format spec comment for the full type-indicator
    // table and general rules (rank, dictionaries, etc. not reimplemented here).
    public static class GeneralFormat
    {
        // Two-pass, buffer-based: pass 1 computes each element's encoded length without
        // allocating its encoded string; pass 2 writes tag + content characters directly
        // into one pre-sized char[] buffer. An initial version allocated a `string[n]` of
        // per-element "<tag><content>" strings, then re-read each one's .Length in a
        // second loop to build the lengths section -- measured (in isolation, 1,000,000
        // all-String elements) at ~93ms just for the per-element string allocations plus
        // ~60ms re-deriving lengths from them, versus ~4ms + ~14ms here for the equivalent
        // work -- see this file's own git history/the csharp-prototype speed-benchmark
        // investigation. Mirrors VFormat.cs's own "write into one buffer, don't allocate
        // per element" approach.
        public static string EncodeGeneral(object[,] values)
        {
            int nr = values.GetUpperBound(0) - values.GetLowerBound(0) + 1;
            int nc = values.GetUpperBound(1) - values.GetLowerBound(1) + 1;
            int n = nr * nc;

            var flat = new object[n];
            var elemLen = new int[n];
            int totalPayloadLen = 0;
            int k = 0;
            // Column-major, matching Julia's array layout and VBA's own SerialiseElement.
            for (int j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++)
            {
                for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
                {
                    object x = values[i, j];
                    flat[k] = x;
                    int len = ElementEncodedLength(x);
                    elemLen[k] = len;
                    totalPayloadLen += len;
                    k++;
                }
            }

            var lens = new StringBuilder(n * 4);
            for (int i = 0; i < n; i++)
            {
                lens.Append(elemLen[i]);
                lens.Append(',');
            }

            var payload = new char[totalPayloadLen];
            int pos = 0;
            for (int i = 0; i < n; i++)
            {
                WriteElement(flat[i], payload, pos);
                pos += elemLen[i];
            }

            string header = (nc == 1)
                ? "*1," + n.ToString(CultureInfo.InvariantCulture)
                : "*2," + nr.ToString(CultureInfo.InvariantCulture) + "," + nc.ToString(CultureInfo.InvariantCulture);
            return header + ";" + lens + ";" + new string(payload);
        }

        public static object[,] DecodeGeneral(string wire, out int nr, out int nc)
        {
            int p1 = wire.IndexOf(';');
            int p2 = wire.IndexOf(';', p1 + 1);

            if (wire[2] != ',') throw new FormatException("Expected a single-digit rank after '*'");
            int rank = wire[1] - '0';
            if (rank != 1) throw new NotSupportedException("GeneralFormat (microbenchmark) only supports rank-1 arrays, got rank " + rank);

            int n = int.Parse(wire.Substring(3, p1 - 3), CultureInfo.InvariantCulture);
            nr = n;
            nc = 1;

            var result = new object[n, 1];
            int m = p1 + 1;   // walks the comma-delimited lengths section
            int pos = p2 + 1; // walks the concatenated payload
            for (int i = 0; i < n; i++)
            {
                int comma = wire.IndexOf(',', m);
                int len = ParseIntFast(wire, m, comma);
                result[i, 0] = DecodeElement(wire, pos, len);
                pos += len;
                m = comma + 1;
            }
            return result;
        }

        private const char PoundSign = (char)163; // Chr(163) in VBA -- the String type-indicator tag
        private static readonly char[] HexDigits = "0123456789ABCDEF".ToCharArray();

        private static int ElementEncodedLength(object x)
        {
            if (x is string) return 1 + ((string)x).Length;
            if (x is double) return 1 + 16;
            throw new NotSupportedException("GeneralFormat (microbenchmark) does not support type " + (x == null ? "null" : x.GetType().Name));
        }

        private static void WriteElement(object x, char[] buf, int pos)
        {
            if (x is string)
            {
                string s = (string)x;
                buf[pos] = PoundSign;
                s.CopyTo(0, buf, pos + 1, s.Length);
                return;
            }
            if (x is double)
            {
                buf[pos] = '#';
                long bits = BitConverter.DoubleToInt64Bits((double)x);
                int baseIdx = pos + 1;
                for (int shift = 56, off = 0; shift >= 0; shift -= 8, off += 2)
                {
                    byte b = (byte)((bits >> shift) & 0xFF);
                    buf[baseIdx + off] = HexDigits[b >> 4];
                    buf[baseIdx + off + 1] = HexDigits[b & 0xF];
                }
                return;
            }
            throw new NotSupportedException("GeneralFormat (microbenchmark) does not support type " + (x == null ? "null" : x.GetType().Name));
        }

        private static object DecodeElement(string wire, int start, int len)
        {
            char tag = wire[start];
            if (tag == PoundSign) return wire.Substring(start + 1, len - 1);
            if (tag == '#')
            {
                long bits = long.Parse(wire.Substring(start + 1, len - 1), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                return BitConverter.Int64BitsToDouble(bits);
            }
            throw new NotSupportedException("GeneralFormat (microbenchmark) does not support type indicator '" + tag + "'");
        }

        // Parses digits between [start, end) without allocating a substring -- VBA's own
        // general-format decode (Unserialise, Case 42) pays an equivalent Mid$/CLng cost
        // per element here, so this keeps the comparison honest rather than taking a
        // shortcut VBA's own code doesn't have.
        private static int ParseIntFast(string s, int start, int end)
        {
            int v = 0;
            for (int i = start; i < end; i++) v = v * 10 + (s[i] - '0');
            return v;
        }
    }
}
