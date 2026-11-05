using System;
using System.Globalization;

namespace JuliaExcel.AddIn
{
    // C# port of the hex encode/decode helpers in modUnserialise.bas/modSerialise.bas
    // (DoubleToHex/HexToDouble, SingleToHex/HexToSingle, and the bulk "V"-format
    // equivalents). All big-endian, matching the wire format exactly. The branch-free
    // byte-pair lookup table used for bulk decode was a deliberate fix found during the
    // csharp-prototype speed-benchmark work: a naive per-nibble decode (3 comparisons per
    // hex character) measured 56-90ms for a 1,000,000-element array versus ~7ms for this
    // table-based approach - see VFormat.cs in csharp-prototype for that investigation.
    public static class Hex
    {
        private static readonly char[] HexDigits = "0123456789ABCDEF".ToCharArray();

        // Indexed by (hiChar << 7) | loChar; both chars are always < 128 for valid hex
        // digits, so the index always fits.
        private static readonly byte[] ByteOfHexPair = BuildBytePairTable();

        private static byte[] BuildBytePairTable()
        {
            var nibble = new int[128];
            for (int i = 0; i <= 9; i++) nibble['0' + i] = i;
            for (int i = 0; i <= 5; i++)
            {
                nibble['A' + i] = 10 + i;
                nibble['a' + i] = 10 + i;
            }
            var t = new byte[128 * 128];
            for (int hi = 0; hi < 128; hi++)
            {
                for (int lo = 0; lo < 128; lo++)
                {
                    t[(hi << 7) | lo] = (byte)((nibble[hi] << 4) | nibble[lo]);
                }
            }
            return t;
        }

        public static string DoubleToHex(double x)
        {
            return BitConverter.DoubleToInt64Bits(x).ToString("X16", CultureInfo.InvariantCulture);
        }

        public static double HexToDouble(string s, int start)
        {
            long bits = 0;
            for (int b = 0; b < 8; b++)
            {
                char hi = s[start + b * 2];
                char lo = s[start + b * 2 + 1];
                bits = (bits << 8) | ByteOfHexPair[(hi << 7) | lo];
            }
            return BitConverter.Int64BitsToDouble(bits);
        }

        public static string SingleToHex(float x)
        {
            int bits = BitConverter.ToInt32(BitConverter.GetBytes(x), 0);
            return ((uint)bits).ToString("X8", CultureInfo.InvariantCulture);
        }

        public static float HexToSingle(string s, int start)
        {
            int bits = 0;
            for (int b = 0; b < 4; b++)
            {
                char hi = s[start + b * 2];
                char lo = s[start + b * 2 + 1];
                bits = (bits << 8) | ByteOfHexPair[(hi << 7) | lo];
            }
            return BitConverter.ToSingle(BitConverter.GetBytes(bits), 0);
        }

        // Bulk port of BulkHexOfDoubleArray (modSerialise.bas): writes directly into a
        // caller-supplied char[] buffer at the given offset, avoiding any per-element
        // string allocation (see VFormat.cs's own notes on why that matters).
        public static void WriteBulkHex(double[] values, char[] buf, int pos)
        {
            for (int i = 0; i < values.Length; i++)
            {
                long bits = BitConverter.DoubleToInt64Bits(values[i]);
                int baseIdx = pos + i * 16;
                for (int shift = 56, off = 0; shift >= 0; shift -= 8, off += 2)
                {
                    byte b = (byte)((bits >> shift) & 0xFF);
                    buf[baseIdx + off] = HexDigits[b >> 4];
                    buf[baseIdx + off + 1] = HexDigits[b & 0xF];
                }
            }
        }

        // Bulk port of BulkDoublesFromHex (modUnserialise.bas): reads n consecutive
        // 16-hex-character elements starting at character position `start` of `wire`.
        public static double[] ReadBulkHex(string wire, int start, int n)
        {
            var result = new double[n];
            for (int i = 0; i < n; i++)
            {
                result[i] = HexToDouble(wire, start + i * 16);
            }
            return result;
        }
    }
}
