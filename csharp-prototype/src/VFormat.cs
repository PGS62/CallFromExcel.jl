using System;

namespace JuliaExcelPrototype
{
    // Wire-format compatible C# port of the "V" fast path only (TrySerialiseArrayAsV /
    // BulkHexOfDoubleArray in vba/JuliaExcel.xlam/VBA/modSerialise.bas, and the 'V' branch
    // of Unserialise / BulkDoublesFromHex in modUnserialise.bas) -- a speed-comparison
    // microbenchmark, not a general wire-format port. Only handles NRx1/1xNC/NRxNC arrays
    // of genuine Doubles, matching the "V" format's own restriction. Format: "V1,<n>;<hex>"
    // for a 1-D array (or an NRx1 column, which VBA's own encoder collapses to the 1-D
    // form), "V2,<nr>,<nc>;<hex>" otherwise. Payload is big-endian hex, 16 chars/element,
    // column-major -- byte-for-byte identical to what modSerialise.bas/modUnserialise.bas
    // produce and consume, verified by round-tripping the same wire string both ways in
    // vba\Test.bas's CSharpVFormatSpeedTest.
    public static class VFormat
    {
        private static readonly char[] HexDigits = "0123456789ABCDEF".ToCharArray();

        public static string EncodeV(object[,] values)
        {
            int nr = values.GetUpperBound(0) - values.GetLowerBound(0) + 1;
            int nc = values.GetUpperBound(1) - values.GetLowerBound(1) + 1;
            int n = nr * nc;
            var buf = new double[n];
            int k = 0;
            // Column-major, matching Julia's array layout and VBA's own TrySerialiseArrayAsV.
            for (int j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++)
            {
                for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
                {
                    buf[k] = (double)values[i, j];
                    k++;
                }
            }

            string header = (nc == 1)
                ? "V1," + n.ToString()
                : "V2," + nr.ToString() + "," + nc.ToString();
            return header + ";" + BulkHexEncode(buf);
        }

        public static double[] DecodeV(string wire, out int nr, out int nc)
        {
            int semi = wire.IndexOf(';');
            string header = wire.Substring(0, semi);
            string[] parts = header.Substring(3).Split(','); // skip "V1,"/"V2," rank prefix

            if (header[1] == '1')
            {
                nr = int.Parse(parts[0]);
                nc = 1;
            }
            else
            {
                nr = int.Parse(parts[0]);
                nc = int.Parse(parts[1]);
            }

            int n = nr * nc;
            int start = semi + 1;
            if (wire.Length - start != n * 16)
            {
                throw new FormatException("'V'-format string has the wrong number of hex characters for " + n + " element(s)");
            }

            var result = new double[n];
            var table = ByteOfHexPair;
            for (int i = 0; i < n; i++)
            {
                int pos = start + i * 16;
                long bits = 0;
                for (int b = 0; b < 8; b++)
                {
                    char hi = wire[pos + b * 2];
                    char lo = wire[pos + b * 2 + 1];
                    bits = (bits << 8) | table[(hi << 7) | lo];
                }
                result[i] = BitConverter.Int64BitsToDouble(bits);
            }
            return result;
        }

        private static string BulkHexEncode(double[] buf)
        {
            int n = buf.Length;
            var chars = new char[n * 16];
            for (int i = 0; i < n; i++)
            {
                long bits = BitConverter.DoubleToInt64Bits(buf[i]);
                int baseIdx = i * 16;
                for (int shift = 56, pos = 0; shift >= 0; shift -= 8, pos += 2)
                {
                    byte b = (byte)((bits >> shift) & 0xFF);
                    chars[baseIdx + pos] = HexDigits[b >> 4];
                    chars[baseIdx + pos + 1] = HexDigits[b & 0xF];
                }
            }
            return new string(chars);
        }

        // Branch-free: one table lookup per hex-character-pair (byte), instead of the
        // three comparisons per hex character a naive per-nibble decode needs. Indexed by
        // (hiChar << 7) | loChar -- both chars are always < 128 for valid hex digits, so
        // the index always fits. An initial version used a 3-branch per-nibble decode
        // (HexNibble, since removed) and measured 56-90ms for a 1,000,000-element decode
        // in isolation (see csharp-prototype's speed-benchmark investigation); this table
        // approach, mirroring BulkHexEncode's own branch-free per-byte lookup, measured
        // 6.5-7ms for the same input -- the per-nibble branching, not COM marshaling or
        // anything else, was the dominant cost.
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
    }
}
