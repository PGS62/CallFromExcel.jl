using System;
using System.Collections.Generic;
using CallFromExcel.AddIn;

namespace CallFromExcel.AddIn.TestHarness
{
    // Self-contained correctness suite for WireFormatEncode/WireFormatDecode -- no Excel,
    // VBA or Julia needed. Includes the exact example strings from modUnserialise.bas's
    // own wire-format spec comment (hand-transcribed, not generated), plus round-trip
    // tests for every type the format supports.
    internal static class WireFormatTests
    {
        private static int _pass;
        private static int _fail;

        public static bool RunAll()
        {
            _pass = 0;
            _fail = 0;

            // --- Exact examples from modUnserialise.bas's wire-format spec comment ---
            Check("#3FF0000000000000 decodes to Double 1", WireFormatDecode.Decode("#3FF0000000000000"), 1.0);
            Check("&1 decodes to Long 1", WireFormatDecode.Decode("&1"), 1);
            Check("pound+Hello decodes to String Hello", WireFormatDecode.Decode(((char)163) + "Hello"), "Hello");
            Check("T decodes to True", WireFormatDecode.Decode("T"), true);
            Check("F decodes to False", WireFormatDecode.Decode("F"), false);
            {
                string wire = "*1,7;2,2,17,1,1,6,6,;%1%2#4008000000000000TF" + (char)163 + "Hello" + (char)163 + "World";
                object decoded = WireFormatDecode.Decode(wire, allowNesting: true);
                object[] arr = (object[])decoded;
                CheckTrue("spec array example length", arr.Length == 7);
                CheckTrue("spec array example [0]", Equals(arr[0], (short)1));
                CheckTrue("spec array example [1]", Equals(arr[1], (short)2));
                CheckTrue("spec array example [2]", Equals(arr[2], 3.0));
                CheckTrue("spec array example [3]", Equals(arr[3], true));
                CheckTrue("spec array example [4]", Equals(arr[4], false));
                CheckTrue("spec array example [5]", Equals(arr[5], "Hello"));
                CheckTrue("spec array example [6]", Equals(arr[6], "World"));
            }
            {
                string wire = "H2;2,3,4,5,;" + (char)163 + "a%10" + (char)163 + "abc%1000";
                var dict = (JuliaDictionary)WireFormatDecode.Decode(wire, allowNesting: true);
                var entries = new List<KeyValuePair<object, object>>(dict.Entries);
                CheckTrue("spec dict example count", entries.Count == 2);
                CheckTrue("spec dict example key0", Equals(entries[0].Key, "a"));
                CheckTrue("spec dict example val0", Equals(entries[0].Value, (short)10));
                CheckTrue("spec dict example key1", Equals(entries[1].Key, "abc"));
                CheckTrue("spec dict example val1", Equals(entries[1].Value, (short)1000));
            }

            // --- Round-trip: every scalar type ---
            RoundTrip("Double", 3.14159);
            RoundTrip("Double negative", -2.5);
            RoundTrip("Double zero", 0.0);
            RoundTrip("String", "Hello, world!");
            RoundTrip("String empty", "");
            RoundTrip("Boolean true", true);
            RoundTrip("Boolean false", false);
            RoundTrip("Int16", (short)12345);
            RoundTrip("Int32", 123456789);
            RoundTrip("Int32 negative", -42);
            RoundTrip("Byte", (byte)200);
            RoundTrip("Single", 1.5f);
            RoundTrip("Int64", 9223372036854775807L);
            RoundTrip("Int64 negative", -9223372036854775808L);
            RoundTrip("Empty (null)", null);
            RoundTrip("Null (DBNull)", DBNull.Value);
            RoundTripDate("Date only", new DateTime(2025, 12, 22));
            RoundTripDate("DateTime with time", new DateTime(2025, 12, 22, 14, 30, 45));
            RoundTripError(new ExcelErrorValue(2007)); // #DIV/0!

            // --- Round-trip: arrays ---
            RoundTripArray1D("1-D Double array", new object[] { 1.0, 2.0, 3.0, 4.0, 5.0 });
            RoundTripArray1D("1-D String array", new object[] { "a", "bb", "ccc" });
            RoundTripArray1D("1-D mixed array", new object[] { 1.0, "two", true, (short)4, null });
            RoundTripArray1D("1-D empty array", new object[0]);

            Test2DArray();
            Test3DArray();
            TestNx1Collapse();
            TestVFormatLargeArray();
            TestDictionaryRoundTrip();
            TestRangeDecode();
            TestLFormatDecode();

            Console.WriteLine();
            Console.WriteLine(_pass + " passed, " + _fail + " failed.");
            return _fail == 0;
        }

        private static void Test2DArray()
        {
            var arr = new object[2, 3];
            arr[0, 0] = 1.0; arr[1, 0] = 2.0;
            arr[0, 1] = "x"; arr[1, 1] = "y";
            arr[0, 2] = true; arr[1, 2] = false;
            string wire = WireFormatEncode.Encode(arr);
            CheckTrue("2-D array header", wire.StartsWith("*2,2,3;"));
            var decoded = (object[,])WireFormatDecode.Decode(wire, allowNesting: true);
            CheckTrue("2-D array round-trip", Equals(decoded[0, 0], 1.0) && Equals(decoded[1, 2], false) && Equals(decoded[0, 1], "x"));
        }

        private static void Test3DArray()
        {
            var arr = Array.CreateInstance(typeof(object), 2, 2, 2);
            int v = 0;
            for (int k = 0; k < 2; k++)
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 2; i++)
                        arr.SetValue((double)(v++), i, j, k);

            string wire = WireFormatEncode.Encode(arr);
            CheckTrue("3-D array takes V path (all Double)", wire.StartsWith("V3,2,2,2;"));
            var decoded = (Array)WireFormatDecode.Decode(wire, allowNesting: true);
            bool ok = true;
            v = 0;
            for (int k = 0; k < 2; k++)
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 2; i++)
                        if (!Equals(decoded.GetValue(i, j, k), (double)(v++))) ok = false;
            CheckTrue("3-D array round-trip (V path)", ok);

            // Force the general (non-V) 3-D path with a mixed-type element.
            arr.SetValue("mixed", 0, 0, 0);
            string wire2 = WireFormatEncode.Encode(arr);
            CheckTrue("3-D array takes general path when mixed", wire2.StartsWith("*3,2,2,2;"));
            var decoded2 = (Array)WireFormatDecode.Decode(wire2, allowNesting: true);
            CheckTrue("3-D array round-trip (general path)", Equals(decoded2.GetValue(0, 0, 0), "mixed") && Equals(decoded2.GetValue(1, 1, 1), 7.0));
        }

        private static void TestNx1Collapse()
        {
            var arr = new object[3, 1];
            arr[0, 0] = 1.0; arr[1, 0] = 2.0; arr[2, 0] = 3.0;
            string wire = WireFormatEncode.Encode(arr);
            CheckTrue("Nx1 collapses to V1 (all Double)", wire.StartsWith("V1,3;"));

            var arrMixed = new object[2, 1];
            arrMixed[0, 0] = "a"; arrMixed[1, 0] = "b";
            string wireMixed = WireFormatEncode.Encode(arrMixed);
            CheckTrue("Nx1 collapses to *1 (mixed)", wireMixed.StartsWith("*1,2;"));

            // Decoding with vectorToColumn=true should turn it back into Nx1.
            var decoded = (object[,])WireFormatDecode.Decode(wireMixed, allowNesting: false, stringLengthLimit: 32768, vectorToColumn: true);
            CheckTrue("Nx1 decode with vectorToColumn", decoded.GetLength(0) == 2 && decoded.GetLength(1) == 1 && Equals(decoded[0, 0], "a"));
        }

        private static void TestVFormatLargeArray()
        {
            const int n = 10000;
            var rnd = new Random(42);
            var arr = new object[n];
            var plain = new double[n];
            for (int i = 0; i < n; i++)
            {
                double v = rnd.NextDouble();
                arr[i] = v;
                plain[i] = v;
            }
            string wire = WireFormatEncode.Encode(arr);
            CheckTrue("Large Double array takes V path", wire.StartsWith("V1,10000;"));
            // V-format decode returns a genuinely-typed double[], not object[] -- see
            // WireFormatDecode.DecodeV's own comment on why that matters for Excel-DNA's
            // return-value marshaling.
            var decoded = (double[])WireFormatDecode.Decode(wire, allowNesting: true);
            bool ok = decoded.Length == n;
            for (int i = 0; i < n && ok; i++) if (decoded[i] != plain[i]) ok = false;
            CheckTrue("Large Double array round-trip", ok);
        }

        private static void TestDictionaryRoundTrip()
        {
            var dict = new JuliaDictionary();
            dict.Add("name", "Julia");
            dict.Add("version", 1.13);
            dict.Add((short)42, true);
            string wire = WireFormatEncode.Encode(dict);
            var decoded = (JuliaDictionary)WireFormatDecode.Decode(wire, allowNesting: true);
            var entries = new List<KeyValuePair<object, object>>(decoded.Entries);
            CheckTrue("Dictionary round-trip count", entries.Count == 3);
            CheckTrue("Dictionary round-trip values", Equals(entries[0].Value, "Julia") && Equals(entries[1].Value, 1.13) && Equals(entries[2].Key, (short)42));

            var emptyDict = new JuliaDictionary();
            string emptyWire = WireFormatEncode.Encode(emptyDict);
            CheckTrue("Empty dictionary wire", emptyWire == "H0;;");
            var emptyDecoded = (JuliaDictionary)WireFormatDecode.Decode(emptyWire, allowNesting: true);
            CheckTrue("Empty dictionary round-trip", emptyDecoded.Count == 0);
        }

        private static void TestRangeDecode()
        {
            // No VBA-side encoder exists for "R" (Julia -> Excel only), so these wire
            // strings are hand-constructed to match encode_for_xl's documented format,
            // mirroring modUnserialise.bas's own worked examples in its spec comment.
            object riDecoded = WireFormatDecode.Decode("RI,5,10,2;"); // 10, 12, 14, 16, 18
            var riArr = (object[])riDecoded;
            CheckTrue("RI range decode", riArr.Length == 5 && Equals(riArr[0], 10L) && Equals(riArr[4], 18L));

            // Beyond 2^53 a Double can't hold every integer: elements must be computed in Int64.
            var bigArr = (object[])WireFormatDecode.Decode("RI,3,1152921504606846977,1;"); // 2^60+1 .. 2^60+3
            CheckTrue("RI range decode beyond 2^53 is exact", Equals(bigArr[0], 1152921504606846977L) && Equals(bigArr[2], 1152921504606846979L));

            var riSheet = WireFormatDecode.Decode("RI,5,10,2;", allowNesting: false, vectorToColumn: true, forWorksheet: true) as double[,];
            CheckTrue("RI range decode for a worksheet is a double column", riSheet != null && riSheet.GetLength(0) == 5 && riSheet[0, 0] == 10 && riSheet[4, 0] == 18);
        }

        // "L" format: Int64 arrays, 16 hex chars per element (Julia's encode_for_xl(::Array{Int64})).
        private static void TestLFormatDecode()
        {
            string vector = "L1,3;" + Hex64(1) + Hex64(-1) + Hex64(long.MaxValue);
            var forVba = WireFormatDecode.Decode(vector, allowNesting: true) as object[];
            CheckTrue("L vector for VBA is boxed Int64", forVba != null && Equals(forVba[0], 1L) && Equals(forVba[1], -1L) && Equals(forVba[2], long.MaxValue));

            var forSheet = WireFormatDecode.Decode(vector, allowNesting: false, vectorToColumn: true, forWorksheet: true) as double[,];
            CheckTrue("L vector for a worksheet is a double column", forSheet != null && forSheet.GetLength(0) == 3 && forSheet.GetLength(1) == 1 && forSheet[1, 0] == -1.0);

            // Column-major: [1 3; 2 4]
            string matrix = "L2,2,2;" + Hex64(1) + Hex64(2) + Hex64(3) + Hex64(4);
            var m = WireFormatDecode.Decode(matrix, allowNesting: false, forWorksheet: true) as double[,];
            CheckTrue("L matrix for a worksheet", m != null && m[0, 0] == 1 && m[1, 0] == 2 && m[0, 1] == 3 && m[1, 1] == 4);
            var mVba = WireFormatDecode.Decode(matrix, allowNesting: true) as object[,];
            CheckTrue("L matrix for VBA", mVba != null && Equals(mVba[0, 1], 3L));

            bool threw = false;
            try { WireFormatDecode.Decode("L1,3;" + Hex64(1), allowNesting: true); } catch (FormatException) { threw = true; }
            CheckTrue("L with too few hex characters is rejected", threw);
        }

        private static string Hex64(long v)
        {
            return v.ToString("X16", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void RoundTrip(string label, object value)
        {
            string wire = WireFormatEncode.Encode(value);
            object decoded = WireFormatDecode.Decode(wire, allowNesting: true);
            CheckTrue(label + " round-trip (wire=" + wire + ")", Equals(decoded, value) || (value == null && decoded == null));
        }

        private static void RoundTripDate(string label, DateTime value)
        {
            string wire = WireFormatEncode.Encode(value);
            object decoded = WireFormatDecode.Decode(wire, allowNesting: true);
            CheckTrue(label + " round-trip (wire=" + wire + ")", decoded is DateTime && (DateTime)decoded == value);
        }

        private static void RoundTripError(ExcelErrorValue value)
        {
            string wire = WireFormatEncode.Encode(value);
            object decoded = WireFormatDecode.Decode(wire, allowNesting: true);
            CheckTrue("Error round-trip (wire=" + wire + ")", decoded is ExcelErrorValue && ((ExcelErrorValue)decoded).Code == value.Code);
        }

        private static void RoundTripArray1D(string label, object[] value)
        {
            string wire = WireFormatEncode.Encode(value);
            object decoded = WireFormatDecode.Decode(wire, allowNesting: true);
            // V-format decode (an all-Double input) returns double[], not object[] -- see
            // WireFormatDecode.DecodeV's own comment on why.
            object[] decodedArr = (decoded is double[] dbl) ? ConvertToObjectArray(dbl) : decoded as object[];
            bool ok = decodedArr != null && decodedArr.Length == value.Length;
            if (ok)
            {
                for (int i = 0; i < value.Length; i++)
                {
                    if (!Equals(decodedArr[i], value[i]) && !(value[i] == null && decodedArr[i] == null)) { ok = false; break; }
                }
            }
            CheckTrue(label + " round-trip (wire=" + wire + ")", ok);
        }

        private static object[] ConvertToObjectArray(double[] values)
        {
            var result = new object[values.Length];
            for (int i = 0; i < values.Length; i++) result[i] = values[i];
            return result;
        }

        private static void Check(string label, object actual, object expected)
        {
            CheckTrue(label, Equals(actual, expected));
        }

        private static void CheckTrue(string label, bool condition)
        {
            if (condition)
            {
                _pass++;
            }
            else
            {
                _fail++;
                Console.WriteLine("FAIL: " + label);
            }
        }
    }
}
