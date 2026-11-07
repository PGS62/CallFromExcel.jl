using System;
using System.Collections.Generic;
using CallFromExcel.AddIn;

namespace CallFromExcel.AddIn.TestHarness
{
    // Cross-validates WireFormatEncode/WireFormatDecode against the REAL Julia process
    // (not just self-consistency): launches Julia, round-trips a battery of values
    // through identity() over HTTP, and checks the result matches the original. This
    // exercises Julia's actual decode_from_xl/encode_for_xl, not a re-implementation of
    // them, so it validates the wire format itself is correctly understood, not just that
    // this C# code agrees with itself.
    internal static class JuliaCrossValidation
    {
        private static int _pass;
        private static int _fail;

        public static bool RunAll()
        {
            _pass = 0;
            _fail = 0;

            Console.WriteLine("Launching Julia ...");
            string launchResult = JuliaLaunch.Launch();
            Console.WriteLine("Launch result: " + launchResult);

            RoundTripViaJulia("Double", 3.14159);
            RoundTripViaJulia("Negative double", -2.5);
            RoundTripViaJulia("Integer-valued double", 42.0);
            RoundTripViaJulia("String", "Hello from C#!");
            RoundTripViaJulia("Empty string", "");
            RoundTripViaJulia("Boolean true", true);
            RoundTripViaJulia("Boolean false", false);
            RoundTripViaJulia("Int32 (Long)", 123456);
            RoundTripViaJulia("Int16 (Integer)", (short)999);
            RoundTripViaJulia("Byte", (byte)250);
            RoundTripViaJulia("Single", 1.5f);
            RoundTripViaJulia("Int64 (LongLong)", 9876543210L);
            RoundTripViaJulia("Null (Empty)", null);
            RoundTripViaJulia("Date only", new DateTime(2025, 12, 22));
            RoundTripViaJulia("DateTime with time", new DateTime(2025, 12, 22, 14, 30, 0));
            RoundTripViaJulia("1-D Double array", new object[] { 1.0, 2.0, 3.0, 4.0 });
            RoundTripViaJulia("1-D String array", new object[] { "a", "bb", "ccc" });
            RoundTripViaJulia("1-D mixed array", new object[] { 1.0, "two", true });

            Console.WriteLine();
            Console.WriteLine(_pass + " passed, " + _fail + " failed (Julia cross-validation).");
            return _fail == 0;
        }

        private static void RoundTripViaJulia(string label, object value)
        {
            try
            {
                string valueEncoded = WireFormatEncode.Encode(value);
                string fnEncoded = WireFormatEncode.Encode("identity");
                string payload = "*1,2;" + fnEncoded.Length + "," + valueEncoded.Length + ",;" + fnEncoded + valueEncoded;

                string responseWire = JuliaHttpClient.Post(payload, "/call");
                object result = WireFormatDecode.Decode(responseWire, allowNesting: true);

                bool ok = DeepEquals(result, value);
                if (ok)
                {
                    _pass++;
                }
                else
                {
                    _fail++;
                    Console.WriteLine("FAIL: " + label + " -- sent wire=" + valueEncoded + " got back wire=" + responseWire);
                }
            }
            catch (Exception ex)
            {
                _fail++;
                Console.WriteLine("FAIL: " + label + " -- exception: " + ex.Message);
            }
        }

        private static bool DeepEquals(object a, object b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            if (a is object[] arrA && b is object[] arrB)
            {
                if (arrA.Length != arrB.Length) return false;
                for (int i = 0; i < arrA.Length; i++)
                {
                    if (!DeepEquals(arrA[i], arrB[i])) return false;
                }
                return true;
            }
            return a.Equals(b);
        }
    }
}
