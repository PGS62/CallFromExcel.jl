using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace JuliaExcel.AddIn
{
    // Represents a VBA error Variant (what CVErr(n) produces / CStr(CVErr(n)) prints as
    // "Error n"). The wire format carries VBA's own XlCVError UI codes (e.g. 2007 for
    // #DIV/0!, 2042 for #N/A) -- NOT ExcelDna.Integration's ExcelError enum, which uses a
    // different, low-level numbering (0, 7, 15, 23, 29, 36, 42). Mapping to ExcelError is
    // deferred to whichever XLL function needs to return this to a worksheet cell.
    public struct ExcelErrorValue
    {
        public readonly int Code;
        public ExcelErrorValue(int code) { Code = code; }
        public override string ToString() { return "Error " + Code; }
    }

    // C# stand-in for VBA's Scripting.Dictionary, returned from JuliaCallVBA/JuliaEvalVBA
    // (step 5 of the plan) when Julia sends back a Dictionary -- the whole reason that path
    // needs to go through a COM-visible class rather than an XLL function in the first
    // place (Application.Run/XLL return types can't carry a live object reference).
    // COM-visible with the same member names VBA users already know from
    // Scripting.Dictionary (Item as the default property, Keys/Items/Exists/Count/Remove),
    // rather than wrapping an actual Scripting.Dictionary COM object -- avoids a dependency
    // on that library being registered, and keeps this self-contained.
    [ComVisible(true)]
    [Guid("2b6e8a1d-4c7f-4e2a-9b3d-5f8e6a1c2d4b")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJuliaDictionary
    {
        [DispId(0)]
        object Item(object key);
        int Count { get; }
        object[] Keys();
        object[] Items();
        bool Exists(object key);
        void Add(object key, object value);
        bool Remove(object key);
    }

    [ComVisible(true)]
    [Guid("7e4c9f2a-1b8d-4a6e-8c3f-9d2b5a7e1f6c")]
    [ProgId("JuliaExcel.AddIn.JuliaDictionary")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJuliaDictionary))]
    public class JuliaDictionary : IJuliaDictionary
    {
        // Preserves insertion order (matching Scripting.Dictionary's own enumeration
        // order) via a simple ordered list rather than a hash-based collection -- lookup
        // speed doesn't matter here; round-trip/enumeration fidelity does. Keyed lookups
        // (Item/Exists/Remove) do a linear scan using object.Equals, which is fine for the
        // realistic sizes a worksheet-facing Dictionary result would have.
        private readonly List<KeyValuePair<object, object>> _entries = new List<KeyValuePair<object, object>>();

        public int Count { get { return _entries.Count; } }

        public void Add(object key, object value)
        {
            _entries.Add(new KeyValuePair<object, object>(key, value));
        }

        public object Item(object key)
        {
            int i = IndexOf(key);
            if (i < 0) throw new KeyNotFoundException("Key not found in JuliaDictionary: " + key);
            return _entries[i].Value;
        }

        public object[] Keys()
        {
            var result = new object[_entries.Count];
            for (int i = 0; i < _entries.Count; i++) result[i] = _entries[i].Key;
            return result;
        }

        public object[] Items()
        {
            var result = new object[_entries.Count];
            for (int i = 0; i < _entries.Count; i++) result[i] = _entries[i].Value;
            return result;
        }

        public bool Exists(object key)
        {
            return IndexOf(key) >= 0;
        }

        public bool Remove(object key)
        {
            int i = IndexOf(key);
            if (i < 0) return false;
            _entries.RemoveAt(i);
            return true;
        }

        public IEnumerable<KeyValuePair<object, object>> Entries { get { return _entries; } }

        private int IndexOf(object key)
        {
            for (int i = 0; i < _entries.Count; i++)
            {
                if (Equals(_entries[i].Key, key)) return i;
            }
            return -1;
        }
    }
}
