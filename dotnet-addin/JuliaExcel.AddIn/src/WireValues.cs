using System.Collections.Generic;

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

    // C# stand-in for VBA's Scripting.Dictionary, used internally for wire-format
    // round-tripping. Preserves insertion order (matching Dictionary's own enumeration
    // order) via a simple ordered list rather than a hash-based collection, since Add
    // order (not lookup speed) is what matters for encode/decode fidelity testing.
    // Whether this should instead be a real COM-visible type exposed to VBA (e.g.
    // wrapping an actual Scripting.Dictionary, or a bespoke COM class with
    // Add/Item/Keys/Exists members) is a decision for when JuliaCallVBA/JuliaEvalVBA are
    // wired up (the August plan's step 5) -- not needed for the wire-format codec itself.
    public class JuliaDictionary
    {
        private readonly List<KeyValuePair<object, object>> _entries = new List<KeyValuePair<object, object>>();

        public int Count { get { return _entries.Count; } }

        public void Add(object key, object value)
        {
            _entries.Add(new KeyValuePair<object, object>(key, value));
        }

        public IEnumerable<KeyValuePair<object, object>> Entries { get { return _entries; } }
    }
}
