using System.Collections.Generic;

namespace CallFromExcel.AddIn
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

    // A decoded Julia Dict, in insertion order. Internal to the wire-format code: results
    // returned to VBA are converted to a real Scripting.Dictionary (ComServer.ForVba).
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
