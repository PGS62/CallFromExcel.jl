using System.Text;

namespace CallFromExcel.AddIn
{
    // C# port of MakeJuliaLiteral (modJuliaLiteral.bas): escape a string so it can be
    // embedded, double-quoted, as a Julia string literal. Sole use: embedding the launch
    // command line in a println statement written into the generated Julia startup script.
    public static class JuliaLiteral
    {
        public static string MakeJuliaLiteral(string x)
        {
            // Order matters: backslash substitution must happen before any substitution
            // that introduces new backslashes (\n, \r, \$, \", \uNNNN), exactly as in the
            // VBA version.
            var sb = new StringBuilder(x.Length + 2);
            sb.Append('"');
            foreach (char c in x)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '$': sb.Append("\\$"); break;
                    case '"': sb.Append("\\\""); break;
                    default:
                        // Julia's "caution" about these characters is a defence against
                        // Trojan Source attacks (bidirectional-formatting Unicode controls)
                        // -- https://github.com/JuliaLang/julia/pull/42918,
                        // https://trojansource.codes/ -- matching the VBA version's two
                        // ChrW(8234..8238)/ChrW(8294..8297) ranges.
                        if ((c >= 8234 && c <= 8238) || (c >= 8294 && c <= 8297))
                        {
                            sb.Append("\\u").Append(((int)c).ToString("x"));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }
    }
}
