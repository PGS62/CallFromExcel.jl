using System.IO;
using System.Net;
using System.Text;

namespace JuliaExcelPrototype
{
    // Minimal C# port of JuliaHttpPost (modHttp.bas) -- enough to measure a real end-to-end
    // round trip, not a full port (no port-refresh-on-stale-connection fallback, no
    // connect-timeout tuning). Uses the classic synchronous HttpWebRequest/HttpWebResponse
    // API rather than HttpClient's async methods: the August plan explicitly flags Excel's
    // threading model as a real risk ("stay synchronous/simple; don't reach for async/await
    // just because it's available") -- HttpWebRequest has no Task/await involved at all, so
    // there's no sync-over-async deadlock risk to reason about when this eventually runs on
    // Excel's own thread via Excel-DNA/COM.
    public static class JuliaHttpClient
    {
        public static string Post(string host, int port, string path, string payload)
        {
            var req = (HttpWebRequest)WebRequest.Create("http://" + host + ":" + port + path);
            req.Method = "POST";
            req.ContentType = "text/plain; charset=utf-8";
            req.KeepAlive = true;
            // HttpWebRequest defaults to sending "Expect: 100-continue" for a POST with a
            // body, which makes the client wait for the server's go-ahead before sending
            // the body at all -- an extra round trip per call that's easy to miss and
            // measured as part of "HTTP time" here, not algorithm time.
            req.ServicePoint.Expect100Continue = false;

            byte[] bytes = Encoding.UTF8.GetBytes(payload);
            req.ContentLength = bytes.Length;
            using (Stream requestStream = req.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }

            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var reader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
