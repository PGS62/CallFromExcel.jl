using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace CallFromExcel.AddIn
{
    // C# port of JuliaHttpPost/TryHttpSend (modHttp.bas). Uses the classic synchronous
    // HttpWebRequest/HttpWebResponse API rather than HttpClient's async methods: the
    // August plan explicitly flags Excel's threading model as a real risk ("stay
    // synchronous/simple; don't reach for async/await just because it's available") --
    // HttpWebRequest has no Task/await involved at all, so there's no sync-over-async
    // deadlock risk to reason about once this runs on Excel's own thread.
    //
    // ServerXMLHTTP's setTimeouts gives VBA's version a short (150ms) resolve/connect
    // timeout but unlimited send/receive timeouts -- a real JuliaEval/JuliaCall can
    // legitimately take a long time to compute and must never be aborted just for being
    // slow to respond once connected. HttpWebRequest has no equivalent split (its Timeout
    // covers the whole synchronous call, connect included), so a short raw TCP connect
    // probe is done first via FastConnectProbe, and the real HttpWebRequest -- issued only
    // once that probe succeeds -- is left with an effectively unlimited timeout.
    public static class JuliaHttpClient
    {
        private const int ConnectTimeoutMs = 150;
        private const string Host = "127.0.0.1";

        public static string Post(string payload, string path)
        {
            int port = JuliaPort.GetJuliaPort();
            if (TryPost(port, path, payload, out string result))
            {
                return result;
            }

            // The cached port didn't respond - it may be stale, e.g. because
            // CallFromExcel.serve_xl was called from Julia on a different port, without this
            // process's cached port being told. Re-read the port file directly (bypassing
            // the cache) and retry once before giving up, matching JuliaHttpPost.
            int refreshedPort = JuliaPort.ReadJuliaPortFromFile();
            if (refreshedPort == 0 || refreshedPort == port)
            {
                throw new InvalidOperationException(JuliaPort.NoConnectionErrorMessage);
            }
            JuliaPort.SetJuliaPort(refreshedPort);
            if (TryPost(refreshedPort, path, payload, out result))
            {
                return result;
            }
            throw new InvalidOperationException(JuliaPort.NoConnectionErrorMessage);
        }

        private static bool TryPost(int port, string path, string payload, out string result)
        {
            result = null;
            if (!FastConnectProbe(port)) return false;

            var req = (HttpWebRequest)WebRequest.Create("http://" + Host + ":" + port + path);
            req.Method = "POST";
            req.ContentType = "text/plain; charset=utf-8";
            req.KeepAlive = true;
            req.ServicePoint.Expect100Continue = false;

            byte[] bytes = Encoding.UTF8.GetBytes(payload);
            req.ContentLength = bytes.Length;
            try
            {
                using (Stream requestStream = req.GetRequestStream())
                {
                    requestStream.Write(bytes, 0, bytes.Length);
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var reader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    result = reader.ReadToEnd();
                    return true;
                }
            }
            catch (WebException ex)
            {
                // A non-200 status (e.g. a 500 from a genuine server-side error) is a real
                // error for the caller to see, not a sign of a stale/dead port -- only a
                // connection-level failure should trigger the stale-port retry above.
                if (ex.Response is HttpWebResponse errorResponse)
                {
                    using (errorResponse)
                    using (var reader = new StreamReader(errorResponse.GetResponseStream(), Encoding.UTF8))
                    {
                        throw new InvalidOperationException("#HTTP " + (int)errorResponse.StatusCode + ": " + errorResponse.StatusDescription + "! " + reader.ReadToEnd());
                    }
                }
                return false;
            }
        }

        private static bool FastConnectProbe(int port)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    IAsyncResult ar = client.BeginConnect(Host, port, null, null);
                    bool connected = ar.AsyncWaitHandle.WaitOne(ConnectTimeoutMs);
                    if (!connected) return false;
                    client.EndConnect(ar);
                    return client.Connected;
                }
            }
            catch (SocketException)
            {
                return false;
            }
        }

        // Port of JuliaIsListening (modHttp.bas): used by JuliaLaunch to check whether a
        // previously-known port is still alive before deciding whether to launch a new
        // Julia session.
        public static bool IsListening(int port, int timeoutMs = 2000)
        {
            if (!FastConnectProbe(port)) return false;
            try
            {
                var req = (HttpWebRequest)WebRequest.Create("http://" + Host + ":" + port + "/eval");
                req.Method = "POST";
                req.ContentType = "text/plain; charset=utf-8";
                req.Timeout = timeoutMs;
                req.ServicePoint.Expect100Continue = false;
                byte[] bytes = Encoding.UTF8.GetBytes("1+1");
                req.ContentLength = bytes.Length;
                using (Stream requestStream = req.GetRequestStream())
                {
                    requestStream.Write(bytes, 0, bytes.Length);
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    return resp.StatusCode == HttpStatusCode.OK;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }
    }
}
