using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace CallFromExcel.AddIn
{
    // HTTP/1.1 client for the Julia server, over TCP connections kept open between calls.
    //
    // Not HttpWebRequest: it has no separate connect timeout, and a connection attempt to a
    // localhost port nobody is listening on takes ~2s to fail on Windows (~4s once its own
    // retry kicks in), so it needed a fresh TCP "probe" connection before every request -
    // which, measured, cost more than the request itself (0.64ms per call vs 0.12ms here).
    // Keeping connections open means a call normally needs no connection set-up at all, and
    // a Julia that has exited shows up immediately as a broken connection. New connections
    // are made with a short connect timeout (ConnectTimeoutMs); reads have none, since a
    // real JuliaEval/JuliaCall can legitimately take a long time to compute.
    //
    // Synchronous on purpose: Excel calls these functions on its own calculation threads,
    // and sync-over-async there risks deadlock. Idle connections are pooled and each call has
    // one to itself, which stays correct even if calls ever arrive concurrently.
    public static class JuliaHttpClient
    {
        private const int ConnectTimeoutMs = 150;
        private const string Host = "127.0.0.1";

        // Every request carries the session token from the port file (JuliaPort.TokenFor);
        // Julia answers 409 to a request without the right one (see _reject_request in
        // comms.jl), e.g. from an Excel whose own Julia has gone and whose port another
        // Excel's Julia has since taken.
        private const int WrongTokenStatus = 409;

        private static readonly object PoolLock = new object();
        private static readonly Stack<Connection> Pool = new Stack<Connection>();

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
            // the cache) and retry once before giving up.
            // The token is compared too: a Julia relaunched on the same port has a new one,
            // which the other copy of this assembly (XLL vs COM server) hasn't cached yet.
            string token = JuliaPort.TokenFor(port);
            int refreshedPort = JuliaPort.ReadJuliaPortFromFile();
            JuliaPort.SetJuliaPort(refreshedPort);
            if (refreshedPort == 0 || (refreshedPort == port && JuliaPort.TokenFor(port) == token))
            {
                throw new InvalidOperationException(JuliaPort.NoConnectionErrorMessage);
            }
            if (TryPost(refreshedPort, path, payload, out result))
            {
                return result;
            }
            throw new InvalidOperationException(JuliaPort.NoConnectionErrorMessage);
        }

        // False means Julia couldn't be reached on this port. A non-200 response is a real
        // error for the caller to see, so it throws instead.
        private static bool TryPost(int port, string path, string payload, out string result)
        {
            byte[] request = BuildRequest(path, payload, JuliaPort.TokenFor(port));

            // A pooled connection may have been closed by a Julia that has since exited (or
            // been replaced by a new session on the same port). That shows up as a failure
            // before any response arrives, in which case Julia never ran the request, so
            // it's safe to send it again on a fresh connection.
            Connection pooled = Rent(port);
            if (pooled != null)
            {
                Response response;
                if (pooled.TrySend(request, out response))
                {
                    return Finish(pooled, response, out result);
                }
                pooled.Dispose();
            }

            Connection fresh = Connection.TryOpen(port);
            if (fresh == null)
            {
                result = null;
                return false;
            }
            Response freshResponse;
            if (!fresh.TrySend(request, out freshResponse))
            {
                fresh.Dispose();
                result = null;
                return false;
            }
            return Finish(fresh, freshResponse, out result);
        }

        private static bool Finish(Connection connection, Response response, out string result)
        {
            if (response.StatusCode == WrongTokenStatus)
            {
                // Whatever holds this port is not this Excel's Julia session: as far as this
                // Excel is concerned, its own Julia is not reachable.
                connection.Dispose();
                result = null;
                return false;
            }
            if (response.KeepAlive) Return(connection); else connection.Dispose();
            if (response.StatusCode != 200)
            {
                throw new InvalidOperationException("#HTTP " + response.StatusCode + ": " + response.ReasonPhrase + "! " + response.Body);
            }
            result = response.Body;
            return true;
        }

        private static byte[] BuildRequest(string path, string payload, string token)
        {
            byte[] body = Encoding.UTF8.GetBytes(payload);
            byte[] header = Encoding.ASCII.GetBytes(
                "POST " + path + " HTTP/1.1\r\nHost: " + Host + "\r\nX-CallFromExcel-Token: " + token +
                "\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: " +
                body.Length.ToString(CultureInfo.InvariantCulture) + "\r\n\r\n");
            var request = new byte[header.Length + body.Length];
            Buffer.BlockCopy(header, 0, request, 0, header.Length);
            Buffer.BlockCopy(body, 0, request, header.Length, body.Length);
            return request;
        }

        private static Connection Rent(int port)
        {
            lock (PoolLock)
            {
                while (Pool.Count > 0)
                {
                    Connection c = Pool.Pop();
                    if (c.Port == port) return c;
                    c.Dispose();
                }
            }
            return null;
        }

        private static void Return(Connection connection)
        {
            lock (PoolLock) Pool.Push(connection);
        }

        private sealed class Response
        {
            public int StatusCode;
            public string ReasonPhrase;
            public string Body;
            public bool KeepAlive;
        }

        private sealed class Connection : IDisposable
        {
            public readonly int Port;
            private readonly TcpClient _client;
            private readonly NetworkStream _stream;
            private byte[] _buffer = new byte[16384];

            private Connection(int port, TcpClient client)
            {
                Port = port;
                _client = client;
                _stream = client.GetStream();
            }

            public void SetReadTimeout(int timeoutMs)
            {
                _stream.ReadTimeout = timeoutMs;
            }

            public static Connection TryOpen(int port)
            {
                var client = new TcpClient { NoDelay = true };
                try
                {
                    IAsyncResult ar = client.BeginConnect(Host, port, null, null);
                    if (!ar.AsyncWaitHandle.WaitOne(ConnectTimeoutMs))
                    {
                        client.Close();
                        return null;
                    }
                    client.EndConnect(ar);
                    return new Connection(port, client);
                }
                catch (SocketException)
                {
                    client.Close();
                    return null;
                }
            }

            // False if the connection failed before any of the response arrived. A failure
            // after that closes the connection and throws.
            public bool TrySend(byte[] request, out Response response)
            {
                response = null;
                int filled = 0;
                try
                {
                    _stream.Write(request, 0, request.Length);
                    int headerEnd = -1;
                    while (headerEnd < 0)
                    {
                        if (filled == _buffer.Length) Array.Resize(ref _buffer, _buffer.Length * 2);
                        int n = _stream.Read(_buffer, filled, _buffer.Length - filled);
                        if (n == 0)
                        {
                            if (filled == 0) return false;
                            throw new IOException("Connection closed while reading the response headers");
                        }
                        filled += n;
                        headerEnd = IndexOfHeaderEnd(_buffer, filled);
                    }
                    response = ParseResponse(filled, headerEnd);
                    return true;
                }
                catch (Exception ex) when ((ex is IOException || ex is SocketException) && filled == 0)
                {
                    return false;
                }
                catch
                {
                    Dispose();
                    throw;
                }
            }

            private Response ParseResponse(int filled, int headerEnd)
            {
                string headers = Encoding.ASCII.GetString(_buffer, 0, headerEnd);
                string[] lines = headers.Split(new[] { "\r\n" }, StringSplitOptions.None);
                string[] statusParts = lines[0].Split(new[] { ' ' }, 3);
                var response = new Response
                {
                    StatusCode = int.Parse(statusParts[1], CultureInfo.InvariantCulture),
                    ReasonPhrase = statusParts.Length > 2 ? statusParts[2] : "",
                    KeepAlive = true
                };

                int contentLength = -1;
                bool chunked = false;
                for (int i = 1; i < lines.Length; i++)
                {
                    int colon = lines[i].IndexOf(':');
                    if (colon < 0) continue;
                    string name = lines[i].Substring(0, colon).Trim();
                    string value = lines[i].Substring(colon + 1).Trim();
                    if (name.Equals("Content-Length", StringComparison.OrdinalIgnoreCase))
                        contentLength = int.Parse(value, CultureInfo.InvariantCulture);
                    else if (name.Equals("Transfer-Encoding", StringComparison.OrdinalIgnoreCase) && value.IndexOf("chunked", StringComparison.OrdinalIgnoreCase) >= 0)
                        chunked = true;
                    else if (name.Equals("Connection", StringComparison.OrdinalIgnoreCase) && value.Equals("close", StringComparison.OrdinalIgnoreCase))
                        response.KeepAlive = false;
                }

                var leftover = new MemoryStream(_buffer, headerEnd + 4, filled - headerEnd - 4, false);
                byte[] body;
                if (chunked)
                {
                    body = ReadChunked(leftover);
                }
                else if (contentLength >= 0)
                {
                    body = new byte[contentLength];
                    int have = leftover.Read(body, 0, contentLength);
                    ReadExactly(body, have, contentLength - have);
                }
                else
                {
                    var all = new MemoryStream();
                    leftover.CopyTo(all);
                    _stream.CopyTo(all);
                    body = all.ToArray();
                    response.KeepAlive = false;
                }
                response.Body = Encoding.UTF8.GetString(body);
                return response;
            }

            private byte[] ReadChunked(MemoryStream leftover)
            {
                var body = new MemoryStream();
                while (true)
                {
                    string sizeLine = ReadLine(leftover);
                    int semi = sizeLine.IndexOf(';');
                    int size = int.Parse(semi >= 0 ? sizeLine.Substring(0, semi) : sizeLine, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    if (size == 0)
                    {
                        while (ReadLine(leftover).Length > 0) { }
                        return body.ToArray();
                    }
                    var chunk = new byte[size];
                    int have = leftover.Read(chunk, 0, size);
                    if (have < size) ReadExactly(chunk, have, size - have);
                    body.Write(chunk, 0, size);
                    ReadLine(leftover);
                }
            }

            private string ReadLine(MemoryStream leftover)
            {
                var line = new StringBuilder();
                while (true)
                {
                    int b = leftover.ReadByte();
                    if (b < 0) b = _stream.ReadByte();
                    if (b < 0) throw new IOException("Connection closed while reading a chunked response");
                    if (b == '\n') return line.ToString().TrimEnd('\r');
                    line.Append((char)b);
                }
            }

            private void ReadExactly(byte[] target, int offset, int count)
            {
                while (count > 0)
                {
                    int n = _stream.Read(target, offset, count);
                    if (n == 0) throw new IOException("Connection closed while reading the response body");
                    offset += n;
                    count -= n;
                }
            }

            private static int IndexOfHeaderEnd(byte[] buffer, int length)
            {
                for (int i = 3; i < length; i++)
                {
                    if (buffer[i] == '\n' && buffer[i - 1] == '\r' && buffer[i - 2] == '\n' && buffer[i - 3] == '\r') return i - 3;
                }
                return -1;
            }

            public void Dispose()
            {
                _client.Close();
            }
        }

        // True if a Julia serving this Excel process is answering on `port`. Used by
        // JuliaLaunch to decide whether a new session is needed, and by JuliaIsRunning.
        // Uses /ping, which has no side effects in Julia (unlike evaluating code, which
        // would overwrite last_question/last_answer), and which, like every endpoint, answers
        // 200 only to a request carrying this Excel's session token.
        public static bool IsListening(int port, int timeoutMs = 2000)
        {
            string token = JuliaPort.TokenFromFile(port);
            if (token.Length == 0) return false;
            Connection connection = Connection.TryOpen(port);
            if (connection == null) return false;
            using (connection)
            {
                try
                {
                    connection.SetReadTimeout(timeoutMs);
                    Response response;
                    if (!connection.TrySend(BuildRequest("/ping", "", token), out response)) return false;
                    return response.StatusCode == 200;
                }
                catch (Exception ex) when (ex is IOException || ex is SocketException || ex is ObjectDisposedException)
                {
                    return false;
                }
            }
        }
    }
}
