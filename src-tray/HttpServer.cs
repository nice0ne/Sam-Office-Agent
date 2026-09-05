using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;

namespace SamOfficeAgent
{
    /// <summary>
    /// In-process user-mode HTTPS static web server.
    /// Uses TcpListener and SslStream so it runs entirely in user mode
    /// without requiring Windows Administrator rights (unlike http.sys).
    /// </summary>
    public class HttpServer : IDisposable
    {
        private TcpListener _listener;
        private Thread _listenThread;
        private volatile bool _isRunning;
        private string _rootDir;
        private X509Certificate2 _certificate;
        private int _port;
        private readonly object _syncLock = new object();

        private static readonly Dictionary<string, string> MimeTypes;

        static HttpServer()
        {
            MimeTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { ".html", "text/html; charset=utf-8" },
                { ".htm",  "text/html; charset=utf-8" },
                { ".js",   "application/javascript; charset=utf-8" },
                { ".mjs",  "application/javascript; charset=utf-8" },
                { ".css",  "text/css; charset=utf-8" },
                { ".json", "application/json; charset=utf-8" },
                { ".png",  "image/png" },
                { ".jpg",  "image/jpeg" },
                { ".jpeg", "image/jpeg" },
                { ".gif",  "image/gif" },
                { ".svg",  "image/svg+xml" },
                { ".ico",  "image/x-icon" },
                { ".woff", "font/woff" },
                { ".woff2","font/woff2" },
                { ".ttf",  "font/ttf" },
                { ".eot",  "application/vnd.ms-fontobject" },
                { ".xml",  "application/xml; charset=utf-8" },
                { ".txt",  "text/plain; charset=utf-8" },
                { ".map",  "application/json; charset=utf-8" },
                { ".wasm", "application/wasm" }
            };
        }

        public bool IsRunning
        {
            get { return _isRunning; }
        }

        public int Port
        {
            get { return _port; }
        }

        public string RootDir
        {
            get { return _rootDir; }
        }

        /// <summary>
        /// Starts the HTTPS web server on the specified port.
        /// </summary>
        /// <param name="port">Port to bind (typically 5173).</param>
        /// <param name="rootDir">Root folder containing static files (e.g. dist/).</param>
        /// <param name="cert">X509Certificate2 with private key for TLS authentication.</param>
        public void Start(int port, string rootDir, X509Certificate2 cert)
        {
            if (cert == null)
            {
                throw new ArgumentNullException("cert");
            }
            if (!cert.HasPrivateKey)
            {
                throw new ArgumentException("Certificate must include a private key for SSL/TLS server authentication.", "cert");
            }
            if (string.IsNullOrEmpty(rootDir))
            {
                throw new ArgumentNullException("rootDir");
            }
            if (!Directory.Exists(rootDir))
            {
                throw new DirectoryNotFoundException("Web root directory not found: " + rootDir);
            }
            if (port <= 0 || port > 65535)
            {
                throw new ArgumentOutOfRangeException("port", "Port must be between 1 and 65535.");
            }

            lock (_syncLock)
            {
                if (_isRunning)
                {
                    Stop();
                }

                _rootDir = Path.GetFullPath(rootDir);
                _certificate = cert;
                _port = port;

                _listener = new TcpListener(IPAddress.Loopback, _port);
                _listener.Start();
                _isRunning = true;

                _listenThread = new Thread(ListenLoop);
                _listenThread.IsBackground = true;
                _listenThread.Name = "HttpServer-Listener";
                _listenThread.Start();
            }
        }

        /// <summary>
        /// Stops the HTTPS web server and frees resources.
        /// </summary>
        public void Stop()
        {
            lock (_syncLock)
            {
                if (!_isRunning)
                {
                    return;
                }

                _isRunning = false;

                try
                {
                    if (_listener != null)
                    {
                        _listener.Stop();
                    }
                }
                catch
                {
                    // Ignore listener stop errors
                }

                try
                {
                    if (_listenThread != null && _listenThread.IsAlive)
                    {
                        _listenThread.Join(1500);
                    }
                }
                catch
                {
                    // Ignore join errors
                }

                _listener = null;
                _listenThread = null;
            }
        }

        public void Dispose()
        {
            Stop();
        }

        private void ListenLoop()
        {
            while (_isRunning)
            {
                try
                {
                    TcpClient client = _listener.AcceptTcpClient();
                    ThreadPool.QueueUserWorkItem(new WaitCallback(ClientWorker), client);
                }
                catch (SocketException)
                {
                    if (!_isRunning) break;
                }
                catch (ObjectDisposedException)
                {
                    if (!_isRunning) break;
                }
                catch (Exception)
                {
                    if (!_isRunning) break;
                }
            }
        }

        private void ClientWorker(object state)
        {
            TcpClient client = (TcpClient)state;
            try
            {
                client.ReceiveTimeout = 10000;
                client.SendTimeout = 10000;

                using (SslStream sslStream = new SslStream(client.GetStream(), false))
                {
                    sslStream.ReadTimeout = 10000;
                    sslStream.WriteTimeout = 10000;

                    // Authenticate as TLS server
                    sslStream.AuthenticateAsServer(
                        _certificate,
                        false,
                        SslProtocols.Tls | SslProtocols.Tls11 | SslProtocols.Tls12,
                        false
                    );

                    ProcessRequest(sslStream);
                }
            }
            catch (Exception)
            {
                // Client disconnection, TLS handshake failure, or socket timeout
            }
            finally
            {
                try
                {
                    client.Close();
                }
                catch
                {
                }
            }
        }

        private void ProcessRequest(SslStream stream)
        {
            // Read HTTP request headers (up to 32KB)
            byte[] buffer = new byte[4096];
            int headerEndIndex = -1;
            string headerText;

            using (MemoryStream ms = new MemoryStream())
            {
                int totalRead = 0;
                while (totalRead < 32768)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                    {
                        break;
                    }

                    ms.Write(buffer, 0, read);
                    totalRead += read;

                    byte[] current = ms.ToArray();
                    headerEndIndex = IndexOfDoubleNewline(current);
                    if (headerEndIndex >= 0)
                    {
                        break;
                    }
                }

                if (headerEndIndex < 0)
                {
                    return;
                }

                byte[] allBytes = ms.ToArray();
                headerText = Encoding.ASCII.GetString(allBytes, 0, headerEndIndex);
            }

            string[] headerLines = headerText.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

            if (headerLines.Length == 0 || string.IsNullOrEmpty(headerLines[0]))
            {
                SendError(stream, 400, "Bad Request");
                return;
            }

            string[] requestLineParts = headerLines[0].Split(' ');
            if (requestLineParts.Length < 2)
            {
                SendError(stream, 400, "Bad Request");
                return;
            }

            string method = requestLineParts[0].Trim().ToUpperInvariant();
            string rawUrl = requestLineParts[1].Trim();

            // Handle CORS preflight OPTIONS request
            if (method == "OPTIONS")
            {
                SendOptionsResponse(stream);
                return;
            }

            if (method != "GET" && method != "HEAD")
            {
                SendError(stream, 405, "Method Not Allowed");
                return;
            }

            // Parse URL path, stripping query string and hash
            int queryIdx = rawUrl.IndexOf('?');
            if (queryIdx >= 0)
            {
                rawUrl = rawUrl.Substring(0, queryIdx);
            }
            int hashIdx = rawUrl.IndexOf('#');
            if (hashIdx >= 0)
            {
                rawUrl = rawUrl.Substring(0, hashIdx);
            }

            string decodedPath;
            try
            {
                decodedPath = Uri.UnescapeDataString(rawUrl);
            }
            catch
            {
                decodedPath = rawUrl;
            }

            // Normalize path separators
            string relPath = decodedPath.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);

            // Directory traversal prevention
            string targetPath = Path.GetFullPath(Path.Combine(_rootDir, relPath));
            string normalizedRoot = _rootDir.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

            if (!targetPath.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(targetPath, _rootDir, StringComparison.OrdinalIgnoreCase))
            {
                SendError(stream, 403, "Forbidden");
                return;
            }

            // If target is a directory or empty path, look for index.html
            if (string.IsNullOrEmpty(relPath) || Directory.Exists(targetPath))
            {
                string candidate = Path.Combine(targetPath, "index.html");
                if (File.Exists(candidate))
                {
                    targetPath = candidate;
                }
            }

            // SPA Routing Fallback: if file does not exist and request has no extension, serve /index.html
            if (!File.Exists(targetPath))
            {
                bool hasExtension = !string.IsNullOrEmpty(Path.GetExtension(relPath));
                if (!hasExtension)
                {
                    string spaFallback = Path.Combine(_rootDir, "index.html");
                    if (File.Exists(spaFallback))
                    {
                        targetPath = spaFallback;
                    }
                }
            }

            if (!File.Exists(targetPath))
            {
                SendError(stream, 404, "Not Found");
                return;
            }

            // Determine content type
            string ext = Path.GetExtension(targetPath);
            string contentType;
            if (!MimeTypes.TryGetValue(ext, out contentType))
            {
                contentType = "application/octet-stream";
            }

            byte[] bodyBytes;
            try
            {
                bodyBytes = File.ReadAllBytes(targetPath);
            }
            catch (Exception ex)
            {
                SendError(stream, 500, "Internal Server Error: " + ex.Message);
                return;
            }

            SendResponse(stream, 200, "OK", contentType, bodyBytes, method == "HEAD");
        }

        private static int IndexOfDoubleNewline(byte[] data)
        {
            for (int i = 0; i < data.Length - 1; i++)
            {
                if (data[i] == '\n' && data[i + 1] == '\n')
                {
                    return i + 2;
                }
                if (i < data.Length - 3 && data[i] == '\r' && data[i + 1] == '\n' && data[i + 2] == '\r' && data[i + 3] == '\n')
                {
                    return i + 4;
                }
            }
            return -1;
        }

        private static void SendOptionsResponse(SslStream stream)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("HTTP/1.1 204 No Content\r\n");
            sb.Append("Content-Length: 0\r\n");
            sb.Append("Connection: close\r\n");
            sb.Append("Access-Control-Allow-Origin: *\r\n");
            sb.Append("Access-Control-Allow-Methods: GET, POST, OPTIONS, HEAD\r\n");
            sb.Append("Access-Control-Allow-Headers: *\r\n");
            sb.Append("Access-Control-Max-Age: 86400\r\n");
            sb.Append("Date: ").Append(DateTime.UtcNow.ToString("R")).Append("\r\n");
            sb.Append("\r\n");

            byte[] headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
            stream.Write(headerBytes, 0, headerBytes.Length);
            stream.Flush();
        }

        private static void SendError(SslStream stream, int statusCode, string message)
        {
            byte[] body = Encoding.UTF8.GetBytes(string.Format("<!DOCTYPE html><html><body><h1>{0} {1}</h1></body></html>", statusCode, message));
            SendResponse(stream, statusCode, message, "text/html; charset=utf-8", body, false);
        }

        private static void SendResponse(SslStream stream, int statusCode, string statusDescription, string contentType, byte[] body, bool headOnly)
        {
            int contentLength = (body != null) ? body.Length : 0;
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("HTTP/1.1 {0} {1}\r\n", statusCode, statusDescription));
            sb.Append(string.Format("Content-Type: {0}\r\n", contentType));
            sb.Append(string.Format("Content-Length: {0}\r\n", contentLength));
            sb.Append("Connection: close\r\n");
            sb.Append("Access-Control-Allow-Origin: *\r\n");
            sb.Append("Access-Control-Allow-Methods: GET, POST, OPTIONS, HEAD\r\n");
            sb.Append("Access-Control-Allow-Headers: *\r\n");
            if (contentType != null && contentType.IndexOf("text/html", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                sb.Append("Cache-Control: no-cache\r\n");
            }
            sb.Append(string.Format("Date: {0}\r\n", DateTime.UtcNow.ToString("R")));
            sb.Append("\r\n");

            byte[] headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
            stream.Write(headerBytes, 0, headerBytes.Length);

            if (!headOnly && body != null && body.Length > 0)
            {
                stream.Write(body, 0, body.Length);
            }

            stream.Flush();
        }
    }
}
