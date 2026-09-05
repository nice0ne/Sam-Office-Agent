using System;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;

namespace SamOfficeAgent
{
    /// <summary>
    /// Test harness for HttpServer: validates start, serve (HTML, CSS, JS, SPA fallback, CORS, security), and stop.
    /// </summary>
    public static class TestProgram
    {
        public static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("TestProgram ready.");
                return 0;
            }

            if (args[0] == "--run-tests")
            {
                if (args.Length < 4)
                {
                    Console.Error.WriteLine("Usage: TestServer.exe --run-tests <pfxPath> <pfxPass> <distDir> [port]");
                    return 1;
                }

                string pfxPath = args[1];
                string pfxPass = args[2];
                string distDir = args[3];
                int port = args.Length > 4 ? int.Parse(args[4]) : 5173;

                return RunTestSuite(pfxPath, pfxPass, distDir, port);
            }

            // Interactive mode: TestServer.exe <pfxPath> <pfxPass> <distDir> [port]
            try
            {
                string pfx = args[0];
                string pass = args[1];
                string dist = args[2];
                int p = args.Length > 3 ? int.Parse(args[3]) : 5173;

                X509Certificate2 cert = new X509Certificate2(pfx, pass, X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
                HttpServer server = new HttpServer();
                server.Start(p, dist, cert);
                Console.WriteLine("Server running on https://127.0.0.1:" + p + " (Press Enter to exit)");
                Console.ReadLine();
                server.Stop();
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 1;
            }
        }

        private static int RunTestSuite(string pfxPath, string pfxPass, string distDir, int port)
        {
            Console.WriteLine("=== Starting HttpServer Test Suite ===");
            Console.WriteLine("Port: " + port);
            Console.WriteLine("Root: " + distDir);
            Console.WriteLine("Cert: " + pfxPath);

            // Bypass SSL certificate validation for self-signed test cert
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            // Enable TLS 1.2 (3072), TLS 1.1, TLS 1.0
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            X509Certificate2 certificate;
            try
            {
                certificate = new X509Certificate2(pfxPath, pfxPass, X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("FAILED: Loading certificate: " + ex.Message);
                return 1;
            }

            HttpServer server = new HttpServer();

            try
            {
                // 1. Start Server
                Console.Write("Test 1: Start server... ");
                server.Start(port, distDir, certificate);
                if (!server.IsRunning)
                {
                    throw new Exception("server.IsRunning is false after Start().");
                }
                Console.WriteLine("PASSED (IsRunning=true)");

                string baseUrl = "https://127.0.0.1:" + port;

                // 2. GET / (serves index.html, 200 OK, CORS header, Content-Type text/html)
                Console.Write("Test 2: GET / ... ");
                HttpWebRequest req1 = (HttpWebRequest)WebRequest.Create(baseUrl + "/");
                using (HttpWebResponse res1 = (HttpWebResponse)req1.GetResponse())
                {
                    AssertEqual((int)res1.StatusCode, 200, "Status 200");
                    AssertContains(res1.ContentType, "text/html", "Content-Type text/html");
                    AssertEqual(res1.Headers["Access-Control-Allow-Origin"], "*", "CORS header");
                    using (StreamReader sr = new StreamReader(res1.GetResponseStream()))
                    {
                        string body = sr.ReadToEnd();
                        AssertTrue(!string.IsNullOrEmpty(body), "Body not empty");
                        AssertContains(body.ToLowerInvariant(), "<html", "Body contains <html>");
                    }
                }
                Console.WriteLine("PASSED");

                // 3. GET /index.html
                Console.Write("Test 3: GET /index.html ... ");
                HttpWebRequest req2 = (HttpWebRequest)WebRequest.Create(baseUrl + "/index.html");
                using (HttpWebResponse res2 = (HttpWebResponse)req2.GetResponse())
                {
                    AssertEqual((int)res2.StatusCode, 200, "Status 200");
                    AssertContains(res2.ContentType, "text/html", "Content-Type text/html");
                }
                Console.WriteLine("PASSED");

                // 4. GET CSS file if present in assets
                string[] cssFiles = Directory.GetFiles(distDir, "*.css", SearchOption.AllDirectories);
                if (cssFiles.Length > 0)
                {
                    string relCss = cssFiles[0].Substring(distDir.Length).Replace('\\', '/');
                    Console.Write("Test 4: GET " + relCss + " ... ");
                    HttpWebRequest reqCss = (HttpWebRequest)WebRequest.Create(baseUrl + relCss);
                    using (HttpWebResponse resCss = (HttpWebResponse)reqCss.GetResponse())
                    {
                        AssertEqual((int)resCss.StatusCode, 200, "Status 200");
                        AssertContains(resCss.ContentType, "text/css", "Content-Type text/css");
                    }
                    Console.WriteLine("PASSED");
                }

                // 5. GET JS file if present in assets
                string[] jsFiles = Directory.GetFiles(distDir, "*.js", SearchOption.AllDirectories);
                if (jsFiles.Length > 0)
                {
                    string relJs = jsFiles[0].Substring(distDir.Length).Replace('\\', '/');
                    Console.Write("Test 5: GET " + relJs + " ... ");
                    HttpWebRequest reqJs = (HttpWebRequest)WebRequest.Create(baseUrl + relJs);
                    using (HttpWebResponse resJs = (HttpWebResponse)reqJs.GetResponse())
                    {
                        AssertEqual((int)resJs.StatusCode, 200, "Status 200");
                        AssertContains(resJs.ContentType, "javascript", "Content-Type javascript");
                    }
                    Console.WriteLine("PASSED");
                }

                // 6. SPA Routing Fallback: GET /chat or /settings (no extension, non-existent file)
                Console.Write("Test 6: SPA Fallback (GET /chat) ... ");
                HttpWebRequest reqSpa = (HttpWebRequest)WebRequest.Create(baseUrl + "/chat");
                using (HttpWebResponse resSpa = (HttpWebResponse)reqSpa.GetResponse())
                {
                    AssertEqual((int)resSpa.StatusCode, 200, "Status 200");
                    AssertContains(resSpa.ContentType, "text/html", "Content-Type text/html");
                    using (StreamReader sr = new StreamReader(resSpa.GetResponseStream()))
                    {
                        string body = sr.ReadToEnd();
                        AssertContains(body.ToLowerInvariant(), "<html", "Body contains <html>");
                    }
                }
                Console.WriteLine("PASSED");

                // 7. Missing file with extension returns 404
                Console.Write("Test 7: 404 for missing file (GET /nonexistent.png) ... ");
                try
                {
                    HttpWebRequest req404 = (HttpWebRequest)WebRequest.Create(baseUrl + "/nonexistent.png");
                    using (HttpWebResponse res404 = (HttpWebResponse)req404.GetResponse())
                    {
                        throw new Exception("Expected 404 but got " + (int)res404.StatusCode);
                    }
                }
                catch (WebException wex)
                {
                    HttpWebResponse res = wex.Response as HttpWebResponse;
                    if (res != null)
                    {
                        AssertEqual((int)res.StatusCode, 404, "Status 404");
                    }
                    else
                    {
                        throw;
                    }
                }
                Console.WriteLine("PASSED");

                // 8. CORS preflight OPTIONS request
                Console.Write("Test 8: OPTIONS / (CORS preflight) ... ");
                HttpWebRequest reqOpt = (HttpWebRequest)WebRequest.Create(baseUrl + "/");
                reqOpt.Method = "OPTIONS";
                using (HttpWebResponse resOpt = (HttpWebResponse)reqOpt.GetResponse())
                {
                    int code = (int)resOpt.StatusCode;
                    AssertTrue(code == 200 || code == 204, "OPTIONS status 200 or 204, got: " + code);
                    AssertEqual(resOpt.Headers["Access-Control-Allow-Origin"], "*", "Access-Control-Allow-Origin");
                    AssertContains(resOpt.Headers["Access-Control-Allow-Methods"], "GET", "Access-Control-Allow-Methods contains GET");
                }
                Console.WriteLine("PASSED");

                // 9. HEAD /index.html (headers only)
                Console.Write("Test 9: HEAD /index.html ... ");
                HttpWebRequest reqHead = (HttpWebRequest)WebRequest.Create(baseUrl + "/index.html");
                reqHead.Method = "HEAD";
                using (HttpWebResponse resHead = (HttpWebResponse)reqHead.GetResponse())
                {
                    AssertEqual((int)resHead.StatusCode, 200, "Status 200");
                    AssertContains(resHead.ContentType, "text/html", "Content-Type text/html");
                    AssertEqual(resHead.Headers["Access-Control-Allow-Origin"], "*", "Access-Control-Allow-Origin");
                }
                Console.WriteLine("PASSED");

                // 10. Directory traversal attack prevention (/../../something)
                Console.Write("Test 10: Directory traversal prevention ... ");
                try
                {
                    HttpWebRequest reqTrav = (HttpWebRequest)WebRequest.Create(baseUrl + "/../../windows/win.ini");
                    using (HttpWebResponse resTrav = (HttpWebResponse)reqTrav.GetResponse())
                    {
                        throw new Exception("Expected 403 or 404 for traversal attempt, but got: " + (int)resTrav.StatusCode);
                    }
                }
                catch (WebException wex)
                {
                    HttpWebResponse res = wex.Response as HttpWebResponse;
                    if (res != null)
                    {
                        int c = (int)res.StatusCode;
                        AssertTrue(c == 403 || c == 404, "Status is 403 or 404, got: " + c);
                    }
                    else
                    {
                        throw;
                    }
                }
                Console.WriteLine("PASSED");

                // 11. Stop Server & verify
                Console.Write("Test 11: Stop server... ");
                server.Stop();
                if (server.IsRunning)
                {
                    throw new Exception("server.IsRunning is true after Stop().");
                }
                Console.WriteLine("PASSED (IsRunning=false)");

                // 12. Verify stopped server refuses connection
                Console.Write("Test 12: Verify stopped server connection refused... ");
                bool connectionFailed = false;
                try
                {
                    HttpWebRequest reqStopped = (HttpWebRequest)WebRequest.Create(baseUrl + "/");
                    reqStopped.Timeout = 2000;
                    using (HttpWebResponse resStopped = (HttpWebResponse)reqStopped.GetResponse())
                    {
                        connectionFailed = false;
                    }
                }
                catch (WebException)
                {
                    connectionFailed = true;
                }
                AssertTrue(connectionFailed, "Connection failed as expected");
                Console.WriteLine("PASSED");

                Console.WriteLine("=== All 12 Tests Passed Successfully! ===");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("\nTEST FAILED: " + ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                return 1;
            }
            finally
            {
                try
                {
                    server.Stop();
                }
                catch
                {
                }
            }
        }

        private static void AssertEqual(object actual, object expected, string message)
        {
            if (!object.Equals(actual, expected))
            {
                throw new Exception(string.Format("Assertion failed: {0}. Expected: [{1}], Actual: [{2}]", message, expected, actual));
            }
        }

        private static void AssertContains(string actual, string substring, string message)
        {
            if (actual == null || actual.IndexOf(substring, StringComparison.OrdinalIgnoreCase) < 0)
            {
                throw new Exception(string.Format("Assertion failed: {0}. String '{1}' does not contain '{2}'", message, actual, substring));
            }
        }

        private static void AssertTrue(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(string.Format("Assertion failed: {0}", message));
            }
        }
    }
}
