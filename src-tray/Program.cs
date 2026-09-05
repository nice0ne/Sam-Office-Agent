using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace SamOfficeAgent
{
    public static class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);

        private const int ATTACH_PARENT_PROCESS = -1;
        private const string GlobalMutexName = @"Global\SamOfficeAgentTrayServerMutex";
        private const string LocalMutexName = @"Local\SamOfficeAgentTrayServerMutex";

        private static Mutex _appMutex;

        [STAThread]
        public static int Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                if (!Console.IsOutputRedirected)
                {
                    AttachConsole(ATTACH_PARENT_PROCESS);
                }

                try
                {
                    Stream outputStream = Console.OpenStandardOutput();
                    if (outputStream != Stream.Null)
                    {
                        StreamWriter standardOutput = new StreamWriter(outputStream, Console.OutputEncoding ?? Encoding.UTF8) { AutoFlush = true };
                        Console.SetOut(standardOutput);
                    }
                }
                catch { }

                try
                {
                    Stream errorStream = Console.OpenStandardError();
                    if (errorStream != Stream.Null)
                    {
                        StreamWriter standardError = new StreamWriter(errorStream, Console.OutputEncoding ?? Encoding.UTF8) { AutoFlush = true };
                        Console.SetError(standardError);
                    }
                }
                catch { }

                string command = args[0].Trim().ToLowerInvariant();

                if (command == "--status")
                {
                    bool isRunning = IsInstanceRunning();
                    Console.WriteLine(string.Format("SamTrayServer status: {0}", isRunning ? "RUNNING" : "STOPPED"));
                    return 0;
                }

                if (command == "--test")
                {
                    if (!TestMutexEnforcement())
                    {
                        Console.Error.WriteLine("SamTrayServer self-test: Mutex enforcement FAILED");
                        return 1;
                    }

                    if (!TestCertificateGeneration())
                    {
                        Console.Error.WriteLine("SamTrayServer self-test: Certificate generation FAILED");
                        return 1;
                    }

                    Console.WriteLine("SamTrayServer self-test: OK");
                    return 0;
                }

                if (command == "--run-server-tests" || command == "--run-tests")
                {
                    if (args.Length < 4)
                    {
                        Console.Error.WriteLine("Usage: SamTrayServer.exe --run-server-tests <pfxPath> <pfxPass> <distDir> [port]");
                        return 1;
                    }
                    string pfx = args[1];
                    string pass = args[2];
                    string dist = args[3];
                    int port = args.Length > 4 ? int.Parse(args[4]) : 5173;
                    return HttpServerTestSuite.Run(pfx, pass, dist, port);
                }

                if (command == "--run-integration-tests")
                {
                    return OfficeIntegrationTestSuite.Run();
                }

                if (command == "--help" || command == "-h" || command == "/?")
                {
                    Console.WriteLine("Sam Office Agent System Tray Server");
                    Console.WriteLine("Usage: SamTrayServer.exe [options]");
                    Console.WriteLine("  (no args)               Run tray server with GUI");
                    Console.WriteLine("  --status                Check if server is running");
                    Console.WriteLine("  --test                  Run instant self-test");
                    Console.WriteLine("  --run-server-tests      Run HTTPS server test suite");
                    Console.WriteLine("  --run-integration-tests Run Office integration test suite");
                    return 0;
                }
            }

            // Interactive Tray Mode
            bool createdNew = AcquireMutex();
            if (!createdNew)
            {
                MessageBox.Show(
                    "Sam Office Agent sudah berjalan di latar belakang (System Tray).",
                    "Sam Office Agent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return 0;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                using (TrayApplicationContext context = new TrayApplicationContext(_appMutex))
                {
                    Application.Run(context);
                }
            }
            finally
            {
                ReleaseMutex();
            }

            return 0;
        }

        internal static bool AcquireMutex()
        {
            bool createdNew = false;
            try
            {
                Mutex m = new Mutex(true, GlobalMutexName, out createdNew);
                if (!createdNew)
                {
                    if (m != null) { m.Close(); }
                    return false;
                }
                _appMutex = m;
                return true;
            }
            catch (Exception)
            {
                // Global mutex creation threw (e.g. access denied), try Local mutex
            }

            try
            {
                Mutex m = new Mutex(true, LocalMutexName, out createdNew);
                if (!createdNew)
                {
                    if (m != null) { m.Close(); }
                    return false;
                }
                _appMutex = m;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static void ReleaseMutex()
        {
            if (_appMutex != null)
            {
                try
                {
                    _appMutex.ReleaseMutex();
                }
                catch { }
                _appMutex.Close();
                _appMutex = null;
            }
        }

        internal static bool TestMutexEnforcement()
        {
            ReleaseMutex();

            // 1. First acquire must succeed
            bool first = AcquireMutex();
            if (!first)
            {
                Console.Error.WriteLine("FAILED TestMutexEnforcement: First AcquireMutex() returned false.");
                return false;
            }

            // 2. Second acquire on a separate thread must return false (single-instance enforcement)
            bool second = true;
            Thread t = new Thread(delegate()
            {
                second = AcquireMutex();
            });
            t.Start();
            t.Join();

            if (second)
            {
                Console.Error.WriteLine("FAILED TestMutexEnforcement: Second AcquireMutex() succeeded while mutex was held!");
                ReleaseMutex();
                return false;
            }

            // 3. Release and verify re-acquire succeeds
            ReleaseMutex();
            bool reacquire = AcquireMutex();
            if (!reacquire)
            {
                Console.Error.WriteLine("FAILED TestMutexEnforcement: Re-acquire after release returned false.");
                return false;
            }

            ReleaseMutex();
            return true;
        }

        internal static bool TestCertificateGeneration()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "SamTrayCertTest_" + Guid.NewGuid().ToString("N"));
            string tempPfx = Path.Combine(tempDir, "test.pfx");
            try
            {
                bool generated = TrayApplicationContext.GenerateSelfSignedCertificate(tempPfx, "testpass123");
                if (!generated || !File.Exists(tempPfx))
                {
                    Console.Error.WriteLine("FAILED TestCertificateGeneration: Certificate was not generated.");
                    return false;
                }

                X509Certificate2 cert = TrayApplicationContext.TryLoadPfx(tempPfx);
                if (cert == null)
                {
                    Console.Error.WriteLine("FAILED TestCertificateGeneration: Generated certificate could not be loaded.");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("FAILED TestCertificateGeneration: " + ex.Message);
                return false;
            }
            finally
            {
                try
                {
                    if (Directory.Exists(tempDir))
                    {
                        Directory.Delete(tempDir, true);
                    }
                }
                catch { }
            }
        }

        private static bool IsInstanceRunning()
        {
            try
            {
                Mutex existing;
                if (Mutex.TryOpenExisting(GlobalMutexName, out existing))
                {
                    existing.Close();
                    return true;
                }
            }
            catch { }

            try
            {
                Mutex existing;
                if (Mutex.TryOpenExisting(LocalMutexName, out existing))
                {
                    existing.Close();
                    return true;
                }
            }
            catch { }

            return false;
        }

        internal static Icon CreateTrayIcon(string baseDir)
        {
            try
            {
                string[] candidates = new string[]
                {
                    Path.Combine(baseDir, "icon-32.png"),
                    Path.Combine(baseDir, "assets", "icon-32.png"),
                    Path.Combine(baseDir, "..", "dist", "icon-32.png"),
                    Path.Combine(baseDir, "dist", "icon-32.png")
                };

                foreach (string path in candidates)
                {
                    if (File.Exists(path))
                    {
                        using (Bitmap bmp = new Bitmap(path))
                        {
                            IntPtr hIcon = bmp.GetHicon();
                            Icon icon = (Icon)Icon.FromHandle(hIcon).Clone();
                            DestroyIcon(hIcon);
                            return icon;
                        }
                    }
                }
            }
            catch { }

            try
            {
                using (Bitmap bmp = new Bitmap(32, 32))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        using (SolidBrush brush = new SolidBrush(Color.FromArgb(37, 99, 235)))
                        {
                            g.FillEllipse(brush, 1, 1, 30, 30);
                        }
                        using (Font font = new Font("Segoe UI", 16, FontStyle.Bold, GraphicsUnit.Pixel))
                        using (SolidBrush textBrush = new SolidBrush(Color.White))
                        {
                            StringFormat sf = new StringFormat();
                            sf.Alignment = StringAlignment.Center;
                            sf.LineAlignment = StringAlignment.Center;
                            g.DrawString("S", font, textBrush, new RectangleF(0, 0, 32, 32), sf);
                        }
                    }
                    IntPtr hIcon = bmp.GetHicon();
                    Icon icon = (Icon)Icon.FromHandle(hIcon).Clone();
                    DestroyIcon(hIcon);
                    return icon;
                }
            }
            catch
            {
                return SystemIcons.Application;
            }
        }
    }

    public class TrayApplicationContext : ApplicationContext
    {
        private NotifyIcon _notifyIcon;
        private ContextMenuStrip _contextMenu;
        private ToolStripMenuItem _statusItem;
        private ToolStripMenuItem _autoStartItem;
        private Font _headerFont;
        private HttpServer _httpServer;
        private Mutex _mutex;
        private string _baseDir;
        private string _manifestPath;

        public TrayApplicationContext(Mutex mutex)
        {
            _mutex = mutex;
            _baseDir = AppDomain.CurrentDomain.BaseDirectory;

            InitializeIntegration();
            InitializeServer();
            InitializeTray();
        }

        private void InitializeIntegration()
        {
            try
            {
                // Look for manifest.xml
                string[] manifestCandidates = new string[]
                {
                    Path.Combine(_baseDir, "manifest.xml"),
                    Path.Combine(_baseDir, "..", "manifest.xml")
                };

                foreach (string candidate in manifestCandidates)
                {
                    if (File.Exists(candidate))
                    {
                        _manifestPath = Path.GetFullPath(candidate);
                        OfficeIntegration.RegisterAddIn(_manifestPath);
                        break;
                    }
                }

                // Look for Developer CA certificate
                string devCaPath = OfficeIntegration.FindDeveloperCaCertPath();
                if (string.IsNullOrEmpty(devCaPath))
                {
                    string fallbackCa = Path.Combine(_baseDir, "..", "certs", "ca.crt");
                    if (File.Exists(fallbackCa))
                    {
                        devCaPath = fallbackCa;
                    }
                }

                if (!string.IsNullOrEmpty(devCaPath) && File.Exists(devCaPath))
                {
                    OfficeIntegration.EnsureCertificateInstalled(devCaPath);
                }
            }
            catch
            {
                // Non-fatal if Office sideloading or CA cert fails on init
            }
        }

        private void InitializeServer()
        {
            try
            {
                string distDir = ResolveDistDirectory();
                X509Certificate2 certificate = ResolveCertificate();

                _httpServer = new HttpServer();
                _httpServer.Start(5173, distDir, certificate);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Peringatan: Gagal menjalankan server HTTPS pada port 5173.\n{0}", ex.Message),
                    "Sam Office Agent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        private string ResolveDistDirectory()
        {
            string[] distCandidates = new string[]
            {
                Path.Combine(_baseDir, "dist"),
                Path.Combine(_baseDir, "..", "dist")
            };

            foreach (string candidate in distCandidates)
            {
                if (Directory.Exists(candidate))
                {
                    return Path.GetFullPath(candidate);
                }
            }

            // Fallback: create empty dist with default index.html
            string fallbackDist = Path.Combine(_baseDir, "dist");
            if (!Directory.Exists(fallbackDist))
            {
                Directory.CreateDirectory(fallbackDist);
                File.WriteAllText(
                    Path.Combine(fallbackDist, "index.html"),
                    "<!DOCTYPE html><html><body><h1>Sam Office Agent</h1><p>Server active.</p></body></html>"
                );
            }
            return fallbackDist;
        }

        private X509Certificate2 ResolveCertificate()
        {
            string[] pfxCandidates = new string[]
            {
                Path.Combine(_baseDir, "certs", "localhost.pfx"),
                Path.Combine(_baseDir, "localhost.pfx"),
                Path.Combine(_baseDir, "bin", "test-server.pfx"),
                Path.Combine(_baseDir, "test-server.pfx"),
                Path.Combine(_baseDir, "..", "bin", "test-server.pfx"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".office-addin-dev-certs", "localhost.pfx")
            };

            foreach (string pfx in pfxCandidates)
            {
                if (File.Exists(pfx))
                {
                    X509Certificate2 cert = TryLoadPfx(pfx);
                    if (cert != null)
                    {
                        return cert;
                    }
                }
            }

            // Auto-generate self-signed cert if missing
            return CreateSelfSignedCertFallback();
        }

        internal static X509Certificate2 TryLoadPfx(string path)
        {
            string[] passwords = new string[] { "testpass123", "", "localhost" };
            foreach (string pass in passwords)
            {
                try
                {
                    return new X509Certificate2(path, pass, X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
                }
                catch { }
            }
            try
            {
                return new X509Certificate2(path);
            }
            catch
            {
                return null;
            }
        }

        internal static bool GenerateSelfSignedCertificate(string pfxPath, string password)
        {
            string certDir = Path.GetDirectoryName(pfxPath);
            if (!string.IsNullOrEmpty(certDir) && !Directory.Exists(certDir))
            {
                Directory.CreateDirectory(certDir);
            }

            string script = string.Format(
                "$cert = New-SelfSignedCertificate -DnsName '127.0.0.1','localhost' -CertStoreLocation 'Cert:\\CurrentUser\\My' -NotAfter (Get-Date).AddYears(1); " +
                "$pwd = ConvertTo-SecureString '{0}' -Force -AsPlainText; " +
                "Export-PfxCertificate -Cert $cert -FilePath '{1}' -Password $pwd | Out-Null; " +
                "Remove-Item ('Cert:\\CurrentUser\\My\\' + $cert.Thumbprint) -Force -ErrorAction SilentlyContinue",
                password.Replace("'", "''"),
                pfxPath.Replace("'", "''")
            );

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("powershell", "-NoProfile -ExecutionPolicy Bypass -Command \"" + script + "\"");
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;
                using (Process proc = Process.Start(psi))
                {
                    proc.WaitForExit(15000);
                }
            }
            catch { }

            return File.Exists(pfxPath);
        }

        private X509Certificate2 CreateSelfSignedCertFallback()
        {
            string certDir = Path.Combine(_baseDir, "certs");
            string pfxPath = Path.Combine(certDir, "localhost.pfx");
            if (GenerateSelfSignedCertificate(pfxPath, "testpass123"))
            {
                return TryLoadPfx(pfxPath);
            }

            throw new InvalidOperationException("Tidak dapat menemukan atau membuat sertifikat SSL untuk localhost.");
        }

        private void InitializeTray()
        {
            _contextMenu = new ContextMenuStrip();

            // 1. Header: "Sam Office Agent - v1.0" (disabled font-bold)
            ToolStripMenuItem headerItem = new ToolStripMenuItem("Sam Office Agent - v1.0");
            headerItem.Enabled = false;
            _headerFont = new Font(headerItem.Font, FontStyle.Bold);
            headerItem.Font = _headerFont;
            _contextMenu.Items.Add(headerItem);

            // 2. Status: "🟢 Server Aktif (Port 5173)"
            _statusItem = new ToolStripMenuItem(
                (_httpServer != null && _httpServer.IsRunning)
                    ? "🟢 Server Aktif (Port 5173)"
                    : "🔴 Server Berhenti"
            );
            _statusItem.Enabled = false;
            _contextMenu.Items.Add(_statusItem);

            // Separator
            _contextMenu.Items.Add(new ToolStripSeparator());

            // 3. "Buka di Browser" -> https://localhost:5173/
            ToolStripMenuItem openBrowserItem = new ToolStripMenuItem("Buka di Browser", null, new EventHandler(OnOpenBrowserClicked));
            _contextMenu.Items.Add(openBrowserItem);

            // 4. "Daftarkan ke Office"
            ToolStripMenuItem registerItem = new ToolStripMenuItem("Daftarkan ke Office", null, new EventHandler(OnRegisterClicked));
            _contextMenu.Items.Add(registerItem);

            // 5. "Copot dari Office"
            ToolStripMenuItem unregisterItem = new ToolStripMenuItem("Copot dari Office", null, new EventHandler(OnUnregisterClicked));
            _contextMenu.Items.Add(unregisterItem);

            // 6. "Mulai saat Windows Menyala" (Checkable)
            _autoStartItem = new ToolStripMenuItem("Mulai saat Windows Menyala");
            _autoStartItem.CheckOnClick = true;
            _autoStartItem.Checked = OfficeIntegration.IsAutoStartEnabled();
            _autoStartItem.Click += new EventHandler(OnAutoStartClicked);
            _contextMenu.Items.Add(_autoStartItem);

            // Separator
            _contextMenu.Items.Add(new ToolStripSeparator());

            // 7. "Keluar"
            ToolStripMenuItem exitItem = new ToolStripMenuItem("Keluar", null, new EventHandler(OnExitClicked));
            _contextMenu.Items.Add(exitItem);

            _contextMenu.Opening += new System.ComponentModel.CancelEventHandler(OnContextMenuOpening);

            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = Program.CreateTrayIcon(_baseDir);
            _notifyIcon.Text = "Sam Office Agent (Aktif di https://localhost:5173)";
            _notifyIcon.ContextMenuStrip = _contextMenu;
            _notifyIcon.Visible = true;
            _notifyIcon.DoubleClick += new EventHandler(OnOpenBrowserClicked);

            _notifyIcon.ShowBalloonTip(
                3000,
                "Sam Office Agent",
                "Sam Office Agent aktif di latar belakang (https://localhost:5173).",
                ToolTipIcon.Info
            );
        }

        private void OnContextMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _autoStartItem.Checked = OfficeIntegration.IsAutoStartEnabled();
            _statusItem.Text = (_httpServer != null && _httpServer.IsRunning)
                ? "🟢 Server Aktif (Port 5173)"
                : "🔴 Server Berhenti";
        }

        private void OnOpenBrowserClicked(object sender, EventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo("https://localhost:5173/") { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Gagal membuka browser: {0}", ex.Message),
                    "Sam Office Agent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        private void OnRegisterClicked(object sender, EventArgs e)
        {
            string manifest = !string.IsNullOrEmpty(_manifestPath)
                ? _manifestPath
                : Path.Combine(_baseDir, "manifest.xml");

            bool ok = OfficeIntegration.RegisterAddIn(manifest);
            if (ok)
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.ShowBalloonTip(3000, "Sam Office Agent", "Manifest berhasil didaftarkan ke Office.", ToolTipIcon.Info);
                }
            }
            else
            {
                MessageBox.Show("Gagal mendaftarkan add-in ke registry Office.", "Sam Office Agent", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnUnregisterClicked(object sender, EventArgs e)
        {
            bool ok = OfficeIntegration.UnregisterAddIn();
            if (ok)
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.ShowBalloonTip(3000, "Sam Office Agent", "Manifest berhasil dicopot dari Office.", ToolTipIcon.Info);
                }
            }
            else
            {
                MessageBox.Show("Gagal mencopot add-in dari registry Office.", "Sam Office Agent", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnAutoStartClicked(object sender, EventArgs e)
        {
            bool enable = _autoStartItem.Checked;
            bool ok = OfficeIntegration.SetAutoStart(enable);
            if (!ok)
            {
                _autoStartItem.Checked = !enable;
                MessageBox.Show("Gagal mengubah konfigurasi startup Windows.", "Sam Office Agent", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnExitClicked(object sender, EventArgs e)
        {
            ExitApplication();
        }

        public void ExitApplication()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }

            if (_headerFont != null)
            {
                _headerFont.Dispose();
                _headerFont = null;
            }

            if (_contextMenu != null)
            {
                _contextMenu.Dispose();
                _contextMenu = null;
            }

            if (_httpServer != null)
            {
                try
                {
                    _httpServer.Stop();
                }
                catch { }
                _httpServer = null;
            }

            ExitThread();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                ExitApplication();
            }
            base.Dispose(disposing);
        }
    }

    public static class HttpServerTestSuite
    {
        public static int Run(string pfxPath, string pfxPass, string distDir, int port)
        {
            Console.WriteLine("=== Starting HttpServer Test Suite ===");
            Console.WriteLine("Port: " + port);
            Console.WriteLine("Root: " + distDir);
            Console.WriteLine("Cert: " + pfxPath);

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
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
                Console.Write("Test 1: Start server... ");
                server.Start(port, distDir, certificate);
                if (!server.IsRunning)
                {
                    throw new Exception("server.IsRunning is false after Start().");
                }
                Console.WriteLine("PASSED (IsRunning=true)");

                string baseUrl = "https://127.0.0.1:" + port;

                Console.Write("Test 2: GET / ... ");
                HttpWebRequest req1 = (HttpWebRequest)WebRequest.Create(baseUrl + "/");
                using (HttpWebResponse res1 = (HttpWebResponse)req1.GetResponse())
                {
                    AssertEqual((int)res1.StatusCode, 200, "Status 200");
                    AssertContains(res1.ContentType, "text/html", "Content-Type text/html");
                    AssertEqual(res1.Headers["Access-Control-Allow-Origin"], "*", "CORS header");
                    AssertEqual(res1.Headers["Cache-Control"], "no-cache", "Cache-Control no-cache for HTML");
                    using (StreamReader sr = new StreamReader(res1.GetResponseStream()))
                    {
                        string body = sr.ReadToEnd();
                        AssertTrue(!string.IsNullOrEmpty(body), "Body not empty");
                        AssertContains(body.ToLowerInvariant(), "<html", "Body contains <html>");
                    }
                }
                Console.WriteLine("PASSED");

                Console.Write("Test 3: GET /index.html ... ");
                HttpWebRequest req2 = (HttpWebRequest)WebRequest.Create(baseUrl + "/index.html");
                using (HttpWebResponse res2 = (HttpWebResponse)req2.GetResponse())
                {
                    AssertEqual((int)res2.StatusCode, 200, "Status 200");
                    AssertContains(res2.ContentType, "text/html", "Content-Type text/html");
                    AssertEqual(res2.Headers["Cache-Control"], "no-cache", "Cache-Control no-cache for HTML");
                }
                Console.WriteLine("PASSED");

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

                Console.Write("Test 6: SPA Fallback (GET /chat) ... ");
                HttpWebRequest reqSpa = (HttpWebRequest)WebRequest.Create(baseUrl + "/chat");
                using (HttpWebResponse resSpa = (HttpWebResponse)reqSpa.GetResponse())
                {
                    AssertEqual((int)resSpa.StatusCode, 200, "Status 200");
                    AssertContains(resSpa.ContentType, "text/html", "Content-Type text/html");
                    AssertEqual(resSpa.Headers["Cache-Control"], "no-cache", "Cache-Control no-cache for HTML");
                    using (StreamReader sr = new StreamReader(resSpa.GetResponseStream()))
                    {
                        string body = sr.ReadToEnd();
                        AssertContains(body.ToLowerInvariant(), "<html", "Body contains <html>");
                    }
                }
                Console.WriteLine("PASSED");

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

                Console.Write("Test 11: Stop server... ");
                server.Stop();
                if (server.IsRunning)
                {
                    throw new Exception("server.IsRunning is true after Stop().");
                }
                Console.WriteLine("PASSED (IsRunning=false)");

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

                Console.WriteLine("=== All 12 HttpServer Tests Passed Successfully! ===");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("\nTEST FAILED: " + ex.Message);
                return 1;
            }
            finally
            {
                try { server.Stop(); } catch { }
            }
        }

        private static void AssertTrue(bool condition, string message)
        {
            if (!condition) throw new Exception("Assertion failed: " + message);
        }

        private static void AssertEqual<T>(T actual, T expected, string label)
        {
            if (!object.Equals(actual, expected))
            {
                throw new Exception(string.Format("{0}: expected '{1}' but got '{2}'", label, expected, actual));
            }
        }

        private static void AssertContains(string text, string substr, string label)
        {
            if (text == null || text.IndexOf(substr, StringComparison.OrdinalIgnoreCase) < 0)
            {
                throw new Exception(string.Format("{0}: expected to contain '{1}' but got '{2}'", label, substr, text));
            }
        }
    }

    public static class OfficeIntegrationTestSuite
    {
        private static int _passedTests = 0;
        private static int _failedTests = 0;

        public static int Run()
        {
            _passedTests = 0;
            _failedTests = 0;

            Console.WriteLine("=== Starting OfficeIntegration Test Suite ===");

            string previousManifest = OfficeIntegration.GetRegisteredManifestPath();
            string previousAutoStart = OfficeIntegration.GetAutoStartPath();

            try
            {
                TestConstants();
                TestOfficeRegistration();
                TestOfficeUnregistration();
                TestAutoStartConfiguration();
                TestCertificateRootQuery();
                TestCertificateStoreLifecycle();
                TestErrorHandling();

                Console.WriteLine();
                if (_failedTests == 0)
                {
                    Console.WriteLine(string.Format("=== All {0} OfficeIntegration Tests Passed Successfully! ===", _passedTests));
                    return 0;
                }
                else
                {
                    Console.WriteLine(string.Format("=== TEST SUITE FAILED: {0} passed, {1} failed ===", _passedTests, _failedTests));
                    return 1;
                }
            }
            finally
            {
                RestoreRegistryState(previousManifest, previousAutoStart);
            }
        }

        private static void Assert(bool condition, string testName, string detail = null)
        {
            if (condition)
            {
                _passedTests++;
                Console.WriteLine(string.Format("  [PASS] {0}", testName));
            }
            else
            {
                _failedTests++;
                Console.WriteLine(string.Format("  [FAIL] {0} - {1}", testName, detail ?? "Assertion failed"));
            }
        }

        private static void TestConstants()
        {
            Console.WriteLine("Test Suite 1: Constants Verification");
            Assert(OfficeIntegration.ManifestGuid == "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23",
                "ManifestGuid matches manifest.xml ID");
            Assert(OfficeIntegration.WefDeveloperKeyPath == @"Software\Microsoft\Office\16.0\WEF\Developer",
                "WefDeveloperKeyPath points to WEF Developer key");
            Assert(OfficeIntegration.WindowsRunKeyPath == @"Software\Microsoft\Windows\CurrentVersion\Run",
                "WindowsRunKeyPath points to CurrentVersion Run key");
            Assert(OfficeIntegration.AutoStartAppName == "SamOfficeAgent",
                "AutoStartAppName is SamOfficeAgent");
        }

        private static void TestOfficeRegistration()
        {
            Console.WriteLine("Test Suite 2: Office Add-In Registration");

            string testManifest = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-manifest.xml");
            string fullExpected = Path.GetFullPath(testManifest);

            bool regResult = OfficeIntegration.RegisterAddIn(testManifest);
            Assert(regResult, "RegisterAddIn returned true");

            bool isRegAny = OfficeIntegration.IsAddInRegistered();
            Assert(isRegAny, "IsAddInRegistered() with no parameter returned true");

            bool isRegPath = OfficeIntegration.IsAddInRegistered(testManifest);
            Assert(isRegPath, "IsAddInRegistered(testManifest) returned true");

            bool isRegWrong = OfficeIntegration.IsAddInRegistered(@"C:\definitely_wrong_path\manifest.xml");
            Assert(!isRegWrong, "IsAddInRegistered with mismatched path returned false");

            string registeredPath = OfficeIntegration.GetRegisteredManifestPath();
            Assert(string.Equals(registeredPath, fullExpected, StringComparison.OrdinalIgnoreCase),
                "GetRegisteredManifestPath matches expected full path");
        }

        private static void TestOfficeUnregistration()
        {
            Console.WriteLine("Test Suite 3: Office Add-In Unregistration");

            bool unregResult = OfficeIntegration.UnregisterAddIn();
            Assert(unregResult, "UnregisterAddIn returned true");

            bool isRegAfter = OfficeIntegration.IsAddInRegistered();
            Assert(!isRegAfter, "IsAddInRegistered() is false after unregister");

            string registeredPathAfter = OfficeIntegration.GetRegisteredManifestPath();
            Assert(registeredPathAfter == null, "GetRegisteredManifestPath() is null after unregister");

            bool unregAgain = OfficeIntegration.UnregisterAddIn();
            Assert(unregAgain, "Second UnregisterAddIn call is idempotent and returns true");
        }

        private static void TestAutoStartConfiguration()
        {
            Console.WriteLine("Test Suite 4: Windows Auto-Start Configuration");

            string testExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SamTrayServer.exe");
            string fullExpected = Path.GetFullPath(testExe);

            bool enableResult = OfficeIntegration.SetAutoStart(true, testExe);
            Assert(enableResult, "SetAutoStart(true) returned true");

            bool isAutoAny = OfficeIntegration.IsAutoStartEnabled();
            Assert(isAutoAny, "IsAutoStartEnabled() returned true");

            bool isAutoPath = OfficeIntegration.IsAutoStartEnabled(testExe);
            Assert(isAutoPath, "IsAutoStartEnabled(testExe) returned true");

            // Verify quote-handling fix
            string quotedExe = string.Format("\"{0}\"", testExe);
            bool isAutoQuoted = OfficeIntegration.IsAutoStartEnabled(quotedExe);
            Assert(isAutoQuoted, "IsAutoStartEnabled works with quoted exePath");

            bool enableQuoted = OfficeIntegration.SetAutoStart(true, quotedExe);
            Assert(enableQuoted, "SetAutoStart works with quoted exePath");

            bool isAutoWrong = OfficeIntegration.IsAutoStartEnabled(@"C:\wrong\different.exe");
            Assert(!isAutoWrong, "IsAutoStartEnabled with mismatched path returned false");

            string autoPath = OfficeIntegration.GetAutoStartPath();
            Assert(string.Equals(autoPath, fullExpected, StringComparison.OrdinalIgnoreCase),
                "GetAutoStartPath matches expected executable path");

            bool disableResult = OfficeIntegration.SetAutoStart(false);
            Assert(disableResult, "SetAutoStart(false) returned true");

            bool isAutoAfter = OfficeIntegration.IsAutoStartEnabled();
            Assert(!isAutoAfter, "IsAutoStartEnabled() is false after disable");

            string autoPathAfter = OfficeIntegration.GetAutoStartPath();
            Assert(autoPathAfter == null, "GetAutoStartPath() is null after disable");

            bool disableAgain = OfficeIntegration.SetAutoStart(false);
            Assert(disableAgain, "Second SetAutoStart(false) call is idempotent and returns true");
        }

        private static void TestCertificateRootQuery()
        {
            Console.WriteLine("Test Suite 5: Developer CA Certificate Root Query");

            string devCaPath = OfficeIntegration.FindDeveloperCaCertPath();
            if (!string.IsNullOrEmpty(devCaPath) && File.Exists(devCaPath))
            {
                Console.WriteLine("  Found Developer CA at: {0}", devCaPath);
                bool isInstalled = OfficeIntegration.IsCertificateInstalled(devCaPath);
                Assert(isInstalled, "Developer CA is installed in CurrentUser\\Root store");

                bool ensureResult = OfficeIntegration.EnsureCertificateInstalled(devCaPath);
                Assert(ensureResult, "EnsureCertificateInstalled(devCaPath) returns true without prompt");
            }
            else
            {
                Console.WriteLine("  [SKIP] Developer CA not found in default paths (~/.office-addin-dev-certs/ca.crt)");
                _passedTests++;
            }
        }

        private static void TestCertificateStoreLifecycle()
        {
            Console.WriteLine("Test Suite 6: Certificate Store Lifecycle (AddressBook non-prompting store)");

            string projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".."));
            string pfxPath = Path.Combine(projectRoot, "bin", "test-server.pfx");
            if (!File.Exists(pfxPath))
            {
                pfxPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-server.pfx");
            }

            if (File.Exists(pfxPath))
            {
                X509Certificate2 testCert = new X509Certificate2(pfxPath, "testpass123");
                StoreName storeName = StoreName.AddressBook;
                StoreLocation storeLoc = StoreLocation.CurrentUser;

                OfficeIntegration.RemoveCertificate(testCert, storeName, storeLoc);

                bool initialPresent = OfficeIntegration.IsCertificateInstalled(testCert, storeName, storeLoc);
                Assert(!initialPresent, "Test certificate initially absent from AddressBook store");

                bool installResult = OfficeIntegration.EnsureCertificateInstalled(testCert, storeName, storeLoc);
                Assert(installResult, "EnsureCertificateInstalled returned true for AddressBook store");

                bool afterInstall = OfficeIntegration.IsCertificateInstalled(testCert, storeName, storeLoc);
                Assert(afterInstall, "IsCertificateInstalled returned true after install");

                bool ensureAgain = OfficeIntegration.EnsureCertificateInstalled(testCert, storeName, storeLoc);
                Assert(ensureAgain, "EnsureCertificateInstalled returns true immediately when already present");

                bool removeResult = OfficeIntegration.RemoveCertificate(testCert, storeName, storeLoc);
                Assert(removeResult, "RemoveCertificate returned true");

                bool afterRemove = OfficeIntegration.IsCertificateInstalled(testCert, storeName, storeLoc);
                Assert(!afterRemove, "IsCertificateInstalled returned false after removal");
            }
            else
            {
                Console.WriteLine("  [SKIP] test-server.pfx not found at: {0}", pfxPath);
                _passedTests++;
            }
        }

        private static void TestErrorHandling()
        {
            Console.WriteLine("Test Suite 7: Error & Edge Case Handling");

            bool badCert1 = OfficeIntegration.EnsureCertificateInstalled(@"C:\definitely_nonexistent_file_12345.crt");
            Assert(!badCert1, "EnsureCertificateInstalled returns false for nonexistent file without throwing");

            bool badCert2 = OfficeIntegration.EnsureCertificateInstalled((string)null);
            Assert(!badCert2, "EnsureCertificateInstalled returns false for null string without throwing");

            bool badCert3 = OfficeIntegration.EnsureCertificateInstalled((X509Certificate2)null);
            Assert(!badCert3, "EnsureCertificateInstalled returns false for null X509Certificate2 without throwing");

            bool badQuery1 = OfficeIntegration.IsCertificateInstalled(@"C:\definitely_nonexistent_file_12345.crt");
            Assert(!badQuery1, "IsCertificateInstalled returns false for nonexistent file without throwing");

            bool badQuery2 = OfficeIntegration.IsCertificateInstalled((string)null);
            Assert(!badQuery2, "IsCertificateInstalled returns false for null string without throwing");

            bool badQuery3 = OfficeIntegration.IsCertificateInstalled((X509Certificate2)null);
            Assert(!badQuery3, "IsCertificateInstalled returns false for null X509Certificate2 without throwing");
        }

        private static void RestoreRegistryState(string previousManifest, string previousAutoStart)
        {
            try
            {
                if (!string.IsNullOrEmpty(previousManifest))
                {
                    OfficeIntegration.RegisterAddIn(previousManifest);
                }
                else
                {
                    OfficeIntegration.UnregisterAddIn();
                }

                if (!string.IsNullOrEmpty(previousAutoStart))
                {
                    OfficeIntegration.SetAutoStart(true, previousAutoStart);
                }
                else
                {
                    OfficeIntegration.SetAutoStart(false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Warning: failed to restore previous registry state: {0}", ex.Message));
            }
        }
    }
}
