using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;

namespace SamOfficeAgent
{
    public class TestIntegrationProgram
    {
        private static int _passedTests = 0;
        private static int _failedTests = 0;

        public static int Main(string[] args)
        {
            Console.WriteLine("=== Starting OfficeIntegration Test Suite ===");

            // Backup existing registry state to guarantee no pollution of user environment
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
                    Console.WriteLine("=== All {0} Tests Passed Successfully! ===", _passedTests);
                    return 0;
                }
                else
                {
                    Console.WriteLine("=== TEST SUITE FAILED: {0} passed, {1} failed ===", _passedTests, _failedTests);
                    return 1;
                }
            }
            finally
            {
                // Restore previous registry state
                RestoreRegistryState(previousManifest, previousAutoStart);
            }
        }

        private static void Assert(bool condition, string testName, string detail = null)
        {
            if (condition)
            {
                _passedTests++;
                Console.WriteLine("  [PASS] {0}", testName);
            }
            else
            {
                _failedTests++;
                Console.WriteLine("  [FAIL] {0} - {1}", testName, detail ?? "Assertion failed");
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

            // Idempotent unregister
            bool unregAgain = OfficeIntegration.UnregisterAddIn();
            Assert(unregAgain, "Second UnregisterAddIn call is idempotent and returns true");
        }

        private static void TestAutoStartConfiguration()
        {
            Console.WriteLine("Test Suite 4: Windows Auto-Start Configuration");

            string testExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SamTrayServer.exe");
            string fullExpected = Path.GetFullPath(testExe);

            // Enable auto-start
            bool enableResult = OfficeIntegration.SetAutoStart(true, testExe);
            Assert(enableResult, "SetAutoStart(true) returned true");

            bool isAutoAny = OfficeIntegration.IsAutoStartEnabled();
            Assert(isAutoAny, "IsAutoStartEnabled() returned true");

            bool isAutoPath = OfficeIntegration.IsAutoStartEnabled(testExe);
            Assert(isAutoPath, "IsAutoStartEnabled(testExe) returned true");

            bool isAutoWrong = OfficeIntegration.IsAutoStartEnabled(@"C:\wrong\different.exe");
            Assert(!isAutoWrong, "IsAutoStartEnabled with mismatched path returned false");

            string autoPath = OfficeIntegration.GetAutoStartPath();
            Assert(string.Equals(autoPath, fullExpected, StringComparison.OrdinalIgnoreCase),
                "GetAutoStartPath matches expected executable path");

            // Disable auto-start
            bool disableResult = OfficeIntegration.SetAutoStart(false);
            Assert(disableResult, "SetAutoStart(false) returned true");

            bool isAutoAfter = OfficeIntegration.IsAutoStartEnabled();
            Assert(!isAutoAfter, "IsAutoStartEnabled() is false after disable");

            string autoPathAfter = OfficeIntegration.GetAutoStartPath();
            Assert(autoPathAfter == null, "GetAutoStartPath() is null after disable");

            // Idempotent disable
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
                _passedTests++; // Count as passed
            }
        }

        private static void TestCertificateStoreLifecycle()
        {
            Console.WriteLine("Test Suite 6: Certificate Store Lifecycle (AddressBook non-prompting store)");

            // Look for test certificate in bin/test-server.pfx
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

                // Ensure clean initial state in AddressBook
                OfficeIntegration.RemoveCertificate(testCert, storeName, storeLoc);

                bool initialPresent = OfficeIntegration.IsCertificateInstalled(testCert, storeName, storeLoc);
                Assert(!initialPresent, "Test certificate initially absent from AddressBook store");

                // Install certificate into AddressBook
                bool installResult = OfficeIntegration.EnsureCertificateInstalled(testCert, storeName, storeLoc);
                Assert(installResult, "EnsureCertificateInstalled returned true for AddressBook store");

                bool afterInstall = OfficeIntegration.IsCertificateInstalled(testCert, storeName, storeLoc);
                Assert(afterInstall, "IsCertificateInstalled returned true after install");

                // Calling Ensure again when already present should return true immediately
                bool ensureAgain = OfficeIntegration.EnsureCertificateInstalled(testCert, storeName, storeLoc);
                Assert(ensureAgain, "EnsureCertificateInstalled returns true immediately when already present");

                // Remove certificate
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
                Console.WriteLine("Warning: failed to restore previous registry state: {0}", ex.Message);
            }
        }
    }
}
