using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;

namespace SamOfficeAgent
{
    /// <summary>
    /// Helper for managing Office Add-in developer sideloading,
    /// CA certificate trust, and Windows startup registry settings.
    /// </summary>
    public class OfficeIntegration
    {
        public const string ManifestGuid = "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23";
        public const string WefDeveloperKeyPath = @"Software\Microsoft\Office\16.0\WEF\Developer";
        public const string WindowsRunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        public const string AutoStartAppName = "SamOfficeAgent";

        #region Office Sideloading (WEF Developer Registry)

        /// <summary>
        /// Registers the Office Add-in manifest in HKCU\Software\Microsoft\Office\16.0\WEF\Developer.
        /// </summary>
        /// <param name="manifestPath">Path to manifest.xml. If null, defaults to manifest.xml in application directory.</param>
        /// <returns>True if registration succeeded.</returns>
        public static bool RegisterAddIn(string manifestPath = null)
        {
            try
            {
                if (string.IsNullOrEmpty(manifestPath))
                {
                    manifestPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "manifest.xml");
                }

                string fullPath = Path.GetFullPath(manifestPath);

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(WefDeveloperKeyPath))
                {
                    if (key == null)
                    {
                        return false;
                    }
                    key.SetValue(ManifestGuid, fullPath, RegistryValueKind.String);
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Removes the Office Add-in registration from HKCU\Software\Microsoft\Office\16.0\WEF\Developer.
        /// </summary>
        /// <returns>True if unregistration succeeded (or was already absent).</returns>
        public static bool UnregisterAddIn()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WefDeveloperKeyPath, true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(ManifestGuid, false);
                    }
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks whether the Office Add-in is registered in WEF\Developer.
        /// </summary>
        /// <param name="manifestPath">Optional specific manifest path to match against.</param>
        /// <returns>True if registered (and matches path if specified).</returns>
        public static bool IsAddInRegistered(string manifestPath = null)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WefDeveloperKeyPath, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object val = key.GetValue(ManifestGuid);
                    if (val == null)
                    {
                        return false;
                    }

                    if (string.IsNullOrEmpty(manifestPath))
                    {
                        return true;
                    }

                    string registered = val.ToString();
                    if (string.Equals(registered, manifestPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    try
                    {
                        string normReg = Path.GetFullPath(registered);
                        string normTarget = Path.GetFullPath(manifestPath);
                        return string.Equals(normReg, normTarget, StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns the registered manifest path from WEF\Developer, or null if not registered.
        /// </summary>
        public static string GetRegisteredManifestPath()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WefDeveloperKeyPath, false))
                {
                    if (key == null)
                    {
                        return null;
                    }
                    object val = key.GetValue(ManifestGuid);
                    return val != null ? val.ToString() : null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        #endregion

        #region Windows Auto-Start (Run Registry)

        /// <summary>
        /// Sets or removes Windows startup registration in HKCU\Software\Microsoft\Windows\CurrentVersion\Run.
        /// </summary>
        /// <param name="enable">True to enable auto-start, false to disable.</param>
        /// <param name="exePath">Path to executable. If null, defaults to current process assembly.</param>
        /// <returns>True if the operation succeeded.</returns>
        public static bool SetAutoStart(bool enable, string exePath = null)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(WindowsRunKeyPath))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    if (enable)
                    {
                        if (string.IsNullOrEmpty(exePath))
                        {
                            exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                        }

                        string fullPath = Path.GetFullPath(exePath);
                        string formatted = fullPath.StartsWith("\"") && fullPath.EndsWith("\"")
                            ? fullPath
                            : string.Format("\"{0}\"", fullPath);

                        key.SetValue(AutoStartAppName, formatted, RegistryValueKind.String);
                    }
                    else
                    {
                        key.DeleteValue(AutoStartAppName, false);
                    }

                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks whether auto-start is configured in HKCU\Software\Microsoft\Windows\CurrentVersion\Run.
        /// </summary>
        /// <param name="exePath">Optional executable path to verify.</param>
        /// <returns>True if registered (and matches exePath if specified).</returns>
        public static bool IsAutoStartEnabled(string exePath = null)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WindowsRunKeyPath, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object val = key.GetValue(AutoStartAppName);
                    if (val == null)
                    {
                        return false;
                    }

                    if (string.IsNullOrEmpty(exePath))
                    {
                        return true;
                    }

                    string registered = val.ToString().Trim('\"');
                    string target = Path.GetFullPath(exePath).Trim('\"');
                    return string.Equals(registered, target, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns the registered auto-start executable path, or null if not registered.
        /// </summary>
        public static string GetAutoStartPath()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WindowsRunKeyPath, false))
                {
                    if (key == null)
                    {
                        return null;
                    }
                    object val = key.GetValue(AutoStartAppName);
                    return val != null ? val.ToString().Trim('\"') : null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        #endregion

        #region Certificate Trust Management

        /// <summary>
        /// Locates the Developer CA certificate file if present in the application certs folder or user profile.
        /// </summary>
        public static string FindDeveloperCaCertPath()
        {
            // 1. Check local application directory certs/ca.crt
            string localCerts = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "certs", "ca.crt");
            if (File.Exists(localCerts))
            {
                return localCerts;
            }

            // 2. Check local application directory ca.crt
            string localCa = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ca.crt");
            if (File.Exists(localCa))
            {
                return localCa;
            }

            // 3. Check user profile ~/.office-addin-dev-certs/ca.crt
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string devCert = Path.Combine(userProfile, ".office-addin-dev-certs", "ca.crt");
            if (File.Exists(devCert))
            {
                return devCert;
            }

            return null;
        }

        /// <summary>
        /// Ensures developer CA certificate is installed in CurrentUser\Root certificate store.
        /// Checks if certificate matching thumbprint or subject is already present before installing.
        /// </summary>
        /// <param name="caCertPath">Path to CA certificate file (.crt or .cer).</param>
        /// <returns>True if certificate is present or was successfully installed.</returns>
        public static bool EnsureCertificateInstalled(string caCertPath)
        {
            return EnsureCertificateInstalled(caCertPath, StoreName.Root, StoreLocation.CurrentUser);
        }

        /// <summary>
        /// Ensures developer CA certificate is installed in the specified certificate store.
        /// </summary>
        public static bool EnsureCertificateInstalled(string caCertPath, StoreName storeName, StoreLocation storeLocation)
        {
            if (string.IsNullOrEmpty(caCertPath))
            {
                return false;
            }

            try
            {
                string fullPath = Path.GetFullPath(caCertPath);
                if (!File.Exists(fullPath))
                {
                    return false;
                }

                X509Certificate2 cert = new X509Certificate2(fullPath);
                return EnsureCertificateInstalled(cert, storeName, storeLocation);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Ensures certificate is installed in CurrentUser\Root.
        /// </summary>
        public static bool EnsureCertificateInstalled(X509Certificate2 cert)
        {
            return EnsureCertificateInstalled(cert, StoreName.Root, StoreLocation.CurrentUser);
        }

        /// <summary>
        /// Ensures certificate is installed in the specified store.
        /// </summary>
        public static bool EnsureCertificateInstalled(X509Certificate2 cert, StoreName storeName, StoreLocation storeLocation)
        {
            if (cert == null)
            {
                return false;
            }

            if (IsCertificateInstalled(cert, storeName, storeLocation))
            {
                return true;
            }

            X509Store store = null;
            try
            {
                store = new X509Store(storeName, storeLocation);
                store.Open(OpenFlags.ReadWrite);
                store.Add(cert);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (store != null)
                {
                    store.Close();
                }
            }
        }

        /// <summary>
        /// Checks if a certificate matching the thumbprint or subject of caCertPath is in CurrentUser\Root.
        /// </summary>
        public static bool IsCertificateInstalled(string caCertPath)
        {
            return IsCertificateInstalled(caCertPath, StoreName.Root, StoreLocation.CurrentUser);
        }

        /// <summary>
        /// Checks if a certificate matching the thumbprint or subject of caCertPath is in the specified store.
        /// </summary>
        public static bool IsCertificateInstalled(string caCertPath, StoreName storeName, StoreLocation storeLocation)
        {
            if (string.IsNullOrEmpty(caCertPath))
            {
                return false;
            }

            try
            {
                string fullPath = Path.GetFullPath(caCertPath);
                if (!File.Exists(fullPath))
                {
                    return false;
                }

                X509Certificate2 cert = new X509Certificate2(fullPath);
                return IsCertificateInstalled(cert, storeName, storeLocation);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if certificate is installed in CurrentUser\Root by thumbprint or subject distinguished name.
        /// </summary>
        public static bool IsCertificateInstalled(X509Certificate2 cert)
        {
            return IsCertificateInstalled(cert, StoreName.Root, StoreLocation.CurrentUser);
        }

        /// <summary>
        /// Checks if certificate is installed in the specified store by thumbprint or subject distinguished name.
        /// </summary>
        public static bool IsCertificateInstalled(X509Certificate2 cert, StoreName storeName, StoreLocation storeLocation)
        {
            if (cert == null)
            {
                return false;
            }

            X509Store store = null;
            try
            {
                store = new X509Store(storeName, storeLocation);
                store.Open(OpenFlags.ReadOnly);

                foreach (X509Certificate2 c in store.Certificates)
                {
                    if (string.Equals(c.Thumbprint, cert.Thumbprint, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    if (!string.IsNullOrEmpty(cert.Subject) &&
                        string.Equals(c.Subject, cert.Subject, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (store != null)
                {
                    store.Close();
                }
            }
        }

        /// <summary>
        /// Removes matching certificate from the specified store.
        /// </summary>
        public static bool RemoveCertificate(X509Certificate2 cert, StoreName storeName, StoreLocation storeLocation)
        {
            if (cert == null)
            {
                return false;
            }

            X509Store store = null;
            try
            {
                store = new X509Store(storeName, storeLocation);
                store.Open(OpenFlags.ReadWrite);

                bool removed = false;
                foreach (X509Certificate2 c in store.Certificates)
                {
                    if (string.Equals(c.Thumbprint, cert.Thumbprint, StringComparison.OrdinalIgnoreCase))
                    {
                        store.Remove(c);
                        removed = true;
                    }
                }

                return removed;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (store != null)
                {
                    store.Close();
                }
            }
        }

        #endregion
    }
}
