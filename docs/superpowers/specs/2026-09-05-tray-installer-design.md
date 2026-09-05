# Standalone Windows Tray Server & Installer Specification for Sam Office Agent

- **Date:** 2026-09-05
- **Status:** Proposed
- **Target OS:** Windows 10 / Windows 11
- **Component:** Desktop Packaging, Local Tray HTTPS Server & Inno Setup Installer

---

## 1. Executive Summary

Sam Office Agent currently runs during development via Node.js (`npm run dev`) with certificates installed via `office-addin-dev-certs` and sideloaded via PowerShell scripts.

For distribution to end-users (clients, colleagues, or enterprise environments without Node.js), this specification designs a self-contained, lightweight, zero-dependency Windows desktop distribution:
1. **`SamTrayServer.exe`**: A native Windows System Tray application written in C# and compiled using the native Windows compiler (`csc.exe`), eliminating external SDK dependencies. It runs an in-process user-mode HTTPS web server (`TcpListener` + `SslStream`) serving the React static build (`dist/`) at `https://localhost:5173/`.
2. **Automated Trust Provisioning**: Automatically imports the local Developer CA certificate into Windows `CurrentUser\Root` store without command prompt prompts, preventing WebView2 security certificate blocks in Excel, Word, and PowerPoint.
3. **Automated Office Sideloading**: Registers `manifest.xml` into the Windows Registry (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`) on startup, making the add-in immediately visible on the Office ribbon.
4. **`SamOfficeAgent-Setup.exe`**: An Inno Setup installer script packaging the server, assets, and manifest into `%LocalAppData%\SamOfficeAgent` with Windows startup shortcuts and clean uninstallation.

---

## 2. Architecture & Component Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                 SamOfficeAgent-Setup.exe                    │
│             (Self-contained Windows Installer)              │
└──────────────────────────────┬──────────────────────────────┘
                               │ Extracts to %LocalAppData%\SamOfficeAgent
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                     %LocalAppData%\SamOfficeAgent           │
│  ├── SamTrayServer.exe (Native Windows Tray + HTTPS Server) │
│  ├── manifest.xml      (Office Universal Manifest)          │
│  ├── certs/            (CA & localhost certificates)        │
│  └── dist/             (Compiled React Application Bundle)  │
└──────────────────────────────┬──────────────────────────────┘
                               │ Runs in background
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                 SamTrayServer.exe (System Tray)             │
│  ├── Context Menu: Status, Open Browser, Sideload, Exit     │
│  ├── SslStream HTTPS Server (Port 5173, User Mode)          │
│  ├── Auto-Imports CA to Cert:\CurrentUser\Root              │
│  └── Auto-Registers Manifest to HKCU:...\WEF\Developer      │
└──────────────────────────────┬──────────────────────────────┘
                               │ Serves HTTPS
                               ▼
┌─────────────────────────────────────────────────────────────┐
│              Microsoft Office Desktop (WebView2)            │
│         Excel / Word / PowerPoint Ribbon Taskpane           │
└─────────────────────────────────────────────────────────────┘
```

---

## 3. Detailed Component Specifications

### 3.1 Tray Server & Web Server (`src-tray/Program.cs`)

#### A. Compilation & Runtime
- **Language**: C# (.NET Framework 4.7.2+).
- **Compiler**: Windows built-in `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.
- **Dependencies**: None. Uses standard assemblies (`System.dll`, `System.Windows.Forms.dll`, `System.Drawing.dll`, `System.Security.dll`).
- **Memory Footprint**: ~10–18 MB RAM (compared to ~80–120 MB for a Node.js process).

#### B. User-Mode HTTPS Web Server (`TcpListener` + `SslStream`)
- **Port**: `5173` (matching `manifest.xml` source location `https://localhost:5173/`).
- **Why `TcpListener` + `SslStream` instead of `HttpListener`**:
  - `System.Net.HttpListener` relies on Windows kernel `http.sys` which requires administrator rights to execute `netsh http add sslcert`.
  - `TcpListener` + `SslStream` runs entirely in user-space without administrator elevation or UAC prompts.
- **MIME Type Handling**:
  - `.html` -> `text/html; charset=utf-8`
  - `.js` -> `application/javascript; charset=utf-8`
  - `.css` -> `text/css; charset=utf-8`
  - `.png` -> `image/png`
  - `.json` -> `application/json; charset=utf-8`
  - `.svg` -> `image/svg+xml`
- **SPA Fallback**: Any non-file route (e.g., `/settings`, `/chat`) serves `index.html`.
- **CORS Headers**: Sends `Access-Control-Allow-Origin: *` to prevent cross-origin issues in WebView2.

#### C. System Tray UI & Menu
- Uses `System.Windows.Forms.NotifyIcon`.
- Tray icon: Bundled embedded icon or loaded from `icon-32.png`.
- Context Menu:
  - 🟢 **Status: Berjalan (https://localhost:5173)** [Disabled item, shows state]
  - 🌐 **Buka di Browser** [Launches default browser to `https://localhost:5173/`]
  - 📋 **Daftarkan ke Office** [Registers manifest to WEF registry]
  - 🗑️ **Copot dari Office** [Removes manifest from WEF registry]
  - ⚙️ **Jalankan saat Windows Menyala** [Checkable, toggles Run registry key]
  - 🔄 **Muat Ulang Server (Restart)**
  - ❌ **Keluar (Exit)**
- Balloon tip notification on start: *"Sam Office Agent aktif di latar belakang."*

---

### 3.2 Certificate Provisioning Routine

1. **Certificate Generation & Loading:**
   - The application checks for `certs/localhost.pfx` (or `localhost.crt` + `localhost.key`).
   - If missing, it automatically creates a secure self-signed PFX using .NET `System.Security.Cryptography.X509Certificates` or copies from `~/.office-addin-dev-certs`.
2. **Root CA Import:**
   - Checks `X509Store(StoreName.Root, StoreLocation.CurrentUser)` for `Developer CA for Microsoft Office Add-ins`.
   - If not found, imports `ca.crt` using `store.Add(cert)`.
   - Once present, Windows, Edge, and Office WebView2 validate the certificate as trusted, completely eliminating the *"The content is blocked because it isn't signed by a valid security certificate"* error.

---

### 3.3 Registry Management & Auto-Sideloading

1. **Office WEF Registration:**
   - Registry path: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
   - Key name: `d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23` (Manifest ID from `manifest.xml`)
   - Value: Full path to installed `manifest.xml` (e.g. `%LocalAppData%\SamOfficeAgent\manifest.xml`)
2. **Auto-Start Registration:**
   - Registry path: `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
   - Key name: `SamOfficeAgent`
   - Value: `"%LocalAppData%\SamOfficeAgent\SamTrayServer.exe"`

---

### 3.4 Installer Specification (`installer/setup.iss`)

- **Tool**: Inno Setup 6 (open-source Windows installer compiler).
- **Application Name**: Sam Office Agent
- **Target Folder**: `{localappdata}\SamOfficeAgent` (does not require administrator privileges).
- **Installer Features**:
  - Modern wizard style (flat, DPI-aware).
  - Automatically kills running `SamTrayServer.exe` before update/install.
  - Copies `dist/`, `manifest.xml`, `certs/`, and `SamTrayServer.exe`.
  - Adds uninstaller to Windows *Installed Apps* (*Add/Remove Programs*).
  - Checkbox: *"Jalankan Sam Office Agent secara otomatis saat Windows mulai"*.
  - Post-install action: Launches `SamTrayServer.exe`.
- **Uninstaller Cleanup**:
  - Closes tray server.
  - Removes `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` key.
  - Removes `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` key.
  - Deletes installation directory.

---

### 3.5 Build Automation Script (`scripts/build-tray-installer.ps1`)

Automates end-to-end packaging into a single command:
1. Runs `npm run build` to generate clean `dist/`.
2. Compiles `src-tray/Program.cs` with `csc.exe` into `dist-tray/SamTrayServer.exe`.
3. Copies `manifest.xml` and SSL certificates into `dist-tray/`.
4. If Inno Setup (`ISCC.exe`) is installed: Compiles `installer/setup.iss` into `release/SamOfficeAgent-Setup.exe`.
5. If Inno Setup is not installed: Packages `dist-tray/` into `release/SamOfficeAgent-Portable.zip` with a 1-click `Install.bat`.

---

## 4. Error Handling & Edge Cases

| Scenario | Handled Behavior |
| :--- | :--- |
| **Port 5173 already in use** | Detects conflict; displays balloon tooltip: *"Port 5173 sedang dipakai oleh aplikasi lain"* and offers restart. |
| **User starts Excel before Tray Server** | Add-in shows taskpane error until Tray Server is started; menu has "Restart Server" to quickly recover. |
| **Untrusted SSL Certificate** | Self-healing: On launch, Tray Server checks cert store and re-applies `ca.crt` to `CurrentUser\Root`. |
| **Multiple Instances** | Uses named `Mutex` (`SamOfficeAgentTrayServerMutex`) so clicking the exe when already running brings up balloon tip instead of failing. |
| **Clean Uninstall** | Uninstaller removes the WEF developer registry entry so Office does not retain orphan add-in icons. |

---

## 5. Verification Plan

1. **C# Compilation**: Verify `csc.exe` compiles `src-tray/Program.cs` without warnings.
2. **Local HTTPS Server Test**: Run `SamTrayServer.exe` and test via `curl -k https://localhost:5173/` and in Chrome/Edge.
3. **Office Sideloading Test**: Verify `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` is populated and Excel Desktop opens with Sam on ribbon.
4. **System Tray Menu Test**: Verify Open Browser, Status, Register/Unregister, and Exit buttons.
5. **Installer & Uninstaller Test**: Execute installer, verify files in `%LocalAppData%\SamOfficeAgent`, verify add-in in Excel, test clean uninstall.
