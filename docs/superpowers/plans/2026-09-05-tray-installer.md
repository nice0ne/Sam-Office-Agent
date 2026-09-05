# Standalone Windows Tray Server & Installer Implementation Plan

English | [Bahasa Indonesia](2026-09-05-tray-installer.id.md)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Package Sam Office Agent into an autonomous Windows System Tray application (`SamTrayServer.exe`) in C# (compiled via native Windows `csc.exe`) with a user-mode HTTPS server, certificate trust provisioning, Office manifest sideloading, Inno Setup script (`setup.iss`), and build automation (`build-tray-installer.ps1`).

**Architecture:** A native C# tray application hosts an in-process user-mode HTTPS web server (`TcpListener` + `SslStream`) serving the React static bundle (`dist/`) at `https://localhost:5173/`. It automatically provisions the local developer CA into the Windows `CurrentUser\Root` certificate store (eliminating WebView2 certificate blocks) and registers `manifest.xml` into the Office WEF registry (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`). An Inno Setup script packages the application into `%LocalAppData%\SamOfficeAgent` with auto-start and uninstallation.

**Tech Stack:** C# (.NET Framework 4.7.2+ compiled via native Windows `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`), WinForms `NotifyIcon`, `System.Net.Security.SslStream`, Inno Setup 6, PowerShell 5.1.

**Spec:** [`docs/superpowers/specs/2026-09-05-tray-installer-design.md`](file:///D:/VIBE-CODING/Sam-Office-Agent/docs/superpowers/specs/2026-09-05-tray-installer-design.md)

## Global Constraints

- Must run with zero external runtime dependencies on the target Windows machine (no Node.js or Python required on client).
- Must use user-mode sockets (`TcpListener` + `SslStream`) rather than `http.sys` so administrator elevation is never required to run.
- Must preserve 100% Bring-Your-Own-Key (BYOK) privacy — all keys remain strictly in browser `localStorage`.
- Server port must bind to `127.0.0.1:5173` matching `<SourceLocation DefaultValue="https://localhost:5173/"/>` in `manifest.xml`.
- Target installation directory is `{localappdata}\SamOfficeAgent` to permit non-admin installation.

---

### Task 1: C# User-Mode HTTPS Web Server (`src-tray/HttpServer.cs`)

**Files:**
- Create: `src-tray/HttpServer.cs`
- Test: `tests/tray/test-server.ps1`

**Interfaces:**
- Consumes: Static files in `dist/` directory, SSL certificate in `.pfx` or `.crt`+`.key` format.
- Produces: `HttpServer` class with `Start(int port, string rootDir, X509Certificate2 cert)` and `Stop()` methods.

- [ ] **Step 1: Write test script for HTTPS server**

Create `tests/tray/test-server.ps1` to test that a C# HTTPS server compiled with `csc.exe` responds with HTTP 200 and serves `index.html` with correct MIME types over SSL.

```powershell
$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"
if (!(Test-Path $binDir)) { New-Item -ItemType Directory -Path $binDir -Force | Out-Null }

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "TestServer.exe"

Write-Host "Compiling TestServer.exe..."
& $csc /nologo /target:exe /out:$outExe "$srcTray\HttpServer.cs" "$srcTray\TestProgram.cs"

if ($LASTEXITCODE -ne 0) {
    Write-Error "Kompilasi C# gagal."
    exit 1
}

Write-Host "TestServer compiled successfully."
```

- [ ] **Step 2: Run test script to verify it fails (RED)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-server.ps1`
Expected: FAIL because `src-tray/HttpServer.cs` and `src-tray/TestProgram.cs` do not exist yet.

- [ ] **Step 3: Implement `src-tray/HttpServer.cs`**

Create `src-tray/HttpServer.cs` implementing `TcpListener` + `SslStream`:
- Listens on `IPAddress.Loopback` (`127.0.0.1`) port `5173`.
- Authenticates server SSL stream with `X509Certificate2`.
- Reads incoming HTTP/1.1 request lines.
- Maps path to `dist/` root directory (falling back to `/index.html` for SPA routing).
- Sends appropriate Content-Type headers (`text/html`, `application/javascript`, `text/css`, `image/png`, `application/json`, etc.).
- Includes `Access-Control-Allow-Origin: *` headers for WebView2.

Create `src-tray/TestProgram.cs` as a minimal entry point to verify server compilation and start/stop.

- [ ] **Step 4: Run test script to verify it passes (GREEN)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-server.ps1`
Expected: PASS with "TestServer compiled successfully."

- [ ] **Step 5: Commit**

```bash
git add src-tray/ tests/tray/
git commit -m "feat(tray): implement user-mode HTTPS static web server in C#"
```

---

### Task 2: Certificate Trust & Office Sideloading Helper (`src-tray/OfficeIntegration.cs`)

**Files:**
- Create: `src-tray/OfficeIntegration.cs`
- Test: `tests/tray/test-integration.ps1`

**Interfaces:**
- Consumes: Windows Certificate Store APIs (`X509Store`), Windows Registry APIs (`Microsoft.Win32.Registry`).
- Produces: `OfficeIntegration` class with:
  - `EnsureCertificateInstalled(string caCertPath)`: Installs developer CA to `CurrentUser\Root` if missing.
  - `RegisterAddIn(string manifestPath)`: Registers manifest in `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
  - `UnregisterAddIn()`: Removes registry key from `WEF\Developer`.
  - `IsAddInRegistered(string manifestPath)`: Checks registration state.
  - `SetAutoStart(bool enable, string exePath)`: Toggles startup in `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`.

- [ ] **Step 1: Write test script for Office Integration**

Create `tests/tray/test-integration.ps1` to test that `OfficeIntegration.cs` compiles cleanly and can query registry and certificate stores.

```powershell
$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "TestIntegration.exe"

& $csc /nologo /target:exe /out:$outExe "$srcTray\OfficeIntegration.cs" "$srcTray\TestIntegrationProgram.cs"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Compilation of OfficeIntegration failed."
    exit 1
}

Write-Host "OfficeIntegration compiled successfully."
```

- [ ] **Step 2: Run test to verify it fails (RED)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-integration.ps1`
Expected: FAIL because `src-tray/OfficeIntegration.cs` does not exist yet.

- [ ] **Step 3: Implement `src-tray/OfficeIntegration.cs`**

Implement:
- `EnsureCertificateInstalled(string caCertPath)`: Checks `X509Store(StoreName.Root, StoreLocation.CurrentUser)` for certificate matching the thumbprint or subject of `ca.crt`. If absent, opens store with `OpenFlags.ReadWrite` and calls `store.Add(cert)`.
- `RegisterAddIn(string manifestPath)`: Opens `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` with write access, sets name `d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23` to `manifestPath`.
- `UnregisterAddIn()`: Deletes value `d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23`.
- `SetAutoStart(bool enable, string exePath)`: Opens `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` and sets/deletes `SamOfficeAgent`.

Create `src-tray/TestIntegrationProgram.cs` testing methods.

- [ ] **Step 4: Run test to verify it passes (GREEN)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-integration.ps1`
Expected: PASS with "OfficeIntegration compiled successfully."

- [ ] **Step 5: Commit**

```bash
git add src-tray/ tests/tray/
git commit -m "feat(tray): implement certificate trust and Office WEF registry integration"
```

---

### Task 3: Windows System Tray Application (`src-tray/Program.cs`)

**Files:**
- Create: `src-tray/Program.cs`
- Modify: Remove `src-tray/TestProgram.cs`, `src-tray/TestIntegrationProgram.cs`
- Test: `tests/tray/test-tray-app.ps1`

**Interfaces:**
- Consumes: `HttpServer` from Task 1, `OfficeIntegration` from Task 2.
- Produces: Complete WinForms System Tray application with context menu and lifecycle management.

- [ ] **Step 1: Write test script for full tray app compilation**

Create `tests/tray/test-tray-app.ps1` to compile `SamTrayServer.exe` with WinForms and Windows Drawing references.

```powershell
$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "SamTrayServer.exe"

$srcFiles = (Get-ChildItem -Path $srcTray -Filter "*.cs" | Select-Object -ExpandProperty FullName) -join " "

Write-Host "Compiling SamTrayServer.exe..."
& $csc /nologo /target:winexe /out:$outExe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll (Get-ChildItem -Path $srcTray -Filter "*.cs" | Select-Object -ExpandProperty FullName)

if ($LASTEXITCODE -ne 0) {
    Write-Error "Kompilasi SamTrayServer.exe gagal."
    exit 1
}

Write-Host "SamTrayServer.exe compiled successfully: $outExe"
```

- [ ] **Step 2: Run test to verify it fails (RED)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-tray-app.ps1`
Expected: FAIL because `src-tray/Program.cs` is not yet created.

- [ ] **Step 3: Implement `src-tray/Program.cs`**

Implement main WinForms Application:
- Named mutex `Global\SamOfficeAgentTrayServerMutex` to enforce single-instance.
- Determines base path (same directory as `SamTrayServer.exe`).
- Initializes `OfficeIntegration` and auto-provisions CA certificate if present in `certs/ca.crt`.
- Auto-registers `manifest.xml` if present in same directory.
- Initializes `HttpServer` with `certs/localhost.pfx` (or generates one dynamically) and root `dist/`.
- Initializes `NotifyIcon` with tooltip: `"Sam Office Agent (Aktif di https://localhost:5173)"`.
- Context menu:
  - Header item: *"Sam Office Agent - v1.0"* (disabled font-bold).
  - Status item: *"🟢 Server Aktif (Port 5173)"*.
  - Item *"Buka di Browser"* -> `Process.Start("https://localhost:5173/")`.
  - Item *"Daftarkan ke Office"* -> calls `OfficeIntegration.RegisterAddIn()`.
  - Item *"Copot dari Office"* -> calls `OfficeIntegration.UnregisterAddIn()`.
  - Item *"Mulai saat Windows Menyala"* (Checkable) -> calls `OfficeIntegration.SetAutoStart()`.
  - Separator.
  - Item *"Keluar"* -> gracefully stops server and calls `Application.Exit()`.
- Runs `Application.Run()`.

Remove temporary test programs.

- [ ] **Step 4: Run test to verify compilation passes (GREEN)**

Run: `powershell -ExecutionPolicy Bypass -File tests/tray/test-tray-app.ps1`
Expected: PASS with "SamTrayServer.exe compiled successfully".

- [ ] **Step 5: Commit**

```bash
git add src-tray/ tests/tray/
git commit -m "feat(tray): implement native Windows System Tray application and context menu"
```

---

### Task 4: Inno Setup Installer Script & Portable Fallback (`installer/setup.iss`)

**Files:**
- Create: `installer/setup.iss`
- Create: `installer/Install.bat`
- Create: `installer/Uninstall.bat`

**Interfaces:**
- Consumes: `dist/`, `manifest.xml`, `certs/`, `bin/SamTrayServer.exe`.
- Produces: Inno Setup installer script for compiling `release/SamOfficeAgent-Setup.exe` and portable batch scripts.

- [ ] **Step 1: Create Inno Setup script `installer/setup.iss`**

Write `installer/setup.iss`:
- `[Setup]`
  - `AppName=Sam Office Agent`
  - `AppVersion=1.0.0`
  - `DefaultDirName={localappdata}\SamOfficeAgent`
  - `PrivilegesRequired=lowest` (no admin elevation required).
  - `OutputBaseFilename=SamOfficeAgent-Setup`
  - `OutputDir=..\release`
  - `CloseApplications=force`
  - `CloseApplicationsFilter=SamTrayServer.exe`
- `[Files]`
  - Include `..\dist-release\*` recursively into `{app}`.
- `[Icons]`
  - Desktop shortcut (optional).
  - Startup shortcut `{userstartup}\Sam Office Agent.lnk` pointing to `{app}\SamTrayServer.exe`.
- `[Run]`
  - `Filename: "{app}\SamTrayServer.exe"; Description: "Jalankan Sam Office Agent sekarang"; Flags: nowait postinstall skipifsilent`
- `[UninstallRun]`
  - Run unregistration helper before deleting files.

- [ ] **Step 2: Create portable fallback `installer/Install.bat` and `installer/Uninstall.bat`**

For users who want 1-click installation without compiling Inno Setup:
- `Install.bat`: Copies files to `%LocalAppData%\SamOfficeAgent`, registers add-in in registry, adds startup shortcut, and launches `SamTrayServer.exe`.
- `Uninstall.bat`: Kills `SamTrayServer.exe`, removes registry keys, and deletes `%LocalAppData%\SamOfficeAgent`.

- [ ] **Step 3: Commit**

```bash
git add installer/
git commit -m "feat(installer): create Inno Setup script and portable batch installers"
```

---

### Task 5: End-to-End Build Automation (`scripts/build-tray-installer.ps1`)

**Files:**
- Create: `scripts/build-tray-installer.ps1`
- Modify: `package.json`

**Interfaces:**
- Consumes: Whole codebase (`src/`, `src-tray/`, `manifest.xml`, `installer/`).
- Produces: `release/SamOfficeAgent-Portable/` and `release/SamOfficeAgent-Setup.exe` (if ISCC available).

- [ ] **Step 1: Create `scripts/build-tray-installer.ps1`**

Write PowerShell automation script:
1. Run `npm run build` to produce clean `dist/`.
2. Ensure SSL certificate files exist:
   - Check `~/.office-addin-dev-certs` for `ca.crt`, `localhost.crt`, `localhost.key`.
   - Convert/export to `localhost.pfx` via PowerShell `New-Object System.Security.Cryptography.X509Certificates.X509Certificate2` or Node script.
3. Compile `src-tray/*.cs` into `SamTrayServer.exe` using `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.
4. Stage distribution directory `dist-release/`:
   - Copy `SamTrayServer.exe`
   - Copy `manifest.xml`
   - Copy `dist/` into `dist-release/dist/`
   - Copy `certs/` into `dist-release/certs/`
   - Copy `installer/Install.bat` and `installer/Uninstall.bat`
5. Check if Inno Setup (`ISCC.exe`) exists:
   - If found, run `& $iscc installer/setup.iss` to produce `release/SamOfficeAgent-Setup.exe`.
   - If not found, zip `dist-release/` into `release/SamOfficeAgent-Portable.zip`.
6. Print summary with file paths and usage instructions.

- [ ] **Step 2: Add scripts in `package.json`**

Add `"build:tray": "powershell -ExecutionPolicy Bypass -File ./scripts/build-tray-installer.ps1"`.

- [ ] **Step 3: Run build automation and verify output**

Run: `npm run build:tray`
Expected:
- Production bundle compiled.
- `SamTrayServer.exe` compiled without errors.
- `release/` folder created with staged package.

- [ ] **Step 4: Run full test suite to ensure zero regressions**

Run: `npx vitest run`
Expected: 175/175 tests pass.

- [ ] **Step 5: Commit**

```bash
git add scripts/ package.json
git commit -m "feat(build): add automated build and packaging pipeline for Windows Tray server"
```

---

## Self-Review Checklist

1. **Spec coverage**:
   - Tray server with C# and `csc.exe` -> Task 1, 3
   - User-mode HTTPS on port 5173 -> Task 1
   - CA certificate trust & Office sideloading -> Task 2
   - Inno Setup script & portable fallback -> Task 4
   - Build automation pipeline -> Task 5
2. **Placeholder scan**: No "TODO", "TBD", or vague placeholders.
3. **Type consistency**: Exact method names and registry keys match throughout.
