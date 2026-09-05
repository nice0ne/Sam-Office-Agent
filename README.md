# 🏢 Sam Office Agent

> **Universal Multi-Agent AI Assistant for Microsoft Office (Excel, Word, PowerPoint) with 100% Client-Side BYOK & Standalone Windows Tray Server**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![React](https://img.shields.io/badge/React-18.3-61dafb?logo=react)](https://reactjs.org/)
[![Office.js](https://img.shields.io/badge/Office.js-Universal-D83B01?logo=microsoftoffice)](https://learn.microsoft.com/office/dev/add-ins/)
[![Vitest](https://img.shields.io/badge/Tests-175%20Passing-brightgreen?logo=vitest)](https://vitest.dev/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11%20%7C%20Office%20365-0078D6?logo=windows)](https://microsoft.com)

English | [Bahasa Indonesia](README.id.md)

---

## 🌟 Overview

**Sam Office Agent** is an enterprise-grade AI productivity assistant built as a universal **Microsoft Office Web Add-in**. Operating inside the side Taskpane of **Microsoft Excel, Word, and PowerPoint** (Desktop Office 365 and Office on the Web), it enables intelligent document automation, formula authoring, table analysis, chart creation, and document generation.

Designed around a strict **100% Client-Side BYOK (Bring Your Own Key)** architecture, Sam Office Agent communicates directly from the user’s sandboxed WebView2 browser context to AI provider APIs. **No intermediary server or third-party proxy ever touches your data or API keys.**

For standalone deployment without Node.js or administrator privileges, the project includes a lightweight **native Windows System Tray Server (`SamTrayServer.exe`)** written in C#, providing user-mode HTTPS serving, automatic SSL certificate trust provisioning, and 1-click Office manifest registration.

---

## ✨ Key Features

### 1. 100% Client-Side BYOK (Zero-Backend Privacy)
- **Direct Streaming**: Streams tokens and tool calls directly from AI providers via browser Fetch APIs.
- **Local Key Storage**: API keys are saved exclusively in browser `localStorage` on the client device.
- **Zero Telemetry**: No third-party relays, logging proxies, or external databases.

### 2. Supported AI Providers
- 🟢 **OpenAI**: GPT-4o, GPT-4o mini, o1, o3-mini
- 🟣 **Anthropic Claude**: Claude 3.7 Sonnet, Claude 3.5 Haiku
- 🔵 **Google Gemini**: Gemini 2.5 Flash, Gemini 2.0 Flash, Gemini 2.0 Pro
- 🟠 **GLM Coding Global (Zhipu AI)**: GLM-4-Plus, GLM-4-Air
- ⚡ **OpenRouter / DeepSeek**: DeepSeek R1, DeepSeek V3, Llama 3.3
- 🦙 **Local Ollama**: Offline local LLMs (Llama 3, Qwen 2.5, DeepSeek-Coder)

### 3. Multi-Host Domain Specialists
- **Excel Specialist (`ExcelAgent`)**:
  - Cell data authoring & dynamic formula generation (`SUM`, `XLOOKUP`, `INDEX/MATCH`, `IF`, etc.).
  - Real-time active worksheet data reading (`readActiveSheetData`) with context compression.
  - Range styling: font weights, cell fills, borders, number formatting (currency, percent, dates).
  - Native Excel chart generation (Column, Line, Pie, Bar).
- **Word Specialist (`WordAgent`)**:
  - Structured content generation, document editing, and section drafting.
  - Formatted paragraphs, headings, bullet lists, and tables.
- **PowerPoint Specialist (`PPTAgent`)**:
  - Slide creation, bulleted content placement, and speaker notes authoring.
- **Sam Coordinator (`SamCoordinator`)**:
  - Host auto-detection (`Office.context.host`) and request routing.
  - Multi-turn conversation memory with thinking animation and Dark Mode synchronization.

### 4. Copilot & Autopilot Modes
- **Copilot Mode (Default)**: Visual Action Cards preview changes (target range, formula, values, formatting) before execution. The user clicks **"Terapkan ke Dokumen" (Apply to Document)** to confirm.
- **Autopilot Mode**: Executes actions directly and proactively into the document without intermediate confirmation, ideal for fast iteration and rapid drafting.

### 5. Standalone Windows System Tray Server (`SamTrayServer.exe`)
- **Zero External Dependencies**: Compiled using native Windows `csc.exe` (.NET Framework 4.7.2+). No Node.js or Python required on client machines.
- **User-Mode HTTPS Server**: Implemented via `TcpListener` and `SslStream` on `https://127.0.0.1:5173/`. Requires **zero Administrator / UAC elevation**.
- **Automated Certificate Trust**: Automatically installs the Developer CA to `CurrentUser\Root` store to eliminate WebView2 certificate warning banners.
- **Office Sideloading**: Automatically registers `manifest.xml` in `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
- **System Tray Controls**: Right-click context menu for Server Status, Open in Browser, Register/Unregister Add-in, Windows Auto-Start, and Graceful Exit.

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                 Sam Office Agent Desktop                    │
├──────────────────────────────┬──────────────────────────────┤
│  Windows System Tray         │  Microsoft Office (WebView2) │
│  (SamTrayServer.exe)         │  Excel / Word / PowerPoint   │
│  ├── Context Menu Controls   │  ├── React 18 + Tailwind UI  │
│  ├── User-Mode SslStream     │  ├── Sam Coordinator Agent   │
│  │   HTTPS (Port 5173)       │  ├── Specialist Agents       │
│  ├── Auto-Cert Provisioning  │  ├── Heuristic Validators    │
│  └── WEF Registry Helper     │  └── Office.js Safe Drivers  │
└──────────────┬───────────────┴──────────────┬───────────────┘
               │ Serves Static Bundle         │ Direct BYOK API
               ▼                              ▼
      https://localhost:5173/     LLM Providers (Client-Side)
      ├── dist/assets/*.js        ├── OpenAI / Anthropic / Gemini
      └── dist/assets/*.css       ├── GLM / OpenRouter / Ollama
```

---

## 🚀 Quick Start for Users (Portable Release)

For end users wanting to use Sam Office Agent without installing development tools:

1. Download **`SamOfficeAgent-Portable.zip`** from the [Releases](../../releases) page.
2. Extract the ZIP file to any folder on your computer.
3. Double-click **`Install.bat`**.
   - The application copies to `%LocalAppData%\SamOfficeAgent`.
   - The developer certificate is registered securely.
   - The add-in is registered into Microsoft Office.
   - `SamTrayServer.exe` starts automatically in your Windows System Tray (near the clock).
4. Open **Microsoft Excel**, **Word**, or **PowerPoint**:
   - Go to the **Home** or **Insert** tab -> **Add-ins** -> **Shared Folder / Developer Add-ins**.
   - Click **Sam Office Agent** to open the taskpane.
5. Click **Settings** (⚙️) in the taskpane header to configure your chosen AI provider API key.

To completely remove the application, simply run **`Uninstall.bat`**.

---

## 💻 Developer Setup & Development Workflow

### Prerequisites
- Windows 10 or 11
- Node.js 18.0 or newer
- Microsoft Office 365 Desktop (or Office on the Web)

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/Sam-Office-Agent/Sam-Office-Agent.git
cd Sam-Office-Agent
npm install
```

### 2. Generate Local Developer SSL Certificates
```bash
npx office-addin-dev-certs install
```

### 3. Run Development Server
```bash
npm run dev
```
The React development server will start on `https://localhost:5173/`. You can open `https://localhost:5173/` in Google Chrome or Microsoft Edge to test in **Mock Browser Mode** (no Office desktop required).

### 4. Sideload Add-in to Office Desktop
In a new PowerShell window:
```bash
npm run sideload
```
To clean up sideloaded registry keys when development is finished:
```bash
npm run sideload:clean
```

### 5. Run Automated Tests
```bash
# Run unit and integration tests (175 tests)
npx vitest run

# Run TypeScript type check
npx tsc --noEmit
```

### 6. Build Standalone Windows Tray Package
```bash
npm run build:tray
```
This automated PowerShell pipeline will:
1. Run `npm run build` to create `dist/`.
2. Provision and convert SSL certificates (`ca.crt`, `localhost.pfx`).
3. Compile native `src-tray/*.cs` using native Windows `csc.exe` into `bin/SamTrayServer.exe`.
4. Stage distribution files in `dist-release/`.
5. Package the portable release into `release/SamOfficeAgent-Portable.zip` (and compile `SamOfficeAgent-Setup.exe` if Inno Setup `ISCC.exe` is installed).

---

## 📁 Project Structure

```text
Sam-Office-Agent/
├── manifest.xml                 # Office universal add-in manifest (Excel, Word, PowerPoint)
├── package.json                 # Project dependencies, scripts, and build triggers
├── tsconfig.json                # Strict TypeScript configuration
├── vite.config.ts               # Vite bundler configuration with HTTPS
├── docs/                        # Architecture specs and SDD implementation plans
├── installer/                   # Inno Setup (.iss) and portable batch scripts (.bat)
├── scripts/                     # Automated PowerShell build & sideload scripts
│   ├── build-tray-installer.ps1 # End-to-end packaging pipeline
│   ├── sideload.ps1             # Developer Office registry sideloading
│   └── sideload-clean.ps1       # Sideload registry cleanup
├── src/                         # React Frontend Application
│   ├── App.tsx                  # Main taskpane container, layout, and mode switcher
│   ├── main.tsx                 # Office.onReady() lifecycle bootstrap
│   ├── agents/                  # Multi-agent coordination and domain specialists
│   │   ├── types.ts             # Agent interfaces, tool definitions, execution context
│   │   ├── coordinator/         # Sam Coordinator agent & prompt orchestrator
│   │   ├── excel/               # Excel specialist agent & tool executors
│   │   ├── word/                # Word specialist agent & tool executors
│   │   └── powerpoint/          # PowerPoint specialist agent & tool executors
│   ├── components/              # UI components (Chat, Actions, Settings, Common)
│   ├── services/                # Storage, Office.js drivers, and LLM adapters
│   │   ├── llm/                 # Direct client-side streaming adapters (OpenAI, Gemini, Claude, etc.)
│   │   ├── office/              # Office.js batch sync drivers & MockOfficeDriver
│   │   └── storage/             # localStorage settings persistence with encryption-ready design
│   └── utils/                   # Formula validators, context compressors, and theme sync
├── src-tray/                    # Standalone C# Windows Tray Application (.NET 4.7.2+)
│   ├── HttpServer.cs            # In-process user-mode HTTPS server (TcpListener + SslStream)
│   ├── OfficeIntegration.cs     # Registry WEF sideloading & CurrentUser\Root CA provisioning
│   └── Program.cs               # WinForms System Tray NotifyIcon, Mutex & CLI entry point
└── tests/                       # Automated test suites
    ├── app.test.tsx             # End-to-end taskpane integration tests
    └── tray/                    # Native C# compilation, server, and integration test scripts
```

---

## 🧪 Test Coverage

The project maintains rigorous test coverage across frontend components, LLM streaming adapters, heuristic formula validators, Office drivers, and native Windows server components:

| Test Suite | Files | Tests | Status |
|---|---|---|---|
| **Formula Validator** | `src/utils/__tests__/formulaValidator.test.ts` | 6 | ✅ Passing |
| **Context Compressor** | `src/utils/__tests__/contextCompressor.test.ts` | 4 | ✅ Passing |
| **Theme & Dark Mode** | `src/utils/__tests__/theme.test.ts` | 10 | ✅ Passing |
| **Settings Storage** | `src/services/storage/__tests__/settingsStorage.test.ts` | 8 | ✅ Passing |
| **Mock Office Driver** | `src/services/office/__tests__/mockDriver.test.ts` | 11 | ✅ Passing |
| **BYOK LLM Adapters** | `src/services/llm/__tests__/llmAdapters.test.ts` | 26 | ✅ Passing |
| **Specialist Agents** | `src/agents/__tests__/specialistAgents.test.ts` | 26 | ✅ Passing |
| **Coordinator Agent** | `src/agents/coordinator/__tests__/samCoordinator.test.ts` | 4 | ✅ Passing |
| **Header & Settings UI** | `src/components/__tests__/headerAndSettings.test.tsx` | 18 | ✅ Passing |
| **Chat & Action Cards** | `src/components/__tests__/chatAndActions.test.tsx` | 30 | ✅ Passing |
| **App Integration** | `tests/app.test.tsx` | 26 | ✅ Passing |
| **Type Definitions** | `tests/types.test.ts` | 6 | ✅ Passing |
| **C# HTTPS Server** | `tests/tray/test-server.ps1` | 12 | ✅ Passing |
| **C# Office Integration** | `tests/tray/test-integration.ps1` | 38 | ✅ Passing |
| **C# Tray App & Mutex** | `tests/tray/test-tray-app.ps1` | 4 | ✅ Passing |
| **Installer Scripts** | `tests/tray/test-installer.ps1` | 6 | ✅ Passing |
| **Total Automated Tests**| **16 Test Suites** | **235 Tests** | **100% Passing** |

---

## 🛡️ Security & Privacy

- **Zero-Admin Elevation**: Neither the installer nor `SamTrayServer.exe` requires Windows Administrator privileges. Everything runs in user-space targeting `HKCU` and `CurrentUser\Root`.
- **Local Key Storage**: API keys are saved exclusively in browser `localStorage`. They are never sent to any server other than the designated AI provider endpoint.
- **Directory Traversal Protection**: The in-process HTTPS server uses strict canonical path normalization (`Path.GetFullPath`) to prevent path traversal attacks (`../`).
- **Batch Synchronization**: Office.js mutations are executed in atomic batches using single `context.sync()` calls to avoid document corruption or race conditions.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
