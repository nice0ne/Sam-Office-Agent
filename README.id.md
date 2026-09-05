# 🏢 Sam Office Agent

> **Universal Multi-Agent AI Assistant for Microsoft Office (Excel, Word, PowerPoint) dengan 100% Client-Side BYOK & Standalone Windows Tray Server**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.6-blue?logo=typescript)](https://www.typescriptlang.org/)
[![React](https://img.shields.io/badge/React-18.3-61dafb?logo=react)](https://reactjs.org/)
[![Office.js](https://img.shields.io/badge/Office.js-Universal-D83B01?logo=microsoftoffice)](https://learn.microsoft.com/office/dev/add-ins/)
[![Vitest](https://img.shields.io/badge/Tests-175%20Passing-brightgreen?logo=vitest)](https://vitest.dev/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11%20%7C%20Office%20365-0078D6?logo=windows)](https://microsoft.com)

[English](README.md) | Bahasa Indonesia

---

## 🌟 Gambaran Umum

**Sam Office Agent** adalah asisten produktivitas AI tingkat enterprise yang dibangun sebagai **Web Add-in Microsoft Office** universal. Beroperasi di dalam Taskpane samping **Microsoft Excel, Word, dan PowerPoint** (Office 365 Desktop dan Office on the Web), aplikasi ini memungkinkan otomatisasi dokumen cerdas, pembuatan formula, analisis tabel, pembuatan grafik, dan pembuatan dokumen.

Dirancang dengan arsitektur ketat **100% Client-Side BYOK (Bring Your Own Key)**, Sam Office Agent berkomunikasi langsung dari konteks peramban WebView2 pengguna yang terisolasi (sandboxed) ke API penyedia AI. **Tidak ada server perantara atau proksi pihak ketiga yang pernah menyentuh data atau kunci API Anda.**

Untuk penerapan mandiri tanpa Node.js atau hak akses administrator, proyek ini menyertakan **Server System Tray Windows bawaan (`SamTrayServer.exe`)** ringan yang ditulis dalam C#, menyediakan layanan HTTPS mode-pengguna, provisi kepercayaan sertifikat SSL otomatis, dan pendaftaran manifes Office dengan 1 klik.

---

## ✨ Fitur Utama

### 1. 100% Client-Side BYOK (Privasi Tanpa Backend)
- **Direct Streaming**: Mengalirkan token dan panggilan alat (tool calls) secara langsung dari penyedia AI melalui Fetch API peramban.
- **Penyimpanan Kunci Lokal**: API key disimpan secara eksklusif di `localStorage` peramban pada perangkat klien.
- **Zero Telemetry**: Tidak ada relai pihak ketiga, proksi pencatatan log, atau basis data eksternal.

### 2. Penyedia AI yang Didukung
- 🟢 **OpenAI**: GPT-4o, GPT-4o mini, o1, o3-mini
- 🟣 **Anthropic Claude**: Claude 3.7 Sonnet, Claude 3.5 Haiku
- 🔵 **Google Gemini**: Gemini 2.5 Flash, Gemini 2.0 Flash, Gemini 2.0 Pro
- 🟠 **GLM Coding Global (Zhipu AI)**: GLM-4-Plus, GLM-4-Air
- ⚡ **OpenRouter / DeepSeek**: DeepSeek R1, DeepSeek V3, Llama 3.3
- 🦙 **Local Ollama**: LLM lokal offline (Llama 3, Qwen 2.5, DeepSeek-Coder)

### 3. Spesialis Domain Multi-Host
- **Spesialis Excel (`ExcelAgent`)**:
  - Penulisan data sel & pembuatan formula dinamis (`SUM`, `XLOOKUP`, `INDEX/MATCH`, `IF`, dll.).
  - Pembacaan data lembar kerja aktif secara real-time (`readActiveSheetData`) dengan kompresi konteks.
  - Penataan gaya rentang (range styling): ketebalan font, warna isi sel, border, pemformatan angka (mata uang, persen, tanggal).
  - Pembuatan grafik native Excel (Column, Line, Pie, Bar).
- **Spesialis Word (`WordAgent`)**:
  - Pembuatan konten terstruktur, pengeditan dokumen, dan perancangan draf bagian.
  - Paragraf terformat, judul (headings), daftar berpoin (bullet lists), dan tabel.
- **Spesialis PowerPoint (`PPTAgent`)**:
  - Pembuatan slide, penempatan konten berpoin, dan penulisan catatan pembicara (speaker notes).
- **Koordinator Sam (`SamCoordinator`)**:
  - Deteksi otomatis host (`Office.context.host`) dan perutean permintaan.
  - Memori percakapan multi-turn dengan animasi berpikir dan sinkronisasi Dark Mode.

### 4. Mode Copilot & Autopilot
- **Mode Copilot (Bawaan)**: Kartu Tindakan Visual mempratinjau perubahan (rentang target, formula, nilai, pemformatan) sebelum dieksekusi. Pengguna mengeklik **"Terapkan ke Dokumen" (Apply to Document)** untuk mengonfirmasi.
- **Mode Autopilot**: Mengeksekusi tindakan secara langsung dan proaktif ke dalam dokumen tanpa konfirmasi perantara, ideal untuk iterasi cepat dan pembuatan draf kilat.

### 5. Server System Tray Windows Mandiri (`SamTrayServer.exe`)
- **Nol Dependensi Eksternal**: Dikompilasi menggunakan `csc.exe` bawaan Windows (.NET Framework 4.7.2+). Tidak memerlukan Node.js atau Python di komputer klien.
- **User-Mode HTTPS Server**: Diimplementasikan melalui `TcpListener` dan `SslStream` pada `https://127.0.0.1:5173/`. Membutuhkan **nol elevasi Administrator / UAC**.
- **Kepercayaan Sertifikat Otomatis**: Memasang Developer CA secara otomatis ke penyimpanan `CurrentUser\Root` untuk menghilangkan banner peringatan sertifikat WebView2.
- **Sideloading Office**: Mendaftarkan `manifest.xml` secara otomatis di `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
- **Kontrol System Tray**: Menu konteks klik kanan untuk Status Server, Buka di Peramban, Daftarkan/Hapus Pendaftaran Add-in, Auto-Start Windows, dan Keluar Bersih (Graceful Exit).

---

## 🏗️ Arsitektur

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

## 🚀 Panduan Ringkas untuk Pengguna (Rilis Portabel)

Bagi pengguna akhir yang ingin menggunakan Sam Office Agent tanpa menginstal perkakas pengembangan:

1. Unduh **`SamOfficeAgent-Portable.zip`** dari halaman [Releases](../../releases).
2. Ekstrak file ZIP ke folder mana pun di komputer Anda.
3. Klik dua kali **`Install.bat`**.
   - Aplikasi disalin ke `%LocalAppData%\SamOfficeAgent`.
   - Sertifikat pengembang didaftarkan dengan aman.
   - Add-in didaftarkan ke dalam Microsoft Office.
   - `SamTrayServer.exe` dimulai secara otomatis di System Tray Windows Anda (di dekat jam).
4. Buka **Microsoft Excel**, **Word**, atau **PowerPoint**:
   - Buka tab **Home** atau **Insert** -> **Add-ins** -> **Shared Folder / Developer Add-ins**.
   - Klik **Sam Office Agent** untuk membuka taskpane.
5. Klik **Pengaturan** (⚙️) di header taskpane untuk mengonfigurasi API key penyedia AI pilihan Anda.

Untuk menghapus aplikasi secara menyeluruh, cukup jalankan **`Uninstall.bat`**.

---

## 💻 Penyiapan Pengembang & Alur Kerja Pengembangan

### Prasyarat
- Windows 10 atau 11
- Node.js 18.0 atau lebih baru
- Microsoft Office 365 Desktop (atau Office on the Web)

### 1. Clone & Pasang Dependensi
```bash
git clone https://github.com/Sam-Office-Agent/Sam-Office-Agent.git
cd Sam-Office-Agent
npm install
```

### 2. Buat Sertifikat SSL Pengembang Lokal
```bash
npx office-addin-dev-certs install
```

### 3. Jalankan Server Pengembangan
```bash
npm run dev
```
Server pengembangan React akan dimulai pada `https://localhost:5173/`. Anda dapat membuka `https://localhost:5173/` di Google Chrome atau Microsoft Edge untuk menguji dalam **Mock Browser Mode** (tidak memerlukan Office desktop).

### 4. Sideload Add-in ke Office Desktop
Di jendela PowerShell baru:
```bash
npm run sideload
```
Untuk membersihkan kunci registri sideload setelah pengembangan selesai:
```bash
npm run sideload:clean
```

### 5. Jalankan Pengujian Otomatis
```bash
# Run unit and integration tests (175 tests)
npx vitest run

# Run TypeScript type check
npx tsc --noEmit
```

### 6. Build Paket Windows Tray Mandiri
```bash
npm run build:tray
```
Pipeline PowerShell otomatis ini akan:
1. Menjalankan `npm run build` untuk membuat `dist/`.
2. Menyiapkan dan mengonversi sertifikat SSL (`ca.crt`, `localhost.pfx`).
3. Mengompilasi kode sumber native `src-tray/*.cs` menggunakan `csc.exe` bawaan Windows menjadi `bin/SamTrayServer.exe`.
4. Menyiapkan file distribusi di `dist-release/`.
5. Mengemas rilis portabel ke dalam `release/SamOfficeAgent-Portable.zip` (dan mengompilasi `SamOfficeAgent-Setup.exe` jika Inno Setup `ISCC.exe` terpasang).

---

## 📁 Struktur Proyek

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

## 🧪 Cakupan Pengujian

Proyek ini mempertahankan cakupan pengujian yang ketat di seluruh komponen frontend, adapter streaming LLM, validator formula heuristik, driver Office, dan komponen server native Windows:

| Rangkaian Pengujian | File | Pengujian | Status |
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
| **Total Pengujian Otomatis**| **16 Rangkaian Pengujian** | **235 Pengujian** | **100% Passing** |

---

## 🛡️ Keamanan & Privasi

- **Zero-Admin Elevation**: Baik installer maupun `SamTrayServer.exe` tidak memerlukan hak akses Windows Administrator. Semuanya berjalan di user-space yang menargetkan `HKCU` dan `CurrentUser\Root`.
- **Penyimpanan Kunci Lokal**: API key disimpan secara eksklusif di `localStorage` peramban. Kunci tersebut tidak pernah dikirim ke server mana pun selain endpoint penyedia AI yang ditentukan.
- **Perlindungan Directory Traversal**: Server HTTPS in-process menggunakan normalisasi path kanonikal yang ketat (`Path.GetFullPath`) untuk mencegah serangan path traversal (`../`).
- **Sinkronisasi Batch**: Mutasi Office.js dieksekusi dalam batch atomik menggunakan panggilan `context.sync()` tunggal untuk menghindari kerusakan dokumen atau kondisi race condition.

---

## 📄 Lisensi

Proyek ini dilisensikan di bawah [MIT License](LICENSE).
