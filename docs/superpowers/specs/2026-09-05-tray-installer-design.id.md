# Spesifikasi Standalone Windows Tray Server & Installer untuk Sam Office Agent

- **Tanggal:** 2026-09-05
- **Status:** Diusulkan (Proposed)
- **Target OS:** Windows 10 / Windows 11
- **Komponen:** Desktop Packaging, Local Tray HTTPS Server & Inno Setup Installer

[English](2026-09-05-tray-installer-design.md) | Bahasa Indonesia

---

## 1. Ringkasan Eksekutif

Saat ini, Sam Office Agent berjalan selama pengembangan melalui Node.js (`npm run dev`) dengan sertifikat yang dipasang via `office-addin-dev-certs` dan di-sideload melalui skrip PowerShell.

Untuk distribusi ke pengguna akhir (klien, rekan kerja, atau lingkungan enterprise tanpa Node.js), spesifikasi ini merancang distribusi desktop Windows yang mandiri (self-contained), ringan, dan tanpa dependensi eksternal (zero-dependency):
1. **`SamTrayServer.exe`**: Aplikasi Windows System Tray native yang ditulis dalam C# dan dikompilasi menggunakan compiler native Windows (`csc.exe`), menghilangkan dependensi SDK eksternal. Aplikasi ini menjalankan in-process user-mode HTTPS web server (`TcpListener` + `SslStream`) yang melayani build statis React (`dist/`) pada `https://localhost:5173/`.
2. **Automated Trust Provisioning**: Mengimpor sertifikat Developer CA lokal secara otomatis ke penyimpanan Windows `CurrentUser\Root` tanpa prompt command prompt, mencegah pemblokiran sertifikat keamanan WebView2 di Excel, Word, dan PowerPoint.
3. **Automated Office Sideloading**: Mendaftarkan `manifest.xml` ke dalam Windows Registry (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`) saat startup, membuat add-in langsung terlihat pada ribbon Office.
4. **`SamOfficeAgent-Setup.exe`**: Skrip installer Inno Setup yang memaketkan server, aset, dan manifes ke dalam `%LocalAppData%\SamOfficeAgent` dengan pintasan startup Windows dan uninstalasi yang bersih.

---

## 2. Arsitektur & Diagram Komponen

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

## 3. Spesifikasi Detail Komponen

### 3.1 Tray Server & Web Server (`src-tray/Program.cs`)

#### A. Kompilasi & Runtime
- **Bahasa**: C# (.NET Framework 4.7.2+).
- **Compiler**: Built-in Windows `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.
- **Dependensi**: Tidak ada. Menggunakan rakitan (assembly) standar (`System.dll`, `System.Windows.Forms.dll`, `System.Drawing.dll`, `System.Security.dll`).
- **Jejak Memori (Memory Footprint)**: ~10–18 MB RAM (dibandingkan dengan ~80–120 MB untuk proses Node.js).

#### B. User-Mode HTTPS Web Server (`TcpListener` + `SslStream`)
- **Port**: `5173` (sesuai dengan lokasi sumber `manifest.xml` `https://localhost:5173/`).
- **Mengapa `TcpListener` + `SslStream` alih-alih `HttpListener`**:
  - `System.Net.HttpListener` bergantung pada Windows kernel `http.sys` yang memerlukan hak akses administrator untuk mengeksekusi `netsh http add sslcert`.
  - `TcpListener` + `SslStream` berjalan sepenuhnya di user-space tanpa elevasi administrator atau prompt UAC.
- **Penanganan Tipe MIME**:
  - `.html` -> `text/html; charset=utf-8`
  - `.js` -> `application/javascript; charset=utf-8`
  - `.css` -> `text/css; charset=utf-8`
  - `.png` -> `image/png`
  - `.json` -> `application/json; charset=utf-8`
  - `.svg` -> `image/svg+xml`
- **SPA Fallback**: Rute non-file apa pun (misalnya `/settings`, `/chat`) menyajikan `index.html`.
- **CORS Headers**: Mengirimkan `Access-Control-Allow-Origin: *` untuk mencegah masalah cross-origin di WebView2.

#### C. UI & Menu System Tray
- Menggunakan `System.Windows.Forms.NotifyIcon`.
- Ikon tray: Ikon tersemat yang dibundel atau dimuat dari `icon-32.png`.
- Menu Konteks:
  - 🟢 **Status: Berjalan (https://localhost:5173)** [Item dinonaktifkan, menampilkan status]
  - 🌐 **Buka di Browser** [Membuka browser default ke `https://localhost:5173/`]
  - 📋 **Daftarkan ke Office** [Mendaftarkan manifes ke registry WEF]
  - 🗑️ **Copot dari Office** [Menghapus manifes dari registry WEF]
  - ⚙️ **Jalankan saat Windows Menyala** [Dapat dicentang, beralih key registry Run]
  - 🔄 **Muat Ulang Server (Restart)**
  - ❌ **Keluar (Exit)**
- Notifikasi balloon tip saat mulai: *"Sam Office Agent aktif di latar belakang."*

---

### 3.2 Rutinitas Penyediaan Sertifikat (Certificate Provisioning Routine)

1. **Pembuatan & Pemuatan Sertifikat:**
   - Aplikasi memeriksa keberadaan `certs/localhost.pfx` (atau `localhost.crt` + `localhost.key`).
   - Jika tidak ada, aplikasi secara otomatis membuat PFX self-signed yang aman menggunakan .NET `System.Security.Cryptography.X509Certificates` atau menyalinnya dari `~/.office-addin-dev-certs`.
2. **Impor Root CA:**
   - Memeriksa `X509Store(StoreName.Root, StoreLocation.CurrentUser)` untuk `Developer CA for Microsoft Office Add-ins`.
   - Jika tidak ditemukan, mengimpor `ca.crt` menggunakan `store.Add(cert)`.
   - Setelah ada, Windows, Edge, dan Office WebView2 memvalidasi sertifikat sebagai tepercaya, sepenuhnya menghilangkan error *"The content is blocked because it isn't signed by a valid security certificate"*.

---

### 3.3 Manajemen Registry & Sideloading Otomatis

1. **Pendaftaran Office WEF:**
   - Path registry: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
   - Nama key: `d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23` (Manifest ID dari `manifest.xml`)
   - Nilai (Value): Path lengkap ke `manifest.xml` yang terpasang (misalnya `%LocalAppData%\SamOfficeAgent\manifest.xml`)
2. **Pendaftaran Auto-Start:**
   - Path registry: `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
   - Nama key: `SamOfficeAgent`
   - Nilai (Value): `"%LocalAppData%\SamOfficeAgent\SamTrayServer.exe"`

---

### 3.4 Spesifikasi Installer (`installer/setup.iss`)

- **Alat**: Inno Setup 6 (compiler installer Windows open-source).
- **Nama Aplikasi**: Sam Office Agent
- **Folder Target**: `{localappdata}\SamOfficeAgent` (tidak memerlukan hak akses administrator).
- **Fitur Installer**:
  - Gaya wizard modern (flat, peka terhadap DPI).
  - Menghentikan (kill) `SamTrayServer.exe` yang sedang berjalan secara otomatis sebelum pembaruan/instalasi.
  - Menyalin `dist/`, `manifest.xml`, `certs/`, dan `SamTrayServer.exe`.
  - Menambahkan uninstaller ke *Installed Apps* (*Add/Remove Programs*) Windows.
  - Kotak centang: *"Jalankan Sam Office Agent secara otomatis saat Windows mulai"*.
  - Aksi pasca-instalasi: Menjalankan `SamTrayServer.exe`.
- **Pembersihan Uninstaller**:
  - Menutup tray server.
  - Menghapus key `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
  - Menghapus key `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`.
  - Menghapus direktori instalasi.

---

### 3.5 Skrip Otomasi Build (`scripts/build-tray-installer.ps1`)

Mengotomatiskan pemaketan end-to-end ke dalam satu perintah:
1. Menjalankan `npm run build` untuk menghasilkan `dist/` yang bersih.
2. Mengompilasi `src-tray/Program.cs` dengan `csc.exe` menjadi `dist-tray/SamTrayServer.exe`.
3. Menyalin `manifest.xml` dan sertifikat SSL ke dalam `dist-tray/`.
4. Jika Inno Setup (`ISCC.exe`) terpasang: Mengompilasi `installer/setup.iss` menjadi `release/SamOfficeAgent-Setup.exe`.
5. Jika Inno Setup tidak terpasang: Memaketkan `dist-tray/` ke dalam `release/SamOfficeAgent-Portable.zip` dengan `Install.bat` 1-klik.

---

## 4. Penanganan Kesalahan & Edge Cases

| Skenario | Perilaku yang Ditangani |
| :--- | :--- |
| **Port 5173 sudah digunakan** | Mendeteksi konflik; menampilkan balloon tooltip: *"Port 5173 sedang dipakai oleh aplikasi lain"* dan menawarkan opsi muat ulang (restart). |
| **Pengguna membuka Excel sebelum Tray Server** | Add-in menampilkan error pada taskpane sampai Tray Server dijalankan; menu memiliki opsi "Muat Ulang Server" untuk pemulihan cepat. |
| **Sertifikat SSL Tidak Tepercaya** | Pemulihan mandiri (self-healing): Saat dijalankan, Tray Server memeriksa penyimpanan sertifikat dan menerapkan ulang `ca.crt` ke `CurrentUser\Root`. |
| **Beberapa Instans Sekaligus (Multiple Instances)** | Menggunakan `Mutex` bernama (`SamOfficeAgentTrayServerMutex`) sehingga mengklik file exe saat sudah berjalan akan memunculkan balloon tip alih-alih gagal. |
| **Uninstalasi Bersih** | Uninstaller menghapus entri registry pengembang WEF sehingga Office tidak menyisakan ikon add-in yang tertinggal (orphan). |

---

## 5. Rencana Verifikasi

1. **Kompilasi C#**: Verifikasi `csc.exe` mengompilasi `src-tray/Program.cs` tanpa peringatan.
2. **Pengujian Web Server HTTPS Lokal**: Jalankan `SamTrayServer.exe` dan uji via `curl -k https://localhost:5173/` dan di Chrome/Edge.
3. **Pengujian Sideloading Office**: Verifikasi `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` terisi dan Excel Desktop terbuka dengan Sam di ribbon.
4. **Pengujian Menu System Tray**: Verifikasi tombol Buka di Browser, Status, Daftarkan/Copot, dan Keluar.
5. **Pengujian Installer & Uninstaller**: Jalankan installer, verifikasi file di `%LocalAppData%\SamOfficeAgent`, verifikasi add-in di Excel, uji uninstalasi bersih.
