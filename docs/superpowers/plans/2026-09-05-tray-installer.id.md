# Rencana Implementasi Server Tray Mandiri & Installer Windows

[English](2026-09-05-tray-installer.md) | Bahasa Indonesia

> **Untuk pekerja agenik:** SUB-KETERAMPILAN WAJIB: Gunakan superpowers:subagent-driven-development (disarankan) atau superpowers:executing-plans untuk mengimplementasikan rencana ini langkah demi langkah.

**Tujuan:** Mengemas Sam Office Agent menjadi aplikasi Windows System Tray mandiri (`SamTrayServer.exe`) berbasis C# (dikompilasi menggunakan `csc.exe` bawaan Windows) dengan server web HTTPS mode pengguna (*user-mode*), penyediaan trust sertifikat, pemuatan lokal (*sideloading*) manifes Office, skrip Inno Setup (`setup.iss`), dan otomatisasi build (`build-tray-installer.ps1`).

**Arsitektur:** Aplikasi tray C# native menjalankan server web HTTPS *user-mode* dalam-proses (`TcpListener` + `SslStream`) yang melayani bundle statis React (`dist/`) pada `https://localhost:5173/`. Aplikasi secara otomatis mendaftarkan sertifikat CA pengembang lokal ke dalam penyimpanan sertifikat Windows `CurrentUser\Root` (menghilangkan blokir sertifikat WebView2) dan mendaftarkan `manifest.xml` ke registry Office WEF (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`). Skrip Inno Setup mengemas aplikasi ke dalam `%LocalAppData%\SamOfficeAgent` dengan fitur auto-start dan uninstalasi.

**Tech Stack:** C# (.NET Framework 4.7.2+ dikompilasi via Windows native `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`), WinForms `NotifyIcon`, `System.Net.Security.SslStream`, Inno Setup 6, PowerShell 5.1.

**Spesifikasi:** [`docs/superpowers/specs/2026-09-05-tray-installer-design.md`](file:///D:/VIBE-CODING/Sam-Office-Agent/docs/superpowers/specs/2026-09-05-tray-installer-design.md)

## Batasan Global

- Harus berjalan tanpa dependensi runtime eksternal pada mesin Windows target (tidak memerlukan Node.js atau Python pada klien).
- Harus menggunakan soket *user-mode* (`TcpListener` + `SslStream`) alih-alih `http.sys` agar elevasi hak administrator tidak pernah diperlukan saat dijalankan.
- Harus mempertahankan 100% privasi *Bring-Your-Own-Key* (BYOK) — seluruh kunci API tetap tersimpan secara ketat di dalam `localStorage` browser.
- Port server harus di-bind ke `127.0.0.1:5173` sesuai dengan `<SourceLocation DefaultValue="https://localhost:5173/"/>` di `manifest.xml`.
- Direktori instalasi target adalah `{localappdata}\SamOfficeAgent` untuk memungkinkan instalasi tanpa hak administrator.

---

### Tugas 1: Server Web HTTPS User-Mode C# (`src-tray/HttpServer.cs`)

**Berkas:**
- Buat: `src-tray/HttpServer.cs`
- Uji: `tests/tray/test-server.ps1`

**Antarmuka:**
- Mengonsumsi: Berkas statis di direktori `dist/`, sertifikat SSL dalam format `.pfx` atau `.crt`+`.key`.
- Menghasilkan: Kelas `HttpServer` dengan metode `Start(int port, string rootDir, X509Certificate2 cert)` dan `Stop()`.

**Ringkasan:**
Tugas ini membangun server web HTTPS *user-mode* di C# menggunakan `TcpListener` dan `SslStream` tanpa memerlukan hak administrator. Server melayani berkas statis dari folder `dist/` pada `127.0.0.1:5173`, mendukung SPA fallback ke `/index.html`, penentuan MIME type yang tepat, serta header CORS untuk WebView2. Pengujian otomatis memverifikasi bahwa server berhasil dikompilasi via `csc.exe` dan merespons permintaan HTTP dengan benar.

---

### Tugas 2: Helper Kepercayaan Sertifikat & Pemuatan Office Sideloading (`src-tray/OfficeIntegration.cs`)

**Berkas:**
- Buat: `src-tray/OfficeIntegration.cs`
- Uji: `tests/tray/test-integration.ps1`

**Antarmuka:**
- Mengonsumsi: API Windows Certificate Store (`X509Store`), API Windows Registry (`Microsoft.Win32.Registry`).
- Menghasilkan: Kelas `OfficeIntegration` dengan:
  - `EnsureCertificateInstalled(string caCertPath)`: Memasang CA pengembang ke `CurrentUser\Root` jika belum ada.
  - `RegisterAddIn(string manifestPath)`: Mendaftarkan manifes ke `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`.
  - `UnregisterAddIn()`: Menghapus nilai registry dari `WEF\Developer`.
  - `IsAddInRegistered(string manifestPath)`: Memeriksa status pendaftaran add-in.
  - `SetAutoStart(bool enable, string exePath)`: Mengaktifkan/menonaktifkan auto-start di `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`.

**Ringkasan:**
Tugas ini mengimplementasikan integrasi sistem operasi untuk penyediaan sertifikat SSL lokal tepercaya dan pendaftaran otomatis Office Add-in. Modul ini memastikan sertifikat CA pengembang terpasang di penyimpanan sertifikat `CurrentUser\Root` agar WebView2 tidak memblokir koneksi HTTPS. Selain itu, modul ini menangani registrasi sideloading manifes Office WEF dan konfigurasi auto-start aplikasi pada saat Windows boot.

---

### Tugas 3: Aplikasi Windows System Tray (`src-tray/Program.cs`)

**Berkas:**
- Buat: `src-tray/Program.cs`
- Modifikasi: Hapus `src-tray/TestProgram.cs`, `src-tray/TestIntegrationProgram.cs`
- Uji: `tests/tray/test-tray-app.ps1`

**Antarmuka:**
- Mengonsumsi: `HttpServer` dari Tugas 1, `OfficeIntegration` dari Tugas 2.
- Menghasilkan: Aplikasi Windows System Tray WinForms lengkap dengan menu konteks dan manajemen siklus hidup.

**Ringkasan:**
Tugas ini menyatukan server web dan integrasi Office ke dalam aplikasi tray WinForms native (`SamTrayServer.exe`) dengan ikon di taskbar notifikasi Windows. Aplikasi menggunakan named mutex untuk menegakkan *single-instance*, menjalankan server HTTPS di latar belakang, dan menampilkan menu konteks interaktif untuk membuka browser, mendaftarkan add-in ke Office, serta keluar secara aman. Seluruh siklus hidup aplikasi berjalan mulus tanpa memunculkan jendela konsol yang mengganggu pengguna.

---

### Tugas 4: Skrip Installer Inno Setup & Opsi Portable Fallback (`installer/setup.iss`)

**Berkas:**
- Buat: `installer/setup.iss`
- Buat: `installer/Install.bat`
- Buat: `installer/Uninstall.bat`

**Antarmuka:**
- Mengonsumsi: `dist/`, `manifest.xml`, `certs/`, `bin/SamTrayServer.exe`.
- Menghasilkan: Skrip installer Inno Setup untuk mengompilasi `release/SamOfficeAgent-Setup.exe` dan skrip batch portabel.

**Ringkasan:**
Tugas ini menyediakan mekanisme pengemasan dan distribusi aplikasi baik melalui installer mandiri maupun skrip batch portabel. Skrip Inno Setup mengemas seluruh aplikasi ke direktori `%LocalAppData%\SamOfficeAgent` dengan izin pengguna biasa tanpa membutuhkan elevasi hak administrator. Berkas batch `Install.bat` dan `Uninstall.bat` disediakan sebagai fallback portabel 1-klik bagi lingkungan yang tidak memiliki kompiler Inno Setup.

---

### Tugas 5: Otomatisasi Build Menyeluruh (End-to-End) (`scripts/build-tray-installer.ps1`)

**Berkas:**
- Buat: `scripts/build-tray-installer.ps1`
- Modifikasi: `package.json`

**Antarmuka:**
- Mengonsumsi: Seluruh basis kode (`src/`, `src-tray/`, `manifest.xml`, `installer/`).
- Menghasilkan: Direktori staging `dist-release/`, arsip `release/SamOfficeAgent-Portable.zip`, dan `release/SamOfficeAgent-Setup.exe` (jika ISCC tersedia).

**Ringkasan:**
Tugas ini membangun pipeline otomatisasi PowerShell untuk mengompilasi dan mengemas seluruh paket rilis dari awal hingga akhir dalam satu perintah. Alur mencakup eksekusi build produksi bundle React, persiapan sertifikat SSL pengembang, kompilasi C# dengan `csc.exe`, penataan direktori rilis, dan pembuatan paket installer atau arsip ZIP portabel. Skrip ini diintegrasikan ke `package.json` sebagai perintah `npm run build:tray` dan diverifikasi dengan suite pengujian lengkap tanpa regresi.

---

## Daftar Periksa Tinjauan Mandiri (Self-Review Checklist)

1. **Cakupan Spesifikasi**:
   - Server tray dengan C# dan `csc.exe` -> Tugas 1, 3
   - HTTPS user-mode pada port 5173 -> Tugas 1
   - Kepercayaan sertifikat CA & sideloading Office -> Tugas 2
   - Skrip Inno Setup & fallback portabel -> Tugas 4
   - Alur otomatisasi build -> Tugas 5
2. **Pemeriksaan Placeholder**: Tidak ada "TODO", "TBD", atau placeholder yang samar.
3. **Konsistensi Tipe**: Nama metode dan kunci registry tepat serta konsisten di seluruh bagian.
