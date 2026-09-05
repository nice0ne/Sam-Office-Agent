@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo     Sam Office Agent - Pemasangan Portable
echo ===================================================
echo.

if /I "%~1"=="--help" (
    echo Penggunaan: %~nx0 [--dry-run]
    echo   [tanpa opsi]  Pasang Sam Office Agent ke %%LocalAppData%%\SamOfficeAgent
    echo   --dry-run     Uji coba sintaks tanpa mengubah berkas atau registri
    exit /b 0
)
if /I "%~1"=="-h" (
    echo Penggunaan: %~nx0 [--dry-run]
    exit /b 0
)
if /I "%~1"=="--dry-run" (
    echo [Dry-Run] Uji coba berhasil. Berkas dan registri tidak diubah.
    exit /b 0
)

set "TARGET_DIR=%LocalAppData%\SamOfficeAgent"
set "SOURCE_DIR=%~dp0"

:: 1. Tentukan direktori sumber berkas instalasi
if exist "%~dp0SamTrayServer.exe" (
    set "SOURCE_DIR=%~dp0"
) else if exist "%~dp0..\dist-release\SamTrayServer.exe" (
    set "SOURCE_DIR=%~dp0..\dist-release\"
) else if exist "%~dp0..\bin\SamTrayServer.exe" (
    set "SOURCE_DIR=%~dp0..\"
)

echo Direktori Sumber : %SOURCE_DIR%
echo Direktori Tujuan : %TARGET_DIR%
echo.

:: 2. Hentikan SamTrayServer jika sedang aktif
echo [1/5] Menghentikan proses SamTrayServer yang sedang berjalan...
taskkill /F /IM SamTrayServer.exe >nul 2>&1
timeout /t 1 /nobreak >nul 2>&1

:: 3. Siapkan direktori tujuan dan salin berkas
echo [2/5] Menyalin berkas ke %TARGET_DIR%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: Salin SamTrayServer.exe
if exist "%SOURCE_DIR%SamTrayServer.exe" (
    copy /Y "%SOURCE_DIR%SamTrayServer.exe" "%TARGET_DIR%\" >nul
) else if exist "%SOURCE_DIR%bin\SamTrayServer.exe" (
    copy /Y "%SOURCE_DIR%bin\SamTrayServer.exe" "%TARGET_DIR%\" >nul
)

:: Salin manifest.xml
if exist "%SOURCE_DIR%manifest.xml" (
    copy /Y "%SOURCE_DIR%manifest.xml" "%TARGET_DIR%\" >nul
)

:: Salin batch scripts
if exist "%SOURCE_DIR%Uninstall.bat" copy /Y "%SOURCE_DIR%Uninstall.bat" "%TARGET_DIR%\" >nul
if exist "%~dp0Uninstall.bat" copy /Y "%~dp0Uninstall.bat" "%TARGET_DIR%\" >nul

:: Salin subfolder dist dan certs
if exist "%SOURCE_DIR%dist" (
    if not exist "%TARGET_DIR%\dist" mkdir "%TARGET_DIR%\dist"
    xcopy "%SOURCE_DIR%dist" "%TARGET_DIR%\dist" /E /I /Y /Q >nul 2>&1
)
if exist "%SOURCE_DIR%certs" (
    if not exist "%TARGET_DIR%\certs" mkdir "%TARGET_DIR%\certs"
    xcopy "%SOURCE_DIR%certs" "%TARGET_DIR%\certs" /E /I /Y /Q >nul 2>&1
)

:: 4. Daftarkan manifest.xml ke Office WEF Developer Registry
echo [3/5] Mendaftarkan manifest ke Microsoft Office (WEF)...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23" /t REG_SZ /d "%TARGET_DIR%\manifest.xml" /f >nul
if errorlevel 1 (
    echo [PERINGATAN] Gagal mendaftarkan manifest ke Registry WEF Developer.
) else (
    echo        -^> WEF Developer Registry terdaftar: d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23
)

:: 5. Tambahkan auto-start Windows Run
echo [4/5] Mengatur mulai otomatis saat Windows menyala...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "SamOfficeAgent" /t REG_SZ /d "\"%TARGET_DIR%\SamTrayServer.exe\"" /f >nul
if errorlevel 1 (
    echo [PERINGATAN] Gagal menambahkan entri Run Registry.
) else (
    echo        -^> Auto-start diatur: HKCU\Software\Microsoft\Windows\CurrentVersion\Run\SamOfficeAgent
)

:: 6. Jalankan SamTrayServer.exe
echo [5/5] Menjalankan Sam Office Agent...
if exist "%TARGET_DIR%\SamTrayServer.exe" (
    start "" "%TARGET_DIR%\SamTrayServer.exe"
    echo.
    echo ===================================================
    echo   Pemasangan selesai! Sam Office Agent telah aktif.
    echo   Ikon aplikasi tersedia di System Tray [sudut kanan bawah].
    echo ===================================================
) else (
    echo [ERROR] Berkas %TARGET_DIR%\SamTrayServer.exe tidak ditemukan.
    exit /b 1
)

exit /b 0
