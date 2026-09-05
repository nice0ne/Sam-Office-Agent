@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo     Sam Office Agent - Pencopotan Pemasangan
echo ===================================================
echo.

if /I "%~1"=="--help" (
    echo Penggunaan: %~nx0 [--dry-run]
    echo   [tanpa opsi]  Copot Sam Office Agent dan bersihkan berkas/registri
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

:: 1. Hentikan SamTrayServer jika sedang aktif
echo [1/4] Menghentikan proses SamTrayServer...
taskkill /F /IM SamTrayServer.exe >nul 2>&1
timeout /t 1 /nobreak >nul 2>&1

:: 2. Hapus pendaftaran Office WEF Developer
echo [2/4] Menghapus pendaftaran dari Microsoft Office (WEF)...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23" /f >nul 2>&1
echo        -^> Pendaftaran WEF Developer telah dihapus.

:: 3. Hapus entri Windows Run auto-start
echo [3/4] Menghapus auto-start Windows...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "SamOfficeAgent" /f >nul 2>&1
echo        -^> Auto-start Windows telah dihapus.

:: 4. Hapus folder aplikasi %LocalAppData%\SamOfficeAgent
echo [4/4] Menghapus direktori berkas %TARGET_DIR%...
if exist "%TARGET_DIR%" (
    :: Jika uninstaller dijalankan dari dalam TARGET_DIR, jadwalkan penghapusan folder setelah exit
    if /I "%~dp0"=="%TARGET_DIR%\" (
        start /b "" cmd /c "timeout /t 1 /nobreak >nul & rmdir /s /q \"%TARGET_DIR%\""
    ) else (
        rmdir /s /q "%TARGET_DIR%" >nul 2>&1
    )
    echo        -^> Direktori berkas telah dihapus.
) else (
    echo        -^> Direktori %TARGET_DIR% sudah tidak ada.
)

echo.
echo ===================================================
echo   Pencopotan selesai! Sam Office Agent telah dibersihkan.
echo ===================================================

exit /b 0
