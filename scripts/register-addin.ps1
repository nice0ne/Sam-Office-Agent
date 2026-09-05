<#
.SYNOPSIS
    Otomatisasi pendaftaran Sam Office Agent ke Microsoft Office (Excel, Word, PowerPoint) di Windows.
.DESCRIPTION
    Script ini mendaftarkan manifest.xml ke Registry Windows (Office Web Extension Framework Developer Key),
    sehingga add-in langsung terdeteksi di Office Desktop tanpa perlu pengaturan manual di Trust Center.
#>

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$manifestPath = Join-Path $projectRoot "manifest.xml"

if (!(Test-Path $manifestPath)) {
    Write-Error "File manifest.xml tidak ditemukan di: $manifestPath"
    exit 1
}

$manifestId = "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23"
$regKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

Write-Host "=== Pendaftaran Otomatis Sam Office Agent ===" -ForegroundColor Cyan

# 1. Pastikan folder registry WEF\Developer ada
if (!(Test-Path $regKeyPath)) {
    Write-Host "[1/2] Membuat Registry Key WEF\Developer..." -ForegroundColor Yellow
    New-Item -Path $regKeyPath -Force | Out-Null
} else {
    Write-Host "[1/2] Registry Key WEF\Developer sudah siap." -ForegroundColor Green
}

# 2. Daftarkan path manifest.xml
Write-Host "[2/2] Mendaftarkan manifest.xml ke Registry..." -ForegroundColor Yellow
Set-ItemProperty -Path $regKeyPath -Name $manifestId -Value $manifestPath -Force

Write-Host "`nBERHASIL! Sam Office Agent telah terdaftar secara otomatis." -ForegroundColor Green
Write-Host "Lokasi Manifes: $manifestPath" -ForegroundColor Gray
Write-Host "`nLangkah selanjutnya:" -ForegroundColor Cyan
Write-Host "1. Pastikan server dev aktif: npm run dev"
Write-Host "2. Buka aplikasi Excel, Word, atau PowerPoint Desktop."
Write-Host "3. Add-in akan langsung muncul di pita Ribbon / tab Home."
