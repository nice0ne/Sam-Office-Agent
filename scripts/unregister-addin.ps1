<#
.SYNOPSIS
    Mencopot pendaftaran Sam Office Agent dari Microsoft Office di Windows.
#>

$ErrorActionPreference = "SilentlyContinue"

$manifestId = "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23"
$regKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

Write-Host "=== Menghapus Pendaftaran Sam Office Agent ===" -ForegroundColor Cyan

if (Test-Path $regKeyPath) {
    Remove-ItemProperty -Path $regKeyPath -Name $manifestId -ErrorAction SilentlyContinue
    Write-Host "Pendaftaran Add-in berhasil dihapus dari Registry Windows." -ForegroundColor Green
} else {
    Write-Host "Registry key tidak ditemukan. Tidak ada yang perlu dihapus." -ForegroundColor Yellow
}
