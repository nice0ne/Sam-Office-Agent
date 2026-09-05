$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"

if (!(Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "SamTrayServer.exe"

$csFiles = Get-ChildItem -Path $srcTray -Filter "*.cs" | Select-Object -ExpandProperty FullName

Write-Host "Compiling SamTrayServer.exe..."
& $csc /nologo /target:winexe /out:$outExe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll $csFiles

if ($LASTEXITCODE -ne 0) {
    Write-Error "Kompilasi SamTrayServer.exe gagal."
    exit 1
}

Write-Host "SamTrayServer.exe compiled successfully: $outExe"

# Verify non-blocking CLI flag execution
Write-Host "Testing non-blocking CLI flags..."
$statusOutput = & $outExe --status
if ($LASTEXITCODE -ne 0) {
    Write-Error "Execution of SamTrayServer.exe --status failed with code $LASTEXITCODE."
    exit 1
}
Write-Host "Status output: $statusOutput"

$testOutput = & $outExe --test
if ($LASTEXITCODE -ne 0) {
    Write-Error "Execution of SamTrayServer.exe --test failed with code $LASTEXITCODE."
    exit 1
}
Write-Host "Test flag output: $testOutput"

Write-Host "All tray app tests passed successfully!"
exit 0
