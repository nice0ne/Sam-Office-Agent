$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"

if (!(Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "TestIntegration.exe"

$csFiles = Get-ChildItem -Path $srcTray -Filter "*.cs" | Select-Object -ExpandProperty FullName
Write-Host "Compiling TestIntegration.exe..."
& $csc /nologo /target:exe /out:$outExe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll $csFiles
if ($LASTEXITCODE -ne 0) {
    Write-Error "Compilation of OfficeIntegration failed."
    exit 1
}

Write-Host "OfficeIntegration compiled successfully."

Write-Host "Executing TestIntegration test suite..."
& $outExe --run-integration-tests
if ($LASTEXITCODE -ne 0) {
    Write-Error "TestIntegration execution failed with exit code $LASTEXITCODE."
    exit 1
}

Write-Host "All OfficeIntegration tests passed successfully!"
