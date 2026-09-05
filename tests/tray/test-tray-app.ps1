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
$statusOutput = (& $outExe --status | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
    Write-Error "Execution of SamTrayServer.exe --status failed with code $LASTEXITCODE."
    exit 1
}
Write-Host "Status output: $statusOutput"
if ($statusOutput -notmatch "SamTrayServer status:") {
    Write-Error "Status output mismatch: $statusOutput"
    exit 1
}

$testOutput = (& $outExe --test | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
    Write-Error "Execution of SamTrayServer.exe --test failed with code $LASTEXITCODE."
    exit 1
}
Write-Host "Test flag output: $testOutput"
if ($testOutput -notmatch "SamTrayServer self-test: OK") {
    Write-Error "Test output mismatch: $testOutput"
    exit 1
}

# Verify mutex detection via --status
Write-Host "Verifying mutex state detection via --status..."
$createdNew = $false
$heldMutex = New-Object System.Threading.Mutex($true, "Global\SamOfficeAgentTrayServerMutex", [ref]$createdNew)
try {
    if ($createdNew) {
        $statusWhenRunning = (& $outExe --status | Out-String).Trim()
        Write-Host "Status while mutex held: $statusWhenRunning"
        if ($statusWhenRunning -notmatch "SamTrayServer status: RUNNING") {
            Write-Error "Expected status RUNNING when mutex is held, got: $statusWhenRunning"
            exit 1
        }
    }
}
finally {
    if ($createdNew -and $heldMutex -ne $null) {
        $heldMutex.ReleaseMutex()
        $heldMutex.Close()
    }
}

$statusAfterRelease = (& $outExe --status | Out-String).Trim()
Write-Host "Status after mutex released: $statusAfterRelease"
if ($statusAfterRelease -notmatch "SamTrayServer status: STOPPED") {
    Write-Error "Expected status STOPPED after mutex release, got: $statusAfterRelease"
    exit 1
}

Write-Host "All tray app tests passed successfully!"
exit 0
