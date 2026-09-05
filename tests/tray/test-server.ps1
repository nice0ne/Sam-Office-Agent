$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$srcTray = Join-Path $projectRoot "src-tray"
$binDir = Join-Path $projectRoot "bin"
$distDir = Join-Path $projectRoot "dist"

if (!(Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $binDir "TestServer.exe"

Write-Host "Compiling TestServer.exe..."
& $csc /nologo /target:exe /out:$outExe "$srcTray\HttpServer.cs" "$srcTray\TestProgram.cs"

if ($LASTEXITCODE -ne 0) {
    Write-Error "Kompilasi C# gagal."
    exit 1
}

Write-Host "TestServer compiled successfully."

# Prepare test certificate for HTTPS verification
$pfxFile = Join-Path $binDir "test-server.pfx"
$pfxPass = "testpass123"

if (!(Test-Path $pfxFile)) {
    Write-Host "Generating self-signed test certificate..."
    $cert = New-SelfSignedCertificate -DnsName "127.0.0.1", "localhost" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(1)
    $securePass = ConvertTo-SecureString -String $pfxPass -Force -AsPlainText
    Export-PfxCertificate -Cert $cert -FilePath $pfxFile -Password $securePass | Out-Null
    Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force -ErrorAction SilentlyContinue
    Write-Host "Test certificate generated at $pfxFile"
}

# Run automated HTTPS server tests
Write-Host "Executing TestServer test suite..."
& $outExe --run-tests "$pfxFile" "$pfxPass" "$distDir" 5173

if ($LASTEXITCODE -ne 0) {
    Write-Error "TestServer verification failed with exit code $LASTEXITCODE."
    exit 1
}

Write-Host "All HTTPS server tests passed successfully!"
exit 0
