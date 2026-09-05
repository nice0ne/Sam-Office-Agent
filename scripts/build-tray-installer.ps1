<#
.SYNOPSIS
    Build and packaging automation for Sam Office Agent Windows System Tray server and installer.
.DESCRIPTION
    Automates the entire end-to-end build pipeline:
    1. Static frontend compilation via 'npm run build'.
    2. SSL developer certificate preparation (certs/ca.crt and certs/localhost.pfx with password 'testpass123').
    3. C# Windows System Tray server compilation into bin/SamTrayServer.exe using csc.exe.
    4. Staging of distribution files into dist-release/.
    5. Packaging release artifacts into release/:
       - If Inno Setup ISCC.exe is detected, compiles installer/setup.iss into release/SamOfficeAgent-Setup.exe.
       - Compresses dist-release/ into release/SamOfficeAgent-Portable.zip.
    6. Formatted summary reporting and usage instructions.
#>

[CmdletBinding()]
param(
    [switch]$SkipFrontend,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    Sam Office Agent - Windows Build & Packaging Pipeline   " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Project Root: $projectRoot`n"

# 0. Clean workspace if requested
if ($Clean) {
    Write-Host "[Clean] Cleaning previous build artifacts..." -ForegroundColor Yellow
    foreach ($dir in @("dist-release", "release", "bin", "certs")) {
        $targetDir = Join-Path $projectRoot $dir
        if (Test-Path $targetDir) {
            Remove-Item -Path $targetDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  -> Removed $dir/"
        }
    }
}

# 1. Build static frontend via npm run build
if (-not $SkipFrontend) {
    Write-Host "[1/5] Building static frontend via npm run build..." -ForegroundColor Cyan
    Push-Location $projectRoot
    try {
        & cmd.exe /c "npm run build"
        if ($LASTEXITCODE -ne 0) {
            throw "Frontend build (npm run build) failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    $distIndex = Join-Path $projectRoot "dist\index.html"
    if (!(Test-Path $distIndex)) {
        throw "Frontend build artifact missing: $distIndex"
    }
    Write-Host "  -> Frontend build completed successfully (dist/ ready)." -ForegroundColor Green
} else {
    Write-Host "[1/5] Skipping frontend build (-SkipFrontend specified)." -ForegroundColor Yellow
}

# 2. Ensure SSL certificate files exist
Write-Host "`n[2/5] Ensuring SSL certificate files (certs/ca.crt & certs/localhost.pfx)..." -ForegroundColor Cyan
$certsDir = Join-Path $projectRoot "certs"
if (!(Test-Path $certsDir)) {
    New-Item -ItemType Directory -Path $certsDir -Force | Out-Null
}

$caCrtPath = Join-Path $certsDir "ca.crt"
$localhostPfxPath = Join-Path $certsDir "localhost.pfx"
$pfxPassword = "testpass123"

# Look for dev certificates in ~/.office-addin-dev-certs
$devCertsDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$devCa = Join-Path $devCertsDir "ca.crt"
$devCrt = Join-Path $devCertsDir "localhost.crt"
$devKey = Join-Path $devCertsDir "localhost.key"

# Copy ca.crt from ~/.office-addin-dev-certs if present
if ((Test-Path $devCa) -and (!(Test-Path $caCrtPath))) {
    Copy-Item -Path $devCa -Destination $caCrtPath -Force
    Write-Host "  -> Copied CA certificate from ~/.office-addin-dev-certs/ca.crt" -ForegroundColor Green
}

# Locate openssl if available on system or Git
$opensslPath = $null
$potentialOpenssl = @(
    "openssl",
    "C:\Program Files\Git\usr\bin\openssl.exe",
    "C:\Program Files (x86)\Git\usr\bin\openssl.exe"
)
foreach ($o in $potentialOpenssl) {
    if (Get-Command $o -ErrorAction SilentlyContinue) {
        $opensslPath = (Get-Command $o).Source
        break
    } elseif (Test-Path $o) {
        $opensslPath = $o
        break
    }
}

# Create or export localhost.pfx
if (!(Test-Path $localhostPfxPath)) {
    $pfxCreated = $false

    # Strategy A: Use openssl with office-addin-dev-certs if available
    if ($opensslPath -and (Test-Path $devCrt) -and (Test-Path $devKey)) {
        Write-Host "  -> Converting localhost.crt and localhost.key to PKCS#12 (.pfx) with OpenSSL..."
        $certfileArg = ""
        if (Test-Path $devCa) {
            $certfileArg = "-certfile `"$devCa`""
        }
        $cmdLine = "`"$opensslPath`" pkcs12 -export -in `"$devCrt`" -inkey `"$devKey`" $certfileArg -out `"$localhostPfxPath`" -passout pass:$pfxPassword"
        & cmd.exe /c $cmdLine
        if (($LASTEXITCODE -eq 0) -and (Test-Path $localhostPfxPath)) {
            $pfxCreated = $true
            Write-Host "  -> localhost.pfx generated via OpenSSL." -ForegroundColor Green
        }
    }

    # Strategy B: Fallback to PowerShell New-SelfSignedCertificate
    if (!$pfxCreated) {
        Write-Host "  -> Generating self-signed certificate for localhost using PowerShell..."
        $cert = New-SelfSignedCertificate -DnsName "127.0.0.1", "localhost" -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(1)
        $securePass = ConvertTo-SecureString -String $pfxPassword -Force -AsPlainText
        Export-PfxCertificate -Cert $cert -FilePath $localhostPfxPath -Password $securePass | Out-Null

        if (!(Test-Path $caCrtPath)) {
            Export-Certificate -Cert $cert -FilePath $caCrtPath | Out-Null
            Write-Host "  -> Exported self-signed public CA certificate to certs/ca.crt." -ForegroundColor Green
        }

        Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force -ErrorAction SilentlyContinue
        $pfxCreated = (Test-Path $localhostPfxPath)
    }

    if (!$pfxCreated) {
        throw "Failed to create or export localhost.pfx certificate."
    }
} else {
    Write-Host "  -> localhost.pfx already present in certs/." -ForegroundColor Green
}

# Ensure ca.crt exists in certs/
if (!(Test-Path $caCrtPath)) {
    if (Test-Path $devCa) {
        Copy-Item -Path $devCa -Destination $caCrtPath -Force
        Write-Host "  -> ca.crt copied from ~/.office-addin-dev-certs." -ForegroundColor Green
    } else {
        $pfxCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($localhostPfxPath, $pfxPassword)
        [System.IO.File]::WriteAllBytes($caCrtPath, $pfxCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        Write-Host "  -> Generated certs/ca.crt from localhost.pfx public key." -ForegroundColor Green
    }
}
Write-Host "  -> SSL certificate files verified: certs/ca.crt & certs/localhost.pfx" -ForegroundColor Green

# 3. Compile src-tray/*.cs into bin/SamTrayServer.exe
Write-Host "`n[3/5] Compiling C# Windows System Tray server (bin/SamTrayServer.exe)..." -ForegroundColor Cyan
$binDir = Join-Path $projectRoot "bin"
if (!(Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}
$outExe = Join-Path $binDir "SamTrayServer.exe"

$cscCandidates = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$cscPath = $null
foreach ($c in $cscCandidates) {
    if (Test-Path $c) {
        $cscPath = $c
        break
    }
}
if (!$cscPath) {
    throw "C# compiler csc.exe was not found in standard Microsoft.NET Framework paths."
}

$srcTrayDir = Join-Path $projectRoot "src-tray"
$csFiles = Get-ChildItem -Path $srcTrayDir -Filter "*.cs" | Select-Object -ExpandProperty FullName
if ($csFiles.Count -eq 0) {
    throw "No C# source files (*.cs) found in $srcTrayDir."
}

Write-Host "  -> Compiler: $cscPath"
Write-Host "  -> Compiling $($csFiles.Count) C# source files..."
& $cscPath /nologo /target:winexe /out:"$outExe" /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll $csFiles

if (($LASTEXITCODE -ne 0) -or (!(Test-Path $outExe))) {
    throw "Compilation of SamTrayServer.exe failed with exit code $LASTEXITCODE."
}
Write-Host "  -> SamTrayServer.exe compiled successfully: $outExe" -ForegroundColor Green

# 4. Stage distribution directory dist-release/
Write-Host "`n[4/5] Staging distribution files into dist-release/..." -ForegroundColor Cyan
$distReleaseDir = Join-Path $projectRoot "dist-release"
if (Test-Path $distReleaseDir) {
    Remove-Item -Path $distReleaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $distReleaseDir -Force | Out-Null

# Copy SamTrayServer.exe
Copy-Item -Path $outExe -Destination (Join-Path $distReleaseDir "SamTrayServer.exe") -Force

# Copy manifest.xml
$manifestPath = Join-Path $projectRoot "manifest.xml"
if (!(Test-Path $manifestPath)) {
    throw "manifest.xml not found at $manifestPath."
}
Copy-Item -Path $manifestPath -Destination (Join-Path $distReleaseDir "manifest.xml") -Force

# Copy dist/
$distDir = Join-Path $projectRoot "dist"
if (!(Test-Path $distDir)) {
    throw "Frontend dist directory not found at $distDir."
}
Copy-Item -Path $distDir -Destination (Join-Path $distReleaseDir "dist") -Recurse -Force

# Copy certs/
Copy-Item -Path $certsDir -Destination (Join-Path $distReleaseDir "certs") -Recurse -Force

# Copy installer scripts
$installerDir = Join-Path $projectRoot "installer"
$installBat = Join-Path $installerDir "Install.bat"
$uninstallBat = Join-Path $installerDir "Uninstall.bat"

if (!(Test-Path $installBat)) { throw "installer/Install.bat not found." }
if (!(Test-Path $uninstallBat)) { throw "installer/Uninstall.bat not found." }

Copy-Item -Path $installBat -Destination (Join-Path $distReleaseDir "Install.bat") -Force
Copy-Item -Path $uninstallBat -Destination (Join-Path $distReleaseDir "Uninstall.bat") -Force

# Also mirror into dist-release/installer/
$stagedInstallerDir = Join-Path $distReleaseDir "installer"
New-Item -ItemType Directory -Path $stagedInstallerDir -Force | Out-Null
Copy-Item -Path $installBat -Destination (Join-Path $stagedInstallerDir "Install.bat") -Force
Copy-Item -Path $uninstallBat -Destination (Join-Path $stagedInstallerDir "Uninstall.bat") -Force

Write-Host "  -> Distribution files staged successfully in $distReleaseDir" -ForegroundColor Green

# 5. Package release artifacts in release/
Write-Host "`n[5/5] Packaging release artifacts into release/..." -ForegroundColor Cyan
$releaseDir = Join-Path $projectRoot "release"
if (!(Test-Path $releaseDir)) {
    New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
}

# 5a. Inno Setup Compiler check
$isccPath = $null
$isccCandidates = @(
    "iscc",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
    "C:\Program Files\Inno Setup 5\ISCC.exe",
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
)
foreach ($p in $isccCandidates) {
    if (Get-Command $p -ErrorAction SilentlyContinue) {
        $isccPath = (Get-Command $p).Source
        break
    } elseif (Test-Path $p) {
        $isccPath = $p
        break
    }
}

$setupExePath = Join-Path $releaseDir "SamOfficeAgent-Setup.exe"
if ($isccPath) {
    Write-Host "  -> Inno Setup Compiler detected: $isccPath" -ForegroundColor Cyan
    $issScript = Join-Path $installerDir "setup.iss"
    Write-Host "  -> Compiling Inno Setup installer..."
    & $isccPath "$issScript"
    if (($LASTEXITCODE -eq 0) -and (Test-Path $setupExePath)) {
        $setupSize = (Get-Item $setupExePath).Length
        Write-Host "  -> Inno Setup installer generated: $setupExePath ($([Math]::Round($setupSize / 1MB, 2)) MB)" -ForegroundColor Green
    } else {
        Write-Warning "Inno Setup compilation finished with code $LASTEXITCODE or output file was not created."
    }
} else {
    Write-Host "  -> Inno Setup compiler (ISCC.exe) not found. Skipping Inno Setup compilation." -ForegroundColor Yellow
}

# 5b. Portable Zip archive
$portableZipPath = Join-Path $releaseDir "SamOfficeAgent-Portable.zip"
if (Test-Path $portableZipPath) {
    Remove-Item -Path $portableZipPath -Force
}
Write-Host "  -> Creating portable zip package: $portableZipPath..."
Compress-Archive -Path (Join-Path $distReleaseDir "*") -DestinationPath $portableZipPath -Force

if (!(Test-Path $portableZipPath)) {
    throw "Failed to create portable zip package at $portableZipPath."
}
$zipSize = (Get-Item $portableZipPath).Length
Write-Host "  -> Portable ZIP package created successfully: $([Math]::Round($zipSize / 1MB, 2)) MB" -ForegroundColor Green

# 6. Summary Report
Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "      Sam Office Agent - Build & Packaging Complete         " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Output Directory  : $releaseDir"
if (Test-Path $setupExePath) {
    Write-Host "Setup Installer   : $setupExePath" -ForegroundColor Cyan
} else {
    Write-Host "Setup Installer   : Skipped (Inno Setup ISCC not installed)" -ForegroundColor Yellow
}
Write-Host "Portable Package  : $portableZipPath" -ForegroundColor Cyan
Write-Host "Staged Directory  : $distReleaseDir"
Write-Host "`nUsage Instructions:" -ForegroundColor Yellow
Write-Host "  1. Portable mode: Extract $portableZipPath and double-click Install.bat."
if (Test-Path $setupExePath) {
    Write-Host "  2. Installer mode: Run $setupExePath to install to %LocalAppData%\SamOfficeAgent."
}
Write-Host "  3. Running server: SamTrayServer.exe starts HTTPS on https://localhost:5173"
Write-Host "                     and displays a system tray icon with status and controls."
Write-Host "============================================================`n" -ForegroundColor Green
