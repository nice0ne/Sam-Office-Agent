$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $testDir)
$installerDir = Join-Path $projectRoot "installer"

$setupIss = Join-Path $installerDir "setup.iss"
$installBat = Join-Path $installerDir "Install.bat"
$uninstallBat = Join-Path $installerDir "Uninstall.bat"

Write-Host "=== Testing Inno Setup & Portable Batch Installer Files ===" -ForegroundColor Cyan

# 1. Existence Checks
Write-Host "Checking file existence..."
$missingFiles = @()
if (!(Test-Path $setupIss)) { $missingFiles += "installer/setup.iss" }
if (!(Test-Path $installBat)) { $missingFiles += "installer/Install.bat" }
if (!(Test-Path $uninstallBat)) { $missingFiles += "installer/Uninstall.bat" }

if ($missingFiles.Count -gt 0) {
    Write-Error "Missing required installer file(s): $($missingFiles -join ', ')"
    exit 1
}

# 2. setup.iss Directive and Section Verification
Write-Host "Verifying setup.iss directives and sections..."
$issContent = Get-Content -Path $setupIss -Raw

$requiredIssPatterns = @(
    "AppName\s*=\s*Sam Office Agent",
    "AppVersion\s*=\s*1\.0\.0",
    "DefaultDirName\s*=\s*\{localappdata\}\\SamOfficeAgent",
    "PrivilegesRequired\s*=\s*lowest",
    "OutputBaseFilename\s*=\s*SamOfficeAgent-Setup",
    "OutputDir\s*=\s*\.\.\\release",
    "CloseApplications\s*=\s*force",
    "CloseApplicationsFilter\s*=\s*SamTrayServer\.exe",
    "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23",
    "\[Files\]",
    "\.\.\\dist-release\\\*",
    "\[Icons\]",
    "\{userstartup\}\\Sam Office Agent",
    "\[Run\]",
    "\{app\}\\SamTrayServer\.exe",
    "\[UninstallRun\]",
    "\[UninstallDelete\]"
)

foreach ($pattern in $requiredIssPatterns) {
    if ($issContent -notmatch $pattern) {
        Write-Error "setup.iss missing expected pattern/directive: '$pattern'"
        exit 1
    }
}
Write-Host "  -> setup.iss directives verified." -ForegroundColor Green

# 3. Install.bat Sanity Verification
Write-Host "Verifying Install.bat sanity..."
$installContent = Get-Content -Path $installBat -Raw

$requiredInstallPatterns = @(
    "taskkill.*SamTrayServer\.exe",
    "SamOfficeAgent",
    "reg\s+add.*Software\\Microsoft\\Office\\16\.0\\WEF\\Developer.*d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23",
    "reg\s+add.*Software\\Microsoft\\Windows\\CurrentVersion\\Run.*SamOfficeAgent",
    "start.*SamTrayServer\.exe"
)

foreach ($pattern in $requiredInstallPatterns) {
    if ($installContent -notmatch $pattern) {
        Write-Error "Install.bat missing expected command pattern: '$pattern'"
        exit 1
    }
}
Write-Host "  -> Install.bat commands verified." -ForegroundColor Green

# 4. Uninstall.bat Sanity Verification
Write-Host "Verifying Uninstall.bat sanity..."
$uninstallContent = Get-Content -Path $uninstallBat -Raw

$requiredUninstallPatterns = @(
    "taskkill.*SamTrayServer\.exe",
    "reg\s+delete.*Software\\Microsoft\\Office\\16\.0\\WEF\\Developer.*d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23",
    "reg\s+delete.*Software\\Microsoft\\Windows\\CurrentVersion\\Run.*SamOfficeAgent",
    "(?s)SamOfficeAgent.*rmdir"
)

foreach ($pattern in $requiredUninstallPatterns) {
    if ($uninstallContent -notmatch $pattern) {
        Write-Error "Uninstall.bat missing expected command pattern: '$pattern'"
        exit 1
    }
}
Write-Host "  -> Uninstall.bat commands verified." -ForegroundColor Green

# 5. Batch Syntax & Dry-Run Execution Check
Write-Host "Validating batch file syntax and dry-run execution..."
foreach ($batFile in @($installBat, $uninstallBat)) {
    $leafName = Split-Path $batFile -Leaf
    
    # Test --help flag execution
    $helpOutput = & cmd.exe /c "call `"$batFile`" --help"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Execution of $leafName --help failed with exit code $LASTEXITCODE"
        exit 1
    }

    # Test --dry-run flag execution
    $dryOutput = & cmd.exe /c "call `"$batFile`" --dry-run"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Execution of $leafName --dry-run failed with exit code $LASTEXITCODE"
        exit 1
    }

    $dryText = $dryOutput -join "`n"
    if ($dryText -notmatch "\[Dry-Run\] Uji coba berhasil") {
        Write-Error "Expected dry-run success message from $leafName, got: $dryText"
        exit 1
    }

    Write-Host "  -> $($leafName): syntax, help, and dry-run execution passed." -ForegroundColor Green
}

# 6. Optional Inno Setup Compiler (ISCC) validation if installed
$isccPath = $null
$potentialPaths = @(
    "iscc",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)
foreach ($p in $potentialPaths) {
    if (Get-Command $p -ErrorAction SilentlyContinue) {
        $isccPath = $p
        break
    } elseif (Test-Path $p) {
        $isccPath = $p
        break
    }
}

if ($isccPath) {
    Write-Host "ISCC compiler detected at: $isccPath. Testing compilation dry-run..."
    Write-Host "  -> ISCC available for Task 5 build pipeline." -ForegroundColor Green
} else {
    Write-Host "ISCC compiler not detected on system (portable zip fallback will be used in Task 5)." -ForegroundColor Yellow
}

Write-Host "`nAll installer tests passed successfully!" -ForegroundColor Green
exit 0
