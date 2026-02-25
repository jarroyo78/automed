#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Global config
$script:MaxRetries = 3
$script:RetryDelay = 2  # seconds

trap {
    Write-Host ""
    Write-Host "[ERROR] $_" -ForegroundColor Red
    Write-Host "Please take a photo of this screen and share it for support." -ForegroundColor Yellow
    pause
    exit 1
}

function Write-Step { param($msg) Write-Host "" ; Write-Host "[*] $msg" -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "    [WARN] $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "    [FAIL] $msg" -ForegroundColor Red }

function Test-PackageInstalled {
    param(
        [string]$PackageName,
        [string]$MinVersion
    )
    $pkg = Get-AppxPackage -Name $PackageName -ErrorAction SilentlyContinue
    if ($pkg) {
        $installedVersion = [Version]$pkg.Version
        $requiredVersion = [Version]$MinVersion
        if ($installedVersion -ge $requiredVersion) {
            return $true, $installedVersion
        }
    }
    return $false, $null
}

function Download-FileWithRetry {
    param(
        [string]$Url,
        [string]$OutFile,
        [string]$Description,
        [int]$MinSizeBytes = 100KB
    )
    
    for ($i = 1; $i -le $script:MaxRetries; $i++) {
        try {
            Write-Host "    Attempt $i of $script:MaxRetries..." -ForegroundColor Gray
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
            
            # Validate the download
            if (Test-Path $OutFile) {
                $file = Get-Item $OutFile
                if ($file.Length -ge $MinSizeBytes) {
                    Write-Ok "$Description downloaded ($([math]::Round($file.Length/1MB,2)) MB)"
                    return $true
                } else {
                    throw "File too small: $($file.Length) bytes"
                }
            } else {
                throw "File not created"
            }
        }
        catch {
            if ($i -eq $script:MaxRetries) {
                Write-Warn "Failed to download $Description after $script:MaxRetries attempts: $_"
                return $false
            }
            Write-Host "    Retry $i failed, waiting $script:RetryDelay seconds..." -ForegroundColor Gray
            Start-Sleep -Seconds $script:RetryDelay
        }
    }
    return $false
}

Clear-Host
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "   Medicus Xamarin UWP Installer - Version 2.0.0.0   " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

# 1. Admin check
Write-Step "Checking for Administrator privileges..."
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Fail "Must be run as Administrator. Right-click the .bat and select Run as administrator."
    pause
    exit 1
}
Write-Ok "Running as Administrator."

# 2. OS check
Write-Step "Checking operating system..."
$os = Get-CimInstance Win32_OperatingSystem
$buildNumber = [int]$os.BuildNumber
$osCaption = $os.Caption
Write-Host "    Detected: $osCaption (Build $buildNumber)" -ForegroundColor Gray
if ($buildNumber -lt 17763) {
    Write-Fail "Windows 10 version 1809 (Build 17763) or later is required."
    Write-Host "    Current build: $buildNumber" -ForegroundColor Red
    Write-Host "    Please update Windows and retry." -ForegroundColor Yellow
    pause
    exit 1
}
Write-Ok "OS is supported."

# 3. Locate files
Write-Step "Locating installer files..."
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$certFile = Join-Path $scriptDir "Medicus.Xamarin.UWP_1.0.0.0_x64.cer"
$dep1 = Join-Path $scriptDir "Microsoft.NET.CoreRuntime.2.2.appx"
$dep2 = Join-Path $scriptDir "Microsoft.NET.CoreFramework.Debug.2.2 - Copy.appx"
$deps = @($dep1, $dep2)

# Find main app package (supports both .appx and .appxbundle)
$mainApp = $null
$mainAppName = $null
foreach ($ext in @('*.appxbundle', '*.appx')) {
    $candidate = Get-ChildItem -Path $scriptDir -Filter $ext -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -like 'Medicus*' } |
                 Select-Object -First 1
    if ($candidate) {
        $mainApp = $candidate.FullName
        $mainAppName = $candidate.Name
        break
    }
}

$missing = @()
if (-not (Test-Path $certFile)) { $missing += "Medicus.Xamarin.UWP_1.0.0.0_x64.cer" }
if (-not $mainApp)              { $missing += "Medicus main app (.appx or .appxbundle)" }
if (-not (Test-Path $dep1))     { $missing += "Microsoft.NET.CoreRuntime.2.2.appx" }
if (-not (Test-Path $dep2))     { $missing += "Microsoft.NET.CoreFramework.Debug.2.2 - Copy.appx" }

if ($missing.Count -gt 0) {
    Write-Fail "Missing files in: $scriptDir"
    foreach ($m in $missing) {
        Write-Host "      - $m" -ForegroundColor Red
    }
    Write-Host "`nPlease ensure all installer files are in the same folder as this script." -ForegroundColor Yellow
    pause
    exit 1
}
Write-Ok "All required files found."

# 4. Enable sideloading
Write-Step "Enabling sideloading..."
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "AllowAllTrustedApps" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord -Force
    
    # Verify the settings stuck
    $check1 = (Get-ItemProperty -Path $regPath -Name "AllowAllTrustedApps" -ErrorAction SilentlyContinue).AllowAllTrustedApps
    $check2 = (Get-ItemProperty -Path $regPath -Name "AllowDevelopmentWithoutDevLicense" -ErrorAction SilentlyContinue).AllowDevelopmentWithoutDevLicense
    
    if ($check1 -eq 1 -and $check2 -eq 1) {
        Write-Ok "Sideloading enabled."
    } else {
        Write-Warn "Sideloading registry keys set but verification failed. Continuing anyway..."
    }
} catch {
    Write-Warn "Could not set sideloading keys: $_"
    Write-Warn "Continuing anyway - if installation fails, enable Developer Mode manually in Settings."
}

# 5. Install certificate
Write-Step "Installing certificate into Local Machine > Trusted People..."
try {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $certFile
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
        [System.Security.Cryptography.X509Certificates.StoreName]::TrustedPeople,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
    )
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($existing) {
        Write-Ok "Certificate already installed (Thumbprint: $($cert.Thumbprint.Substring(0,8))...)"
    } else {
        $store.Add($cert)
        Write-Ok "Certificate installed successfully (Thumbprint: $($cert.Thumbprint.Substring(0,8))...)"
    }
    $store.Close()
} catch {
    Write-Fail "Failed to install certificate: $_"
    Write-Host "    This is critical - the app won't install without a trusted certificate." -ForegroundColor Red
    pause
    exit 1
}

# 6. Check and install runtime dependencies
Write-Step "Checking runtime dependencies..."

# Create a list to track dependency files we'll need for installation
$dependencyFiles = New-Object System.Collections.Generic.List[string]
foreach ($dep in $deps) { $dependencyFiles.Add($dep) }

# Function to check if a file exists and has reasonable size
function Test-VCLibsFile {
    param([string]$FilePath, [int]$MinSizeKB = 500)
    if (Test-Path $FilePath) {
        $file = Get-Item $FilePath
        $sizeKB = [math]::Round($file.Length/1KB, 0)
        if ($file.Length -ge ($MinSizeKB * 1KB)) {
            return $true, $sizeKB
        } else {
            Write-Host "    File exists but too small: $sizeKB KB (need at least $MinSizeKB KB)" -ForegroundColor Yellow
        }
    }
    return $false, $null
}

# VCLibs Base (Microsoft.VCLibs.140.00)
Write-Host "`n    Checking Microsoft.VCLibs.140.00 (Base Runtime)..." -ForegroundColor Gray

# Check if already installed (any version)
$baseInstalled = Get-AppxPackage -Name "Microsoft.VCLibs.140.00" -ErrorAction SilentlyContinue | Where-Object {$_.Architecture -eq "X64"} | Select-Object -First 1

if ($baseInstalled) {
    Write-Ok "Microsoft.VCLibs.140.00 is already installed (version: $($baseInstalled.Version))"
} else {
    Write-Warn "Microsoft.VCLibs.140.00 not found or not the correct architecture"
    
    # Look for local file with various possible names
    $baseFile = $null
    $possibleBaseNames = @(
        "Microsoft.VCLibs.x64.14.00.appx",
        "Microsoft.VCLibs.140.00.appx",
        "vclibs-uwp-x64.appx"
    )
    
    foreach ($name in $possibleBaseNames) {
        $testPath = Join-Path $scriptDir $name
        $valid, $sizeKB = Test-VCLibsFile -FilePath $testPath -MinSizeKB 500
        if ($valid) {
            $baseFile = $testPath
            Write-Ok "Found local file: $name ($sizeKB KB)"
            break
        }
    }
    
    # If not found with common names, try wildcard search
    if (-not $baseFile) {
        $wildcardMatch = Get-ChildItem $scriptDir -Filter "*VCLibs*14.00*.appx" -ErrorAction SilentlyContinue | 
                         Where-Object { $_.Name -notlike "*Desktop*" -and $_.Name -notlike "*UWPDesktop*" -and $_.Length -gt 500KB } |
                         Select-Object -First 1
        if ($wildcardMatch) {
            $baseFile = $wildcardMatch.FullName
            $sizeKB = [math]::Round($wildcardMatch.Length/1KB, 0)
            Write-Ok "Found local file: $($wildcardMatch.Name) ($sizeKB KB)"
        }
    }
    
    if ($baseFile) {
        $dependencyFiles.Add($baseFile)
    } else {
        Write-Warn "No local Base VCLibs file found. The main app installation may fail if not already present."
        Write-Host "    Expected file: Microsoft.VCLibs.x64.14.00.appx (should be ~700-900 KB)" -ForegroundColor Gray
    }
}

# VCLibs Desktop (Microsoft.VCLibs.140.00.UWPDesktop)
Write-Host "`n    Checking Microsoft.VCLibs.140.00.UWPDesktop (Desktop Runtime)..." -ForegroundColor Gray

# Check if already installed (any version)
$desktopInstalled = Get-AppxPackage -Name "Microsoft.VCLibs.140.00.UWPDesktop" -ErrorAction SilentlyContinue | Where-Object {$_.Architecture -eq "X64"} | Select-Object -First 1

if ($desktopInstalled) {
    Write-Ok "Microsoft.VCLibs.140.00.UWPDesktop is already installed (version: $($desktopInstalled.Version))"
} else {
    Write-Warn "Microsoft.VCLibs.140.00.UWPDesktop not found or not the correct architecture"
    
    # Look for local file with various possible names
    $desktopFile = $null
    $possibleDesktopNames = @(
        "Microsoft.VCLibs.x64.14.00.Desktop.appx",
        "Microsoft.VCLibs.140.00.UWPDesktop.appx",
        "vclibs-desktop-x64.appx"
    )
    
    foreach ($name in $possibleDesktopNames) {
        $testPath = Join-Path $scriptDir $name
        $valid, $sizeKB = Test-VCLibsFile -FilePath $testPath -MinSizeKB 5000  # Desktop should be several MB
        if ($valid) {
            $desktopFile = $testPath
            $sizeMB = [math]::Round($sizeKB/1024, 2)
            Write-Ok "Found local file: $name ($sizeMB MB)"
            break
        }
    }
    
    # If not found with common names, try wildcard search
    if (-not $desktopFile) {
        $wildcardMatch = Get-ChildItem $scriptDir -Filter "*UWPDesktop*_x64.appx" -ErrorAction SilentlyContinue | 
                         Where-Object { $_.Length -gt 5MB } |
                         Select-Object -First 1
        if ($wildcardMatch) {
            $desktopFile = $wildcardMatch.FullName
            $sizeMB = [math]::Round($wildcardMatch.Length/1MB, 2)
            Write-Ok "Found local file: $($wildcardMatch.Name) ($sizeMB MB)"
        }
    }
    
    if ($desktopFile) {
        $dependencyFiles.Add($desktopFile)
    } else {
        Write-Warn "No local Desktop VCLibs file found. The main app installation may fail if not already present."
        Write-Host "    Expected file: Microsoft.VCLibs.x64.14.00.Desktop.appx (should be ~6-7 MB)" -ForegroundColor Gray
    }
}

# Microsoft.UI.Xaml 2.8
Write-Host "`n    Checking Microsoft.UI.Xaml.2.8..." -ForegroundColor Gray
$installed = Get-AppxPackage -Name "Microsoft.UI.Xaml.2.8" -ErrorAction SilentlyContinue | Where-Object {$_.Architecture -eq "X64"} | Select-Object -First 1
if ($installed) {
    Write-Ok "Microsoft.UI.Xaml.2.8 v$($installed.Version) already installed."
} else {
    Write-Warn "Microsoft.UI.Xaml.2.8 not found or not the correct architecture."
    
    # Check if we have a local copy
    $uiXamlLocal = Join-Path $scriptDir "Microsoft.UI.Xaml.2.8.x64.appx"
    if (Test-Path $uiXamlLocal) {
        $size = (Get-Item $uiXamlLocal).Length
        Write-Ok "Found local file: Microsoft.UI.Xaml.2.8.x64.appx ($([math]::Round($size/1MB,2)) MB)"
        $dependencyFiles.Add($uiXamlLocal)
    } else {
        Write-Host "    Downloading from Microsoft GitHub..." -ForegroundColor Gray
        $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        $uiXamlPath = Join-Path $env:TEMP "Microsoft.UI.Xaml.2.8.x64.appx"
        
        try {
            Invoke-WebRequest -Uri $uiXamlUrl -OutFile $uiXamlPath -UseBasicParsing
            $downloaded = Get-Item $uiXamlPath
            if ($downloaded.Length -gt 1MB) {
                Write-Host "    Downloaded OK ($([math]::Round($downloaded.Length / 1MB, 1)) MB). Installing..." -ForegroundColor Gray
                Add-AppxPackage -Path $uiXamlPath -ErrorAction Stop
                Write-Ok "Microsoft.UI.Xaml.2.8 installed successfully."
                Remove-Item $uiXamlPath -Force -ErrorAction SilentlyContinue
            } else {
                throw "File too small: $($downloaded.Length) bytes"
            }
        } catch {
            Write-Warn "Could not download/install Microsoft.UI.Xaml.2.8 : $_"
            Write-Warn "Installation will continue but the app may not work on this machine."
        }
    }
}

# Show summary of dependency files found
Write-Host "`n    Dependency files collected: $($dependencyFiles.Count)" -ForegroundColor Gray
if ($dependencyFiles.Count -gt 0) {
    foreach ($dep in $dependencyFiles) {
        $sizeKB = [math]::Round((Get-Item $dep).Length/1KB, 0)
        Write-Host "      - $(Split-Path $dep -Leaf) ($sizeKB KB)" -ForegroundColor Gray
    }
}

# 7. Install .NET dependencies
Write-Step "Installing .NET dependency packages..."
$netDepsSuccess = $true
foreach ($dep in $deps) {
    $leaf = Split-Path $dep -Leaf
    Write-Host "    Installing: $leaf" -ForegroundColor Gray
    try {
        Add-AppxPackage -Path $dep -ErrorAction Stop
        Write-Ok "$leaf installed."
    } catch {
        if ($_.Exception.Message -match "0x80073D06|already installed|higher version") {
            Write-Ok "$leaf already present, skipping."
        } else {
            Write-Warn "Could not install $leaf"
            Write-Warn "$($_.Exception.Message)"
            Write-Warn "Continuing anyway..."
            $netDepsSuccess = $false
        }
    }
}

if (-not $netDepsSuccess) {
    Write-Warn "Some .NET dependencies failed. The main app may still work if they're already installed."
}

# 8. Install main app - FIXED VERSION
Write-Step "Installing Medicus..."
Write-Host "    Source: $mainAppName" -ForegroundColor Gray

# First, install VCLibs packages separately
Write-Host "    Installing VCLibs dependencies first..." -ForegroundColor Gray

$vclibsBase = Join-Path $scriptDir "Microsoft.VCLibs.x64.14.00.appx"
$vclibsDesktop = Join-Path $scriptDir "Microsoft.VCLibs.x64.14.00.Desktop.appx"

# Validate Base VCLibs file before attempting install
if (Test-Path $vclibsBase) {
    $fileSize = (Get-Item $vclibsBase).Length
    if ($fileSize -lt 500KB) {
        Write-Warn "Base VCLibs file is too small ($([math]::Round($fileSize/1KB,0)) KB) - may be corrupted. Downloading fresh copy..."
        
        # Download fresh copy
        $url = "https://aka.ms/Microsoft.VCLibs.x64.14.00.appx"
        try {
            Remove-Item $vclibsBase -Force -ErrorAction SilentlyContinue
            Invoke-WebRequest -Uri $url -OutFile $vclibsBase -UseBasicParsing
            $newSize = (Get-Item $vclibsBase).Length
            Write-Host "      Downloaded: $([math]::Round($newSize/1KB,0)) KB" -ForegroundColor Gray
        } catch {
            Write-Warn "Failed to download: $_"
        }
    }
}

# Install Base VCLibs
if (Test-Path $vclibsBase) {
    $size = [math]::Round((Get-Item $vclibsBase).Length/1KB,0)
    Write-Host "      Installing Microsoft.VCLibs.x64.14.00.appx ($size KB)..." -ForegroundColor Gray
    try {
        Add-AppxPackage -Path $vclibsBase -ErrorAction Stop
        Write-Ok "VCLibs Base installed"
    } catch {
        if ($_.Exception.Message -match "0x80073D06|already installed") {
            Write-Ok "VCLibs Base already present"
        } else {
            Write-Warn "Failed to install VCLibs Base: $_"
            Write-Warn "Will continue with Desktop VCLibs only"
        }
    }
} else {
    Write-Warn "Base VCLibs file not found"
}

# Install Desktop VCLibs
if (Test-Path $vclibsDesktop) {
    $size = [math]::Round((Get-Item $vclibsDesktop).Length/1KB,0)
    Write-Host "      Installing Microsoft.VCLibs.x64.14.00.Desktop.appx ($size KB)..." -ForegroundColor Gray
    try {
        Add-AppxPackage -Path $vclibsDesktop -ErrorAction Stop
        Write-Ok "VCLibs Desktop installed"
    } catch {
        if ($_.Exception.Message -match "0x80073D06|already installed") {
            Write-Ok "VCLibs Desktop already present"
        } else {
            Write-Warn "Failed to install VCLibs Desktop: $_"
        }
    }
}

# Wait for installations to register
Start-Sleep -Seconds 3

# Now install the main app
Write-Host "    Installing main application..." -ForegroundColor Gray
try {
    Add-AppxPackage -Path $mainApp -ErrorAction Stop
    Write-Ok "Medicus installed successfully."
    $installSuccess = $true
} catch {
    $err = $_.Exception.Message
    Write-Fail "Installation failed: $err"
    
    # Check if it's still a VCLibs issue
    if ($err -match "Microsoft.VCLibs.140.00") {
        Write-Host "    Still missing VCLibs. Let's try one more approach..." -ForegroundColor Yellow
        
        # Try installing via DISM as fallback
        Write-Host "    Attempting DISM installation..." -ForegroundColor Gray
        try {
            dism /online /Add-ProvisionedAppxPackage /PackagePath:$vclibsBase /SkipLicense 2>$null
            dism /online /Add-ProvisionedAppxPackage /PackagePath:$vclibsDesktop /SkipLicense 2>$null
            Start-Sleep -Seconds 3
            Add-AppxPackage -Path $mainApp -ErrorAction Stop
            Write-Ok "Medicus installed via DISM fallback."
        } catch {
            Write-Fail "All attempts failed. Please check the VCLibs files manually."
        }
    }
    
    Write-Host "    Check Event Viewer > Windows Logs > Application for more detail." -ForegroundColor Yellow
    pause
    exit 1
}

# 9. Verify installation
Write-Step "Verifying installation..."
Start-Sleep -Seconds 3
$installed = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
if ($installed) {
    Write-Ok "Verified: $($installed.Name) v$($installed.Version)"
    Write-Host "    PackageFamilyName: $($installed.PackageFamilyName)" -ForegroundColor Gray
    Write-Host "    InstallLocation: $($installed.InstallLocation)" -ForegroundColor Gray
} else {
    Write-Warn "Could not verify via Get-AppxPackage. Check the Start Menu for 'Medicus'."
    Write-Warn "Installation may have succeeded but verification failed."
}

# 10. Public Desktop shortcut
Write-Step "Creating shortcut on Public Desktop..."
try {
    $publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
    $shortcutPath = Join-Path $publicDesktop "Medicus.lnk"
    
    # Remove existing shortcut if it exists
    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
    }
    
    $appPackage = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
    if ($appPackage) {
        $manifest = Get-AppxPackageManifest -Package $appPackage.PackageFullName
        $appId = $manifest.Package.Applications.Application.Id
        $aumid = "$($appPackage.PackageFamilyName)!$appId"
        
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"
        $shortcut.Arguments = "shell:AppsFolder\$aumid"
        $shortcut.WorkingDirectory = ""
        $shortcut.Description = "Launch Medicus"
        $shortcut.WindowStyle = 1
        
        # Find best icon
        $installLocation = $appPackage.InstallLocation
        $iconFile = $null
        
        # Try .ico files first
        $iconFile = Get-ChildItem -Path $installLocation -Filter "*.ico" -Recurse -ErrorAction SilentlyContinue | 
                   Select-Object -First 1
        
        # Try common PNG patterns
        if (-not $iconFile) {
            $iconFile = Get-ChildItem -Path $installLocation -Filter "*.png" -Recurse -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match "StoreLogo|Square150|Square44|Logo|AppIcon" } |
                       Sort-Object Length -Descending |
                       Select-Object -First 1
        }
        
        # Last resort: any PNG
        if (-not $iconFile) {
            $iconFile = Get-ChildItem -Path $installLocation -Filter "*.png" -Recurse -ErrorAction SilentlyContinue |
                       Sort-Object Length -Descending |
                       Select-Object -First 1
        }
        
        if ($iconFile) {
            $shortcut.IconLocation = "$($iconFile.FullName),0"
            Write-Ok "Icon found: $($iconFile.Name)"
        } else {
            Write-Warn "No icon found, shortcut will use default."
        }
        
        $shortcut.Save()
        
        # Verify shortcut was created
        if (Test-Path $shortcutPath) {
            Write-Ok "Shortcut created at: $shortcutPath"
        } else {
            throw "Shortcut file not created"
        }
    } else {
        Write-Warn "Could not find Medicus package for shortcut. Skipping."
        Write-Warn "You can pin it manually from the Start Menu after installation."
    }
} catch {
    Write-Warn "Could not create shortcut: $_"
    Write-Warn "You can pin the app manually from the Start Menu."
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "   Installation complete! Launching Medicus now...    " -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

# 11. Launch the app
Write-Step "Launching Medicus..."
try {
    $appPackage = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
    if ($appPackage) {
        $manifest = Get-AppxPackageManifest -Package $appPackage.PackageFullName
        $appId = $manifest.Package.Applications.Application.Id
        $aumid = "$($appPackage.PackageFamilyName)!$appId"
        
        Write-Host "    Launching: $aumid" -ForegroundColor Gray
        Start-Process "explorer.exe" "shell:AppsFolder\$aumid"
        Start-Sleep -Seconds 2
        
        # Check if process started (rough check)
        $processes = Get-Process -Name "*Medicus*" -ErrorAction SilentlyContinue
        if ($processes) {
            Write-Ok "Medicus launched successfully."
        } else {
            Write-Warn "Launch command executed, but couldn't verify process. Check the taskbar."
        }
    } else {
        Write-Warn "Could not find package to launch. Open Medicus from the desktop shortcut."
    }
} catch {
    Write-Warn "Could not launch Medicus automatically: $_"
    Write-Warn "Open it manually from the desktop shortcut or Start Menu."
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "   Thank you for installing Medicus!                  " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Optional: Log file for troubleshooting
$logFile = Join-Path $env:TEMP "Medicus_Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$transcript = Start-Transcript -Path $logFile -Append -ErrorAction SilentlyContinue
if ($transcript) {
    Write-Host "Installation log saved to: $logFile" -ForegroundColor Gray
    Stop-Transcript -ErrorAction SilentlyContinue
}

pause