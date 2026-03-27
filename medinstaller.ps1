#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# GitHub repo configuration
$script:RepoBase = "https://raw.githubusercontent.com/jarroyo78/automed/main"
$script:TempDir = Join-Path $env:TEMP "MedicusInstaller_$(Get-Date -Format 'yyyyMMddHHmmss')"
$script:MaxRetries = 3
$script:RetryDelay = 2

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

function Test-PackageInstalled {
    param(
        [string]$PackageName,
        [string]$MinVersion
    )
    
    $pkgs = Get-AppxPackage -Name $PackageName -ErrorAction SilentlyContinue
    if (-not $pkgs) { return $false, $null }
    
    $pkg = $pkgs | Where-Object { $_.Architecture -eq "X64" } | Select-Object -First 1
    if (-not $pkg) { $pkg = $pkgs | Select-Object -First 1 }
    
    try {
        if (-not $pkg.Version) { return $false, $null }
        
        $versionString = $pkg.Version.ToString()
        if ($versionString -match '(\d+\.\d+\.\d+\.\d+)') {
            $versionString = $matches[1]
        }
        
        $installedVersion = [Version]$versionString
        $requiredVersion = [Version]$MinVersion
        
        if ($installedVersion -ge $requiredVersion) {
            return $true, $installedVersion
        } else {
            return $false, $installedVersion
        }
    }
    catch {
        return $false, $null
    }
}

Clear-Host
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "   Medicus Xamarin UWP Installer - GitHub Edition    " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

# 1. Admin check
Write-Step "Checking for Administrator privileges..."
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Fail "Must be run as Administrator. Restarting with elevation..."
    
    # Self-elevate script
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process PowerShell -Verb RunAs -ArgumentList $arguments
    exit
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
    pause
    exit 1
}
Write-Ok "OS is supported."

# 3. Create temp folder and download files
Write-Step "Setting up installation environment..."
New-Item -ItemType Directory -Path $script:TempDir -Force | Out-Null
Write-Host "    Using temp folder: $script:TempDir" -ForegroundColor Gray

# List of required files
$requiredFiles = @(
    @{ File = "Medicus.Xamarin.UWP_1.0.0.0_x64.cer"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Medicus.Xamarin.UWP_1.0.0.0_x64.cer"; MinSize = 800 },  # 800 bytes (actual is 811)
    @{ File = "Microsoft.NET.CoreRuntime.2.2.appx"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Microsoft.NET.CoreRuntime.2.2.appx"; MinSize = 5MB },  # 5MB works
    @{ File = "Microsoft.NET.CoreFramework.Debug.2.2 - Copy.appx"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Microsoft.NET.CoreFramework.Debug.2.2%20-%20Copy.appx"; MinSize = 7MB },  # 7MB works
    @{ File = "Microsoft.VCLibs.x64.14.00.appx"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Microsoft.VCLibs.x64.14.00.appx"; MinSize = 800KB },  # 800KB works
    @{ File = "Microsoft.VCLibs.x64.14.00.Desktop.appx"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Microsoft.VCLibs.x64.14.00.Desktop.appx"; MinSize = 6MB },  # 6MB works
    @{ File = "Medicus.Xamarin.UWP_1.0.411.0_x64.appx"; Url = "https://media.githubusercontent.com/media/jarroyo78/automed/main/Medicus.Xamarin.UWP_1.0.411.0_x64.appx"; MinSize = 90MB }  # 90MB works
)

Write-Host "    Downloading required files from GitHub..." -ForegroundColor Cyan
$downloadSuccess = $true

foreach ($file in $requiredFiles) {
    $filePath = Join-Path $script:TempDir $file.File
    Write-Host "      Downloading $($file.File)..." -ForegroundColor Gray
    
    $success = Download-FileWithRetry -Url $file.Url -OutFile $filePath -Description $file.File -MinSizeBytes $file.MinSize
    if (-not $success) {
        $downloadSuccess = $false
    }
}

if (-not $downloadSuccess) {
    Write-Fail "Failed to download some required files. Check your internet connection."
    pause
    exit 1
}

Write-Ok "All required files downloaded successfully."

# Set file paths
$certFile = Join-Path $script:TempDir "Medicus.Xamarin.UWP_1.0.0.0_x64.cer"
$dep1 = Join-Path $script:TempDir "Microsoft.NET.CoreRuntime.2.2.appx"
$dep2 = Join-Path $script:TempDir "Microsoft.NET.CoreFramework.Debug.2.2 - Copy.appx"
$deps = @($dep1, $dep2)
$mainApp = Join-Path $script:TempDir "Medicus.Xamarin.UWP_1.0.411.0_x64.appx"
$mainAppName = "Medicus.Xamarin.UWP_1.0.411.0_x64.appx"
$vclibsBase = Join-Path $script:TempDir "Microsoft.VCLibs.x64.14.00.appx"
$vclibsDesktop = Join-Path $script:TempDir "Microsoft.VCLibs.x64.14.00.Desktop.appx"

# 4. Enable sideloading
Write-Step "Enabling sideloading..."
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "AllowAllTrustedApps" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Type DWord -Force
    Write-Ok "Sideloading enabled."
} catch {
    Write-Warn "Could not set sideloading keys: $_"
    Write-Warn "Continuing anyway..."
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
        Write-Ok "Certificate already installed."
    } else {
        $store.Add($cert)
        Write-Ok "Certificate installed successfully."
    }
    $store.Close()
} catch {
    Write-Fail "Failed to install certificate: $_"
    pause
    exit 1
}

# 6. Check runtime dependencies
Write-Step "Checking runtime dependencies..."
$dependencyFiles = New-Object System.Collections.Generic.List[string]
foreach ($dep in $deps) { $dependencyFiles.Add($dep) }

# VCLibs Base
Write-Host "`n    Checking Microsoft.VCLibs.140.00 (Base Runtime)..." -ForegroundColor Gray
$baseInstalled, $baseVersion = Test-PackageInstalled -PackageName "Microsoft.VCLibs.140.00" -MinVersion "14.0.33519.0"
if ($baseInstalled) {
    Write-Ok "Microsoft.VCLibs.140.00 v$baseVersion already installed."
} else {
    Write-Warn "Microsoft.VCLibs.140.00 not found or outdated."
    if (Test-Path $vclibsBase) {
        $size = [math]::Round((Get-Item $vclibsBase).Length/1KB, 0)
        Write-Ok "Found local file: Microsoft.VCLibs.x64.14.00.appx ($size KB)"
        $dependencyFiles.Add($vclibsBase)
    }
}

# VCLibs Desktop
Write-Host "`n    Checking Microsoft.VCLibs.140.00.UWPDesktop (Desktop Runtime)..." -ForegroundColor Gray
$desktopInstalled, $desktopVersion = Test-PackageInstalled -PackageName "Microsoft.VCLibs.140.00.UWPDesktop" -MinVersion "14.0.33728.0"
if ($desktopInstalled) {
    Write-Ok "Microsoft.VCLibs.140.00.UWPDesktop v$desktopVersion already installed."
} else {
    Write-Warn "Microsoft.VCLibs.140.00.UWPDesktop not found or outdated."
    if (Test-Path $vclibsDesktop) {
        $size = [math]::Round((Get-Item $vclibsDesktop).Length/1MB, 2)
        Write-Ok "Found local file: Microsoft.VCLibs.x64.14.00.Desktop.appx ($size MB)"
        $dependencyFiles.Add($vclibsDesktop)
    }
}

# Microsoft.UI.Xaml 2.8
Write-Host "`n    Checking Microsoft.UI.Xaml.2.8..." -ForegroundColor Gray
$uiXamlInstalled, $uiXamlVersion = Test-PackageInstalled -PackageName "Microsoft.UI.Xaml.2.8" -MinVersion "8.2208.12001.0"
if ($uiXamlInstalled) {
    Write-Ok "Microsoft.UI.Xaml.2.8 v$uiXamlVersion already installed."
} else {
    Write-Warn "Microsoft.UI.Xaml.2.8 not found. Installing..."
    $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
    $uiXamlPath = Join-Path $env:TEMP "Microsoft.UI.Xaml.2.8.x64.appx"
    
    if (Download-FileWithRetry -Url $uiXamlUrl -OutFile $uiXamlPath -Description "Microsoft.UI.Xaml.2.8" -MinSizeBytes 1MB) {
        try {
            Add-AppxPackage -Path $uiXamlPath -ErrorAction Stop
            Write-Ok "Microsoft.UI.Xaml.2.8 installed successfully."
        } catch {
            if ($_.Exception.Message -match "0x80073D06|already installed") {
                Write-Ok "Microsoft.UI.Xaml.2.8 already present."
            } else {
                Write-Warn "Installation failed: $_"
            }
        }
        Remove-Item $uiXamlPath -Force -ErrorAction SilentlyContinue
    }
}

# 7. Check AppX deployment service
Write-Step "Checking AppX deployment service..."
try {
    $svc = Get-Service -Name "AppXSvc" -ErrorAction Stop
    if ($svc.Status -ne "Running") {
        Write-Warn "AppXSvc is not running. Starting it..."
        Start-Service -Name "AppXSvc" -ErrorAction Stop
        Write-Ok "AppXSvc started."
    } else {
        Write-Ok "AppXSvc is running."
    }
} catch {
    Write-Warn "Could not verify AppXSvc: $_"
}

# 8. Install .NET dependencies
Write-Step "Installing .NET dependency packages..."
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
            Write-Warn "Could not install $leaf : $_"
        }
    }
}

# 9. Install VCLibs separately first
Write-Step "Installing VCLibs framework packages..."
$vclibsInstalled = @($baseInstalled, $desktopInstalled)
$vclibsFiles = @($vclibsBase, $vclibsDesktop)
foreach ($i in 0..($vclibsFiles.Count - 1)) {
    $vclib = $vclibsFiles[$i]
    if ($vclibsInstalled[$i]) {
        Write-Ok "$(Split-Path $vclib -Leaf) already installed, skipping."
        continue
    }
    if (Test-Path $vclib) {
        $name = Split-Path $vclib -Leaf
        Write-Host "    Installing $name..." -ForegroundColor Gray
        try {
            Add-AppxPackage -Path $vclib -ErrorAction Stop
            Write-Ok "$name installed"
        } catch {
            if ($_.Exception.Message -match "0x80073D06|already installed") {
                Write-Ok "$name already present"
            } else {
                Write-Warn "Failed to install $name : $_"
            }
        }
    }
}
Start-Sleep -Seconds 3

# 10. Install main app
Write-Step "Installing Medicus..."
Write-Host "    Source: $mainAppName" -ForegroundColor Gray

try {
    Add-AppxPackage -Path $mainApp -ErrorAction Stop
    Write-Ok "Medicus installed successfully."
    $installSuccess = $true
} catch {
    $error = $_.Exception.Message
    Write-Fail "Installation failed: $error"
    
    # Log file for troubleshooting
    $logFile = Join-Path $env:TEMP "Medicus_Error_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $error | Out-File -FilePath $logFile
    Write-Host "    Error log saved to: $logFile" -ForegroundColor Yellow
    
    pause
    exit 1
}

# 11. Verify installation
Write-Step "Verifying installation..."
Start-Sleep -Seconds 3
$installed = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
if ($installed) {
    Write-Ok "Verified: $($installed.Name) v$($installed.Version)"
} else {
    Write-Warn "Could not verify installation. Check the Start Menu for 'Medicus'."
}

# 12. Create desktop shortcut
Write-Step "Creating shortcut on Public Desktop..."
try {
    $publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
    $shortcutPath = Join-Path $publicDesktop "Medicus.lnk"
    
    $appPackage = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
    if ($appPackage) {
        $manifest = Get-AppxPackageManifest -Package $appPackage.PackageFullName
        $appId = $manifest.Package.Applications.Application.Id
        $aumid = "$($appPackage.PackageFamilyName)!$appId"
        
        # Extract the app icon to a public location so the shortcut can read it
        $iconDir = "C:\ProgramData\Medicus"
        $iconPath = "$iconDir\medicus.ico"
        if (-not (Test-Path $iconDir)) { New-Item -ItemType Directory -Path $iconDir | Out-Null }

        $srcPng = Join-Path $appPackage.InstallLocation "Assets\Square150x150Logo.scale-200.png"
        if (-not (Test-Path $srcPng)) {
            # Fallback to any available square logo
            $srcPng = Get-ChildItem (Join-Path $appPackage.InstallLocation "Assets") -Filter "Square150x150Logo*" |
                      Select-Object -First 1 -ExpandProperty FullName
        }

        if ($srcPng -and (Test-Path $srcPng)) {
            Add-Type -AssemblyName System.Drawing
            $img = [System.Drawing.Image]::FromFile($srcPng)
            $sizes = @(256, 128, 64, 48, 32, 16)
            $ms = New-Object System.IO.MemoryStream
            $bw = New-Object System.IO.BinaryWriter($ms)
            $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
            $dataOffset = 6 + 16 * $sizes.Count
            $imageData = @()
            foreach ($size in $sizes) {
                $bmp = New-Object System.Drawing.Bitmap($size, $size)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.InterpolationMode = 'HighQualityBicubic'
                $g.DrawImage($img, 0, 0, $size, $size)
                $g.Dispose()
                $imgMs = New-Object System.IO.MemoryStream
                $bmp.Save($imgMs, [System.Drawing.Imaging.ImageFormat]::Png)
                $bmp.Dispose()
                $imageData += ,$imgMs.ToArray()
                $imgMs.Dispose()
            }
            $offset = $dataOffset
            foreach ($i in 0..($sizes.Count - 1)) {
                $sz = $sizes[$i]; $data = $imageData[$i]
                $bw.Write([byte]($sz -band 0xFF)); $bw.Write([byte]($sz -band 0xFF))
                $bw.Write([byte]0); $bw.Write([byte]0)
                $bw.Write([uint16]1); $bw.Write([uint16]32)
                $bw.Write([uint32]$data.Length); $bw.Write([uint32]$offset)
                $offset += $data.Length
            }
            foreach ($data in $imageData) { $bw.Write($data) }
            $bw.Flush()
            [System.IO.File]::WriteAllBytes($iconPath, $ms.ToArray())
            $ms.Dispose(); $img.Dispose()
        }

        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"
        $shortcut.Arguments = "shell:AppsFolder\$aumid"
        $shortcut.WindowStyle = 1
        if (Test-Path $iconPath) { $shortcut.IconLocation = "$iconPath,0" }
        $shortcut.Save()
        Write-Ok "Shortcut created at: $shortcutPath"
    }
} catch {
    Write-Warn "Could not create shortcut: $_"
}

# 13. Cleanup
Write-Step "Cleaning up temporary files..."
try {
    Remove-Item $script:TempDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Ok "Temporary files removed."
} catch {
    Write-Warn "Could not remove temp files: $_"
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "   Installation complete! Launching Medicus now...    " -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

# 14. Launch the app
try {
    $appPackage = Get-AppxPackage | Where-Object { $_.Name -like '*Medicus*' -or $_.Name -like '*medicus*' } | Select-Object -First 1
    if ($appPackage) {
        $manifest = Get-AppxPackageManifest -Package $appPackage.PackageFullName
        $appId = $manifest.Package.Applications.Application.Id
        $aumid = "$($appPackage.PackageFamilyName)!$appId"
        Start-Process "explorer.exe" "shell:AppsFolder\$aumid"
        Write-Ok "Medicus launched."
    }
} catch {
    Write-Warn "Could not launch Medicus automatically."
}

Write-Host ""
pause