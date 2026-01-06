# PowerShell script to create Outlook Phishing Reporter installer (EXE and MSI)
# This script packages the built VSTO add-in and creates installers

param(
    [string]$Configuration = "Release",
    [string]$Platform = "Any CPU",
    [switch]$BuildMSI = $false
)

$ErrorActionPreference = "Stop"

# Define paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseDir = Join-Path $scriptDir "OutlookPhishingReporter\bin\Release"
$outputDir = Join-Path $scriptDir "OutlookPhishingReporterSetup\Release"
$msiDir = Join-Path $scriptDir "OutlookPhishingReporterSetup\MSISource"
$distDir = Join-Path $scriptDir "OutlookPhishingReporterSetup\Distribution"

Write-Host "`n=================================" -ForegroundColor Cyan
Write-Host "Outlook Phishing Reporter Installer Builder" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check if Release folder exists
if (-not (Test-Path $releaseDir)) {
    Write-Host "ERROR: Release folder not found at $releaseDir" -ForegroundColor Red
    Write-Host "Please build the solution first!" -ForegroundColor Red
    exit 1
}

Write-Host "[+] Release build folder found: $releaseDir" -ForegroundColor Green

# Check for required files
$requiredFiles = @(
    "OutlookPhishingReporter.dll",
    "OutlookPhishingReporter.vsto",
    "OutlookPhishingReporter.dll.manifest"
)

foreach ($file in $requiredFiles) {
    $filePath = Join-Path $releaseDir $file
    if (-not (Test-Path $filePath)) {
        Write-Host "ERROR: Required file not found: $file" -ForegroundColor Red
        exit 1
    }
    Write-Host "[+] Found: $file" -ForegroundColor Green
}

# Create output directories
@($outputDir, $msiDir, $distDir) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

Write-Host ""
Write-Host "Copying files to output directory..." -ForegroundColor Yellow

$files = Get-ChildItem -Path $releaseDir -Include "*.dll", "*.vsto", "*.manifest", "*.pdb"
foreach ($file in $files) {
    Copy-Item -Path $file.FullName -Destination $outputDir -Force
    Copy-Item -Path $file.FullName -Destination $msiDir -Force
    Write-Host "  [+] Copied: $($file.Name)" -ForegroundColor Green
}

# Create improved setup.bat
$setupBatch = @"
@echo off
REM Outlook Phishing Reporter - Setup Script
REM This script installs the Outlook add-in with registry configuration

setlocal enabledelayedexpansion

echo.
echo ===================================================
echo Outlook Phishing Reporter Setup
echo ===================================================
echo.

REM Check for admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: This installer requires administrator privileges.
    echo Please right-click and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

REM Get the script directory
set SCRIPTDIR=%~dp0

REM Define registry paths
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set DLLPATH=%SCRIPTDIR%OutlookPhishingReporter.dll

echo [*] Installing Outlook Phishing Reporter add-in...
echo [*] Installation path: %SCRIPTDIR%
echo.

REM Check if required files exist
if not exist "%SCRIPTDIR%OutlookPhishingReporter.dll" (
    echo ERROR: OutlookPhishingReporter.dll not found!
    pause
    exit /b 1
)

REM Create the registry key for the add-in
reg add "%REGPATH%" /f >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to create registry key
    pause
    exit /b 1
)

REM Set LoadBehavior to 3 (load at startup)
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1

REM Set the description
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

REM Set the path to the manifest
reg add "%REGPATH%" /v "Manifest" /t REG_SZ /d "%SCRIPTDIR%OutlookPhishingReporter.dll.manifest" /f >nul 2>&1

echo [+] Add-in registry entry created
echo.

REM Prompt for admin email
:EMAILPROMPT
set ADMINEMAIL=
echo [REQUIRED] Configuration
echo.
set /p ADMINEMAIL="Enter the admin email address for phishing reports: "

if "!ADMINEMAIL!"=="" (
    echo.
    echo ERROR: Admin email is required for the add-in to function
    echo.
    goto EMAILPROMPT
)

REM Simple email validation
echo !ADMINEMAIL! | findstr /r "^[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]*\.[a-zA-Z]{2,}$" >nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Invalid email address format
    echo Please enter a valid email address (e.g., security@company.com)
    echo.
    goto EMAILPROMPT
)

REM Set the admin email in the registry
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "!ADMINEMAIL!" /f >nul 2>&1

echo.
echo [+] Admin email configured: !ADMINEMAIL!
echo.

REM Optional: Ask for custom icon
:ICONPROMPT
set CUSTOMIZE=
set /p CUSTOMIZE="[OPTIONAL] Do you want to set a custom icon? (Y/N): "

if /i "!CUSTOMIZE!"=="Y" (
    set CUSTOMICON=
    set /p CUSTOMICON="Enter the path to your custom icon (PNG/ICO): "
    
    if exist "!CUSTOMICON!" (
        reg add "%REGPATH%" /v "CustomIconPath" /t REG_SZ /d "!CUSTOMICON!" /f >nul 2>&1
        echo [+] Custom icon configured
    ) else (
        echo ERROR: Icon file not found at !CUSTOMICON!
    )
    echo.
)

REM Optional: Ask for custom button label
:LABELPROMPT
set CUSTOMLABEL=
set /p CUSTOMLABEL="[OPTIONAL] Do you want to customize the button label? (leave blank for default): "

if not "!CUSTOMLABEL!"=="" (
    reg add "%REGPATH%" /v "ButtonLabel" /t REG_SZ /d "!CUSTOMLABEL!" /f >nul 2>&1
    echo [+] Custom button label configured: !CUSTOMLABEL!
    echo.
)

echo.
echo ===================================================
echo Installation Complete!
echo ===================================================
echo.
echo [+] Add-in installed successfully
echo [+] Configuration saved to registry
echo.
echo IMPORTANT: Please restart Microsoft Outlook for the changes to take effect.
echo.
echo The "Report Phishing" button will appear in:
echo   - Mail ribbon (main mail view)
echo   - Message toolbar (when reading an email)
echo.
echo To uninstall: Delete the registry key at:
echo   %REGPATH%
echo.
pause
"@

$setupBatchPath = Join-Path $outputDir "setup.bat"
Set-Content -Path $setupBatchPath -Value $setupBatch -Encoding ASCII
Write-Host "[+] Created: setup.bat" -ForegroundColor Green

# Create WiX MSI template if requested
if ($BuildMSI) {
    Write-Host ""
    Write-Host "Creating MSI template..." -ForegroundColor Yellow

    $wxsContent = '<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product 
    Id="*"
    Name="Outlook Phishing Reporter"
    Language="1033"
    Version="1.0.0.0"
    Manufacturer="Your Organization"
    UpgradeCode="12345678-1234-1234-1234-123456789ABC">
    
    <Package 
      Id="*"
      Keywords="Outlook,Add-in,Phishing"
      Description="Outlook Phishing Reporter Add-in"
      InstallerVersion="200"
      Languages="1033"
      Compressed="yes"
      InstallScope="perUser" />

    <Media Id="1" Cabinet="OutlookPhishingReporter.cab" EmbedCab="yes" />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="LocalAppDataFolder">
        <Directory Id="INSTALLFOLDER" Name="OutlookPhishingReporter" />
      </Directory>
    </Directory>

    <Feature Id="ProductFeature" Title="Outlook Phishing Reporter" Level="1">
      <Description>Report phishing emails with one click</Description>
      <ComponentRef Id="ProductComponent" />
      <ComponentRef Id="RegistryComponent" />
    </Feature>

    <DirectoryRef Id="INSTALLFOLDER">
      <Component Id="ProductComponent" Guid="*">
        <File Id="DLLFile" Source="OutlookPhishingReporter.dll" KeyPath="yes" />
        <File Id="VSTOFile" Source="OutlookPhishingReporter.vsto" />
        <File Id="ManifestFile" Source="OutlookPhishingReporter.dll.manifest" />
      </Component>
    </DirectoryRef>

    <DirectoryRef Id="TARGETDIR">
      <Component Id="RegistryComponent" Guid="*">
        <RegistryKey Root="HKCU" Key="Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter">
          <RegistryValue Type="string" Name="Description" Value="Outlook Phishing Reporter" KeyPath="yes" />
          <RegistryValue Type="integer" Name="LoadBehavior" Value="3" />
          <RegistryValue Type="string" Name="Manifest" Value="[INSTALLFOLDER]OutlookPhishingReporter.dll.manifest" />
        </RegistryKey>
      </Component>
    </DirectoryRef>
  </Product>
</Wix>'

    $wxsPath = Join-Path $msiDir "Product.wxs"
    Set-Content -Path $wxsPath -Value $wxsContent -Encoding UTF8
    Write-Host "[+] Created: Product.wxs (WiX template)" -ForegroundColor Green

    # Create MSI build instructions
    $msiInstructions = @"
# Building MSI Installer for Outlook Phishing Reporter

## Prerequisites
1. Install WiX Toolset v3.11 or later
   - Download: https://github.com/wixtoolset/wix3/releases

2. Install Visual Studio 2019 or later (optional)

## Build Steps

1. Open Command Prompt in OutlookPhishingReporterSetup\MSISource
2. Run these commands:

```batch
REM Compile WiX source file
"C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" Product.wxs

REM Link to create MSI
"C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

## Deployment

Silent installation:
```batch
msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```

With log file:
```batch
msiexec /i OutlookPhishingReporter.msi /l*v install.log
```

## Group Policy Deployment
1. Place MSI on network share: \\SERVER\Software\OutlookPhishingReporter.msi
2. Create GPO for deployment
3. Configure registry preferences for AdminEmail settings

## Support
For issues, contact your IT department
"@

    $instructionsPath = Join-Path $msiDir "MSI_BUILD_INSTRUCTIONS.md"
    Set-Content -Path $instructionsPath -Value $msiInstructions -Encoding UTF8
    Write-Host "[+] Created: MSI_BUILD_INSTRUCTIONS.md" -ForegroundColor Green
}

# Create comprehensive installation guide
$installGuide = @"
# Outlook Phishing Reporter - Installation Guide

## Installation Methods

### Method 1: Batch File Setup (Recommended for End Users)
1. Right-click `setup.bat`
2. Select "Run as Administrator"
3. Enter your organization's security email address
4. Optionally configure custom icon and button label
5. Restart Microsoft Outlook

### Method 2: PowerShell Setup (For IT Administrators)
`powershell.exe -ExecutionPolicy Bypass -File Install.ps1`

### Method 3: Manual Registry Setup
Run as Administrator:

`powershell
`$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
New-Item -Path `$regPath -Force | Out-Null
New-ItemProperty -Path `$regPath -Name "AdminEmail" -Value "security@company.com" -PropertyType String -Force
New-ItemProperty -Path `$regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path `$regPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force
New-ItemProperty -Path `$regPath -Name "Manifest" -Value "C:\Path\To\OutlookPhishingReporter.dll.manifest" -PropertyType String -Force
`

### Method 4: MSI Setup (Enterprise)
`msiexec /i OutlookPhishingReporter.msi /quiet /norestart`

## System Requirements
- Windows 7 SP1 or later (Windows 10/11 recommended)
- Microsoft Outlook 2016 or later
- .NET Framework 4.7.2 or later
- Administrator privileges for installation

## Configuration

### Required Settings
- **AdminEmail**: Email address where phishing reports will be sent
  - Example: security@company.com
  - Must be a valid email address

### Optional Settings
- **CustomIconPath**: Path to custom icon file (PNG/ICO)
  - Example: C:\Company\Branding\icon.png
  - Overrides default icon

- **ButtonLabel**: Custom label for the ribbon button
  - Example: "Report Threat" or "Report Spam"
  - Default: "Report Phishing"

All settings are stored in the Windows Registry at:
`HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`

## Post-Installation Verification

1. **Check Registry Entries**
   - Run: `regedit`
   - Navigate to: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
   - Verify LoadBehavior = 3

2. **Verify in Outlook**
   - Open Microsoft Outlook
   - Look for "Report Phishing" button in ribbon
   - Should appear in both mail list and message reading views

3. **Test Functionality**
   - Select an email
   - Click "Report Phishing" button
   - Verify email forwards to AdminEmail
   - Check that original email moves to Deleted Items

## Troubleshooting

### Button doesn't appear in Outlook
1. Verify registry entries exist at: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
2. Ensure LoadBehavior = 3
3. Restart Outlook completely (close all windows)
4. Check log file: %TEMP%\OutlookPhishingReporter.log

### Configuration error message
1. Verify AdminEmail format is correct: user@domain.com
2. Re-run setup.bat and re-enter email
3. Check network connectivity

### Custom icon not loading
1. Verify file path is correct
2. Ensure file is valid PNG or ICO format
3. Verify user has read permissions on file
4. Check log file for details

## Uninstallation

### Via Registry (Recommended)
`powershell
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -Force -ErrorAction SilentlyContinue
`

Restart Outlook.

### Via Registry Editor
1. Run: regedit
2. Navigate to: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins
3. Delete: OutlookPhishingReporter folder
4. Restart Outlook

### Via MSI
`msiexec /x OutlookPhishingReporter.msi /quiet`

## Support and Documentation

For detailed information, see:
- README.md - Project overview
- QUICK_START.md - Quick installation steps
- MSI_BUILD_INSTRUCTIONS.md - Enterprise deployment

For issues, check: %TEMP%\OutlookPhishingReporter.log

## Version Information
- Version: 1.0
- Target: Outlook 2016+
- Framework: .NET 4.7.2
"@

$installGuidePath = Join-Path $outputDir "INSTALLATION_GUIDE.md"
Set-Content -Path $installGuidePath -Value $installGuide -Encoding UTF8
Write-Host "[+] Created: INSTALLATION_GUIDE.md" -ForegroundColor Green

# Create a summary
Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "Build Summary" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration: $Configuration" -ForegroundColor Yellow
Write-Host "Output Directory: $outputDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "Generated Files:" -ForegroundColor Yellow
Get-ChildItem -Path $outputDir | ForEach-Object {
    Write-Host "  [+] $($_.Name) ($([Math]::Round($_.Length / 1KB, 2)) KB)" -ForegroundColor Green
}
Write-Host ""
Write-Host "[+] Build completed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Installation Options:" -ForegroundColor Cyan
Write-Host "  1. EXE: Run setup.bat as Administrator" -ForegroundColor Cyan
Write-Host "  2. MSI: Build using WiX Toolset (if enabled)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Test: Run $outputDir\setup.bat" -ForegroundColor Yellow
Write-Host "  2. Review: INSTALLATION_GUIDE.md" -ForegroundColor Yellow

if ($BuildMSI) {
    Write-Host "  3. MSI: See $msiDir\MSI_BUILD_INSTRUCTIONS.md" -ForegroundColor Yellow
}

Write-Host ""
