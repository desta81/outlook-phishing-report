# Build MSI Installer for Outlook Phishing Reporter

## Prerequisites

### Step 1: Install WiX Toolset

WiX (Windows Installer XML) is required to build the MSI installer.

**Option A: Download from WiX Official Website**
1. Visit: https://wixtoolset.org/
2. Download WiX Toolset v3.14 (or latest v3.x version)
3. Run the installer
4. Complete installation
5. Restart your computer if prompted

**Option B: Download from GitHub Releases**
1. Visit: https://github.com/wixtoolset/wix3/releases
2. Download the latest WiX v3 installer (e.g., `wix314.exe`)
3. Run the installer
4. Follow setup wizard
5. Restart computer

**Option C: Install via Package Manager (Advanced)**

Using Chocolatey:
```
choco install wixtoolset
```

### Step 2: Verify Installation

After installation, verify WiX Toolset was installed correctly:

```cmd
candle.exe -version
light.exe -version
```

Both commands should return version information.

---

## Building the MSI

### Step 1: Open Command Prompt

1. Press `Win + R`
2. Type: `cmd`
3. Press Enter

### Step 2: Navigate to MSI Source Directory

```cmd
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource"
```

### Step 3: Build the MSI

Run these commands in order:

**Step 3a: Compile WiX Source**
```cmd
"C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" Product.wxs
```

This creates `Product.wixobj` file.

**Step 3b: Link to Create MSI**
```cmd
"C:\Program Files (x86)\WiX Toolset v3.14\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

This creates the final `OutlookPhishingReporter.msi` file.

### Expected Output

If successful, you should see:
```
WiX Toolset Linker version X.X.X.X
Copyright (c) .NET Foundation and contributors. All rights reserved.

[MSI] Updating summary information...
[MSI] Creating media cabinet file...
[MSI] Committing media cabinet file...
[MSI] Successfully created OutlookPhishingReporter.msi
```

---

## Troubleshooting

### Error: "candle.exe not found"

**Cause:** WiX Toolset is not installed or not in PATH

**Solution:**
1. Install WiX Toolset from https://wixtoolset.org/
2. Use full path to candle.exe: `C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe`

### Error: "Product.wxs not found"

**Cause:** You're in the wrong directory

**Solution:**
```cmd
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource"
```

### Error: "Distribution folder not found"

**Cause:** Distribution folder doesn't exist

**Solution:**
```cmd
mkdir ..\Distribution
```

Then run the `light.exe` command again.

### Error: "Product.wixobj not found"

**Cause:** The candle.exe compilation failed

**Solution:**
1. Check for errors in the candle.exe output
2. Verify Product.wxs file exists
3. Ensure you're in the MSISource directory

---

## Quick Build Script

Instead of running commands manually, you can use this batch script:

**CreateFile:** `build-msi.bat` in the MSISource directory

```batch
@echo off
setlocal

REM Build MSI Installer for Outlook Phishing Reporter
REM This script compiles the WiX source and creates the MSI

cls
echo.
echo ===================================================
echo Building Outlook Phishing Reporter MSI
echo ===================================================
echo.

REM Set WiX path (adjust version if different)
set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.14\bin
set PROJECT_DIR=%~dp0
set OUTPUT_DIR=%PROJECT_DIR%..\Distribution

REM Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" (
    echo Creating output directory...
    mkdir "%OUTPUT_DIR%"
)

REM Verify WiX tools exist
if not exist "%WIX_PATH%\candle.exe" (
    echo.
    echo ERROR: WiX Toolset not found at: %WIX_PATH%
    echo.
    echo Please install WiX Toolset from: https://wixtoolset.org/
    echo.
    pause
    exit /b 1
)

REM Compile WiX source
echo [1/2] Compiling WiX source file...
cd /d "%PROJECT_DIR%"
"%WIX_PATH%\candle.exe" Product.wxs

if errorlevel 1 (
    echo.
    echo ERROR: Compilation failed!
    pause
    exit /b 1
)

REM Link to create MSI
echo [2/2] Linking and creating MSI...
"%WIX_PATH%\light.exe" -out "%OUTPUT_DIR%\OutlookPhishingReporter.msi" Product.wixobj

if errorlevel 1 (
    echo.
    echo ERROR: MSI creation failed!
    pause
    exit /b 1
)

REM Success
echo.
echo ===================================================
echo Build Complete!
echo ===================================================
echo.
echo MSI File: %OUTPUT_DIR%\OutlookPhishingReporter.msi
echo.
echo Next Steps:
echo   1. Test the MSI on a test machine
echo   2. Deploy via Group Policy or SCCM
echo   3. See: MSI_BUILD_INSTRUCTIONS.md for deployment
echo.
pause
```

---

## Using the Built MSI

### Silent Installation
```cmd
msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```

### Installation with User Interface
```cmd
msiexec /i OutlookPhishingReporter.msi
```

### Installation with Log
```cmd
msiexec /i OutlookPhishingReporter.msi /l*v install.log
```

### Uninstall
```cmd
msiexec /x OutlookPhishingReporter.msi /quiet
```

---

## MSI Customization

If you need to customize the MSI, edit `Product.wxs`:

### Change Product Name
```xml
<Product 
    Name="Custom Product Name"
    ...>
```

### Change Manufacturer
```xml
<Product 
    Manufacturer="Your Company Name"
    ...>
```

### Change Installation Path
```xml
<Directory Id="INSTALLFOLDER" Name="CustomFolderName" />
```

After editing, rebuild the MSI using the same steps above.

---

## Enterprise Deployment

### Group Policy Deployment

1. **Place MSI on Network Share:**
   ```
   \\SERVER\Software\OutlookPhishingReporter\OutlookPhishingReporter.msi
   ```

2. **Create Group Policy:**
   - Open `gpedit.msc` on domain controller
   - Navigate to: Computer Configuration ? Software Settings ? Software Installation
   - Right-click ? New ? Package
   - Select the MSI from network share
   - Choose deployment type:
     - **Assigned:** Install automatically
     - **Published:** Available for user installation

3. **Configure Registry via Group Policy:**
   - Use Group Policy Preferences
   - Set registry values for AdminEmail, etc.
   - Apply to target OUs

### SCCM Deployment

1. **Create Application:**
   - New Application
   - Add deployment type: MSI
   - Set installation command: `msiexec /i OutlookPhishingReporter.msi /quiet`

2. **Configure Detection:**
   - Registry path: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
   - Verify key exists

3. **Deploy:**
   - Create deployment to device collection
   - Set deadline and enforcement

---

## Next Steps

1. **Install WiX Toolset** (if not already installed)
2. **Build the MSI** using the commands above
3. **Test the MSI** on a test machine
4. **Deploy** to your organization using Group Policy or SCCM

---

## Support

For issues:
1. Check WiX Toolset is installed correctly
2. Verify you're in the correct directory
3. Review the build script output for errors
4. See MSI_BUILD_INSTRUCTIONS.md for additional help

---

**Status:** Ready to build! Install WiX Toolset and follow the steps above.
