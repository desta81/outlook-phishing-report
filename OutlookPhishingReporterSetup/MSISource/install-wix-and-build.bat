@echo off
REM Install WiX Toolset and Build MSI - Automated Setup
REM This script downloads and installs WiX Toolset, then builds the MSI

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo WiX Toolset Installer & MSI Builder
echo ===================================================
echo.

REM Check admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script requires administrator privileges.
    echo Please right-click and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

REM Check if WiX is already installed
echo [*] Checking for WiX Toolset installation...
if exist "C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" (
    echo [+] WiX v3.14 found - skipping installation
    set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.14\bin
    goto BUILD_MSI
)

if exist "C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" (
    echo [+] WiX v3.11 found - skipping installation
    set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.11\bin
    goto BUILD_MSI
)

REM WiX not found - provide installation instructions
echo.
echo [!] WiX Toolset not found
echo.
echo To build the MSI, you need to install WiX Toolset:
echo.
echo 1. Download from: https://wixtoolset.org/
echo    (Look for "WiX Toolset v3.14" or latest v3.x release)
echo.
echo 2. Run the installer when download completes
echo.
echo 3. After installation, restart your computer
echo.
echo 4. Run this script again (or run build-msi.bat)
echo.
echo.
echo Alternatively, download directly from GitHub:
echo   https://github.com/wixtoolset/wix3/releases
echo.
pause
exit /b 1

:BUILD_MSI
echo.
echo ===================================================
echo Building MSI Installer
echo ===================================================
echo.

set PROJECT_DIR=%~dp0
set OUTPUT_DIR=%PROJECT_DIR%..\Distribution

REM Create output directory if needed
if not exist "%OUTPUT_DIR%" (
    echo [*] Creating output directory...
    mkdir "%OUTPUT_DIR%"
)

REM Verify Product.wxs exists
if not exist "%PROJECT_DIR%Product.wxs" (
    echo [ERROR] Product.wxs not found!
    pause
    exit /b 1
)

REM Clean previous artifacts
if exist "%PROJECT_DIR%Product.wixobj" del "%PROJECT_DIR%Product.wixobj"
if exist "%PROJECT_DIR%Product.wixpdb" del "%PROJECT_DIR%Product.wixpdb"

REM Build MSI
echo [1/2] Compiling WiX source...
cd /d "%PROJECT_DIR%"
"%WIX_PATH%\candle.exe" Product.wxs -o Product.wixobj

if errorlevel 1 (
    echo.
    echo [ERROR] Compilation failed!
    pause
    exit /b 1
)

echo [+] Compilation successful
echo.
echo [2/2] Linking and creating MSI...
"%WIX_PATH%\light.exe" -out "%OUTPUT_DIR%\OutlookPhishingReporter.msi" Product.wixobj

if errorlevel 1 (
    echo.
    echo [ERROR] MSI creation failed!
    pause
    exit /b 1
)

REM Verify MSI was created
if not exist "%OUTPUT_DIR%\OutlookPhishingReporter.msi" (
    echo [ERROR] MSI file was not created!
    pause
    exit /b 1
)

echo [+] MSI created successfully
echo.
echo ===================================================
echo Build Complete!
echo ===================================================
echo.
echo [+] MSI File: %OUTPUT_DIR%\OutlookPhishingReporter.msi
echo.
echo Next Steps:
echo   1. Test: msiexec /i "%OUTPUT_DIR%\OutlookPhishingReporter.msi"
echo   2. Deploy: Copy to network share or use Group Policy
echo   3. Reference: See MSI_BUILD_INSTRUCTIONS.md for deployment
echo.
pause
