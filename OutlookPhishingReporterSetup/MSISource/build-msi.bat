@echo off
REM Build MSI Installer for Outlook Phishing Reporter
REM This script automates the WiX compilation and linking process

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo Outlook Phishing Reporter - MSI Build Script
echo ===================================================
echo.

REM Set working directory
set PROJECT_DIR=%~dp0
set OUTPUT_DIR=%PROJECT_DIR%..\Distribution
set PRODUCT_WXS=%PROJECT_DIR%Product.wxs

echo [*] Working Directory: %PROJECT_DIR%
echo [*] Output Directory: %OUTPUT_DIR%
echo.

REM Create output directory if needed
if not exist "%OUTPUT_DIR%" (
    echo [*] Creating output directory...
    mkdir "%OUTPUT_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create output directory
        pause
        exit /b 1
    )
    echo [+] Output directory created
)

REM Find WiX installation
echo [*] Searching for WiX Toolset...
set WIX_PATH=

REM Check common installation paths
if exist "C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" (
    set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.14\bin
    echo [+] Found WiX v3.14
) else if exist "C:\Program Files (x86)\WiX Toolset v3.13\bin\candle.exe" (
    set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.13\bin
    echo [+] Found WiX v3.13
) else if exist "C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" (
    set WIX_PATH=C:\Program Files (x86)\WiX Toolset v3.11\bin
    echo [+] Found WiX v3.11
) else if exist "C:\Program Files\WiX Toolset v3.14\bin\candle.exe" (
    set WIX_PATH=C:\Program Files\WiX Toolset v3.14\bin
    echo [+] Found WiX v3.14 ^(x64^)
) else (
    echo.
    echo ERROR: WiX Toolset not found!
    echo.
    echo Please install WiX Toolset from:
    echo   https://wixtoolset.org/
    echo.
    echo Or download from GitHub:
    echo   https://github.com/wixtoolset/wix3/releases
    echo.
    pause
    exit /b 1
)

echo [+] WiX Path: %WIX_PATH%
echo.

REM Verify Product.wxs exists
if not exist "%PRODUCT_WXS%" (
    echo ERROR: Product.wxs not found at: %PRODUCT_WXS%
    pause
    exit /b 1
)

REM Clean previous build artifacts
echo [*] Cleaning previous build artifacts...
if exist "%PROJECT_DIR%Product.wixobj" del "%PROJECT_DIR%Product.wixobj"
if exist "%PROJECT_DIR%Product.wixpdb" del "%PROJECT_DIR%Product.wixpdb"
echo [+] Cleaned

echo.
echo ===================================================
echo Building MSI Installer
echo ===================================================
echo.

REM Step 1: Compile WiX source
echo [1/2] Compiling WiX source file...
cd /d "%PROJECT_DIR%"
"%WIX_PATH%\candle.exe" Product.wxs -o Product.wixobj

if errorlevel 1 (
    echo.
    echo [ERROR] Compilation failed!
    echo.
    echo Check the error messages above for details.
    echo.
    pause
    exit /b 1
)

echo [+] Compilation successful
echo.

REM Step 2: Link to create MSI
echo [2/2] Linking and creating MSI...
"%WIX_PATH%\light.exe" -out "%OUTPUT_DIR%\OutlookPhishingReporter.msi" Product.wixobj -v

if errorlevel 1 (
    echo.
    echo [ERROR] MSI creation failed!
    echo.
    echo Check the error messages above for details.
    echo.
    pause
    exit /b 1
)

echo [+] Linking successful
echo.

REM Verify MSI was created
if not exist "%OUTPUT_DIR%\OutlookPhishingReporter.msi" (
    echo ERROR: MSI file was not created!
    pause
    exit /b 1
)

REM Get file size
for %%A in ("%OUTPUT_DIR%\OutlookPhishingReporter.msi") do (
    set MSI_SIZE=%%~zA
)

echo ===================================================
echo Build Complete!
echo ===================================================
echo.
echo [+] MSI File Created Successfully
echo.
echo Location: %OUTPUT_DIR%\OutlookPhishingReporter.msi
echo Size: %MSI_SIZE% bytes
echo.
echo ===================================================
echo Next Steps
echo ===================================================
echo.
echo 1. Test Installation on a test machine:
echo    msiexec /i OutlookPhishingReporter.msi
echo.
echo 2. Silent Installation (for enterprise):
echo    msiexec /i OutlookPhishingReporter.msi /quiet /norestart
echo.
echo 3. Deployment via Group Policy:
echo    See: MSI_BUILD_INSTRUCTIONS.md
echo.
echo ===================================================
echo.
pause
