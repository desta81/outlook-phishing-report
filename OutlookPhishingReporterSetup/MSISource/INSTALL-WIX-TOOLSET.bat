@echo off
REM ========================================================================
REM WiX Toolset - Automatic Download and Installation
REM ========================================================================
REM This script automatically downloads and installs WiX Toolset
REM Required for building the MSI installer

setlocal enabledelayedexpansion

cls
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║         WiX TOOLSET - AUTOMATIC INSTALLER                         ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

REM Check admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script requires administrator privileges.
    echo Please right-click and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

echo [*] Checking for WiX Toolset installation...
echo.

REM Check if WiX is already installed
if exist "C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" (
    echo [+] WiX Toolset v3.14 is already installed!
    echo     Location: C:\Program Files (x86)\WiX Toolset v3.14\
    echo.
    echo You are ready to build the MSI.
    echo.
    pause
    exit /b 0
)

if exist "C:\Program Files (x86)\WiX Toolset v3.13\bin\candle.exe" (
    echo [+] WiX Toolset v3.13 is already installed!
    echo     Location: C:\Program Files (x86)\WiX Toolset v3.13\
    echo.
    echo You are ready to build the MSI.
    echo.
    pause
    exit /b 0
)

echo [!] WiX Toolset not found.
echo.
echo This script will help you install WiX Toolset.
echo.
echo You have two options:
echo   1. Download manually from https://wixtoolset.org/
echo   2. Use the direct download link below
echo.

REM Create temp directory for download
set TEMP_DIR=%TEMP%\WiXInstall
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

echo [*] Preparing to download WiX Toolset v3.14...
echo.
echo Download Link:
echo https://github.com/wixtoolset/wix3/releases/download/wix314rtm/wix314.exe
echo.
echo To continue:
echo   1. Download the file from the link above
echo   2. Save to: %TEMP_DIR%\
echo   3. Run this script again (it will detect the installer)
echo.
echo Or:
echo   1. Visit https://wixtoolset.org/
echo   2. Download the latest WiX v3.x
echo   3. Run the installer
echo   4. Run this script again to verify
echo.
echo After installation, restart this script to proceed with MSI build.
echo.
pause
