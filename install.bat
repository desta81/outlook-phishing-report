@echo off
REM Outlook Phishing Reporter Installer
REM This script will launch the PowerShell installer with elevated privileges

setlocal enabledelayedexpansion

REM Check if running as Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~f0' -Verb RunAs"
    exit /b
)

echo.
echo ============================================
echo Outlook Phishing Reporter Installer
echo ============================================
echo.

REM Get script directory
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM Run PowerShell installer
echo Running installer...
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Install.ps1" -Action Install

if %errorLevel% equ 0 (
    echo.
    echo Installation completed successfully!
) else (
    echo.
    echo Installation failed with error code %errorLevel%
)

pause
