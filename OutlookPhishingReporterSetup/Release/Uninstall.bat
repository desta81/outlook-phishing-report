@echo off
REM Outlook Phishing Reporter - Uninstallation Script
REM This script removes the add-in from Outlook

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo Outlook Phishing Reporter - Uninstaller
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

REM Confirm uninstallation
echo.
echo WARNING: This will remove the Outlook Phishing Reporter add-in
echo.
set /p CONFIRM="Are you sure you want to uninstall? (Y/N): "

if /i NOT "!CONFIRM!"=="Y" (
    echo.
    echo Uninstallation cancelled.
    echo.
    pause
    exit /b 0
)

REM Close Outlook
echo.
echo [*] Closing Outlook...
taskkill /IM OUTLOOK.EXE /F >nul 2>&1

REM Wait for Outlook to close
timeout /t 2 /nobreak >nul

REM Remove registry entry
echo [*] Removing registry entries...
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter

reg delete "%REGPATH%" /f >nul 2>&1

if errorlevel 1 (
    echo.
    echo [!] Warning: Could not delete registry key
    echo     The add-in may not have been installed
    echo.
) else (
    echo [+] Registry entries removed
)

REM Success message
echo.
echo ===================================================
echo Uninstallation Complete!
echo ===================================================
echo.
echo [+] The Outlook Phishing Reporter add-in has been removed.
echo.
echo You may safely delete the installation files if desired.
echo.
echo To reinstall, run: Install.vbs
echo.
pause
