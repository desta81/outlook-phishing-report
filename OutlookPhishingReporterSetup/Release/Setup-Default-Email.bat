@echo off
REM Outlook Phishing Reporter - Setup for yosi.desta@outlook.co.il
REM This script configures the plugin with the security email address

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo Outlook Phishing Reporter - Setup Configuration
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

REM Close Outlook
echo [*] Closing Outlook...
taskkill /IM OUTLOOK.EXE /F >nul 2>&1
timeout /t 2 /nobreak >nul

REM Registry path
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set ADMINEMAIL=yosi.desta@outlook.co.il

REM Create registry entries
echo [*] Configuring registry...

reg add "%REGPATH%" /f >nul 2>&1
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "%ADMINEMAIL%" /f >nul 2>&1
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

echo.
echo ===================================================
echo Configuration Complete!
echo ===================================================
echo.
echo [+] Admin Email Configured: %ADMINEMAIL%
echo.
echo Next Steps:
echo   1. Restart Microsoft Outlook
echo   2. You will see "Report phishing" button in toolbar
echo   3. Click button to report suspicious emails
echo.
echo All phishing reports will be sent to: %ADMINEMAIL%
echo.
pause
