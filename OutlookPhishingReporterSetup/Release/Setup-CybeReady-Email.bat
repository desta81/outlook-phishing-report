@echo off
REM Outlook Phishing Reporter - Setup for report@phishcage.cybeready.net
REM This script configures the plugin with the phishing report email address

setlocal enabledelayedexpansion

cls
echo.
echo ╔═════════════════════════════════════════════════════════════════════╗
echo ║  Outlook Phishing Reporter - Setup Configuration                   ║
echo ║  Email: report@phishcage.cybeready.net                             ║
echo ╚═════════════════════════════════════════════════════════════════════╝
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
echo [+] Outlook closed
echo.

REM Registry path
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set ADMINEMAIL=report@phishcage.cybeready.net

REM Create registry entries
echo [*] Configuring registry...

reg add "%REGPATH%" /f >nul 2>&1
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "%ADMINEMAIL%" /f >nul 2>&1
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

echo [+] Configuration saved
echo.

REM Verify configuration
echo [*] Verifying configuration...
for /f "tokens=3" %%a in ('reg query "%REGPATH%" /v "AdminEmail"') do (
    set SAVED_EMAIL=%%a
)

if "%SAVED_EMAIL%"=="%ADMINEMAIL%" (
    echo [+] Email verified: %ADMINEMAIL%
) else (
    echo [!] Warning: Email may not be correctly saved
    echo [!] Expected: %ADMINEMAIL%
    echo [!] Got: %SAVED_EMAIL%
)
echo.

echo ╔═════════════════════════════════════════════════════════════════════╗
echo ║                    CONFIGURATION COMPLETE!                         ║
echo ╚═════════════════════════════════════════════════════════════════════╝
echo.
echo [+] Configuration error has been fixed
echo.
echo IMPORTANT - NEXT STEP:
echo   1. Close this window
echo   2. Restart Microsoft Outlook
echo   3. The configuration will be applied
echo   4. "Report phishing" button will work
echo.
echo Email configured: %ADMINEMAIL%
echo All reports will be sent to: %ADMINEMAIL%
echo.
pause
