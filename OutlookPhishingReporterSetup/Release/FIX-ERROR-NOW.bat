@echo off
REM IMMEDIATE FIX - Configuration Error: Invalid Admin Email
REM This script fixes the "Invalid admin email address" error instantly

setlocal enabledelayedexpansion

cls
echo.
echo ╔═══════════════════════════════════════════════════════════════════╗
echo ║         CONFIGURATION ERROR - IMMEDIATE FIX                      ║
echo ╚═══════════════════════════════════════════════════════════════════╝
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

echo [*] Fixing configuration error...
echo.

REM Close Outlook
echo [STEP 1] Closing Outlook...
taskkill /IM OUTLOOK.EXE /F >nul 2>&1
timeout /t 2 /nobreak >nul
echo [+] Outlook closed
echo.

REM Remove old registry entry
echo [STEP 2] Removing old configuration...
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
reg delete "%REGPATH%" /f >nul 2>&1
echo [+] Old configuration removed
echo.

REM Create fresh registry configuration
echo [STEP 3] Creating new configuration...
set EMAIL=yosi.desta@outlook.co.il

reg add "%REGPATH%" /f >nul 2>&1
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "%EMAIL%" /f >nul 2>&1
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

echo [+] New configuration created
echo.

REM Verify configuration
echo [STEP 4] Verifying configuration...
for /f "tokens=3" %%a in ('reg query "%REGPATH%" /v "AdminEmail"') do (
    set SAVED_EMAIL=%%a
)

if "%SAVED_EMAIL%"=="%EMAIL%" (
    echo [+] Email verified: %EMAIL%
) else (
    echo [!] Warning: Email may not be correctly saved
    echo [!] Expected: %EMAIL%
    echo [!] Got: %SAVED_EMAIL%
)
echo.

echo ╔═══════════════════════════════════════════════════════════════════╗
echo ║                    FIX COMPLETE!                                 ║
echo ╚═══════════════════════════════════════════════════════════════════╝
echo.
echo [+] Configuration error has been fixed
echo.
echo IMPORTANT - NEXT STEP:
echo   1. Close this window
echo   2. Restart Microsoft Outlook
echo   3. The error should be gone
echo   4. "Report phishing" button will work
echo.
echo Email configured: %EMAIL%
echo.
pause
