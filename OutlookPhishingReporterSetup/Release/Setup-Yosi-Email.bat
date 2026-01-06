@echo off
REM ========================================================================
REM Setup Email Address for Outlook Phishing Reporter
REM Email: yosi.desta@outlook.co.il
REM ========================================================================

setlocal enabledelayedexpansion

cls
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║   SETUP EMAIL ADDRESS - YOSI.DESTA@OUTLOOK.CO.IL                 ║
echo ╚════════════════════════════════════════════════════════════════════╝
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

echo [+] Running as Administrator
echo.

REM Email to set
set EMAIL=yosi.desta@outlook.co.il
set REGISTRY_PATH=HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set VALUE_NAME=AdminEmail

echo [*] Setting up email address...
echo     Email: %EMAIL%
echo.

REM Add to registry
echo [*] Writing to registry...
reg add "%REGISTRY_PATH%" /v "%VALUE_NAME%" /d "%EMAIL%" /f

if %errorlevel% neq 0 (
    echo [ERROR] Failed to add registry value!
    echo.
    pause
    exit /b 1
)

echo [+] Email address added to registry!
echo.

REM Verify
echo [*] Verifying...
reg query "%REGISTRY_PATH%" /v "%VALUE_NAME%" >nul 2>&1

if %errorlevel% equ 0 (
    echo [+] Verification successful!
    echo.
    for /f "tokens=3" %%A in ('reg query "%REGISTRY_PATH%" /v "%VALUE_NAME%"') do (
        set SAVED_EMAIL=%%A
    )
    echo Registry Path: %REGISTRY_PATH%
    echo Value Name:   %VALUE_NAME%
    echo Value Data:   %SAVED_EMAIL%
    echo.
) else (
    echo [ERROR] Verification failed!
    echo.
    pause
    exit /b 1
)

echo ════════════════════════════════════════════════════════════════════
echo.
echo [+] EMAIL SETUP COMPLETE!
echo.
echo Next Steps:
echo   1. Close Outlook (if open)
echo   2. Close Visual Studio debug session (Shift + F5)
echo   3. Restart debug in Visual Studio (F5)
echo   4. Select an email
echo   5. Click 'Report phishing' button
echo   6. Email will be sent to: %EMAIL%
echo.
echo ════════════════════════════════════════════════════════════════════
echo.
pause
