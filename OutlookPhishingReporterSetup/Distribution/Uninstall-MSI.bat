@echo off
REM Uninstall Outlook Phishing Reporter MSI
REM This script uninstalls the MSI package

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo Outlook Phishing Reporter - MSI Uninstaller
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

REM Display instructions
echo.
echo This will uninstall the Outlook Phishing Reporter MSI package
echo.
echo Options:
echo   1. Uninstall with UI
echo   2. Silent uninstall (no prompts)
echo   3. Uninstall with log file
echo.

set /p OPTION="Choose an option (1-3): "

if "!OPTION!"=="1" (
    echo.
    echo [*] Starting MSI uninstallation with UI...
    msiexec /x "OutlookPhishingReporter.msi" /norestart
    if !errorlevel! equ 0 (
        echo [+] Uninstallation completed successfully
    ) else (
        echo [ERROR] Uninstallation failed
    )
    goto END
)

if "!OPTION!"=="2" (
    echo.
    echo [*] Starting silent MSI uninstallation...
    msiexec /x "OutlookPhishingReporter.msi" /quiet /norestart
    if !errorlevel! equ 0 (
        echo [+] Uninstallation completed successfully
    ) else (
        echo [ERROR] Uninstallation failed
    )
    goto END
)

if "!OPTION!"=="3" (
    echo.
    echo [*] Starting MSI uninstallation with logging...
    msiexec /x "OutlookPhishingReporter.msi" /norestart /l*v uninstall.log
    if !errorlevel! equ 0 (
        echo [+] Uninstallation completed successfully
        echo [+] Log file: uninstall.log
    ) else (
        echo [ERROR] Uninstallation failed
        echo [+] Check log file: uninstall.log
    )
    goto END
)

echo ERROR: Invalid option
goto END

:END
echo.
echo ===================================================
echo.
pause
