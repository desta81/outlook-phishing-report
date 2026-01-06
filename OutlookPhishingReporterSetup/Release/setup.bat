@echo off
REM Outlook Phishing Reporter - Setup Script
REM This script installs the Outlook add-in with registry configuration

setlocal enabledelayedexpansion

echo.
echo ===================================================
echo Outlook Phishing Reporter Setup
echo ===================================================
echo.

REM Check for admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: This installer requires administrator privileges.
    echo Please right-click and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

REM Get the script directory
set SCRIPTDIR=%~dp0

REM Define registry paths
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set DLLPATH=%SCRIPTDIR%OutlookPhishingReporter.dll

echo [*] Installing Outlook Phishing Reporter add-in...
echo [*] Installation path: %SCRIPTDIR%
echo.

REM Check if required files exist
if not exist "%SCRIPTDIR%OutlookPhishingReporter.dll" (
    echo ERROR: OutlookPhishingReporter.dll not found!
    pause
    exit /b 1
)

REM Create the registry key for the add-in
reg add "%REGPATH%" /f >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to create registry key
    pause
    exit /b 1
)

REM Set LoadBehavior to 3 (load at startup)
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1

REM Set the description
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

REM Set the path to the manifest
reg add "%REGPATH%" /v "Manifest" /t REG_SZ /d "%SCRIPTDIR%OutlookPhishingReporter.dll.manifest" /f >nul 2>&1

echo [+] Add-in registry entry created
echo.

REM Prompt for admin email
:EMAILPROMPT
set ADMINEMAIL=
echo [REQUIRED] Configuration
echo.
set /p ADMINEMAIL="Enter the admin email address for phishing reports: "

if "!ADMINEMAIL!"=="" (
    echo.
    echo ERROR: Admin email is required for the add-in to function
    echo.
    goto EMAILPROMPT
)

REM Simple email validation
echo !ADMINEMAIL! | findstr /r "^[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]*\.[a-zA-Z]{2,}$" >nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Invalid email address format
    echo Please enter a valid email address (e.g., security@company.com)
    echo.
    goto EMAILPROMPT
)

REM Set the admin email in the registry
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "!ADMINEMAIL!" /f >nul 2>&1

echo.
echo [+] Admin email configured: !ADMINEMAIL!
echo.

REM Optional: Ask for custom icon
:ICONPROMPT
set CUSTOMIZE=
set /p CUSTOMIZE="[OPTIONAL] Do you want to set a custom icon? (Y/N): "

if /i "!CUSTOMIZE!"=="Y" (
    set CUSTOMICON=
    set /p CUSTOMICON="Enter the path to your custom icon (PNG/ICO): "
    
    if exist "!CUSTOMICON!" (
        reg add "%REGPATH%" /v "CustomIconPath" /t REG_SZ /d "!CUSTOMICON!" /f >nul 2>&1
        echo [+] Custom icon configured
    ) else (
        echo ERROR: Icon file not found at !CUSTOMICON!
    )
    echo.
)

REM Optional: Ask for custom button label
:LABELPROMPT
set CUSTOMLABEL=
set /p CUSTOMLABEL="[OPTIONAL] Do you want to customize the button label? (leave blank for default): "

if not "!CUSTOMLABEL!"=="" (
    reg add "%REGPATH%" /v "ButtonLabel" /t REG_SZ /d "!CUSTOMLABEL!" /f >nul 2>&1
    echo [+] Custom button label configured: !CUSTOMLABEL!
    echo.
)

echo.
echo ===================================================
echo Installation Complete!
echo ===================================================
echo.
echo [+] Add-in installed successfully
echo [+] Configuration saved to registry
echo.
echo IMPORTANT: Please restart Microsoft Outlook for the changes to take effect.
echo.
echo The "Report Phishing" button will appear in:
echo   - Mail ribbon (main mail view)
echo   - Message toolbar (when reading an email)
echo.
echo To uninstall: Delete the registry key at:
echo   %REGPATH%
echo.
pause
