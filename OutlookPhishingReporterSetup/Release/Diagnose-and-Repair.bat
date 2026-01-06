@echo off
REM Outlook Phishing Reporter - Diagnostic & Repair Tool
REM This script diagnoses and fixes configuration issues

setlocal enabledelayedexpansion

cls
echo.
echo ===================================================
echo Outlook Phishing Reporter - Diagnostic Tool
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

echo [*] Running diagnostics...
echo.

REM Registry path
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter

REM Check if registry key exists
echo [STEP 1] Checking registry key...
reg query "%REGPATH%" >nul 2>&1
if errorlevel 1 (
    echo [!] Registry key NOT FOUND
    echo.
    goto REPAIR
) else (
    echo [+] Registry key found
)

REM Check AdminEmail value
echo.
echo [STEP 2] Checking AdminEmail configuration...
reg query "%REGPATH%" /v "AdminEmail" >nul 2>&1
if errorlevel 1 (
    echo [!] AdminEmail value NOT SET
    echo.
    goto REPAIR
) else (
    echo [+] AdminEmail value found
    echo.
    for /f "tokens=3" %%a in ('reg query "%REGPATH%" /v "AdminEmail"') do (
        set ADMIN_EMAIL=%%a
        echo [*] Current AdminEmail: !ADMIN_EMAIL!
    )
)

REM Check LoadBehavior
echo.
echo [STEP 3] Checking LoadBehavior...
reg query "%REGPATH%" /v "LoadBehavior" >nul 2>&1
if errorlevel 1 (
    echo [!] LoadBehavior value NOT SET
    echo.
    goto REPAIR
) else (
    for /f "tokens=3" %%a in ('reg query "%REGPATH%" /v "LoadBehavior"') do (
        echo [+] LoadBehavior: %%a
    )
)

REM Check Description
echo.
echo [STEP 4] Checking Description...
reg query "%REGPATH%" /v "Description" >nul 2>&1
if errorlevel 1 (
    echo [!] Description value NOT SET
) else (
    for /f "tokens=3*" %%a in ('reg query "%REGPATH%" /v "Description"') do (
        echo [+] Description: %%a %%b
    )
)

REM Check Manifest
echo.
echo [STEP 5] Checking Manifest path...
reg query "%REGPATH%" /v "Manifest" >nul 2>&1
if errorlevel 1 (
    echo [!] Manifest value NOT SET
) else (
    for /f "tokens=3*" %%a in ('reg query "%REGPATH%" /v "Manifest"') do (
        echo [+] Manifest found
    )
)

echo.
echo ===================================================
echo Diagnostics Complete!
echo ===================================================
echo.
echo All configuration values are present and valid.
echo.
pause
exit /b 0

:REPAIR
echo.
echo ===================================================
echo Repair Wizard
echo ===================================================
echo.

REM Close Outlook
echo [*] Closing Outlook...
taskkill /IM OUTLOOK.EXE /F >nul 2>&1
timeout /t 2 /nobreak >nul

REM Delete existing registry entry
echo [*] Removing old configuration...
reg delete "%REGPATH%" /f >nul 2>&1

REM Prompt for admin email
:EMAIL_PROMPT
set ADMINEMAIL=
echo.
echo [REQUIRED] Configure Admin Email
echo.
set /p ADMINEMAIL="Enter admin email address (e.g., security@company.com): "

if "!ADMINEMAIL!"=="" (
    echo.
    echo ERROR: Email is required
    echo.
    goto EMAIL_PROMPT
)

REM Validate email format
echo !ADMINEMAIL! | findstr /r "^[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]*\.[a-zA-Z]{2,}$" >nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Invalid email format
    echo Please use format: user@domain.com
    echo.
    goto EMAIL_PROMPT
)

REM Create registry entries
echo.
echo [*] Creating registry configuration...

set SCRIPTDIR=%~dp0

reg add "%REGPATH%" /f >nul 2>&1
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "!ADMINEMAIL!" /f >nul 2>&1
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f >nul 2>&1

if exist "%SCRIPTDIR%OutlookPhishingReporter.dll.manifest" (
    reg add "%REGPATH%" /v "Manifest" /t REG_SZ /d "%SCRIPTDIR%OutlookPhishingReporter.dll.manifest" /f >nul 2>&1
)

echo [+] Configuration saved
echo.
echo ===================================================
echo Repair Complete!
echo ===================================================
echo.
echo [+] Admin Email: !ADMINEMAIL!
echo.
echo IMPORTANT: Please restart Microsoft Outlook now
echo.
echo The configuration error should be resolved.
echo.
pause
