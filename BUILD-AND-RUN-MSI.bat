@echo off
REM ========================================================================
REM Complete MSI Build and Run Script
REM ========================================================================
REM This script builds the DLL, builds the MSI, and runs the installer

setlocal enabledelayedexpansion

cls
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║    OUTLOOK PHISHING REPORTER - BUILD AND RUN MSI INSTALLER       ║
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

REM Get workspace root
cd /d "%~dp0"
set ROOT_DIR=%CD%

echo [*] Workspace Root: %ROOT_DIR%
echo.

REM Step 1: Build DLL
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║ STEP 1: BUILD PLUGIN DLL                                          ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

echo [*] Building DLL...
call BUILD-DLL.bat

if %errorlevel% neq 0 (
    echo.
    echo ERROR: DLL build failed!
    echo.
    pause
    exit /b 1
)

echo [+] DLL build completed
echo.

REM Step 2: Verify DLL
echo [*] Verifying DLL...
if not exist "OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll" (
    echo ERROR: DLL not found!
    pause
    exit /b 1
)
echo [+] DLL verified
echo.

REM Step 3: Build MSI
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║ STEP 2: BUILD MSI INSTALLER                                       ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

echo [*] Building MSI...
cd /d "%ROOT_DIR%\OutlookPhishingReporterSetup\MSISource"
call build-msi.bat

if %errorlevel% neq 0 (
    echo.
    echo ERROR: MSI build failed!
    echo.
    pause
    exit /b 1
)

echo [+] MSI build completed
echo.

REM Step 4: Verify MSI
echo [*] Verifying MSI...
if not exist "%ROOT_DIR%\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi" (
    echo ERROR: MSI not found!
    pause
    exit /b 1
)

for %%A in ("%ROOT_DIR%\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi") do (
    set MSI_SIZE=%%~zA
)

echo [+] MSI verified
echo     Location: %ROOT_DIR%\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
echo     Size: %MSI_SIZE% bytes
echo.

REM Step 5: Ask to run installer
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║ STEP 3: RUN INSTALLER                                             ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo Ready to run the MSI installer!
echo.
echo Choose your installation type:
echo.
echo   1 = Normal Installation (with wizard)
echo   2 = Silent Installation (automatic, no interaction)
echo   3 = Skip Installation (just build MSI)
echo   4 = Exit
echo.

set /p CHOICE="Select option (1-4): "

if "%CHOICE%"=="1" (
    echo.
    echo [*] Running normal installation wizard...
    echo.
    cd /d "%ROOT_DIR%\OutlookPhishingReporterSetup\Distribution"
    start OutlookPhishingReporter.msi
    goto :COMPLETE
)

if "%CHOICE%"=="2" (
    echo.
    echo [*] Running silent installation...
    echo.
    cd /d "%ROOT_DIR%\OutlookPhishingReporterSetup\Distribution"
    msiexec /i OutlookPhishingReporter.msi /quiet /norestart
    echo.
    echo [+] Silent installation completed
    goto :COMPLETE
)

if "%CHOICE%"=="3" (
    echo.
    echo [*] Skipping installation
    goto :COMPLETE
)

if "%CHOICE%"=="4" (
    echo.
    echo [*] Exiting
    pause
    exit /b 0
)

echo ERROR: Invalid choice!
pause
exit /b 1

:COMPLETE
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    BUILD COMPLETE!                                ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
echo [+] DLL built successfully
echo [+] MSI created successfully
echo.
echo Location: %ROOT_DIR%\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
echo.
echo Next steps:
echo   1. If not installed, run: OutlookPhishingReporter.msi
echo   2. Restart Outlook
echo   3. Verify "Report phishing" button appears
echo   4. Test report functionality
echo.
echo For silent installation, use:
echo   msiexec /i OutlookPhishingReporter.msi /quiet /norestart
echo.
pause
