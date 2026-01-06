@echo off
REM ========================================================================
REM Outlook Phishing Reporter - Build DLL File
REM ========================================================================
REM This script builds the plugin DLL using MSBuild

setlocal enabledelayedexpansion

cls
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║   OUTLOOK PHISHING REPORTER - BUILD DLL                           ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

REM Get the project directory
cd /d "%~dp0"

echo [*] Workspace Directory: %CD%
echo.

REM Check if project file exists
if not exist "OutlookPhishingReporter\OutlookPhishingReporter.csproj" (
    echo ERROR: Project file not found!
    echo Expected: OutlookPhishingReporter\OutlookPhishingReporter.csproj
    echo.
    pause
    exit /b 1
)

echo [*] Building Release Configuration...
echo.

REM Build using msbuild
msbuild OutlookPhishingReporter\OutlookPhishingReporter.csproj /p:Configuration=Release /p:Platform="AnyCPU"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed!
    echo.
    echo Troubleshooting:
    echo   1. Ensure Visual Studio is installed
    echo   2. Check for missing references
    echo   3. Verify .NET Framework 4.7.2 is installed
    echo   4. Run from Visual Studio Developer Command Prompt
    echo.
    pause
    exit /b 1
)

echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                      BUILD SUCCESSFUL!                            ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

REM Verify DLL exists
set DLL_PATH=OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll

if exist "%DLL_PATH%" (
    for %%A in ("%DLL_PATH%") do (
        set DLL_SIZE=%%~zA
    )
    echo [+] DLL File Created Successfully
    echo.
    echo Location: %DLL_PATH%
    echo Size:     %DLL_SIZE% bytes
    echo.
    echo ╔════════════════════════════════════════════════════════════════════╗
    echo ║                       NEXT STEPS                                  ║
    echo ╚════════════════════════════════════════════════════════════════════╝
    echo.
    echo 1. Build MSI Installer (Optional):
    echo    Navigate to: OutlookPhishingReporterSetup\MSISource\
    echo    Run: build-msi.bat
    echo.
    echo 2. Test the Plugin:
    echo    Run: RUN-TEST-SUITE.bat
    echo.
    echo 3. Install Directly:
    echo    Copy DLL to: %%APPDATA%%\Microsoft\Addins\
    echo    Restart Outlook
    echo.
    echo DLL Ready for:
    echo   - MSI packaging
    echo   - Direct installation
    echo   - Enterprise deployment
    echo   - Testing
    echo.
) else (
    echo [ERROR] DLL file not found!
    echo Expected: %DLL_PATH%
    echo.
)

pause
