@echo off
REM ========================================================================
REM Outlook Phishing Reporter - Complete Test Suite
REM ========================================================================
REM This script tests all aspects of the plugin installation and functionality
REM

setlocal enabledelayedexpansion

cls
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║   OUTLOOK PHISHING REPORTER - COMPLETE TEST SUITE                 ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

REM Test execution tracking
set TESTS_PASSED=0
set TESTS_FAILED=0
set TESTS_TOTAL=0

REM =======================================================================
REM TEST 1: Check Administrator Privileges
REM =======================================================================
echo [TEST 1/10] Checking Administrator Privileges...
set TESTS_TOTAL=!TESTS_TOTAL!+1
net session >nul 2>&1
if %errorlevel% equ 0 (
    echo [PASS] Administrator privileges confirmed
    set TESTS_PASSED=!TESTS_PASSED!+1
) else (
    echo [FAIL] Administrator privileges required
    echo        Please run as Administrator
    set TESTS_FAILED=!TESTS_FAILED!+1
    goto SKIP_TESTS
)
echo.

REM =======================================================================
REM TEST 2: Check .NET Framework
REM =======================================================================
echo [TEST 2/10] Checking .NET Framework 4.7.2...
set TESTS_TOTAL=!TESTS_TOTAL!+1
reg query "HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=3" %%a in ('reg query "HKLM\Software\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release') do (
        set RELEASE=%%a
    )
    if !RELEASE! geq 461808 (
        echo [PASS] .NET Framework 4.7.2 or later detected
        set TESTS_PASSED=!TESTS_PASSED!+1
    ) else (
        echo [WARN] .NET Framework older than 4.7.2
        echo        Recommend upgrading to .NET Framework 4.7.2+
    )
) else (
    echo [FAIL] .NET Framework not found
    set TESTS_FAILED=!TESTS_FAILED!+1
)
echo.

REM =======================================================================
REM TEST 3: Check Outlook Installation
REM =======================================================================
echo [TEST 3/10] Checking Microsoft Outlook Installation...
set TESTS_TOTAL=!TESTS_TOTAL!+1
reg query "HKLM\Software\Microsoft\Office" >nul 2>&1
if %errorlevel% equ 0 (
    echo [PASS] Microsoft Office detected
    set TESTS_PASSED=!TESTS_PASSED!+1
) else (
    echo [FAIL] Microsoft Office not found
    set TESTS_FAILED=!TESTS_FAILED!+1
)
echo.

REM =======================================================================
REM TEST 4: Check Plugin Registry Configuration
REM =======================================================================
echo [TEST 4/10] Checking Plugin Registry Configuration...
set TESTS_TOTAL=!TESTS_TOTAL!+1
set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
reg query "!REGPATH!" >nul 2>&1
if %errorlevel% equ 0 (
    echo [PASS] Plugin registry key found
    set TESTS_PASSED=!TESTS_PASSED!+1
    
    REM Check AdminEmail value
    echo   [SUBTEST] Checking AdminEmail value...
    for /f "tokens=3" %%a in ('reg query "!REGPATH!" /v "AdminEmail"') do (
        set ADMIN_EMAIL=%%a
    )
    if not "!ADMIN_EMAIL!"=="" (
        echo     [PASS] AdminEmail: !ADMIN_EMAIL!
    ) else (
        echo     [WARN] AdminEmail not configured
    )
) else (
    echo [FAIL] Plugin registry key not found
    echo        Plugin may not be installed
    set TESTS_FAILED=!TESTS_FAILED!+1
)
echo.

REM =======================================================================
REM TEST 5: Check LoadBehavior Setting
REM =======================================================================
echo [TEST 5/10] Checking LoadBehavior Setting...
set TESTS_TOTAL=!TESTS_TOTAL!+1
for /f "tokens=3" %%a in ('reg query "!REGPATH!" /v "LoadBehavior"') do (
    set LOAD_BEHAVIOR=%%a
)
if "!LOAD_BEHAVIOR!"=="3" (
    echo [PASS] LoadBehavior = 3 ^(Auto-load enabled^)
    set TESTS_PASSED=!TESTS_PASSED!+1
) else if "!LOAD_BEHAVIOR!"=="" (
    echo [FAIL] LoadBehavior not set
    set TESTS_FAILED=!TESTS_FAILED!+1
) else (
    echo [WARN] LoadBehavior = !LOAD_BEHAVIOR! ^(expected 3^)
)
echo.

REM =======================================================================
REM TEST 6: Check Outlook Process
REM =======================================================================
echo [TEST 6/10] Checking Outlook Process...
set TESTS_TOTAL=!TESTS_TOTAL!+1
tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>NUL | find /I /N "OUTLOOK.EXE">NUL
if %errorlevel% equ 0 (
    echo [INFO] Outlook is currently running
    echo        Some tests cannot be performed while Outlook is running
) else (
    echo [PASS] Outlook is not running
    echo        Good for full configuration testing
    set TESTS_PASSED=!TESTS_PASSED!+1
)
echo.

REM =======================================================================
REM TEST 7: Check Plugin DLL
REM =======================================================================
echo [TEST 7/10] Checking Plugin DLL...
set TESTS_TOTAL=!TESTS_TOTAL!+1
if exist "%APPDATA%\Microsoft\Addins\OutlookPhishingReporter.dll" (
    echo [PASS] Plugin DLL found
    set TESTS_PASSED=!TESTS_PASSED!+1
) else (
    echo [INFO] DLL not in standard location
    echo        Plugin may be installed elsewhere or not installed
)
echo.

REM =======================================================================
REM TEST 8: Check Log Directory
REM =======================================================================
echo [TEST 8/10] Checking Log Directory...
set TESTS_TOTAL=!TESTS_TOTAL!+1
set LOG_DIR=%LOCALAPPDATA%\OutlookPhishingReporter\Logs
if exist "!LOG_DIR!" (
    echo [PASS] Log directory exists
    echo        Location: !LOG_DIR!
    set TESTS_PASSED=!TESTS_PASSED!+1
    
    REM Count log files
    for /f %%a in ('dir /b "!LOG_DIR!\*.log" 2^>nul ^| find /c /v ""') do set LOG_COUNT=%%a
    if "!LOG_COUNT!"=="" set LOG_COUNT=0
    echo        Log files: !LOG_COUNT!
) else (
    echo [INFO] Log directory not yet created
    echo        Will be created on first use
)
echo.

REM =======================================================================
REM TEST 9: Check Email Configuration
REM =======================================================================
echo [TEST 9/10] Checking Email Configuration...
set TESTS_TOTAL=!TESTS_TOTAL!+1
if not "!ADMIN_EMAIL!"=="" (
    echo [PASS] Admin email configured: !ADMIN_EMAIL!
    set TESTS_PASSED=!TESTS_PASSED!+1
    
    REM Basic email validation
    echo !ADMIN_EMAIL! | findstr /r "^[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]*\.[a-zA-Z]{2,}$" >nul
    if %errorlevel% equ 0 (
        echo [PASS] Email format is valid
    ) else (
        echo [WARN] Email format may be invalid
    )
) else (
    echo [FAIL] Admin email not configured
    echo        Need to run: FIX-ERROR-NOW.bat
    set TESTS_FAILED=!TESTS_FAILED!+1
)
echo.

REM =======================================================================
REM TEST 10: Check Plugin Manifest
REM =======================================================================
echo [TEST 10/10] Checking Plugin Manifest...
set TESTS_TOTAL=!TESTS_TOTAL!+1
for /f "tokens=3*" %%a in ('reg query "!REGPATH!" /v "Manifest"') do (
    set MANIFEST=%%a %%b
)
if not "!MANIFEST!"=="" (
    echo [PASS] Plugin manifest configured
    echo        Path: !MANIFEST!
    set TESTS_PASSED=!TESTS_PASSED!+1
    
    if exist "!MANIFEST!" (
        echo [PASS] Manifest file exists
    ) else (
        echo [WARN] Manifest file not found at configured path
    )
) else (
    echo [INFO] Manifest not configured in registry
)
echo.

:SKIP_TESTS
REM =======================================================================
REM TEST SUMMARY
REM =======================================================================
echo.
echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                          TEST SUMMARY                             ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.

setlocal enabledelayedexpansion
set /a TOTAL=!TESTS_PASSED!+!TESTS_FAILED!

if !TESTS_PASSED! gtr 0 (
    echo [PASS] Tests Passed: !TESTS_PASSED!
)
if !TESTS_FAILED! gtr 0 (
    echo [FAIL] Tests Failed: !TESTS_FAILED!
)
echo [INFO] Tests Total:  !TOTAL!
echo.

REM =======================================================================
REM RESULTS AND RECOMMENDATIONS
REM =======================================================================
if !TESTS_FAILED! equ 0 (
    echo ✓ ALL TESTS PASSED - Plugin appears to be correctly installed
    echo.
    echo Recommendations:
    echo   1. Open Outlook
    echo   2. Look for "Report phishing" button in ribbon
    echo   3. Select a test email
    echo   4. Click the button to verify it works
    echo   5. Check that email is reported successfully
    echo.
) else (
    echo ✗ SOME TESTS FAILED - Please address the issues above
    echo.
    echo Troubleshooting Steps:
    echo   1. Run FIX-ERROR-NOW.bat to fix configuration
    echo   2. Verify .NET Framework 4.7.2 is installed
    echo   3. Verify Outlook is properly installed
    echo   4. Restart your computer
    echo   5. Try reinstalling the plugin
    echo.
)

echo ╔════════════════════════════════════════════════════════════════════╗
echo ║                    END OF TEST SUITE                              ║
echo ╚════════════════════════════════════════════════════════════════════╝
echo.
pause
