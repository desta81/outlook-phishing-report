@echo off
REM Manual Installation Script for Outlook Phishing Reporter
REM This script installs the add-in using the built files

setlocal

echo.
echo ================================================
echo   Outlook Phishing Reporter - Manual Install
echo ================================================
echo.

REM Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo Installing Outlook Phishing Reporter Add-in...
echo.

REM Get paths
set SOURCE_DIR=%~dp0OutlookPhishingReporter\bin\Release
set INSTALL_DIR=%ProgramFiles%\OutlookPhishingReporter

REM Create installation directory
echo Creating installation directory...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy files
echo Copying add-in files...
xcopy /Y /I /Q "%SOURCE_DIR%\*.dll" "%INSTALL_DIR%\"
xcopy /Y /I /Q "%SOURCE_DIR%\*.vsto" "%INSTALL_DIR%\"
xcopy /Y /I /Q "%SOURCE_DIR%\*.manifest" "%INSTALL_DIR%\"

REM Configure registry
echo Configuring registry...
powershell -Command "$regPath = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter'; if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }; New-ItemProperty -Path $regPath -Name 'Description' -Value 'Report phishing emails to security team' -PropertyType String -Force | Out-Null; New-ItemProperty -Path $regPath -Name 'FriendlyName' -Value 'Outlook Phishing Reporter' -PropertyType String -Force | Out-Null; New-ItemProperty -Path $regPath -Name 'LoadBehavior' -Value 3 -PropertyType DWORD -Force | Out-Null; New-ItemProperty -Path $regPath -Name 'Manifest' -Value '%INSTALL_DIR%\OutlookPhishingReporter.vsto|vstolocal' -PropertyType String -Force | Out-Null; New-ItemProperty -Path $regPath -Name 'AdminEmail' -Value 'yosi.desta@outlook.co.il' -PropertyType String -Force | Out-Null; Write-Host 'Registry configured successfully' -ForegroundColor Green"

echo.
echo ================================================
echo   Installation Complete!
echo ================================================
echo.
echo Please restart Microsoft Outlook to load the add-in.
echo.
echo The 'Report Phishing' button will appear in:
echo   - Mail ribbon (when viewing inbox)
echo   - Message ribbon (when reading emails)
echo.
pause
