#!/usr/bin/env powershell
# Quick Start Guide - Outlook Phishing Reporter

Write-Host "`n" -NoNewline
Write-Host "????????????????????????????????????????????????????????????????" -ForegroundColor Cyan
Write-Host "  Outlook Phishing Reporter - Installer Quick Reference" -ForegroundColor Cyan
Write-Host "????????????????????????????????????????????????????????????????`n" -ForegroundColor Cyan

# EXE Installation
Write-Host "?? EXE INSTALLER (Recommended for most users)" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  Location: OutlookPhishingReporterSetup\Release\setup.bat" -ForegroundColor Yellow
Write-Host ""
Write-Host "  How to Install:" -ForegroundColor Cyan
Write-Host "    1. Open File Explorer" -ForegroundColor White
Write-Host "    2. Navigate to: OutlookPhishingReporterSetup\Release\" -ForegroundColor White
Write-Host "    3. Right-click setup.bat" -ForegroundColor White
Write-Host "    4. Select 'Run as Administrator'" -ForegroundColor White
Write-Host "    5. Enter your organization's security email" -ForegroundColor White
Write-Host "    6. Optionally configure custom icon and label" -ForegroundColor White
Write-Host "    7. Restart Microsoft Outlook" -ForegroundColor White
Write-Host ""
Write-Host "  Expected Result:" -ForegroundColor Cyan
Write-Host "    ? 'Report Phishing' button appears in Outlook ribbon" -ForegroundColor White
Write-Host "    ? Button available in mail list and message views" -ForegroundColor White
Write-Host "    ? Settings stored in Windows Registry" -ForegroundColor White
Write-Host ""

# MSI Installation
Write-Host "?? MSI INSTALLER (For Enterprise Deployment)" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  Source Location: OutlookPhishingReporterSetup\MSISource\" -ForegroundColor Yellow
Write-Host "  Template File: Product.wxs" -ForegroundColor Yellow
Write-Host ""
Write-Host "  How to Build:" -ForegroundColor Cyan
Write-Host "    1. Install WiX Toolset v3.11+" -ForegroundColor White
Write-Host "       Download: https://wixtoolset.org/" -ForegroundColor White
Write-Host "    2. Open Command Prompt" -ForegroundColor White
Write-Host "    3. Navigate to: OutlookPhishingReporterSetup\MSISource\" -ForegroundColor White
Write-Host "    4. Run these commands:" -ForegroundColor White
Write-Host ""
Write-Host "       ""C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe"" Product.wxs" -ForegroundColor Gray
Write-Host "       ""C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe"" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj" -ForegroundColor Gray
Write-Host ""
Write-Host "  How to Deploy:" -ForegroundColor Cyan
Write-Host "    Silent:  msiexec /i OutlookPhishingReporter.msi /quiet /norestart" -ForegroundColor Gray
Write-Host "    With UI: msiexec /i OutlookPhishingReporter.msi" -ForegroundColor Gray
Write-Host "    With Log: msiexec /i OutlookPhishingReporter.msi /l*v install.log" -ForegroundColor Gray
Write-Host ""
Write-Host "  For Group Policy or SCCM deployment:" -ForegroundColor Cyan
Write-Host "    See: OutlookPhishingReporterSetup\MSISource\MSI_BUILD_INSTRUCTIONS.md" -ForegroundColor White
Write-Host ""

# Configuration
Write-Host "??  CONFIGURATION" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  Registry Path:" -ForegroundColor Cyan
Write-Host "    HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -ForegroundColor Gray
Write-Host ""
Write-Host "  Required Settings:" -ForegroundColor Cyan
Write-Host "    AdminEmail (string) = security@company.com" -ForegroundColor White
Write-Host ""
Write-Host "  Auto-Configured:" -ForegroundColor Cyan
Write-Host "    LoadBehavior (DWORD) = 3" -ForegroundColor White
Write-Host "    Description (string) = Outlook Phishing Reporter" -ForegroundColor White
Write-Host ""
Write-Host "  Optional Settings:" -ForegroundColor Cyan
Write-Host "    CustomIconPath (string) = C:\path\to\icon.png" -ForegroundColor White
Write-Host "    ButtonLabel (string) = Report Threat" -ForegroundColor White
Write-Host ""

# Verification
Write-Host "? VERIFICATION CHECKLIST" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  After installation, verify:" -ForegroundColor Cyan
Write-Host "    ? Outlook shows 'Report Phishing' button in ribbon" -ForegroundColor White
Write-Host "    ? Button appears in both mail list and message views" -ForegroundColor White
Write-Host "    ? Click button on test email - it should forward correctly" -ForegroundColor White
Write-Host "    ? Check registry: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -ForegroundColor White
Write-Host "    ? AdminEmail value is set correctly" -ForegroundColor White
Write-Host "    ? LoadBehavior value = 3" -ForegroundColor White
Write-Host ""

# Troubleshooting
Write-Host "?? QUICK TROUBLESHOOTING" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  Button doesn't appear:" -ForegroundColor Cyan
Write-Host "    1. Restart Outlook completely" -ForegroundColor White
Write-Host "    2. Verify registry entries exist" -ForegroundColor White
Write-Host "    3. Ensure LoadBehavior = 3" -ForegroundColor White
Write-Host "    4. Check: %TEMP%\OutlookPhishingReporter.log" -ForegroundColor White
Write-Host ""
Write-Host "  Invalid email error:" -ForegroundColor Cyan
Write-Host "    1. Verify email format: user@domain.com" -ForegroundColor White
Write-Host "    2. No spaces or special characters (except . _ %)" -ForegroundColor White
Write-Host "    3. Run setup.bat again" -ForegroundColor White
Write-Host ""

# Documentation
Write-Host "?? DOCUMENTATION FILES" -ForegroundColor Green
Write-Host "???????????????????????????????????????????????????????????" -ForegroundColor Green
Write-Host ""
Write-Host "  Location: OutlookPhishingReporterSetup\Release\" -ForegroundColor Yellow
Write-Host ""
Write-Host "    INSTALLATION_GUIDE.md" -ForegroundColor Cyan
Write-Host "      ? Complete installation & configuration guide" -ForegroundColor Gray
Write-Host ""
Write-Host "    QUICK_START.md" -ForegroundColor Cyan
Write-Host "      ? Quick start for end users" -ForegroundColor Gray
Write-Host ""
Write-Host "    README.md" -ForegroundColor Cyan
Write-Host "      ? Project overview and features" -ForegroundColor Gray
Write-Host ""
Write-Host "  Location: OutlookPhishingReporterSetup\MSISource\" -ForegroundColor Yellow
Write-Host ""
Write-Host "    MSI_BUILD_INSTRUCTIONS.md" -ForegroundColor Cyan
Write-Host "      ? Step-by-step MSI build guide" -ForegroundColor Gray
Write-Host ""
Write-Host "  Location: OutlookPhishingReporterSetup\" -ForegroundColor Yellow
Write-Host ""
Write-Host "    BUILD_COMPLETION_SUMMARY.md" -ForegroundColor Cyan
Write-Host "      ? Complete build summary & troubleshooting" -ForegroundColor Gray
Write-Host ""

# Success message
Write-Host "? BUILD COMPLETE AND READY FOR DEPLOYMENT!" -ForegroundColor Green
Write-Host ""
Write-Host "Next Step: Run setup.bat as Administrator to test installation" -ForegroundColor Yellow
Write-Host ""
Write-Host "????????????????????????????????????????????????????????????????`n" -ForegroundColor Cyan
