# Outlook Phishing Reporter - Installation Guide

## Installation Methods

### Method 1: Batch File Setup (Recommended for End Users)
1. Right-click setup.bat
2. Select "Run as Administrator"
3. Enter your organization's security email address
4. Optionally configure custom icon and button label
5. Restart Microsoft Outlook

### Method 2: PowerShell Setup (For IT Administrators)
powershell.exe -ExecutionPolicy Bypass -File Install.ps1

### Method 3: Manual Registry Setup
Run as Administrator:

powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
New-Item -Path $regPath -Force | Out-Null
New-ItemProperty -Path $regPath -Name "AdminEmail" -Value "security@company.com" -PropertyType String -Force
New-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path $regPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force
New-ItemProperty -Path $regPath -Name "Manifest" -Value "C:\Path\To\OutlookPhishingReporter.dll.manifest" -PropertyType String -Force


### Method 4: MSI Setup (Enterprise)
msiexec /i OutlookPhishingReporter.msi /quiet /norestart

## System Requirements
- Windows 7 SP1 or later (Windows 10/11 recommended)
- Microsoft Outlook 2016 or later
- .NET Framework 4.7.2 or later
- Administrator privileges for installation

## Configuration

### Required Settings
- **AdminEmail**: Email address where phishing reports will be sent
  - Example: security@company.com
  - Must be a valid email address

### Optional Settings
- **CustomIconPath**: Path to custom icon file (PNG/ICO)
  - Example: C:\Company\Branding\icon.png
  - Overrides default icon

- **ButtonLabel**: Custom label for the ribbon button
  - Example: "Report Threat" or "Report Spam"
  - Default: "Report Phishing"

All settings are stored in the Windows Registry at:
HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter

## Post-Installation Verification

1. **Check Registry Entries**
   - Run: egedit
   - Navigate to: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
   - Verify LoadBehavior = 3

2. **Verify in Outlook**
   - Open Microsoft Outlook
   - Look for "Report Phishing" button in ribbon
   - Should appear in both mail list and message reading views

3. **Test Functionality**
   - Select an email
   - Click "Report Phishing" button
   - Verify email forwards to AdminEmail
   - Check that original email moves to Deleted Items

## Troubleshooting

### Button doesn't appear in Outlook
1. Verify registry entries exist at: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
2. Ensure LoadBehavior = 3
3. Restart Outlook completely (close all windows)
4. Check log file: %TEMP%\OutlookPhishingReporter.log

### Configuration error message
1. Verify AdminEmail format is correct: user@domain.com
2. Re-run setup.bat and re-enter email
3. Check network connectivity

### Custom icon not loading
1. Verify file path is correct
2. Ensure file is valid PNG or ICO format
3. Verify user has read permissions on file
4. Check log file for details

## Uninstallation

### Via Registry (Recommended)
powershell
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -Force -ErrorAction SilentlyContinue


Restart Outlook.

### Via Registry Editor
1. Run: regedit
2. Navigate to: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins
3. Delete: OutlookPhishingReporter folder
4. Restart Outlook

### Via MSI
msiexec /x OutlookPhishingReporter.msi /quiet

## Support and Documentation

For detailed information, see:
- README.md - Project overview
- QUICK_START.md - Quick installation steps
- MSI_BUILD_INSTRUCTIONS.md - Enterprise deployment

For issues, check: %TEMP%\OutlookPhishingReporter.log

## Version Information
- Version: 1.0
- Target: Outlook 2016+
- Framework: .NET 4.7.2
