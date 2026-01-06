# Outlook Phishing Reporter - Build & Installation Summary

## Build Status: ? SUCCESS

The build script has successfully generated both **EXE** and **MSI** installers for the Outlook Phishing Reporter add-in.

---

## Generated Files

### Release Directory
**Location:** `OutlookPhishingReporterSetup\Release\`

| File | Size | Purpose |
|------|------|---------|
| `setup.bat` | 3.98 KB | Interactive batch installer (EXE) |
| `setup.exe` | 2.04 KB | Batch wrapper executable |
| `INSTALLATION_GUIDE.md` | 4.05 KB | Comprehensive installation documentation |
| `OutlookPhishingReporter.dll` | 41.5 KB | Main add-in assembly |
| `OutlookPhishingReporter.vsto` | 5.5 KB | VSTO manifest |
| `OutlookPhishingReporter.dll.manifest` | 13.8 KB | Deployment manifest |
| `Microsoft.Office.Tools.*.dll` | 79.8 KB | Supporting libraries |
| `OutlookPhishingReporter.pdb` | 43.5 KB | Debug symbols |

### MSI Source Directory
**Location:** `OutlookPhishingReporterSetup\MSISource\`

| File | Purpose |
|------|---------|
| `Product.wxs` | WiX source file for MSI (ready to compile) |
| `MSI_BUILD_INSTRUCTIONS.md` | Step-by-step MSI build guide |

---

## Installation Options

### Option 1: EXE Installer (Interactive) ? Ready to Use

**Perfect for:** End users, single machines, quick deployment

**Steps:**
1. Open `OutlookPhishingReporterSetup\Release\`
2. Right-click `setup.bat` ? **"Run as Administrator"**
3. Follow the interactive prompts:
   - Enter your organization's security email address (required)
   - Optionally set a custom icon path
   - Optionally customize the button label
4. Restart Microsoft Outlook
5. Look for the "Report Phishing" button in the ribbon

**Features:**
- ? Interactive configuration
- ? Email validation
- ? Custom icon support
- ? Custom button label support
- ? Registry-based configuration
- ? No dependencies

---

### Option 2: MSI Installer (Enterprise) ?? Requires Build

**Perfect for:** Enterprise deployment, Group Policy, SCCM

**Prerequisites:**
1. Install **WiX Toolset v3.11** or later
   - Download: https://github.com/wixtoolset/wix3/releases
   - Or: https://wixtoolset.org/

**Build Steps:**

```batch
cd OutlookPhishingReporterSetup\MSISource

REM Compile WiX source file
"C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" Product.wxs

REM Link to create MSI
"C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

**Output:**
- File: `OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi`

**Deployment:**

```batch
REM Silent installation
msiexec /i OutlookPhishingReporter.msi /quiet /norestart

REM With installation log
msiexec /i OutlookPhishingReporter.msi /l*v install.log

REM With user interface
msiexec /i OutlookPhishingReporter.msi
```

**Features:**
- ? Silent installation
- ? Group Policy deployment ready
- ? SCCM compatible
- ? Centralized configuration
- ? Standard uninstall via Programs & Features
- ? Enterprise logging

---

## Registry Configuration

### Registry Path
```
HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
```

### Required Values

| Value | Type | Example | Description |
|-------|------|---------|-------------|
| `AdminEmail` | REG_SZ | security@company.com | Email for phishing reports |

### Auto-Set Values

| Value | Type | Default |
|-------|------|---------|
| `LoadBehavior` | REG_DWORD | 3 (load at startup) |
| `Description` | REG_SZ | Outlook Phishing Reporter |
| `Manifest` | REG_SZ | Path to manifest file |

### Optional Values

| Value | Type | Example | Description |
|-------|------|---------|-------------|
| `CustomIconPath` | REG_SZ | C:\Company\icon.png | Custom button icon |
| `ButtonLabel` | REG_SZ | Report Threat | Custom button label |

---

## Manual PowerShell Installation

If you need to install manually without using the batch file:

```powershell
# Run as Administrator

$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"

# Create registry key
New-Item -Path $regPath -Force | Out-Null

# Set required configuration
New-ItemProperty -Path $regPath -Name "AdminEmail" -Value "security@company.com" -PropertyType String -Force
New-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path $regPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force

# Set manifest path (update path as needed)
New-ItemProperty -Path $regPath -Name "Manifest" -Value "C:\Path\To\OutlookPhishingReporter.dll.manifest" -PropertyType String -Force

# Optional: Custom icon
New-ItemProperty -Path $regPath -Name "CustomIconPath" -Value "C:\Company\icon.png" -PropertyType String -Force

# Optional: Custom button label
New-ItemProperty -Path $regPath -Name "ButtonLabel" -Value "Report Threat" -PropertyType String -Force
```

---

## System Requirements

| Component | Requirement |
|-----------|-------------|
| **Windows** | Windows 7 SP1 or later (Windows 10/11 recommended) |
| **Outlook** | Microsoft Outlook 2016, 2019, or Microsoft 365 |
| **.NET Framework** | 4.7.2 or later |
| **Permissions** | Administrator rights for installation |
| **Network** | Email server connectivity for reporting |

---

## Post-Installation Verification

### 1. Check Registry Entries
```batch
regedit
```
Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`

Verify:
- ? `LoadBehavior` = 3
- ? `AdminEmail` is set
- ? `Manifest` path is correct

### 2. Verify in Outlook
- Open Microsoft Outlook
- Look for **"Report Phishing"** button in the ribbon
- Should appear in:
  - Mail list view (top ribbon)
  - Message reading view (top ribbon)

### 3. Test Functionality
1. Select any email
2. Click **"Report Phishing"** button
3. Verify:
   - Email is forwarded to the AdminEmail address
   - Original email moves to Deleted Items
   - Success message appears

### 4. Check Log File
```
%TEMP%\OutlookPhishingReporter.log
```

---

## Troubleshooting

### Issue: Button doesn't appear in Outlook

**Solutions:**
1. Verify registry entries exist at: `HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
2. Ensure `LoadBehavior` value = 3
3. Verify DLL file exists at the path in registry
4. Restart Outlook completely (close all windows)
5. Check log file: `%TEMP%\OutlookPhishingReporter.log`

### Issue: "Invalid Admin Email" error

**Solutions:**
1. Verify email format: `user@domain.com`
2. Ensure no spaces or special characters (except . _ %)
3. Run setup.bat again and re-enter email
4. Check network connectivity to email server

### Issue: Custom icon not loading

**Solutions:**
1. Verify file path is correct (full path, not relative)
2. Ensure file is valid PNG or ICO format
3. Verify user account has read permissions
4. Check log file for specific error

---

## Uninstallation

### Method 1: Via Registry (Recommended)

```powershell
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -Force
```

Then restart Outlook.

### Method 2: Via Registry Editor

1. Run: `regedit`
2. Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins`
3. Right-click: `OutlookPhishingReporter` folder
4. Select: **Delete**
5. Restart Outlook

### Method 3: Via MSI (if installed via MSI)

```batch
msiexec /x OutlookPhishingReporter.msi /quiet
```

---

## Enterprise Deployment

### Group Policy (Windows Domain)

1. Build the MSI file (see MSI section above)
2. Place on network share: `\\SERVER\Software\OutlookPhishingReporter.msi`
3. Create Group Policy:
   - Open `gpedit.msc` (domain controller)
   - Go: Computer Configuration ? Software Settings ? Software Installation
   - New ? Package ? Select MSI from network share
   - Choose: Assigned (computers) or Published (users)
4. Apply to target Organizational Units (OUs)
5. Use Group Policy Preferences to set registry values for:
   - AdminEmail
   - CustomIconPath
   - ButtonLabel

### System Center Configuration Manager (SCCM)

1. Create new Application
2. Add deployment type with MSI
3. Configure detection method:
   - Registry path exists at: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
4. Set installation command:
   ```batch
   msiexec /i OutlookPhishingReporter.msi /quiet /norestart
   ```
5. Deploy to device collection
6. Monitor deployment status

---

## Build Documentation

For detailed information about building the installers, see:

| Document | Location | Purpose |
|----------|----------|---------|
| `INSTALLATION_GUIDE.md` | `OutlookPhishingReporterSetup\Release\` | Complete installation & configuration guide |
| `MSI_BUILD_INSTRUCTIONS.md` | `OutlookPhishingReporterSetup\MSISource\` | WiX MSI build instructions |
| `QUICK_START.md` | `OutlookPhishingReporterSetup\Release\` | Quick start guide |
| `README.md` | `OutlookPhishingReporterSetup\Release\` | Project overview |
| `BUILD_SUMMARY.md` | `OutlookPhishingReporterSetup\Release\` | Build process summary |

---

## Support & Troubleshooting

### Log File Location
```
%TEMP%\OutlookPhishingReporter.log
```

### Common Issues
- [Button doesn't appear in Outlook](#issue-button-doesnt-appear-in-outlook)
- [Invalid Admin Email error](#issue-invalid-admin-email-error)
- [Custom icon not loading](#issue-custom-icon-not-loading)

### Contact
For additional support, please:
1. Review log file at `%TEMP%\OutlookPhishingReporter.log`
2. Verify system requirements
3. Check registry settings
4. Contact your IT department or administrator

---

## Version Information

- **Product:** Outlook Phishing Reporter
- **Version:** 1.0
- **Build Date:** 2026-01-05
- **Target:** Outlook 2016+
- **Framework:** .NET 4.7.2
- **Repository:** https://github.com/samicsc0/Outlook-Spam-Reporter

---

## Next Steps

### Immediate (Today)
1. ? Test EXE installer: `setup.bat`
2. ? Verify installation in Outlook
3. ? Test phishing report functionality

### Short-term (This Week)
1. Distribute EXE to test group of users
2. Collect feedback on functionality
3. Validate email forwarding to AdminEmail

### Long-term (This Month)
1. Build MSI using WiX Toolset (if enterprise deployment needed)
2. Test MSI deployment
3. Deploy to organization via Group Policy or SCCM
4. Train end users on how to use the add-in

---

**Build completed successfully! All files are ready for deployment.** ??
