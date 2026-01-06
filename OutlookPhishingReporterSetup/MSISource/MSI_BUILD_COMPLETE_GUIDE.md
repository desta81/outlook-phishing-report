# ?? MSI BUILD - Complete Step-by-Step Guide

## Overview

This guide walks you through building and deploying the MSI installer for Outlook Phishing Reporter.

---

## Phase 1: Install WiX Toolset (One-time Setup)

### Step 1: Download WiX Toolset

**Option A: Official Website (Recommended)**
1. Go to: https://wixtoolset.org/
2. Click "Download"
3. Download WiX Toolset v3.14 (or latest)
4. Save the installer (e.g., `wix314.exe`)

**Option B: GitHub Releases**
1. Go to: https://github.com/wixtoolset/wix3/releases
2. Find latest v3.x release
3. Download the `.exe` file
4. Save the installer

### Step 2: Install WiX Toolset

1. Double-click the downloaded installer
2. Follow the installation wizard
3. Accept license agreement
4. Choose installation location (default is fine)
5. Click "Install"
6. Wait for installation to complete
7. **Restart your computer**

### Step 3: Verify Installation

1. Open Command Prompt (Win + R ? cmd)
2. Type: `candle.exe -version`
3. If you see version info, installation is successful
4. Type: `light.exe -version`
5. If you see version info, you're ready to build

---

## Phase 2: Build the MSI (Quick Method)

### Easiest Way: Run Build Script

1. **Navigate to MSI Source folder:**
   ```
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
   ```

2. **Right-click `build-msi.bat`**

3. **Select "Run as Administrator"**

4. **Wait for build to complete**
   - The script will:
     - Find WiX Toolset installation
     - Compile Product.wxs
     - Link to create MSI
     - Display status and next steps
   - You'll see: `[+] MSI File Created Successfully`

5. **Press any key to close the window**

The MSI is now created at:
```
OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

---

## Phase 2 Alternative: Manual Build

If you prefer to run commands manually:

### Step 1: Open Command Prompt

1. Press `Win + R`
2. Type: `cmd`
3. Press Enter

### Step 2: Navigate to MSI Source

```cmd
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource"
```

### Step 3: Create Distribution Folder

```cmd
mkdir ..\Distribution
```

### Step 4: Compile WiX Source

```cmd
"C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" Product.wxs
```

**Expected output:**
```
WiX Toolset Compiler version X.X.X.X
...
```

### Step 5: Link to Create MSI

```cmd
"C:\Program Files (x86)\WiX Toolset v3.14\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

**Expected output:**
```
WiX Toolset Linker version X.X.X.X
[MSI] Successfully created OutlookPhishingReporter.msi
```

---

## Phase 3: Verify MSI Created

### Check File Exists

1. Open File Explorer
2. Navigate to: `OutlookPhishingReporterSetup\Distribution\`
3. Verify `OutlookPhishingReporter.msi` exists
4. File should be ~2-5 MB in size

---

## Phase 4: Test MSI Installation

### Test 1: Interactive Installation

1. **Navigate to:** `OutlookPhishingReporterSetup\Distribution\`

2. **Double-click:** `OutlookPhishingReporter.msi`

3. **Installation Wizard opens:**
   - Click "Next" to proceed
   - Choose installation path (or use default)
   - Click "Install"
   - Wait for completion

4. **Verify Installation:**
   - Check registry: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
   - Open Outlook
   - Look for "Report Phishing" button

### Test 2: Silent Installation

Open Command Prompt as Administrator:

```cmd
msiexec /i "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi" /quiet /norestart
```

**Expected:** No user interface, silent installation

### Test 3: Installation with Log

```cmd
msiexec /i "OutlookPhishingReporter.msi" /l*v install.log
```

**Check log file:** Open `install.log` to see installation details

---

## Phase 5: Enterprise Deployment

### Option 1: Group Policy Deployment

**For IT Administrators:**

1. **Place MSI on Network Share:**
   ```
   \\SERVER\Software\OutlookPhishingReporter\OutlookPhishingReporter.msi
   ```

2. **Create Group Policy on Domain Controller:**
   - Open `gpedit.msc`
   - Navigate to: Computer Configuration ? Software Settings ? Software Installation
   - Right-click ? New ? Package
   - Select MSI from network share
   - Choose:
     - **Assigned** = Automatic installation
     - **Published** = User can choose to install

3. **Configure Registry via Group Policy Preferences:**
   - Create preference to set AdminEmail registry value
   - Link policy to target OUs

4. **Apply to Users/Computers:**
   - Policy applies at next logon or reboot

### Option 2: SCCM Deployment

**For System Center Configuration Manager:**

1. **Create Application:**
   - New Application
   - Import from MSI
   - Configure deployment type

2. **Set Detection Method:**
   - Registry path: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
   - Value name: `AdminEmail`

3. **Configure Requirements:**
   - OS: Windows 7 or later
   - Office: Outlook 2016 or later
   - .NET Framework: 4.7.2 or later

4. **Create Deployment:**
   - Target: Device collection
   - Purpose: Required or Available
   - Deadline: Set as needed

5. **Monitor Deployment:**
   - Check success rate
   - Review deployment logs

---

## Phase 6: Distribute to End Users

### Option 1: Share on Network

1. Place MSI on network share
2. Send link to users
3. Users run: `msiexec /i \\server\share\OutlookPhishingReporter.msi`

### Option 2: Email Distribution

1. Attach MSI to email
2. Include instructions:
   ```
   1. Extract the MSI file
   2. Right-click OutlookPhishingReporter.msi
   3. Select "Install"
   4. Follow wizard
   5. Restart Outlook
   ```

### Option 3: Self-Service Portal

1. Upload MSI to internal IT portal
2. Users download and install
3. Include deployment guide on portal

---

## Troubleshooting

### Problem: "WiX Toolset not found"

**Solution:**
1. Verify WiX installation completed
2. Restart computer after installation
3. Check installation path matches
4. Reinstall WiX if needed

### Problem: "Product.wxs not found"

**Solution:**
1. Verify you're in the correct directory:
   ```
   OutlookPhishingReporterSetup\MSISource\
   ```
2. Check that `Product.wxs` file exists
3. Copy the full path: `cd "C:\Users\...\MSISource"`

### Problem: "Distribution folder not found"

**Solution:**
```cmd
mkdir ..\Distribution
```

Then run the build again.

### Problem: MSI Installation Fails

1. **Check permissions:** Run as Administrator
2. **Check Outlook:** Outlook 2016+ must be installed
3. **Check .NET:** Framework 4.7.2+ must be installed
4. **Review log:** Run with `/l*v install.log` to see details

### Problem: Button doesn't appear after installation

1. Restart Outlook completely
2. Check registry entries created
3. Verify LoadBehavior = 3
4. Check `%TEMP%\OutlookPhishingReporter.log`

---

## File Locations

### Build Files
- **Source:** `OutlookPhishingReporterSetup\MSISource\Product.wxs`
- **Build Script:** `OutlookPhishingReporterSetup\MSISource\build-msi.bat`
- **Output:** `OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi`

### Configuration Files
- **Build Instructions:** This file (MSI_BUILD_INSTRUCTIONS.md)
- **Build Guide:** `BUILD_MSI_GUIDE.md`
- **Installation Guide:** `..\Release\INSTALLATION_GUIDE.md`

---

## Success Indicators

? **Successful Build:**
- `Product.wixobj` created (compilation)
- `OutlookPhishingReporter.msi` created (linking)
- No error messages

? **Successful Installation:**
- Registry key created at `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
- Outlook loads without errors
- "Report Phishing" button appears

? **Successful Deployment:**
- Users receive MSI
- Installation completes silently
- Button appears in all user instances of Outlook

---

## Quick Reference

| Task | Command |
|------|---------|
| Build MSI | `build-msi.bat` (double-click) |
| Manual Compile | `candle.exe Product.wxs` |
| Manual Link | `light.exe -out .msi Product.wixobj` |
| Test Install | `msiexec /i OutlookPhishingReporter.msi` |
| Silent Install | `msiexec /i OutlookPhishingReporter.msi /quiet` |
| Uninstall | `msiexec /x OutlookPhishingReporter.msi` |
| View Log | `msiexec /i OutlookPhishingReporter.msi /l*v install.log` |

---

## Next Steps

1. ? **Install WiX Toolset** (if not already done)
2. ? **Run build-msi.bat** to create the MSI
3. ? **Test MSI** on a test machine
4. ? **Deploy** via Group Policy or SCCM
5. ? **Verify** add-in loads on user machines

---

## Support

For issues or questions:
1. Review troubleshooting section above
2. Check WiX documentation: https://wixtoolset.org/
3. Review build output for error messages
4. Check installation log: `install.log`

---

**Status:** Ready to build! Follow Phase 1 (if needed) and Phase 2 above.
