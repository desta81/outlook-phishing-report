# 📦 MSI BUILD - COMPLETE STEP-BY-STEP GUIDE

## 🎯 OVERVIEW

This guide will walk you through building a professional Windows Installer package (MSI) for enterprise deployment of the Outlook Phishing Reporter.

---

## 📋 PREREQUISITES CHECKLIST

Before starting, verify you have:

- [ ] Administrator access on your computer
- [ ] WiX Toolset v3.11+ installed
- [ ] Visual Studio (to build plugin .DLL)
- [ ] Plugin compiled (OutlookPhishingReporter.dll)
- [ ] All source files present

---

## 🔧 STEP 1: INSTALL WIX TOOLSET (One-Time)

### **What is WiX?**
WiX (Windows Installer XML) is a free toolset for creating Windows installer packages (MSI files).

### **Download**

**From Official Site:**
```
1. Go to: https://wixtoolset.org/
2. Click: Download
3. Look for: WiX v3.14 (or latest v3.x)
4. Click download link
```

**Or from GitHub:**
```
1. Go to: https://github.com/wixtoolset/wix3/releases
2. Find latest v3.x release
3. Download installer .exe
```

### **Install WiX**

**Windows Installation:**
```
1. Run the downloaded installer
2. Accept license terms
3. Choose installation path (default is fine)
4. Installation path: C:\Program Files (x86)\WiX Toolset v3.14\
5. Click Install
6. Wait for installation to complete
7. Click Finish
8. Restart computer (if prompted)
```

**Verify Installation:**
```cmd
# Open Command Prompt
where candle.exe

# Should show path like:
# C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe
```

---

## 🔨 STEP 2: PREPARE PLUGIN

### **Build DLL**

```
1. Open Visual Studio
2. Open project: OutlookPhishingReporter.csproj
3. Build → Build Solution
4. Verify no errors (should have "Build succeeded")
5. Creates: OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
```

### **Verify DLL Exists**

```
File should be at:
OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll

If not found:
  - Rebuild solution
  - Check for compilation errors
  - Verify .csproj settings
```

---

## 📂 STEP 3: VERIFY MSI SOURCE FILES

### **Check MSISource Folder**

```
Navigate to: OutlookPhishingReporterSetup\MSISource\

Should contain:
  ✓ build-msi.bat
  ✓ Product.wxs
  ✓ [Other WiX files if any]
```

### **Check Product.wxs**

```
This file defines your MSI:
  - Package name
  - Version
  - Installation paths
  - Registry entries
  - Included files
```

---

## 🏗️ STEP 4: BUILD THE MSI

### **Method 1: Double-Click (Easiest)**

**Steps:**
```
1. Open File Explorer
2. Navigate to: OutlookPhishingReporterSetup\MSISource\
3. Find: build-msi.bat
4. Double-click: build-msi.bat
5. Windows command prompt opens
6. Watch build progress
7. Wait for "Build Complete!" message
8. Press any key to close
```

### **Method 2: Right-Click as Administrator**

**Steps:**
```
1. Navigate to: OutlookPhishingReporterSetup\MSISource\
2. Right-click: build-msi.bat
3. Select: Run as Administrator
4. Command prompt opens with admin rights
5. Follow same process as Method 1
```

### **Method 3: Command Line**

**Steps:**
```
1. Open Command Prompt as Administrator
2. Navigate to: cd OutlookPhishingReporterSetup\MSISource\
3. Run: build-msi.bat
4. Wait for completion
```

---

## 📊 BUILD PROGRESS MESSAGES

When you run the build script, you'll see:

```
===================================================
Outlook Phishing Reporter - MSI Build Script
===================================================

[*] Working Directory: C:\...\OutlookPhishingReporterSetup\MSISource\
[*] Output Directory: C:\...\OutlookPhishingReporterSetup\Distribution\

[*] Searching for WiX Toolset...
[+] Found WiX v3.14
[+] WiX Path: C:\Program Files (x86)\WiX Toolset v3.14\bin

[*] Cleaning previous build artifacts...
[+] Cleaned

===================================================
Building MSI Installer
===================================================

[1/2] Compiling WiX source file...
[+] Compilation successful

[2/2] Linking and creating MSI...
[+] Linking successful

===================================================
Build Complete!
===================================================

[+] MSI File Created Successfully

Location: C:\...\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
Size: 8576000 bytes

===================================================
Next Steps
===================================================

1. Test Installation on a test machine...
2. Silent Installation (for enterprise)...
3. Deployment via Group Policy...
```

---

## ✅ STEP 5: VERIFY BUILD SUCCESS

### **Check 1: Output File Exists**

```
Navigate to: OutlookPhishingReporterSetup\Distribution\

Should see: OutlookPhishingReporter.msi
File size: Typically 5-10 MB
```

### **Check 2: File Properties**

```
Right-click MSI file
Select: Properties
Check:
  - File size > 1 MB (should be 5-10 MB)
  - Modified date = today's date
  - Version info (if present)
```

### **Check 3: Test on Test Machine**

```
Recommended: Use isolated test environment

1. Copy MSI to test machine
2. Double-click: OutlookPhishingReporter.msi
3. Follow installation wizard
4. Choose installation options
5. Click Install
6. Installation completes
7. Verify in Control Panel → Programs → Programs and Features
8. Check Outlook for "Report phishing" button
9. Uninstall to test uninstallation
```

---

## 🚀 STEP 6: DEPLOY MSI

### **Option 1: Manual Deployment (Small Groups)**

```
1. Copy MSI file to network share
2. Users download and run
3. Or provide USB drive
4. Users run: msiexec /i OutlookPhishingReporter.msi
```

**Pros:** Simple, full control  
**Cons:** Requires user action  

### **Option 2: Silent Installation (Automated)**

```cmd
# Run without user interaction:
msiexec /i OutlookPhishingReporter.msi /quiet /norestart

# Or with logging:
msiexec /i OutlookPhishingReporter.msi /quiet /norestart /l* install.log
```

**Pros:** No user input needed  
**Cons:** Less user control  

### **Option 3: Group Policy (Enterprise)**

```
1. Copy MSI to Group Policy folder:
   \\server\NETLOGON\Policies\

2. Open Group Policy Editor (gpedit.msc)
3. Navigate to:
   Computer Configuration → Software Settings → Software Installation
4. Right-click → New → Package
5. Select your MSI file
6. Choose deployment method:
   - Assigned: Installs automatically
   - Published: Users can install from Add/Remove Programs
7. Link GPO to organizational unit
8. Users receive installation at next login
```

**Pros:** Centralized management  
**Cons:** Requires Active Directory  

### **Option 4: SCCM (Enterprise)**

```
1. Upload MSI to SCCM
2. Create application package
3. Define deployment parameters
4. Create deployment
5. Target device collections
6. Automatic installation on target machines
7. Monitor installation status
8. Get usage reports
```

**Pros:** Full control, reporting  
**Cons:** Requires SCCM expertise  

---

## 🔧 TROUBLESHOOTING

### **Issue: "WiX Toolset not found"**

**Error Message:**
```
ERROR: WiX Toolset not found!
```

**Solution:**
```
1. Verify WiX installation:
   Check: C:\Program Files (x86)\WiX Toolset v3.14\bin\
2. If not found, install WiX from https://wixtoolset.org/
3. Restart computer
4. Try build again
```

### **Issue: "Product.wxs not found"**

**Error Message:**
```
ERROR: Product.wxs not found at: ...
```

**Solution:**
```
1. Navigate to: OutlookPhishingReporterSetup\MSISource\
2. Verify Product.wxs exists
3. If missing, restore from source control or backup
4. Try build again
```

### **Issue: "Compilation failed"**

**Error Message:**
```
[ERROR] Compilation failed!
```

**Solution:**
```
1. Check error messages above
2. Common causes:
   - Syntax error in Product.wxs
   - Missing referenced files
   - Invalid XML structure
3. Edit Product.wxs to fix issues
4. Try build again
```

### **Issue: "MSI creation failed"**

**Error Message:**
```
[ERROR] MSI creation failed!
```

**Solution:**
```
1. Check error messages
2. Verify Product.wixobj was created
3. Check output directory permissions
4. Run as Administrator
5. Try build again
```

### **Issue: Very Small MSI File**

**Problem:**
```
MSI file is < 1 MB (should be 5-10 MB)
```

**Solution:**
```
1. DLL may not be included
2. Check Product.wxs references DLL
3. Verify DLL exists and is correct version
4. Rebuild plugin
5. Run build again
```

---

## 📝 BUILD CHECKLIST

Before building, verify:

- [ ] WiX Toolset installed
- [ ] Plugin built (DLL exists)
- [ ] Product.wxs file present
- [ ] MSISource folder accessible
- [ ] Administrator access available
- [ ] Distribution folder exists (or will be created)

During build:

- [ ] build-msi.bat runs without errors
- [ ] Compilation successful message appears
- [ ] Linking successful message appears
- [ ] Build Complete message shown

After build:

- [ ] MSI file exists in Distribution folder
- [ ] File size 5-10 MB
- [ ] File is recent (today's date)
- [ ] Test installation succeeds

---

## 🎯 BUILD TIMELINE

| Task | Time | Notes |
|------|------|-------|
| **Install WiX** | 10-15 min | One-time only |
| **Build Plugin DLL** | 2-3 min | If not already built |
| **Navigate to folder** | 1 min | MSISource directory |
| **Run build script** | 5-10 min | Compile & link |
| **Verify output** | 2 min | Check file exists |
| **Test MSI** | 5-10 min | Optional on test machine |
| **Total** | 25-40 min | First time ~40 min |

---

## 🎉 SUCCESS INDICATORS

When build is successful:

✅ No error messages  
✅ "Build Complete!" message shown  
✅ MSI file appears in Distribution folder  
✅ MSI file size 5-10 MB  
✅ MSI can be installed on test machine  
✅ Plugin works after installation  

---

## 📞 GETTING HELP

If you encounter issues:

1. **Check Error Messages**
   - Read error output carefully
   - Google the specific error
   - Check WiX documentation

2. **Check Prerequisites**
   - WiX installed correctly
   - Plugin built successfully
   - All files present
   - Admin rights available

3. **Try Alternatives**
   - Restart computer
   - Run as Administrator explicitly
   - Try command line method
   - Rebuild plugin first

4. **Get Help**
   - Check WiX forums: https://stackoverflow.com/questions/tagged/wix
   - Review build log
   - Ask for help with specific error messages

---

## 🎓 NEXT STEPS

After successful MSI build:

1. **Test Installation**
   - Test machine setup
   - Install from MSI
   - Verify functionality
   - Uninstall cleanly

2. **Prepare Deployment**
   - Choose deployment method
   - Prepare documentation
   - Plan rollout schedule
   - Communicate with users

3. **Deploy**
   - Copy MSI to deployment location
   - Configure Group Policy (if using)
   - Monitor installation
   - Support users

4. **Monitor**
   - Check logs
   - Monitor success rate
   - Address issues
   - Collect feedback

---

**You're ready to build your MSI!** 📦

Start with: **OutlookPhishingReporterSetup\MSISource\build-msi.bat**
