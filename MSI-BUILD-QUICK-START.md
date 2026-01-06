# 📦 MSI BUILD - QUICK START GUIDE

## 🎯 WHAT IS MSI?

**MSI** = Windows Installer Package

This allows you to:
- ✅ Enterprise deployment (Group Policy, SCCM)
- ✅ Silent installation (no user interaction)
- ✅ Automatic configuration
- ✅ Professional distribution
- ✅ Easy uninstallation

---

## ⚡ QUICK BUILD (5 MINUTES)

### **Prerequisites**

Before building, ensure you have:
- ✅ WiX Toolset installed (free, open source)
- ✅ Administrator access
- ✅ Plugin compiled (.DLL ready)
- ✅ Product.wxs file present

### **Check If You Have WiX**

**Option 1: Command Line**
```cmd
where candle.exe
```

If you see a path, WiX is installed. If not, install it.

**Option 2: Check Directly**
```
Look for: C:\Program Files (x86)\WiX Toolset v3.14\
```

---

## 🚀 BUILD MSI NOW

### **Step 1: Navigate to Build Directory**
```
Open: OutlookPhishingReporterSetup\MSISource\
```

### **Step 2: Run Build Script**
```
Double-click: build-msi.bat
Or right-click → Run as Administrator
```

### **Step 3: Wait for Build**
```
You'll see progress messages:
  [*] Searching for WiX Toolset...
  [+] Found WiX v3.14
  [1/2] Compiling WiX source file...
  [+] Compilation successful
  [2/2] Linking and creating MSI...
  [+] Linking successful
  
  Build Complete!
```

### **Step 4: MSI is Ready**
```
Output location: OutlookPhishingReporterSetup\Distribution\
File: OutlookPhishingReporter.msi
```

---

## 📋 WHAT YOU NEED FIRST

### **1. WiX Toolset (Required)**

**Download from:**
```
https://wixtoolset.org/
(Look for: WiX Toolset v3.14 or latest v3.x)
```

**Or from GitHub:**
```
https://github.com/wixtoolset/wix3/releases
```

**Installation:**
```
1. Download installer
2. Run installer
3. Follow setup wizard
4. Default installation path is fine
5. Restart if prompted
```

### **2. Plugin Built (Required)**

```
1. Open Visual Studio
2. Open: OutlookPhishingReporter.csproj
3. Build → Build Solution
4. Should complete with no errors
5. Creates: OutlookPhishingReporter.dll
```

### **3. Product.wxs File (Included)**

```
Already present in: OutlookPhishingReporterSetup\MSISource\
File: Product.wxs
This defines the MSI configuration
```

---

## 🛠️ BUILD OPTIONS

### **Option 1: Quick Build (Recommended)**
```
Double-click: build-msi.bat
Automatically finds WiX
Creates MSI
Shows completion message
```

### **Option 2: Command Line**
```cmd
cd OutlookPhishingReporterSetup\MSISource\
build-msi.bat
```

### **Option 3: Manual Build**
```cmd
REM Step 1: Compile
"C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" Product.wxs -o Product.wixobj

REM Step 2: Link
"C:\Program Files (x86)\WiX Toolset v3.14\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

---

## 🎉 WHAT YOU GET

### **Output File**
```
Location: OutlookPhishingReporterSetup\Distribution\
File: OutlookPhishingReporter.msi
Size: ~5-10 MB
```

### **What MSI Includes**
- ✅ Plugin .DLL
- ✅ Registry entries
- ✅ Configuration for email address
- ✅ Installation/Uninstallation support
- ✅ Enterprise deployment ready

---

## 🚀 AFTER BUILD - DEPLOY MSI

### **Option 1: Manual Install (Testing)**
```cmd
msiexec /i OutlookPhishingReporter.msi
```

### **Option 2: Silent Install (Enterprise)**
```cmd
msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```

### **Option 3: Group Policy (Enterprise)**
```
1. Copy MSI to Group Policy folder
2. Create GPO
3. Assign application
4. Users receive at next login
```

### **Option 4: SCCM (Enterprise)**
```
1. Upload MSI to SCCM
2. Create deployment package
3. Deploy to device collections
4. Automatic installation
```

---

## ✅ BUILD VERIFICATION

After build completes:

**Check 1: File Exists**
```
Navigate to: OutlookPhishingReporterSetup\Distribution\
Should see: OutlookPhishingReporter.msi
```

**Check 2: File Size**
```
Should be: 5-10 MB (typical)
If < 1 MB: Build may have failed
If > 50 MB: Includes extra files
```

**Check 3: Test Install**
```
1. Double-click: OutlookPhishingReporter.msi
2. Follow installation wizard
3. Verify installation completes
4. Check Outlook for button
5. Uninstall to clean up
```

---

## 🆘 TROUBLESHOOTING

### **Error: "WiX Toolset not found"**

**Solution:**
```
1. Install WiX Toolset from https://wixtoolset.org/
2. Default installation path: C:\Program Files (x86)\WiX Toolset v3.14\
3. Restart computer if prompted
4. Try build again
```

### **Error: "Product.wxs not found"**

**Solution:**
```
1. Navigate to: OutlookPhishingReporterSetup\MSISource\
2. Verify Product.wxs exists
3. If missing, restore from repository
4. Try build again
```

### **Error: "Compilation failed"**

**Solution:**
```
1. Check error message in output
2. Verify all files in MSISource folder exist
3. Try manually opening Product.wxs to check syntax
4. Rebuild plugin in Visual Studio
5. Try build again
```

### **Error: "MSI creation failed"**

**Solution:**
```
1. Check WiX version compatibility
2. Verify Product.wixobj was created
3. Check Output directory has write permissions
4. Run as Administrator
5. Try build again
```

---

## 📊 BUILD TIMELINE

| Step | Time | Details |
|------|------|---------|
| **Prerequisites** | 5 min | Install WiX (one-time) |
| **Build Prep** | 1 min | Navigate to folder |
| **Compilation** | 1-2 min | Compile Product.wxs |
| **Linking** | 1-2 min | Create MSI file |
| **Verification** | 1 min | Check output |
| **Total** | 5-10 min | First build slightly slower |

---

## 🎯 NEXT STEPS

### **After Successful Build**

1. **Copy MSI**
   ```
   From: OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
   To: Network share or deployment location
   ```

2. **Test Installation**
   ```
   1. Test machine (isolated from production)
   2. Run: msiexec /i OutlookPhishingReporter.msi
   3. Verify installation
   4. Test functionality
   5. Uninstall and clean
   ```

3. **Deploy to Users**
   ```
   Option A: Manual - users run MSI
   Option B: Group Policy - automatic deployment
   Option C: SCCM - managed deployment
   ```

4. **Monitor**
   ```
   Check logs: %LocalAppData%\OutlookPhishingReporter\Logs\
   Monitor installation on target machines
   Address any issues
   ```

---

## 📝 BUILD COMMANDS

### **Quick Reference**

```batch
REM Navigate to MSI folder
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\

REM Run build
build-msi.bat

REM Or run as administrator
powershell -Command "Start-Process build-msi.bat -Verb runAs"
```

---

## 🔐 MSIS CAN INCLUDE DEFAULTS

The Product.wxs can be configured with:
- ✅ Default email address
- ✅ Default icon
- ✅ Default button label
- ✅ Registry configuration
- ✅ Automatic setup

---

## 💡 PRO TIPS

✅ **Build Once, Deploy Many**
- Build MSI once
- Deploy to hundreds of users
- No rebuild needed per user

✅ **Pre-Configuration**
- Configure defaults in Product.wxs
- Email set for all installations
- No user input needed

✅ **Silent Deployment**
- Users don't see wizard
- Automatic installation
- Zero user interaction

✅ **Easy Uninstall**
- MSI handles uninstall
- Clean registry removal
- Professional experience

---

## 🎉 YOU'RE READY TO BUILD!

**Right now:**
1. Ensure WiX is installed
2. Navigate to: OutlookPhishingReporterSetup\MSISource\
3. Double-click: build-msi.bat
4. Wait 5-10 minutes
5. MSI created in Distribution folder
6. Ready to deploy!

---

**Build your MSI now!** 📦
