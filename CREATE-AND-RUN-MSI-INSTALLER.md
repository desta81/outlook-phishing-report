# 📦 CREATE AND RUN MSI INSTALLER - COMPLETE GUIDE

## 🎯 WHAT YOU'LL GET

After following this guide, you will have:
- ✅ **OutlookPhishingReporter.msi** - Professional installer
- ✅ Ready to deploy to users
- ✅ Works with Group Policy, SCCM, manual installation
- ✅ Automatic plugin configuration
- ✅ Registry setup included

---

## ⚡ QUICK START (5 MINUTES)

### **Step 1: Build Plugin DLL First**
```
Right-click: BUILD-DLL.bat
Select: Run as Administrator
Wait: 1-2 minutes for "Build successful" message
Output: OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
```

### **Step 2: Install WiX Toolset (First Time Only)**
```
Download: https://wixtoolset.org/
Download: WiX v3.14 (or latest v3.x)
Install: Run installer with default settings
Verify: C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe exists
```

### **Step 3: Build MSI**
```
Right-click: OutlookPhishingReporterSetup\MSISource\build-msi.bat
Select: Run as Administrator
Wait: 5-10 minutes
Output: OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

### **Step 4: Run MSI**
```
Navigate to: OutlookPhishingReporterSetup\Distribution\
Double-click: OutlookPhishingReporter.msi
Follow: Installation wizard
Restart: Outlook
Done! ✓
```

---

## 🔍 PREREQUISITES CHECK

Before building MSI, verify you have:

### **Checklist**
```
☐ Visual Studio installed
☐ .NET Framework 4.7.2 installed
☐ WiX Toolset v3.11+ installed
☐ Plugin DLL built (OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll)
☐ Product.wxs exists (OutlookPhishingReporterSetup\MSISource\Product.wxs)
☐ Admin rights to run build script
```

### **Check DLL Built**
```cmd
Navigate to: OutlookPhishingReporter\bin\Release\
Should see: OutlookPhishingReporter.dll (> 100 KB)
```

### **Check WiX Installed**
```cmd
Navigate to: C:\Program Files (x86)\WiX Toolset v3.14\bin\
Should see: candle.exe and light.exe
```

---

## 🚀 BUILD MSI STEP-BY-STEP

### **Method 1: Batch Script (Easiest)**

**Step 1: Navigate to Build Directory**
```
Open File Explorer
Go to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
```

**Step 2: Run Build Script**
```
Right-click: build-msi.bat
Select: Run as Administrator
```

**Step 3: Watch Build Progress**
```
You'll see:
  [*] Searching for WiX Toolset...
  [+] Found WiX v3.14
  [*] Cleaning previous build artifacts...
  [1/2] Compiling WiX source file...
  [+] Compilation successful
  [2/2] Linking and creating MSI...
  [+] Linking successful
  
  Build Complete!
  [+] MSI File Created Successfully
```

**Step 4: Verify MSI Created**
```
Navigate to: OutlookPhishingReporterSetup\Distribution\
Should see: OutlookPhishingReporter.msi (5-10 MB)
```

---

### **Method 2: Command Prompt**

**Step 1: Open Command Prompt as Administrator**
```cmd
Press: Win + R
Type: cmd
Press: Ctrl + Shift + Enter
Or right-click cmd → Run as Administrator
```

**Step 2: Navigate to MSISource**
```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
```

**Step 3: Run Build**
```cmd
build-msi.bat
```

**Step 4: Check Output**
```
When complete, MSI will be in:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

---

### **Method 3: PowerShell**

**Step 1: Open PowerShell as Administrator**
```powershell
Right-click Start → Windows PowerShell (Admin)
```

**Step 2: Navigate and Run**
```powershell
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource"
.\build-msi.bat
```

---

## 📊 BUILD OUTPUT

### **Location**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\
```

### **Files Created**
```
OutlookPhishingReporter.msi    (5-10 MB)    ← Main installer file
OutlookPhishingReporter.msi.log            (optional, debug info)
```

### **Verification**
```
✓ MSI file exists
✓ Size 5-10 MB
✓ Created today
✓ Can be executed
```

---

## 🏃 RUN MSI INSTALLER

### **Method 1: Double-Click (Easiest)**

**Step 1: Navigate to MSI**
```
Open File Explorer
Go to: OutlookPhishingReporterSetup\Distribution\
Find: OutlookPhishingReporter.msi
```

**Step 2: Run Installer**
```
Double-click: OutlookPhishingReporter.msi
```

**Step 3: Follow Wizard**
```
Welcome screen → click Next
Select installation directory → click Next
Click Install
Wait for installation to complete
Click Finish
```

**Step 4: Complete**
```
Restart Outlook
Plugin should appear in ribbon
"Report phishing" button available
```

---

### **Method 2: Command Line - Normal Install**

```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\

msiexec /i OutlookPhishingReporter.msi
```

**What happens:**
- Installation wizard appears
- User selects options
- Click Install
- Installation completes

---

### **Method 3: Command Line - Silent Install**

```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\

msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```

**What happens:**
- No wizard shown
- Automatic installation
- No user interaction needed
- Computer not restarted
- Perfect for enterprise deployment

---

### **Method 4: Command Line - Silent with Logging**

```cmd
msiexec /i OutlookPhishingReporter.msi /quiet /norestart /l* install.log
```

**What happens:**
- Silent installation
- Detailed log created (install.log)
- Useful for troubleshooting

---

## ✅ VERIFY INSTALLATION

### **Check 1: Plugin Appears in Outlook**

```
1. Open Outlook
2. Look at Home tab in ribbon
3. Should see "Report phishing" button
4. Should be in a group named "Report phishing"
```

### **Check 2: Registry Configured**

```powershell
# Run in PowerShell as Administrator:
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -Name AdminEmail
```

Should show:
```
AdminEmail : yosi.desta@outlook.co.il
(or your configured email)
```

### **Check 3: Test Report Function**

```
1. Select a test email
2. Click "Report phishing" button
3. Confirmation dialog appears
4. Click OK
5. Email should be reported
6. No error message
```

---

## 🆘 TROUBLESHOOTING

### **Error: WiX Toolset Not Found**

**Solution:**
```
1. Download WiX from https://wixtoolset.org/
2. Install to: C:\Program Files (x86)\WiX Toolset v3.14\
3. Restart computer
4. Try build again
```

### **Error: DLL Not Found**

**Solution:**
```
1. Build plugin first: BUILD-DLL.bat
2. Verify: OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll exists
3. Try MSI build again
```

### **Error: Product.wxs Not Found**

**Solution:**
```
1. Verify file exists: OutlookPhishingReporterSetup\MSISource\Product.wxs
2. If missing, restore from source control
3. Try build again
```

### **Installation Fails**

**Solution:**
```
1. Close Outlook completely
2. Try installation again
3. If still fails, try silent install:
   msiexec /i OutlookPhishingReporter.msi /quiet /norestart
4. Check Windows Event Viewer for errors
```

### **Button Not Appearing After Install**

**Solution:**
```
1. Restart Outlook completely
2. Or: Run Setup-CybeReady-Email.bat to reconfigure
3. Restart Outlook again
```

---

## 📋 INSTALLATION TYPES

### **Interactive Installation** (Users)
```cmd
msiexec /i OutlookPhishingReporter.msi
```
- Shows wizard
- User selects options
- User confirms installation
- Takes 2-3 minutes

### **Silent Installation** (Enterprise)
```cmd
msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```
- No wizard
- Automatic
- No user interaction
- Takes 30-60 seconds
- Perfect for Group Policy or SCCM

### **Unattended Installation** (Remote)
```cmd
msiexec /i OutlookPhishingReporter.msi /qn /norestart
```
- No wizard
- No interactive mode
- Fully automated
- For remote deployment

---

## 🎯 FULL WORKFLOW

### **Complete Build and Run**

```batch
REM Step 1: Build DLL
BUILD-DLL.bat

REM Step 2: Wait for "Build successful"

REM Step 3: Build MSI
cd OutlookPhishingReporterSetup\MSISource\
build-msi.bat

REM Step 4: Wait for "Build Complete!"

REM Step 5: Run MSI
cd ..\Distribution\
OutlookPhishingReporter.msi

REM Step 6: Follow installation wizard

REM Step 7: Restart Outlook

REM Step 8: Verify button appears in ribbon

REM Done!
```

---

## 📊 BUILD TIMELINE

| Step | Time | Details |
|------|------|---------|
| Build DLL | 1-2 min | Compile plugin |
| Build MSI | 5-10 min | Create installer |
| Run Setup | 2-3 min | User interaction |
| Restart Outlook | 1-2 min | Plugin loads |
| **Total** | **10-20 min** | First time |

---

## 🎉 SUCCESS INDICATORS

✅ MSI file created (5-10 MB)  
✅ Installation completes without errors  
✅ "Report phishing" button appears in Outlook  
✅ Button clickable and functional  
✅ Registry properly configured  
✅ Test report sends successfully  

---

## 📝 COPY/PASTE COMMANDS

### **Build MSI**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat
```

### **Run Normal Install**
```
msiexec /i "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi"
```

### **Run Silent Install**
```
msiexec /i "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi" /quiet /norestart
```

---

## 🚀 START NOW!

### **Quick Steps**

1. **Build DLL:**
   ```
   Right-click: BUILD-DLL.bat → Run as Administrator
   ```

2. **Build MSI:**
   ```
   Right-click: OutlookPhishingReporterSetup\MSISource\build-msi.bat → Run as Administrator
   ```

3. **Run MSI:**
   ```
   Double-click: OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
   ```

4. **Follow wizard and restart Outlook**

**Done!** 🎉

---

**Your MSI installer is ready to build and deploy!** 📦✅
