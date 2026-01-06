# 🚀 FULL PATH TO RUN MSI BUILD

## 📍 YOUR WORKSPACE LOCATION

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
```

---

## 📦 MSI BUILD FULL PATH

### **Build Script Location**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat
```

### **Output Location (Where MSI Gets Saved)**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

---

## ⚡ QUICK RUN - 3 METHODS

### **Method 1: File Explorer (Easiest)**

**Step 1: Open File Explorer**
```
Press: Win + E
```

**Step 2: Navigate to:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
```

**Step 3: Run Build**
```
Right-click: build-msi.bat
Select: Run as Administrator
```

**Step 4: Wait for Completion**
```
Script runs 5-10 minutes
Displays success message
```

---

### **Method 2: Command Prompt (CD Method)**

**Step 1: Open Command Prompt as Administrator**
```
Press: Win + R
Type: cmd
Press: Ctrl + Shift + Enter
Or right-click cmd → Run as Administrator
```

**Step 2: Navigate to Build Directory**
```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
```

**Step 3: Run Build**
```cmd
build-msi.bat
```

**Step 4: Wait for Completion**
```
Watch progress
MSI created in Distribution folder
```

---

### **Method 3: PowerShell (Advanced)**

**Step 1: Open PowerShell as Administrator**
```
Right-click Start Menu → Windows PowerShell (Admin)
```

**Step 2: Navigate to Build Directory**
```powershell
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\"
```

**Step 3: Run Build**
```powershell
.\build-msi.bat
```

**Step 4: Wait for Completion**
```
Watch progress
MSI created in Distribution folder
```

---

## 🎯 FULL PATH REFERENCE

### **Build Script**
```
Full Path: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat
Filename: build-msi.bat
Folder: MSISource
```

### **WiX Source File**
```
Full Path: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs
Filename: Product.wxs
Folder: MSISource
```

### **Output MSI**
```
Full Path: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
Filename: OutlookPhishingReporter.msi
Folder: Distribution
Size: 5-10 MB (typical)
```

### **Plugin DLL Source**
```
Full Path: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
Filename: OutlookPhishingReporter.dll
Folder: Release
```

---

## 🔧 PREREQUISITES BEFORE BUILDING

### **1. WiX Toolset Installation**
```
Download: https://wixtoolset.org/
Install to: C:\Program Files (x86)\WiX Toolset v3.14\
Verify: candle.exe and light.exe exist
```

### **2. Build Plugin DLL**
```
Open: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj
Build: Build → Build Solution
Output: OutlookPhishingReporter.dll
```

### **3. Check Product.wxs**
```
Location: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs
Must exist: Yes
```

---

## 📋 BUILD PROCESS SUMMARY

### **Complete Build Command**
```batch
REM Navigate to MSISource
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\

REM Run build script
build-msi.bat

REM Output will be in:
REM C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

---

## ✅ AFTER BUILD - DEPLOYMENT

### **Test Installation**
```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\

msiexec /i OutlookPhishingReporter.msi
```

### **Silent Installation** (Enterprise)
```cmd
msiexec /i "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi" /quiet /norestart
```

### **Copy to Network**
```
From: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
To: Network share for enterprise deployment
```

---

## 📊 DIRECTORY STRUCTURE

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
├── OutlookPhishingReporter\
│   ├── bin\
│   │   └── Release\
│   │       └── OutlookPhishingReporter.dll ← Plugin DLL
│   └── OutlookPhishingReporter.csproj ← Build this first
│
└── OutlookPhishingReporterSetup\
    ├── MSISource\
    │   ├── build-msi.bat ← Run this to build MSI
    │   ├── Product.wxs ← MSI configuration
    │   └── [other WiX files]
    │
    └── Distribution\
        └── OutlookPhishingReporter.msi ← Output goes here
```

---

## 🚀 QUICK START COMMANDS

### **Copy/Paste Ready Commands**

**For Command Prompt:**
```batch
cd /d C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

**For PowerShell:**
```powershell
cd 'C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource'
.\build-msi.bat
```

---

## ⏱️ BUILD TIMELINE

| Step | Time | Details |
|------|------|---------|
| **Navigate to folder** | 1 min | Open directory |
| **Check prerequisites** | 2 min | WiX, DLL built |
| **Run build-msi.bat** | 5-10 min | Compile & link |
| **Verify MSI created** | 1 min | Check Distribution folder |
| **Total** | 9-15 min | Full build time |

---

## 🎉 SUCCESS INDICATOR

When build completes successfully:

```
✓ MSI file appears in:
  C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi

✓ File size: 5-10 MB

✓ Message shown:
  "Build Complete!"
  "[+] MSI File Created Successfully"
  "Location: ...Distribution\OutlookPhishingReporter.msi"
```

---

## 🔍 VERIFY BUILD OUTPUT

**Navigate to Output Folder:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\
```

**Should See:**
```
- OutlookPhishingReporter.msi (5-10 MB)
- File created today
- Can be double-clicked to install
```

---

## 💡 TROUBLESHOOTING

### **If Build Fails: WiX Not Found**
```
1. Check: C:\Program Files (x86)\WiX Toolset v3.14\bin\
2. If missing: Download from https://wixtoolset.org/
3. Install WiX
4. Restart computer
5. Try build again
```

### **If Build Fails: Product.wxs Not Found**
```
1. Verify file exists:
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs
2. If missing: Restore from repository
3. Try build again
```

### **If DLL Not Found**
```
1. Build plugin first:
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj
2. Build → Build Solution in Visual Studio
3. Try MSI build again
```

---

## 📝 COPY/PASTE ALL PATHS

```
Workspace Root:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\

MSI Build Script:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat

MSI Output:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi

Plugin Source:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj

Plugin DLL:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll

WiX Source:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs
```

---

## 🎯 READY TO BUILD!

**Option 1: File Explorer (Easiest)**
```
1. Press Win + E
2. Navigate to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
3. Right-click: build-msi.bat
4. Select: Run as Administrator
5. Wait 5-10 minutes
6. MSI ready in Distribution folder
```

**Option 2: Command Prompt**
```
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

**Option 3: PowerShell**
```powershell
cd 'C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource'
.\build-msi.bat
```

---

**All paths provided! Ready to build your MSI!** 📦✅
