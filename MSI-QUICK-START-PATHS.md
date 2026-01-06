# 🎯 MSI BUILD - FULL PATHS QUICK REFERENCE

## 📍 YOUR WORKSPACE LOCATION

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
```

---

## 🚀 BUILD SCRIPT - FULL PATH

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat
```

**How to Run:**
```
Method 1 (Easiest):
  Right-click → Run as Administrator

Method 2 (Command Prompt):
  cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
  build-msi.bat

Method 3 (PowerShell):
  cd 'C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource'
  .\build-msi.bat
```

---

## 📦 OUTPUT MSI - FULL PATH

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

**When Built:**
- Location: Distribution folder
- Size: 5-10 MB
- Ready to: Install or deploy
- Name: OutlookPhishingReporter.msi

---

## 📋 ALL IMPORTANT PATHS

| Item | Full Path |
|------|-----------|
| **Workspace Root** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\` |
| **Plugin Project** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj` |
| **Plugin DLL** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll` |
| **MSI Build Script** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat` |
| **WiX Config** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs` |
| **Output Directory** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\` |
| **MSI Output** | `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi` |

---

## ⚡ STEP-BY-STEP BUILD

### **Step 1: Prerequisites**
```
✓ WiX Toolset installed: https://wixtoolset.org/
✓ Plugin DLL built: Build Solution in Visual Studio
✓ Product.wxs exists: MSISource folder
✓ Admin rights: Required
```

### **Step 2: Navigate to Build Directory**
```
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
```

### **Step 3: Run Build Script**
```
build-msi.bat
Or: Right-click → Run as Administrator
```

### **Step 4: Wait for Completion**
```
Takes: 5-10 minutes
Watch: Compilation and linking progress
Expected: "Build Complete!" message
```

### **Step 5: Verify Output**
```
Navigate to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\
Should see: OutlookPhishingReporter.msi (5-10 MB)
```

---

## 🔗 COPY/PASTE COMMANDS

### **For Command Prompt (Admin)**
```batch
cd /d C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

### **For PowerShell (Admin)**
```powershell
cd 'C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource'
.\build-msi.bat
```

### **Test Install (After Build)**
```batch
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution
msiexec /i OutlookPhishingReporter.msi
```

### **Silent Install (Enterprise)**
```batch
msiexec /i "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi" /quiet /norestart
```

---

## 🎯 QUICK ACTIONS

### **Action 1: Open Build Folder in File Explorer**
```
Press: Win + E
Copy/Paste in address bar:
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
Press: Enter
```

### **Action 2: Run Build Now**
```
In the MSISource folder above:
Right-click: build-msi.bat
Select: Run as Administrator
```

### **Action 3: Check Output**
```
After build completes:
Navigate to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\
You should see: OutlookPhishingReporter.msi
```

---

## 🔍 FOLDER STRUCTURE

```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
│
├── OutlookPhishingReporter\                    (Plugin source)
│   ├── bin\Release\
│   │   └── OutlookPhishingReporter.dll         ← DLL OUTPUT
│   └── OutlookPhishingReporter.csproj          ← Build this first
│
└── OutlookPhishingReporterSetup\
    ├── MSISource\                              (MSI build)
    │   ├── build-msi.bat                       ← RUN THIS
    │   ├── Product.wxs                         ← Config
    │   └── [other WiX files]
    │
    └── Distribution\                           (MSI output)
        └── OutlookPhishingReporter.msi         ← FINAL OUTPUT
```

---

## ✅ SUCCESS CHECKLIST

- [ ] WiX Toolset installed
- [ ] Plugin DLL built
- [ ] Navigated to MSISource folder
- [ ] Ran build-msi.bat as Administrator
- [ ] Build completed successfully
- [ ] MSI file created in Distribution folder
- [ ] File size 5-10 MB
- [ ] Ready to deploy

---

## 🚀 READY TO BUILD!

### **Easiest Method (One Click)**

1. **Open File Explorer:** Press `Win + E`

2. **Navigate to:**
   ```
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\
   ```

3. **Right-click:** `build-msi.bat`

4. **Select:** `Run as Administrator`

5. **Wait:** 5-10 minutes

6. **Check Output:**
   ```
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
   ```

---

## 📞 TROUBLESHOOTING

**Build fails - WiX not found?**
```
1. Install from: https://wixtoolset.org/
2. Restart computer
3. Try again
```

**Build fails - DLL not found?**
```
1. Build plugin: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj
2. Build → Build Solution in Visual Studio
3. Try MSI build again
```

**MSI not created?**
```
1. Check Distribution folder
2. Look for: OutlookPhishingReporter.msi
3. Check file size (should be 5-10 MB)
4. If missing, check build errors
```

---

## 🎉 COMPLETE PATH REFERENCE

### **Copy These Exact Paths**

**Build Script:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\build-msi.bat
```

**Output MSI:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

**Plugin DLL:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
```

---

**All paths provided! Ready to build your MSI!** 📦✅
