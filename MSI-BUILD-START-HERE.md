# 🚀 MSI BUILD - START HERE

## ⚠️ ISSUE DETECTED

The MSI build script had a WiX detection issue. This guide will help you build the MSI successfully.

---

## ✅ SOLUTION - 3 SIMPLE STEPS

### **Step 1: Install WiX Toolset (One-Time)**

**Download WiX:**
```
Go to: https://wixtoolset.org/
Download: WiX v3.14 (or latest v3.x)
```

**Installation Path:**
```
Default: C:\Program Files (x86)\WiX Toolset v3.14\
```

**Verify Installation:**
```
Open File Explorer
Navigate to: C:\Program Files (x86)\WiX Toolset v3.14\bin\
Should see: candle.exe and light.exe
```

---

### **Step 2: Build Plugin DLL**

**Open Visual Studio:**
```
File → Open → Project/Solution
Select: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\OutlookPhishingReporter.csproj
```

**Build the Plugin:**
```
Build → Build Solution
Wait for: "Build succeeded" message
Output: OutlookPhishingReporter.dll created
```

**Verify DLL:**
```
Location: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
Should exist and be > 100 KB
```

---

### **Step 3: Build MSI**

**Open Command Prompt as Administrator:**
```
Right-click Start Menu
Select: Command Prompt (Admin)
Or: PowerShell (Admin)
```

**Run Build Command:**
```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

**Wait for Completion:**
```
Script will:
  1. Check for WiX
  2. Compile WiX source
  3. Link to create MSI
  4. Display success message

Takes: 5-10 minutes
```

**Check Output:**
```
Location: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\
File: OutlookPhishingReporter.msi (5-10 MB)
```

---

## 🔧 TROUBLESHOOTING

### **Error: WiX Not Found**

**Solution:**
```
1. Download WiX from https://wixtoolset.org/
2. Install to default location
3. Verify: C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe exists
4. Restart computer
5. Try build again
```

### **Error: DLL Not Found**

**Solution:**
```
1. Build plugin in Visual Studio first
2. Build → Build Solution
3. Verify DLL created: bin\Release\OutlookPhishingReporter.dll
4. Try MSI build again
```

### **Error: Product.wxs Not Found**

**Solution:**
```
1. Verify file exists:
   C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource\Product.wxs
2. If missing, restore from repository
3. Try build again
```

---

## 📋 COMPLETE CHECKLIST

Before running build, verify:

- [ ] WiX Toolset installed
  - [ ] Location: C:\Program Files (x86)\WiX Toolset v3.14\
  - [ ] Files exist: candle.exe, light.exe

- [ ] Plugin DLL built
  - [ ] Location: OutlookPhishingReporter\bin\Release\
  - [ ] File: OutlookPhishingReporter.dll
  - [ ] Size: > 100 KB

- [ ] MSI source files ready
  - [ ] Product.wxs exists: MSISource\Product.wxs
  - [ ] build-msi.bat exists: MSISource\build-msi.bat

- [ ] Command prompt open
  - [ ] As Administrator: Yes
  - [ ] In correct directory: MSISource folder

---

## 🎯 BUILD NOW

**Copy/Paste Command:**
```batch
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource && build-msi.bat
```

**Or Step by Step:**
```batch
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

---

## ✨ EXPECTED OUTPUT

```
===================================================
Outlook Phishing Reporter - MSI Build Script
===================================================

[*] Working Directory: ...MSISource\
[*] Output Directory: ...Distribution\

[*] Searching for WiX Toolset...
[+] Found WiX v3.14
[+] WiX Path: C:\Program Files (x86)\WiX Toolset v3.14\bin

[1/2] Compiling WiX source file...
[+] Compilation successful

[2/2] Linking and creating MSI...
[+] Linking successful

===================================================
Build Complete!
===================================================

[+] MSI File Created Successfully
Location: ...Distribution\OutlookPhishingReporter.msi
Size: [SIZE] bytes
```

---

## 🎉 WHEN BUILD SUCCEEDS

**MSI Created:**
```
C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi
```

**Next Steps:**
```
1. Test installation:
   Double-click OutlookPhishingReporter.msi
   
2. Or silent install:
   msiexec /i OutlookPhishingReporter.msi /quiet /norestart
   
3. Deploy to network:
   Copy MSI to shared location for enterprise deployment
```

---

## 📞 NEED MORE HELP?

**See detailed guides:**
- MSI-BUILD-COMPLETE-GUIDE.md (comprehensive)
- MSI-BUILD-FULL-PATHS.md (all paths)
- MSI-QUICK-START-PATHS.md (quick reference)

---

**Ready to build! Follow the 3 steps above.** 🚀
