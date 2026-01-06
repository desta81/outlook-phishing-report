# 🔨 BUILD OUTLOOK PHISHING REPORTER - DLL & EXE GUIDE

## ⚠️ IMPORTANT - PROJECT TYPE

**This project is a VSTO Outlook Add-in**

It builds as:
- ✅ **DLL File** (OutlookPhishingReporter.dll) - Main plugin
- ✅ **Installer** (EXE setup.exe) - Installation wizard

**Not a standalone EXE application**

---

## 🎯 BUILD OPTIONS

### **Option 1: Build DLL Only** (Quickest)
```
Build plugin as DLL
Time: 1-2 minutes
Output: OutlookPhishingReporter.dll
Use for: Direct installation or MSI packaging
```

### **Option 2: Build with ClickOnce Installer** (Recommended)
```
Build DLL + create installer EXE
Time: 3-5 minutes
Output: setup.exe + installer package
Use for: User installation, enterprise deployment
```

### **Option 3: Build MSI Package** (Enterprise)
```
Build DLL + create MSI installer
Time: 10-15 minutes
Output: OutlookPhishingReporter.msi
Use for: Group Policy, SCCM, enterprise deployment
```

---

## 🚀 QUICK BUILD - DLL FILE

### **Step 1: Open Visual Studio**
```
Open: OutlookPhishingReporter.csproj
Or File → Open Project
Select: OutlookPhishingReporter folder
```

### **Step 2: Select Release Configuration**
```
Top toolbar dropdown: Debug → Release
This optimizes the build for production
```

### **Step 3: Build the Project**
```
Menu: Build → Build Solution
Or: Right-click project → Build
Or: Press: Ctrl + Shift + B
```

### **Step 4: Wait for Build to Complete**
```
Watch Output window:
  ========== Build: 1 succeeded ==========
Output shows path to DLL file
```

### **Step 5: Verify DLL Created**
```
Location: OutlookPhishingReporter\bin\Release\
File: OutlookPhishingReporter.dll
Size: Should be > 100 KB
```

---

## 📦 BUILD WITH INSTALLER - EXE FILE

### **Step 1: Open Project Properties**
```
Right-click project → Properties
Or: Project → [ProjectName] Properties
```

### **Step 2: Go to Publish Tab**
```
Click: Publish tab
Or: Menu → Build → Publish [ProjectName]
```

### **Step 3: Configure Publishing**
```
Publish Location:
  Browse to: OutlookPhishingReporterSetup\Distribution\

Publishing Folder:
  Will create setup.exe here
```

### **Step 4: Click Publish**
```
Publishes the project
Creates:
  - setup.exe (installer)
  - application.manifest
  - OutlookPhishingReporter.dll.manifest
  - .deploy files
```

### **Step 5: Verify Setup.exe Created**
```
Location: OutlookPhishingReporterSetup\Distribution\
File: setup.exe
Size: 500+ KB
```

---

## 📊 BUILD OUTPUTS

### **DLL Build Output**
```
Location: OutlookPhishingReporter\bin\Release\
File: OutlookPhishingReporter.dll
Size: 100+ KB
Usage: Plugin core file
```

### **ClickOnce Installer Output**
```
Location: OutlookPhishingReporterSetup\Distribution\
Files:
  - setup.exe (installer executable)
  - OutlookPhishingReporter.dll.manifest
  - application.manifest
  - OutlookPhishingReporter.dll
  - Various .deploy files
```

### **MSI Installer Output**
```
Location: OutlookPhishingReporterSetup\Distribution\
File: OutlookPhishingReporter.msi
Size: 5-10 MB
Usage: Enterprise deployment
```

---

## 🛠️ BUILD CONFIGURATION

### **Current Configuration**

**From .csproj file:**
```
OutputType: Library (DLL)
TargetFramework: .NET Framework 4.7.2
AssemblyName: OutlookPhishingReporter
OfficeApplication: Outlook
```

**This means:**
- ✅ Builds as DLL add-in
- ✅ Not a standalone application
- ✅ Requires Outlook to run
- ✅ Registers with Windows registry

---

## 📋 STEP-BY-STEP BUILD GUIDE

### **Method 1: Visual Studio UI (Easiest)**

**Step 1: Open Solution**
```
File → Open → Project/Solution
Navigate to: OutlookPhishingReporter.csproj
Click: Open
```

**Step 2: Set Release Configuration**
```
Top toolbar: Debug dropdown → Release
```

**Step 3: Build**
```
Build → Build Solution
Or: Ctrl + Shift + B
```

**Step 4: Check Output**
```
Output window shows:
  ========== Build: 1 succeeded ==========
  
Location: OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
```

---

### **Method 2: Command Line**

**Open Command Prompt:**
```cmd
cd C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter
```

**Build Release:**
```cmd
msbuild OutlookPhishingReporter\OutlookPhishingReporter.csproj /p:Configuration=Release
```

**Build Debug:**
```cmd
msbuild OutlookPhishingReporter\OutlookPhishingReporter.csproj /p:Configuration=Debug
```

---

### **Method 3: PowerShell**

**Open PowerShell as Administrator:**
```powershell
cd "C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter"

# Build Release
msbuild "OutlookPhishingReporter\OutlookPhishingReporter.csproj" /p:Configuration=Release

# Or build Debug
msbuild "OutlookPhishingReporter\OutlookPhishingReporter.csproj" /p:Configuration=Debug
```

---

## ✅ BUILD VERIFICATION

### **Check 1: DLL File Exists**
```
Location: OutlookPhishingReporter\bin\Release\
File: OutlookPhishingReporter.dll
Size: > 100 KB
Date: Today
```

### **Check 2: No Build Errors**
```
Output window shows:
  ========== Build: 1 succeeded ==========
  0 errors, 0 warnings
```

### **Check 3: PDB File (Optional)**
```
Location: OutlookPhishingReporter\bin\Release\
File: OutlookPhishingReporter.pdb
Used for: Debugging symbols
```

---

## 🔍 BUILD CONFIGURATIONS

### **Debug Configuration**
```
Purpose: Development and testing
Optimization: Off
Symbols: Full debugging
Output: OutlookPhishingReporter\bin\Debug\
```

### **Release Configuration**
```
Purpose: Production deployment
Optimization: On
Symbols: PDB only
Output: OutlookPhishingReporter\bin\Release\
```

---

## 📦 AFTER BUILD - NEXT STEPS

### **Option 1: Install DLL Directly**
```
1. Use registry scripts to configure
2. Copy DLL to Outlook addins folder
3. Register with Windows
4. Restart Outlook
```

### **Option 2: Create MSI Installer**
```
1. Build DLL (done above)
2. Navigate to: OutlookPhishingReporterSetup\MSISource\
3. Run: build-msi.bat
4. Output: OutlookPhishingReporter.msi
```

### **Option 3: Use ClickOnce Installer**
```
1. Build in Visual Studio
2. Project → Properties → Publish
3. Configure publish settings
4. Click: Publish
5. Output: setup.exe in Distribution folder
```

---

## 🎯 RECOMMENDED WORKFLOW

### **For Testing**
```
1. Build → Build Solution (Visual Studio)
2. Run: RUN-TEST-SUITE.bat
3. Verify in Outlook
4. Fix any issues
5. Repeat
```

### **For Production**
```
1. Build Release configuration
2. Build → Publish (creates setup.exe)
3. Test on clean machine
4. Create MSI for enterprise deployment
5. Test MSI installation
6. Deploy to users
```

---

## 🆘 TROUBLESHOOTING

### **Build Fails - Missing References**

**Solution:**
```
1. Right-click project → Properties
2. Go to References tab
3. Check for yellow warning triangles
4. Right-click missing reference → Resolve
5. Rebuild
```

### **Build Fails - Visual Studio Version**

**Solution:**
```
1. Open with Visual Studio 2022 or later
2. Or: Use MSBuild command line
3. msbuild OutlookPhishingReporter.csproj /p:Configuration=Release
```

### **DLL Not Created**

**Solution:**
```
1. Check Output window for errors
2. Verify .csproj file syntax
3. Clean solution: Build → Clean Solution
4. Rebuild: Build → Build Solution
5. Check bin\Release folder manually
```

### **Build Takes Too Long**

**Solution:**
```
1. Close unnecessary programs
2. Disable antivirus temporarily
3. Use Release configuration (faster)
4. Check disk space (> 1 GB needed)
```

---

## 📂 OUTPUT LOCATIONS

### **Debug Build**
```
OutlookPhishingReporter\bin\Debug\
├── OutlookPhishingReporter.dll
├── OutlookPhishingReporter.pdb
└── [other dependencies]
```

### **Release Build**
```
OutlookPhishingReporter\bin\Release\
├── OutlookPhishingReporter.dll (production)
├── OutlookPhishingReporter.pdb
└── [other dependencies]
```

### **Published (Installer)**
```
OutlookPhishingReporterSetup\Distribution\
├── setup.exe (installer)
├── OutlookPhishingReporter.dll.manifest
├── application.manifest
└── [deploy files]
```

---

## 🎉 SUCCESS INDICATORS

✅ **Build succeeded** message in Output window  
✅ **0 errors** reported  
✅ **DLL file created** in bin\Release folder  
✅ **File size > 100 KB**  
✅ **Builds in < 5 minutes**  

---

## 📋 QUICK COMMAND REFERENCE

### **Visual Studio Shortcuts**
```
Ctrl + Shift + B       = Build Solution
Ctrl + Shift + F5      = Rebuild Solution
Ctrl + Shift + D       = Publish
F7                     = Build
```

### **Command Line**
```
msbuild OutlookPhishingReporter.csproj /p:Configuration=Release
msbuild OutlookPhishingReporter.csproj /p:Configuration=Debug
```

### **PowerShell**
```powershell
msbuild "OutlookPhishingReporter\OutlookPhishingReporter.csproj" /p:Configuration=Release
```

---

## 🚀 BUILD NOW!

### **Quick Build (1-2 minutes)**
```
1. Open OutlookPhishingReporter.csproj in Visual Studio
2. Change to Release configuration
3. Build → Build Solution
4. Check: OutlookPhishingReporter\bin\Release\OutlookPhishingReporter.dll
```

### **Full Package (5-10 minutes)**
```
1. Build DLL (as above)
2. Build → Publish
3. Check: OutlookPhishingReporterSetup\Distribution\setup.exe
```

---

**Your project is ready to build!** 🔨✅
