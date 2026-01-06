# ?? MSI BUILD - PREREQUISITES & NEXT STEPS

## ?? STATUS: WiX TOOLSET REQUIRED

To build the MSI installer, you need to install **WiX Toolset** first.

---

## ?? INSTALLATION STEPS (15 minutes)

### Step 1: Download WiX Toolset (2-5 min)
- **Visit:** https://wixtoolset.org/
- **Click:** Download
- **Select:** WiX Toolset v3.14 (or latest v3.x)
- **Save:** The installer file

### Step 2: Install WiX Toolset (3-5 min)
1. **Double-click** the downloaded installer
2. **Follow** the installation wizard
3. **Accept** all default settings
4. **Click** Install
5. **Wait** for installation to complete

### Step 3: Restart Computer (Required)
- **Restart** your computer to complete installation

### Step 4: Build the MSI (3 min)
Navigate to: `OutlookPhishingReporterSetup\MSISource\`

**Option A - Automatic Build:**
- Right-click: `build-msi.bat`
- Select: **Run as Administrator**
- MSI will be created automatically

**Option B - Combined Install & Build:**
- Right-click: `install-wix-and-build.bat`
- Select: **Run as Administrator**
- This script helps with WiX installation and builds MSI

---

## ?? WHAT TO DO NOW

### Immediate Actions
1. Visit: **https://wixtoolset.org/**
2. Download: **WiX Toolset v3.14** (or latest)
3. Run: **The installer**
4. Restart: **Your computer**

### After Restart
1. Navigate to: `OutlookPhishingReporterSetup\MSISource\`
2. Run: `build-msi.bat` (as Administrator)
3. MSI will be created at: `Distribution\OutlookPhishingReporter.msi`

---

## ?? KEY FILES

| File | Purpose | Location |
|------|---------|----------|
| **build-msi.bat** | MSI build script | MSISource\ |
| **install-wix-and-build.bat** | WiX install helper | MSISource\ |
| **RUN_MSI_BUILD.md** | Detailed instructions | MSISource\ |

---

## ?? QUICK SUMMARY

```
1. Go to: https://wixtoolset.org/
2. Download WiX Toolset
3. Install it
4. Restart computer
5. Run: OutlookPhishingReporterSetup\MSISource\build-msi.bat
6. MSI created in: Distribution\ folder
```

---

## ? BUILD TIME BREAKDOWN

| Step | Time |
|------|------|
| Download WiX | 2-5 min |
| Install WiX | 3-5 min |
| Restart | 1-2 min |
| Build MSI | 3 min |
| **Total** | **~15 min** |

---

## ?? FULL INSTRUCTIONS

For detailed step-by-step instructions, see:
- **RUN_MSI_BUILD.md** (in MSISource folder)
- **MSI_BUILD_COMPLETE_GUIDE.md** (for enterprise deployment)

---

## ?? TIPS

- **Admin Rights:** Run build script as Administrator
- **Disk Space:** Ensure 500 MB free space
- **Internet:** Needed to download WiX (first time only)
- **Restart:** Essential after WiX installation

---

## ?? AFTER MSI IS BUILT

Once the MSI is successfully created:

1. **Test Installation:**
   ```
   msiexec /i Distribution\OutlookPhishingReporter.msi
   ```

2. **Deploy to Users:**
   - Use Group Policy
   - Use SCCM
   - Share on network

3. **See Deployment Guide:**
   - `MSI_BUILD_INSTRUCTIONS.md`
   - `MSI_BUILD_COMPLETE_GUIDE.md`

---

## ? SUCCESS INDICATORS

After following these steps, you should:
- ? Have WiX Toolset installed
- ? Have built the MSI successfully
- ? Have MSI file at: `Distribution\OutlookPhishingReporter.msi`
- ? Be ready to deploy to users

---

**Next Step: Download WiX Toolset from https://wixtoolset.org/**

*Estimated total time: 15 minutes*
