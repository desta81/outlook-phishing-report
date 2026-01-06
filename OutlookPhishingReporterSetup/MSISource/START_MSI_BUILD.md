# ?? MSI BUILD - Ready to Start

## ? What You Have

You now have everything needed to build the MSI installer:

1. ? **Product.wxs** - WiX source template (pre-configured)
2. ? **build-msi.bat** - Automated build script
3. ? **BUILD_MSI_GUIDE.md** - Detailed instructions
4. ? **MSI_BUILD_COMPLETE_GUIDE.md** - Step-by-step guide

---

## ?? Quick Start (5 Minutes)

### For Experts Only:

1. **Install WiX Toolset:** https://wixtoolset.org/
2. **Navigate to:** `OutlookPhishingReporterSetup\MSISource\`
3. **Run:** `build-msi.bat`
4. **Done!** MSI created at `Distribution\OutlookPhishingReporter.msi`

---

## ?? Full Instructions (15 Minutes)

### Step 1: Install WiX Toolset (5 minutes)

1. Visit: https://wixtoolset.org/
2. Download: WiX Toolset v3.14 (or latest v3.x)
3. Run the installer
4. Follow setup wizard
5. **Restart computer**

**Verify Installation:**
```cmd
candle.exe -version
```

Should show version number.

### Step 2: Build the MSI (3 minutes)

**Method A: Automatic (Easiest)**
1. Navigate to: `OutlookPhishingReporterSetup\MSISource\`
2. Right-click: `build-msi.bat`
3. Select: "Run as Administrator"
4. Wait for completion message
5. Done!

**Method B: Manual**
```cmd
cd OutlookPhishingReporterSetup\MSISource
"C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" Product.wxs
"C:\Program Files (x86)\WiX Toolset v3.14\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

### Step 3: Verify (2 minutes)

1. Open: `OutlookPhishingReporterSetup\Distribution\`
2. Look for: `OutlookPhishingReporter.msi` file
3. Size should be: 2-5 MB

### Step 4: Test (5 minutes)

1. Double-click the MSI
2. Click "Install"
3. Restart Outlook
4. Look for "Report Phishing" button

---

## ?? File Locations

| File | Purpose | Location |
|------|---------|----------|
| **build-msi.bat** | Build script | MSISource\ |
| **Product.wxs** | WiX template | MSISource\ |
| **OutlookPhishingReporter.msi** | Built MSI | Distribution\ |
| **BUILD_MSI_GUIDE.md** | Quick guide | MSISource\ |
| **MSI_BUILD_COMPLETE_GUIDE.md** | Full guide | MSISource\ |
| **MSI_BUILD_INSTRUCTIONS.md** | Enterprise guide | MSISource\ |

---

## ?? Requirements

| Item | Requirement |
|------|-------------|
| **WiX Toolset** | v3.11 or later |
| **Windows** | Windows 7 or later |
| **Disk Space** | 100 MB free |
| **Admin Rights** | Required for build |

---

## ?? Next Steps

1. **If WiX not installed:**
   - Download from: https://wixtoolset.org/
   - Install and restart computer

2. **To build MSI:**
   - Read: `BUILD_MSI_COMPLETE_GUIDE.md` (Phase 1 & 2)
   - OR Run: `build-msi.bat`

3. **To deploy:**
   - Read: `MSI_BUILD_INSTRUCTIONS.md` (Phase 5 & 6)
   - Deploy via Group Policy or SCCM

---

## ?? Need Help?

**For Build Issues:**
- Review: `BUILD_MSI_GUIDE.md` > Troubleshooting
- Check WiX installation
- Run build script with Administrator rights

**For Deployment Questions:**
- Review: `MSI_BUILD_INSTRUCTIONS.md`
- Check Group Policy documentation
- Check SCCM deployment guides

---

## ? Summary

| Phase | Status | Time |
|-------|--------|------|
| 1. Install WiX | If needed | 5-10 min |
| 2. Build MSI | Ready now | 3 min |
| 3. Test MSI | Ready now | 5 min |
| 4. Deploy | Ready now | Varies |

---

**Status: MSI BUILD SYSTEM COMPLETE AND READY! ??**

**Start here:** 

? Run `build-msi.bat` in `OutlookPhishingReporterSetup\MSISource\` folder

OR

? Read `MSI_BUILD_COMPLETE_GUIDE.md` for step-by-step instructions

---

*Last Updated: January 5, 2026*
