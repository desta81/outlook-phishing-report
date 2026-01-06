# ?? RUN MSI BUILD - Complete Instructions

## ?? IMPORTANT: WiX Toolset Not Found

WiX Toolset is **not installed** on your system. You need to install it first before building the MSI.

---

## ?? Step 1: Install WiX Toolset

### Option A: Automatic Installation Helper (Recommended)

1. **Navigate to:** `OutlookPhishingReporterSetup\MSISource\`
2. **Right-click:** `install-wix-and-build.bat`
3. **Select:** "Run as Administrator"
4. **Follow prompts** to download and install WiX
5. **Restart computer** when prompted
6. **Script will automatically build MSI** after installation

### Option B: Manual Download & Installation

**Step 1: Download WiX Toolset**
1. Visit: https://wixtoolset.org/
2. Click "Download"
3. Download WiX v3.14 (or latest v3.x)
4. Save the installer (e.g., `wix314.exe`)

**Step 2: Install**
1. Double-click the WiX installer
2. Follow the installation wizard
3. Accept all defaults
4. Click "Install"
5. Wait for installation to complete

**Step 3: Restart Computer**
- Installation may require restart
- Restart your computer to complete installation

**Step 4: Build the MSI**
- Navigate to: `OutlookPhishingReporterSetup\MSISource\`
- Right-click: `build-msi.bat`
- Select: "Run as Administrator"
- Wait for build to complete

---

## ?? Step 2: Build the MSI

### After WiX is Installed

**Method 1: Automatic Build (Easiest)**
```
1. Navigate to: OutlookPhishingReporterSetup\MSISource\
2. Right-click: build-msi.bat
3. Select: Run as Administrator
4. Wait for success message
```

**Method 2: Combined Install & Build**
```
1. Navigate to: OutlookPhishingReporterSetup\MSISource\
2. Right-click: install-wix-and-build.bat
3. Select: Run as Administrator
4. Script handles installation and build
```

**Method 3: Manual Commands**
```cmd
cd OutlookPhishingReporterSetup\MSISource
"C:\Program Files (x86)\WiX Toolset v3.14\bin\candle.exe" Product.wxs
"C:\Program Files (x86)\WiX Toolset v3.14\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
```

---

## ? Verify Build Success

After running the build script, check:

1. **No error messages** displayed
2. **File created:** `OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi`
3. **File size:** 2-5 MB
4. **Success message** from script

---

## ?? What Happens During Build

### Compilation Phase
- WiX compiler (candle.exe) processes `Product.wxs`
- Creates `Product.wixobj` object file
- Takes ~1-2 seconds

### Linking Phase
- WiX linker (light.exe) links object file
- Creates final MSI file in Distribution folder
- Takes ~1-2 seconds

### Output
```
? Product.wixobj created
? OutlookPhishingReporter.msi created
? Build complete!
```

---

## ?? Quick Reference

| Task | Steps |
|------|-------|
| **Install WiX** | Visit https://wixtoolset.org/ ? Download ? Install ? Restart |
| **Build MSI** | Run `build-msi.bat` as Administrator |
| **Verify** | Check `Distribution\OutlookPhishingReporter.msi` exists |
| **Test** | `msiexec /i OutlookPhishingReporter.msi` |
| **Deploy** | See MSI_BUILD_INSTRUCTIONS.md |

---

## ?? Troubleshooting

### "WiX not found" Error

**Solution:**
1. Download WiX from: https://wixtoolset.org/
2. Install WiX Toolset v3.14
3. **Restart your computer**
4. Try again

### "Product.wxs not found" Error

**Solution:**
1. Make sure you're in the `MSISource` folder
2. Verify `Product.wxs` file exists
3. Try again

### "Distribution folder not found" Error

**Solution:**
- The script creates it automatically
- Or manually create it: `mkdir ..\Distribution`

### Build still fails

**Solution:**
1. Run as Administrator
2. Check Windows event logs
3. Try manual commands above
4. Review build script output for specific errors

---

## ?? File Locations

| File | Purpose | Location |
|------|---------|----------|
| **build-msi.bat** | MSI build script | MSISource\ |
| **install-wix-and-build.bat** | WiX install + build | MSISource\ |
| **Product.wxs** | WiX source template | MSISource\ |
| **OutlookPhishingReporter.msi** | Built MSI (output) | Distribution\ |

---

## ?? Steps Summary

### Quick Path (30 minutes)
1. Download WiX: https://wixtoolset.org/
2. Install WiX
3. Restart computer
4. Run `build-msi.bat`
5. Done! MSI ready for deployment

### Automated Path (15 minutes)
1. Run `install-wix-and-build.bat`
2. Follow prompts
3. Restart when prompted
4. Done! MSI automatically built

---

## ? Next Steps After Build

1. **Test MSI:**
   ```
   msiexec /i "OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi"
   ```

2. **Deploy to Users:**
   - See: `MSI_BUILD_INSTRUCTIONS.md` > Deployment section
   - Group Policy
   - SCCM
   - Network share

3. **Verify Installation:**
   - Check Outlook for "Report Phishing" button
   - Verify registry entries created
   - Test with sample email

---

## ?? Support

### For WiX Installation Issues
- Visit: https://wixtoolset.org/
- Check system requirements
- Ensure administrator rights

### For Build Issues
- Review build script output
- Check file permissions
- Try running as Administrator
- See MSI_BUILD_GUIDE.md for troubleshooting

### For Deployment Questions
- See: MSI_BUILD_INSTRUCTIONS.md
- Review enterprise deployment patterns
- Check Group Policy or SCCM documentation

---

## ?? You're Ready!

Choose your path:

**Option 1:** Install WiX manually, then run `build-msi.bat`  
**Option 2:** Run `install-wix-and-build.bat` for everything  
**Option 3:** Read `MSI_BUILD_COMPLETE_GUIDE.md` for detailed instructions

---

**Once WiX is installed, building the MSI takes only 3 minutes!**

Start with: **Download WiX from https://wixtoolset.org/**
