# ??? UNINSTALLATION GUIDE

## Overview

This guide explains how to uninstall the Outlook Phishing Reporter add-in from your system.

---

## ? QUICK UNINSTALL (30 Seconds)

### Option 1: Double-Click Uninstaller (Easiest)

1. Navigate to: `OutlookPhishingReporterSetup\Release\`
2. Double-click: **`Uninstall.vbs`**
3. Click: **"YES"** to confirm uninstallation
4. Done! ?

**The add-in is now removed from your system.**

---

## ?? UNINSTALLATION METHODS

### Method 1: VBScript Uninstaller (Recommended)

**File:** `Uninstall.vbs`

**Steps:**
1. Double-click: `Uninstall.vbs`
2. Confirm when asked
3. Outlook closes automatically
4. Registry entries removed
5. Complete!

**Time:** Less than 1 minute  
**Difficulty:** Very Easy

---

### Method 2: Batch File Uninstaller

**File:** `Uninstall.bat`

**Steps:**
1. Right-click: `Uninstall.bat`
2. Select: "Run as Administrator"
3. Confirm when asked
4. Outlook closes automatically
5. Registry entries removed
6. Complete!

**Time:** Less than 1 minute  
**Difficulty:** Very Easy

---

### Method 3: PowerShell Uninstaller

**File:** `Uninstall.ps1`

**Steps:**
1. Open PowerShell as Administrator
2. Navigate to: `OutlookPhishingReporterSetup\Release\`
3. Run: `.\Uninstall.ps1`
4. Confirm when asked
5. Registry entries removed
6. Complete!

**Alternative (force without prompting):**
```powershell
.\Uninstall.ps1 -Force
```

**Time:** Less than 1 minute  
**Difficulty:** Easy

---

### Method 4: Manual Registry Deletion

**Manual Steps:**

1. Press: `Win + R`
2. Type: `regedit`
3. Click: OK

4. Navigate to:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
   ```

5. Right-click: `OutlookPhishingReporter` folder
6. Select: **Delete**
7. Confirm deletion
8. Close Registry Editor

9. Close Outlook if open
10. Reopen Outlook

**Time:** 2-3 minutes  
**Difficulty:** Intermediate

---

### Method 5: Group Policy Removal (Enterprise)

**For IT Administrators removing via Group Policy:**

1. Open `gpedit.msc` on domain controller
2. Navigate to: Computer Configuration ? Software Settings ? Software Installation
3. Find: "Outlook Phishing Reporter"
4. Right-click ? "All Tasks" ? **"Remove"**
5. Choose: "Immediately uninstall software from computers"
6. Apply policy to target OUs
7. Policy removes add-in from all users

**Time:** 5-10 minutes  
**Difficulty:** Advanced

---

## ? VERIFICATION CHECKLIST

After uninstalling, verify the removal:

- [ ] Outlook opens without errors
- [ ] "Report Phishing" button no longer appears in ribbon
- [ ] Registry key deleted: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
- [ ] No leftover files (optional - can be deleted)

---

## ?? WHAT GETS REMOVED

The uninstallation process removes:

? **Registry Entries**
- Add-in registry key
- Configuration settings
- Email address settings
- Custom icon settings
- Button label settings

?? **NOT Removed** (Manual cleanup if desired):
- Installation files (.dll, .vsto, .manifest)
- Documentation files
- This folder itself

---

## ?? WHAT STAYS AFTER UNINSTALL

These files remain after uninstallation (can be manually deleted if desired):

- `OutlookPhishingReporter.dll`
- `OutlookPhishingReporter.vsto`
- `OutlookPhishingReporter.dll.manifest`
- `OutlookPhishingReporter.pdb`
- Documentation and guide files
- Installer scripts

**To completely clean up:**
1. Delete the entire: `OutlookPhishingReporterSetup` folder
2. Empty Recycle Bin

---

## ?? COMMON SCENARIOS

### I want to uninstall and reinstall

1. Run: `Uninstall.vbs`
2. Later, run: `Install.vbs` to reinstall

### I want to change the configuration

**No need to uninstall!**
- Run: `QuickFix.bat` to reconfigure email
- Or run: `Install.vbs` again

### I'm replacing the add-in with a new version

1. Run: `Uninstall.vbs`
2. Install new version when ready

### I want to remove it from all users (Enterprise)

Use Group Policy method (Method 5) to remove from all computers at once.

---

## ?? TROUBLESHOOTING

### Error: "Uninstaller requires administrator privileges"

**Solution:**
- Right-click the uninstaller
- Select: "Run as Administrator"

### Error: "Could not remove registry entries"

**Solution:**
- Make sure Outlook is completely closed
- Run uninstaller again
- Or use Method 4 (Manual Registry Deletion)

### Button still appears after uninstalling

**Solution:**
1. Restart Outlook completely
2. Close all Outlook windows
3. Reopen Outlook
4. If button still appears, manually delete registry key (Method 4)

### "Add-in may not be installed" message

**This is OK!**
- Means the registry entry wasn't found
- The add-in was either never installed or already removed
- You can safely ignore this message

---

## ?? FILES PROVIDED

| File | Purpose | Difficulty |
|------|---------|-----------|
| **Uninstall.vbs** | VBScript uninstaller | Very Easy |
| **Uninstall.bat** | Batch file uninstaller | Very Easy |
| **Uninstall.ps1** | PowerShell uninstaller | Easy |
| **UNINSTALL_GUIDE.md** | This guide | Reference |

---

## ?? QUICK REFERENCE

| Task | Command | Time |
|------|---------|------|
| **Quick uninstall** | Double-click Uninstall.vbs | <1 min |
| **Admin uninstall** | Right-click Uninstall.bat | <1 min |
| **PowerShell uninstall** | ./Uninstall.ps1 | <1 min |
| **Manual deletion** | Via regedit | 2-3 min |
| **Enterprise removal** | Group Policy | 5-10 min |

---

## ?? TIME ESTIMATES

- **Quick removal (VBS):** Less than 1 minute
- **Registry method:** 2-3 minutes
- **Complete cleanup:** 5 minutes (includes deleting files)

---

## ?? SAFETY NOTES

? **Safe to uninstall anytime**
- No system files modified
- No critical dependencies
- Easy to reinstall

?? **Before uninstalling:**
- Close Outlook (uninstaller will do this)
- Save any open emails
- No active Outlook tasks

---

## ?? SUPPORT

### If uninstallation fails:

1. **Verify Outlook is closed:**
   - Ctrl + Alt + Delete
   - Task Manager
   - End any OUTLOOK.EXE processes

2. **Try again:**
   - Run uninstaller as Administrator
   - Use Method 4 (Manual Registry)

3. **Manual cleanup:**
   ```
   reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" /f
   ```

---

## ? SUMMARY

**Uninstalling is simple:**

1. **Use Method 1 (VBScript)** - Fastest
2. **Confirm when asked** - Double-check you want to remove
3. **Restart Outlook** - (Optional, but recommended)
4. **Done!** - Add-in is removed

---

**Choose your uninstallation method above and follow the steps.** ?

All methods are safe and will completely remove the add-in.
