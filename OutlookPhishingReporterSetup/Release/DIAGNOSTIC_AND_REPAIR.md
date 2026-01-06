# 🔧 DIAGNOSTIC & REPAIR GUIDE

## The Error You're Seeing

```
Configuration error: Invalid admin email address. Please contact your administrator.
```

This error means the **AdminEmail registry value is missing or invalid**.

---

## 🚀 QUICK FIX (Choose One)

### Option 1: Automatic Repair (Easiest - Recommended)

**Batch File Version:**
```
1. Right-click: Diagnose-and-Repair.bat
2. Select: Run as Administrator
3. Enter email when prompted
4. Restart Outlook
Done!
```

**PowerShell Version:**
```
1. Open PowerShell as Administrator
2. Run: .\Diagnose-and-Repair.ps1 -Repair
3. Enter email when prompted
4. Restart Outlook
Done!
```

**Time:** 2 minutes

---

### Option 2: Manual Registry Fix

1. **Press:** `Win + R`
2. **Type:** `regedit`
3. **Navigate to:**
   ```
   HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
   ```
4. **Create/Edit these values:**
   - **AdminEmail** (String): `security@company.com`
   - **LoadBehavior** (DWORD): `3`
   - **Description** (String): `Outlook Phishing Reporter`

5. **Restart Outlook**

**Time:** 3-5 minutes

---

### Option 3: Re-run Installer

1. **Navigate to:** `OutlookPhishingReporterSetup\Release\`
2. **Double-click:** `Install.vbs`
3. **Enter email** when prompted
4. **Restart Outlook**

**Time:** 3 minutes

---

## 🔍 WHAT THE DIAGNOSTIC TOOLS DO

### Diagnose-and-Repair.bat

**Diagnostic Mode (No Arguments):**
- Checks if registry key exists
- Verifies AdminEmail is set
- Verifies LoadBehavior = 3
- Validates Description
- Confirms Manifest path

**Repair Mode (Automatic):**
- Closes Outlook
- Removes old configuration
- Prompts for valid email
- Recreates all registry values
- Restarts configuration

### Diagnose-and-Repair.ps1

**Diagnostic Mode (No Arguments):**
```powershell
.\Diagnose-and-Repair.ps1
```
- Runs all diagnostics
- Reports status
- No changes made

**Repair Mode (With -Repair Flag):**
```powershell
.\Diagnose-and-Repair.ps1 -Repair
```
- Runs diagnostics
- If issues found, launches repair
- Fixes configuration
- Requires admin rights

---

## 📋 VALID EMAIL FORMATS

### ✅ CORRECT
- `security@company.com`
- `admin@yourdomain.org`
- `phishing.report@example.co.uk`
- `user+tag@domain.com`

### ❌ INCORRECT
- `security` (no @domain)
- `@company.com` (no username)
- `admin @ company.com` (spaces)
- `admin@` (no extension)

---

## 🔧 WHAT GETS CHECKED/FIXED

| Item | Purpose | Required |
|------|---------|----------|
| **AdminEmail** | Where reports get sent | YES |
| **LoadBehavior** | Auto-load add-in | YES (=3) |
| **Description** | Display name | YES |
| **Manifest** | Add-in location | Optional |

---

## 🆘 IF REPAIR FAILS

### Step 1: Manual Registry Cleanup
```
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" /f
```

### Step 2: Close Outlook Completely
- Ctrl + Alt + Delete
- Task Manager
- End all OUTLOOK.EXE processes

### Step 3: Re-run Repair
```
Diagnose-and-Repair.bat
or
.\Diagnose-and-Repair.ps1 -Repair
```

### Step 4: Restart Outlook
- Close all Outlook windows
- Reopen Outlook
- Error should be gone

---

## 📝 TROUBLESHOOTING

### "Administrator privileges required"
**Solution:** Right-click → "Run as Administrator"

### "Email format invalid"
**Solution:** Use format: `user@domain.com`

### "Outlook won't close"
**Solution:** 
1. Ctrl + Alt + Delete
2. Task Manager
3. Kill OUTLOOK.EXE
4. Run repair again

### "Registry key already exists" (during repair)
**Solution:** Script will automatically replace it

### "Permission denied"
**Solution:** 
1. Close Outlook completely
2. Run as Administrator
3. Try again

---

## ✨ SUCCESS INDICATORS

After running repair, verify:
- ✅ No error when Outlook opens
- ✅ "Report phishing" button appears
- ✅ Button is clickable
- ✅ Test button on sample email

---

## 📊 DIAGNOSTIC CHECKLIST

Run diagnostic to check:
- [ ] Registry key exists
- [ ] AdminEmail is set
- [ ] AdminEmail format is valid
- [ ] LoadBehavior = 3
- [ ] Description is set
- [ ] Manifest path valid

---

## 🎯 WHICH TOOL TO USE?

**If you're not sure:** Use `Diagnose-and-Repair.bat`
- Easy to run (double-click)
- Works on all Windows versions
- Clear messages
- Automatic repair

**If you prefer PowerShell:** Use `Diagnose-and-Repair.ps1 -Repair`
- More detailed output
- Advanced filtering
- Option to diagnose only (no repair)

**If you like manual control:** Use Registry Editor
- Full control
- Can verify each step
- No automation

---

## 🚀 RECOMMENDED APPROACH

1. **Try Automatic Repair First:**
   ```
   Right-click: Diagnose-and-Repair.bat
   Select: Run as Administrator
   ```

2. **If it works:**
   ```
   Restart Outlook
   Verify button appears
   Done!
   ```

3. **If it doesn't work:**
   ```
   Try Option 3: Re-run Install.vbs
   Or Use Manual Registry Fix (Option 2)
   ```

---

## 📞 SUPPORT

If diagnostic tools don't help:
1. Check log: `%LocalAppData%\OutlookPhishingReporter\Logs\add-in.log`
2. Verify .NET Framework 4.7.2 is installed
3. Verify Outlook version is 2016 or later
4. Try manual registry approach
5. Contact IT support with diagnostic output

---

**The diagnostic tools can usually fix this in 1-2 minutes!** 🎉
