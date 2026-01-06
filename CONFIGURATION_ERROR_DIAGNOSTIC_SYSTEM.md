# 🆘 CONFIGURATION ERROR - COMPLETE DIAGNOSTIC & REPAIR SYSTEM

## ⚠️ THE ERROR

```
Configuration error: Invalid admin email address. Please contact your administrator.
```

## ✅ THE SOLUTION

Professional diagnostic and repair tools have been created to automatically detect and fix this issue.

---

## 🚀 QUICKEST FIX (2 MINUTES)

### Step 1: Navigate
```
Go to: OutlookPhishingReporterSetup\Release\
```

### Step 2: Run Repair Tool
```
Right-click: Diagnose-and-Repair.bat
Select: Run as Administrator
```

### Step 3: Configure
```
When prompted, enter your security email:
Example: security@company.com
```

### Step 4: Restart
```
Close Outlook completely
Reopen Outlook
Error should be gone!
```

---

## 📋 TOOLS PROVIDED

### 1. Diagnose-and-Repair.bat (Recommended)
**What it does:**
- Scans your registry configuration
- Identifies issues automatically
- Prompts for fix if needed
- Rebuilds registry values
- Validates configuration

**How to use:**
```
Right-click → Run as Administrator
Follow prompts
Restart Outlook
```

**Best for:** Everyone - Simple, clear, effective

### 2. Diagnose-and-Repair.ps1 (Advanced)
**What it does:**
- Detailed diagnostic output
- PowerShell validation
- Automatic or manual repair mode
- Comprehensive error handling

**How to use:**
```
PowerShell as Administrator:
.\Diagnose-and-Repair.ps1 -Repair
Follow prompts
Restart Outlook
```

**Best for:** IT professionals, advanced users

### 3. DIAGNOSTIC_AND_REPAIR.md (Reference)
**What it contains:**
- Detailed troubleshooting guide
- Valid email format examples
- Alternative fix methods
- FAQ and support

---

## 🔧 WHAT GETS CHECKED & FIXED

### Diagnostic Checks
- ✅ Registry key exists
- ✅ AdminEmail value configured
- ✅ AdminEmail format valid
- ✅ LoadBehavior = 3
- ✅ Description set correctly
- ✅ Manifest path valid

### Automatic Fixes
- ✅ Creates missing registry key
- ✅ Sets valid AdminEmail value
- ✅ Configures LoadBehavior = 3
- ✅ Sets Description
- ✅ Updates Manifest path

---

## 📊 WHY THIS ERROR HAPPENS

The error occurs when **AdminEmail is missing or invalid** in the Windows Registry.

### Common Causes
1. Incomplete installation
2. Registry value deleted
3. Invalid email format entered
4. Registry corruption
5. Outlook configuration reset

### Solution
The diagnostic tools automatically:
1. Detect the issue
2. Prompt for valid email
3. Rebuild configuration
4. Verify settings
5. Resolve the error

---

## ✨ VALID EMAIL FORMATS

### ✅ CORRECT
```
security@company.com
admin@yourdomain.org
phishing@domain.co.uk
user.name@company.net
first+last@domain.com
```

### ❌ INCORRECT
```
security (no @domain)
@company.com (no username)
admin @ company.com (spaces)
admin@ (missing extension)
user (just text)
```

---

## 🎯 STEP-BY-STEP REPAIR

### Using Batch File (Easiest)

**Step 1: Close Outlook**
- Close all Outlook windows
- Make sure Outlook is not running

**Step 2: Run Diagnostic Tool**
- Navigate to: `OutlookPhishingReporterSetup\Release\`
- Right-click: `Diagnose-and-Repair.bat`
- Select: **Run as Administrator**

**Step 3: View Diagnostic Results**
- Tool checks your configuration
- Shows what's missing
- Offers to repair

**Step 4: Enter Email**
- Provide security email address
- Example: `security@company.com`
- Tool validates format

**Step 5: Complete**
- Repair finishes
- Registry updated
- Tool closes

**Step 6: Restart Outlook**
- Open Outlook
- Click "Report phishing" button
- No error should appear!

---

## 🆘 IF THE TOOL DOESN'T WORK

### Alternative 1: Re-run Installer
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Double-click: Install.vbs
3. Enter email when prompted
4. Restart Outlook
```

### Alternative 2: Manual Registry Fix
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail value
5. Enter: security@company.com
6. Restart Outlook
```

### Alternative 3: PowerShell Command
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
Remove-Item -Path $regPath -Force -ErrorAction SilentlyContinue
New-Item -Path $regPath -Force | Out-Null
New-ItemProperty -Path $regPath -Name "AdminEmail" -Value "security@company.com" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
```

---

## 📈 SUCCESS VERIFICATION

After repair, verify these things:

- ✅ Outlook opens without error dialog
- ✅ No "Configuration error" message
- ✅ "Report phishing" button appears in ribbon
- ✅ Button is clickable
- ✅ Test on a sample email
- ✅ Email forwards to security team
- ✅ Email moved to Deleted Items

---

## 🎓 HOW IT WORKS

### Diagnostic Mode
```
Check Registry Key
  ↓
Check AdminEmail
  ↓
Check LoadBehavior
  ↓
Check Description
  ↓
Check Manifest
  ↓
Report Results
```

### Repair Mode
```
Close Outlook
  ↓
Delete Old Config
  ↓
Prompt for Email
  ↓
Validate Email
  ↓
Create Registry Entries
  ↓
Report Success
  ↓
Prompt Restart
```

---

## 📞 TROUBLESHOOTING

### "Administrator privileges required"
**Problem:** Script needs admin rights  
**Solution:** Right-click → "Run as Administrator"

### "Email format invalid"
**Problem:** Email doesn't match standard format  
**Solution:** Use format: `user@domain.com`

### "Outlook won't close"
**Problem:** Outlook process still running  
**Solution:** 
1. Ctrl + Alt + Delete
2. Task Manager
3. End all OUTLOOK.EXE processes
4. Run tool again

### "Registry key already exists"
**Problem:** Old config still there  
**Solution:** Tool will replace it automatically

### "Permission denied"
**Problem:** Can't write to registry  
**Solution:**
1. Run as Administrator
2. Close Outlook completely
3. Try again

---

## 📝 DOCUMENTATION

| File | Purpose |
|------|---------|
| **Diagnose-and-Repair.bat** | Main repair tool (batch) |
| **Diagnose-and-Repair.ps1** | PowerShell repair tool |
| **DIAGNOSTIC_AND_REPAIR.md** | Comprehensive guide |
| **QUICK_FIX_GUIDE.md** | Quick reference |
| **CONFIG_ERROR_COMPLETE_SOLUTION.md** | Alternative solutions |

---

## 🎯 RECOMMENDED APPROACH

**Priority 1 (Fastest - 2 minutes):**
```
Run: Diagnose-and-Repair.bat as Administrator
```

**Priority 2 (If first fails - 3 minutes):**
```
Run: Install.vbs and enter email again
```

**Priority 3 (If both fail - 5 minutes):**
```
Manual registry edit via regedit
```

---

## 🎉 SUMMARY

**The diagnostic and repair system:**
- ✅ Automatically detects the problem
- ✅ Validates your configuration
- ✅ Prompts for email address
- ✅ Rebuilds registry values
- ✅ Verifies the fix
- ✅ Takes 2-3 minutes
- ✅ No technical knowledge required

**Ready to fix the error?**
→ Right-click `Diagnose-and-Repair.bat`
→ Select "Run as Administrator"
→ Follow the prompts!

---

**Professional diagnostic and repair tools are ready!** 🚀

The error can typically be fixed in under 2 minutes using the provided tools.
