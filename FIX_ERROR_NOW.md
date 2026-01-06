# 🎯 IMMEDIATE ACTION GUIDE

## YOU ARE HERE

You're seeing this error in Outlook:
```
Configuration error: Invalid admin email address. Please contact your administrator.
```

## DO THIS NOW (2 MINUTES)

### Step 1: Locate the Tool
```
Navigate to: OutlookPhishingReporterSetup\Release\
Look for: Diagnose-and-Repair.bat
```

### Step 2: Run the Tool
```
Right-click: Diagnose-and-Repair.bat
Select: Run as Administrator
```

### Step 3: Follow Prompts
```
The tool will:
- Close Outlook
- Check your configuration
- Ask for security email address
- Fix the problem
- Tell you when done
```

### Step 4: Restart Outlook
```
Close Outlook
Reopen Outlook
Error should be GONE!
```

---

## IF YOU CAN'T FIND THE FILE

### Search for it:
```
Windows Key + S
Type: Diagnose-and-Repair.bat
Press Enter
```

### Or navigate manually:
```
1. Open File Explorer
2. Go to: C:\Users\[YourName]\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
3. Open: OutlookPhishingReporterSetup
4. Open: Release
5. Find: Diagnose-and-Repair.bat
```

---

## WHAT TO ENTER WHEN ASKED

### When it asks for email address:

**Enter something like:**
```
security@company.com
admin@yourdomain.com
phishing@example.org
```

**NOT:**
```
security
@company.com
admin @ company (with spaces)
```

---

## WHAT HAPPENS NEXT

### During repair:
- Outlook closes automatically (don't panic!)
- Registry gets updated
- Configuration gets fixed
- Takes 1-2 minutes

### After repair:
- Tool shows success message
- Close the tool window
- Restart Outlook
- Error should be gone!

---

## VERIFY IT WORKED

After restarting Outlook:

1. ✅ Check: No error dialog appears
2. ✅ Check: "Report phishing" button visible
3. ✅ Check: Button is clickable
4. ✅ Test: Click button on a test email
5. ✅ Success: Email forwards to security team

---

## IF IT STILL DOESN'T WORK

### Try This:

1. Run the tool again
2. This time, when it asks for email, use a DIFFERENT valid email
3. Restart Outlook again

---

## ALTERNATIVE QUICK FIX

If you prefer NOT to run the tool:

### Use PowerShell instead:

```powershell
# Open PowerShell as Administrator
# Then run:
.\Diagnose-and-Repair.ps1 -Repair
```

---

## MANUAL FIX (If tools don't work)

1. Press: `Win + R`
2. Type: `regedit`
3. Go to: `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
4. Right-click: `AdminEmail`
5. Select: **Modify**
6. Enter: `security@company.com`
7. Click: **OK**
8. Close: Registry Editor
9. Restart: Outlook

---

## ⏰ EXPECTED TIME

- Tool runs: 2 minutes
- Restart Outlook: 1 minute
- **Total: 3 minutes**

---

## 📞 IF YOU NEED HELP

1. Check: `DIAGNOSTIC_AND_REPAIR.md` (detailed guide)
2. Review: `CONFIG_ERROR_COMPLETE_SOLUTION.md` (all options)
3. Try: All three tool options provided
4. Contact: IT support with tool output

---

## 🎉 THAT'S IT!

The error will be fixed in minutes. The tools handle everything.

**Now go run: Diagnose-and-Repair.bat** ✓

---

**Questions? Read the complete guide at:**
`OutlookPhishingReporterSetup\Release\DIAGNOSTIC_AND_REPAIR.md`
