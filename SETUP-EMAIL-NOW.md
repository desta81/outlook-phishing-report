# 📧 SETUP EMAIL ADDRESS - QUICK GUIDE

## ✅ STEP 1: CLOSE OUTLOOK

**Action:**
```
Close all Outlook windows
Stop debug session (Shift + F5)
```

---

## ✅ STEP 2: ADD EMAIL TO REGISTRY

The plugin reads the email from Windows Registry. We need to add it.

**Option A: Automatic (Easiest)**

Run this script:
```
OutlookPhishingReporterSetup\Release\SETUP-YOSI-EMAIL.md
```

Or use:
```
OutlookPhishingReporterSetup\Release\Setup-Default-Email.bat
```

**Option B: Manual Setup**

Follow these steps:

1. **Press:** Win + R
2. **Type:** regedit
3. **Press:** Enter

4. **Navigate to:**
```
HKEY_CURRENT_USER
→ Software
→ Microsoft
→ Office
→ Outlook
→ Addins
→ OutlookPhishingReporter
```

5. **Look for value:** AdminEmail

6. **If AdminEmail exists:**
   - Right-click it
   - Select: Modify
   - Change value to: yosi.desta@outlook.co.il
   - Click: OK

7. **If AdminEmail doesn't exist:**
   - Right-click in empty space
   - Select: New → String Value
   - Name it: AdminEmail
   - Set value to: yosi.desta@outlook.co.il
   - Click: OK

8. **Close regedit**

---

## ✅ STEP 3: VERIFY EMAIL SAVED

**Action:**
```
1. Still in regedit
2. Navigate to: OutlookPhishingReporter key
3. Look for: AdminEmail
4. Value should be: yosi.desta@outlook.co.il
5. Close regedit
```

---

## ✅ STEP 4: RESTART DEBUG

**Action:**

1. **In Visual Studio:**
   - Press: F5 (Start Debugging)
   - Outlook opens

2. **Wait for Outlook to fully load**

---

## ✅ STEP 5: TEST WITH EMAIL

**Action:**

1. **Select an email in inbox**
2. **Click: "Report phishing" button**
3. **Dialog appears with confirmation**
4. **Click: OK**

**Expected Results:**
- ✓ Email forwarded to: yosi.desta@outlook.co.il
- ✓ Email moved to Deleted Items
- ✓ Success message shown
- ✓ No error about missing email

---

## 🔧 ALTERNATIVE: USE BATCH SCRIPT

**If you prefer automatic setup:**

Run as Administrator:
```
OutlookPhishingReporterSetup\Release\Setup-Default-Email.bat
```

This script will:
- ✓ Stop Visual Studio debug
- ✓ Set email in registry
- ✓ Verify it's saved
- ✓ Restart debug
- ✓ Open Outlook

---

## 📝 WHAT'S HAPPENING

**Why the error occurs:**
- Plugin looks for email in registry
- Key: `AdminEmail` in Outlook Addins registry
- If not found → error "no email found"

**How we fix it:**
- Add/update `AdminEmail` value in registry
- Set it to: yosi.desta@outlook.co.il
- Plugin reads it on next load
- Email is sent to that address

---

## 🎯 QUICK COPY-PASTE (Manual Registry)

**If doing manually:**

```
1. Win + R
2. regedit
3. Enter
4. Navigate to: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
5. Right-click → New → String Value
6. Name: AdminEmail
7. Value: yosi.desta@outlook.co.il
8. OK
9. Close regedit
10. Restart debug (F5)
```

---

## ✅ VERIFY IT'S WORKING

**After setup:**

1. Close Outlook (Shift + F5)
2. Restart debug (F5)
3. Outlook opens
4. Select email
5. Click button
6. Should work without error

---

**Do this now and report back!** 🚀
