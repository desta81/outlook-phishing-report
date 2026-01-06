# 🔧 FIX EMAIL CONFIGURATION - DEBUG MODE

## 🎯 PROBLEM

The plugin is reading the email from Windows Registry but it's either:
1. Not saved correctly
2. Not being read by the plugin
3. Invalid format issue

## ✅ SOLUTION - 3 STEPS

### **STEP 1: Stop Debug Session**

In Visual Studio:
```
Press: Shift + F5 (Stop Debugging)
Wait 5 seconds
Close Outlook completely
```

---

### **STEP 2: Set Email via Registry (Manual)**

**Action:**

1. Press: `Win + R`
2. Type: `regedit`
3. Press: `Enter`

**Navigate to:**
```
HKEY_CURRENT_USER
→ Software
→ Microsoft
→ Office
→ Outlook
→ Addins
→ OutlookPhishingReporter
```

**In this folder, do this:**

1. Right-click in empty space
2. Select: `New` → `String Value`
3. Name: `AdminEmail`
4. Press: `Enter`
5. Double-click the new `AdminEmail` entry
6. Value: `yosi.desta@outlook.co.il`
7. Click: `OK`
8. Close regedit

**Verify it looks like:**
```
Name:  AdminEmail
Value: yosi.desta@outlook.co.il
```

---

### **STEP 3: Reload in Debug Mode**

**Action:**

1. In Visual Studio make sure Outlook is completely closed
2. Clean the solution:
   ```
   Build → Clean Solution
   Wait for completion
   ```

3. Rebuild the solution:
   ```
   Build → Rebuild Solution
   Wait for "Build succeeded"
   ```

4. Start debug:
   ```
   Press: F5 (Debug → Start Debugging)
   ```

5. Outlook opens

6. Test the button:
   - Select email
   - Click: "Report phishing" button
   - Should work now!

---

## 🆘 IF IT STILL DOESN'T WORK

### **Check the Log File**

The plugin creates a log file. Check what error it shows:

**Log Location:**
```
C:\Users\YosiD\AppData\Local\OutlookPhishingReporter\Logs\add-in.log
```

**Action:**
1. Open File Explorer
2. Navigate to: `C:\Users\YosiD\AppData\Local\OutlookPhishingReporter\Logs\`
3. Open: `add-in.log`
4. Look at the last messages
5. Tell me what you see

---

## 🔍 DEBUG THE REGISTRY

**Verify email is really saved:**

1. Open: `regedit`
2. Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter`
3. Look for: `AdminEmail`
4. Check the value is: `yosi.desta@outlook.co.il`
5. If empty or different → need to fix

**If value is wrong:**
- Right-click it
- Select: `Modify`
- Change to: `yosi.desta@outlook.co.il`
- Click: `OK`

---

## 🎯 QUICK CHECKLIST

- [ ] Stopped debug (Shift + F5)
- [ ] Closed Outlook
- [ ] Opened regedit
- [ ] Navigated to: OutlookPhishingReporter key
- [ ] Created/Updated: AdminEmail = yosi.desta@outlook.co.il
- [ ] Closed regedit
- [ ] Cleaned solution (Build → Clean)
- [ ] Rebuilt solution (Build → Rebuild)
- [ ] Pressed F5 to debug
- [ ] Outlook opened
- [ ] Selected email
- [ ] Clicked button
- [ ] ✓ Works or still error?

---

## 📋 REPORT BACK

After trying this, tell me:

1. ✓ Does the button work now?
2. ✓ Or do you still get the error?
3. ✓ What does the log file say?
4. ✓ What does the registry value show?

---

**Do this now and report back!** 🚀
