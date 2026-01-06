# 🔍 DEBUG MODE EXECUTION - COMPLETE GUIDE

## 🎯 WHAT WE'RE DOING

1. Run project in **Debug mode** in Visual Studio
2. Let Outlook open with the plugin
3. Check if "Report phishing" button appears
4. Verify button functionality
5. Troubleshoot if needed

---

## ✅ STEP 1: CLOSE OUTLOOK (CRITICAL)

**Action:**
```
1. Look at taskbar (bottom of screen)
2. Close ALL Outlook windows
3. Wait 5 seconds
4. Verify Outlook is gone from taskbar
```

**Verify in Task Manager:**
```
Press: Ctrl + Shift + Esc
Search for: Outlook.exe
If found, select and click: End Task
```

---

## ✅ STEP 2: OPEN PROJECT IN VISUAL STUDIO

**Action:**

**Option A: Already have VS open**
```
File → Open → Project/Solution
Navigate to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
Select: OutlookPhishingReporter.csproj
Click: Open
```

**Option B: Open from File Explorer**
```
Navigate to: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
Double-click: OutlookPhishingReporter.csproj
(Visual Studio opens automatically)
```

**Verify:**
- Visual Studio opens
- Project loads
- Solution Explorer shows "OutlookPhishingReporter" project

---

## ✅ STEP 3: SET DEBUG CONFIGURATION

**Action:**

Look at top toolbar:

1. Find dropdown that says: "Debug" or "Release"
2. Click the dropdown
3. Select: **Debug**

**Expected:**
```
Toolbar shows: Debug (not Release)
Platform shows: AnyCPU
```

---

## ✅ STEP 4: START DEBUGGING

**Action:**

**Option A: Using Keyboard**
```
Press: F5
(Starts debugging immediately)
```

**Option B: Using Menu**
```
Debug → Start Debugging
Or: Debug → Start Without Debugging (Ctrl + F5)
```

**Option C: Using Toolbar Button**
```
Look for: Green play button (►) in toolbar
Click it
```

**What Happens:**
```
Visual Studio:
  1. Builds the project
  2. Compiles code
  3. Registers add-in
  4. Launches Outlook with debugger attached
```

---

## 📊 STEP 5: WAIT FOR BUILD (1-2 MINUTES)

**You'll see:**
```
Output window at bottom:
  Build started...
  Compiling...
  .... (progress messages)
  Build succeeded
  
Visual Studio will then:
  Prepare debug session
  Launch Outlook
```

**If Errors Appear:**
- Stop here
- Tell me the error message
- Don't continue

**If Build Succeeds:**
- Continue to next step

---

## 📊 STEP 6: OUTLOOK LAUNCHES (WAIT 10-15 SECONDS)

**You'll see:**
```
Outlook splash screen appears
Outlook loading...
Main Outlook window opens
```

**Wait for:**
- Outlook fully loads
- Inbox appears
- No loading/spinning indicators
- Ready to use

**Don't close Visual Studio:**
- Keep it open in background
- Keep debugger running

---

## 🔍 STEP 7: CHECK FOR BUTTON IN OUTLOOK

**Action:**

**Step 7A: Look at Outlook Ribbon**
```
1. Outlook window is active
2. Look at top of window
3. Find tabs: Home, Send/Receive, Folder, View, etc.
4. Select: Home tab (click it)
5. Look for: "Report phishing" button
```

**Expected Location:**
```
Home tab → Look at ribbon
Should see a button/icon labeled "Report phishing"
Probably near other security-related buttons
```

**Step 7B: What It Should Look Like**
```
┌─────────────────────────────────────────┐
│ Home │ Send/Receive │ Folder │ View │   │
├─────────────────────────────────────────┤
│ New │ Delete │ Spam │ [Report phishing] │  ← BUTTON HERE
│ Message │  │ │ │
└─────────────────────────────────────────┘
```

---

## ✅ STEP 8: VERIFY BUTTON EXISTS

**Success Indicators:**

✓ Button is visible in ribbon
✓ Button is labeled "Report phishing"
✓ Button has an icon
✓ Button is clickable (not grayed out)

**If Button Appears:**
→ Go to STEP 9 (Test It)

**If Button NOT Visible:**
→ Go to STEP 11 (Troubleshooting)

---

## 🧪 STEP 9: TEST THE BUTTON

**Action:**

**Step 9A: Select a Test Email**
```
1. Look at Inbox
2. Select any email
   (Click on it once to select)
3. Verify email is selected (highlighted)
```

**Step 9B: Click the Button**
```
1. Look at ribbon
2. Find: "Report phishing" button
3. Click it
```

**Expected:**
```
Dialog box appears with:
  Title: "Report Phishing Email"
  Message: "You are about to report this email as phishing..."
  Buttons: [OK] [Cancel]
```

**Step 9C: Confirm Report**
```
1. Click: OK (to proceed)
Or:
2. Click: Cancel (to abort)
```

**If Click OK:**
```
Another dialog appears:
  "Thank you for reporting this email as phishing. 
   It has been moved to the Deleted Items folder."

Email is:
  - Forwarded to security team
  - Moved to trash
  - Report complete ✓
```

---

## 🎊 STEP 10: SUCCESS VERIFICATION

**Check List:**

- [ ] Outlook opened from Visual Studio
- [ ] Button visible in Home tab ribbon
- [ ] Button labeled "Report phishing"
- [ ] Button clickable
- [ ] Dialog appeared when clicked
- [ ] Email was reported
- [ ] No errors occurred
- [ ] ✓ SUCCESS!

**You've Verified:**
✓ Plugin loads correctly
✓ Button appears in ribbon
✓ Button is functional
✓ Dialog works
✓ Reporting works

---

## 🆘 STEP 11: IF BUTTON DOESN'T APPEAR

**Troubleshooting:**

### **Issue 1: Button not visible immediately**
```
Solution:
1. Wait another 10 seconds
2. Or close/reopen Outlook
3. Or restart Visual Studio debug session
```

### **Issue 2: Outlook didn't launch**
```
Solution:
1. Make sure Outlook is closed
2. Stop debug session (press Shift + F5)
3. Close Visual Studio
4. Manually open Outlook
5. Try again
```

### **Issue 3: Build failed with errors**
```
Solution:
1. Read error message in Output window
2. Common errors:
   - Missing .NET Framework 4.7.2
   - Missing Office interop assemblies
   - Visual Studio extensions not installed
3. Tell me the error
```

### **Issue 4: Build succeeded but button missing**
```
Solution:
1. Close Outlook in Visual Studio
   (Stop debug session: Shift + F5)
2. Run fix script:
   OutlookPhishingReporterSetup\Release\FIX-ERROR-NOW.bat
3. Restart debug session (F5)
```

### **Issue 5: Visual Studio shows "add-in didn't load"**
```
Solution:
1. This may appear in a dialog
2. If it does, click OK to dismiss
3. Close and restart Outlook
4. Try again
```

---

## 📊 VISUAL STUDIO DEBUG FEATURES

### **While Debugging:**

**Stop Debugging:**
```
Press: Shift + F5
Or: Debug → Stop Debugging
Outlook closes
Debug session ends
```

**Pause Debugging:**
```
Press: Ctrl + Alt + Break
Execution pauses
You can inspect variables
```

**Set Breakpoints:**
```
Click in left margin of code
Red dot appears
Execution will pause when line is hit
```

---

## 📝 COMMON DEBUG OUTPUTS

### **Watch Output Window for:**

**Success Messages:**
```
Build started...
Build succeeded.
Starting Outlook...
Add-in loaded successfully
```

**Error Messages:**
```
Build failed
Error: Could not find...
Exception: ...
```

**Application Logs:**
```
Right-click on email...
Log file created in:
C:\Users\YosiD\AppData\Local\OutlookPhishingReporter\Logs\
```

---

## 🎯 QUICK COMMANDS

### **During Debug Session:**

| Action | Key |
|--------|-----|
| Start Debug | F5 |
| Stop Debug | Shift + F5 |
| Pause Debug | Ctrl + Alt + Break |
| Step Over | F10 |
| Step Into | F11 |
| Restart | Ctrl + Shift + F5 |
| Build | Ctrl + Shift + B |

---

## 📋 COMPLETE CHECKLIST

Track your progress:

- [ ] Outlook closed
- [ ] Visual Studio opened with project
- [ ] Debug configuration selected
- [ ] F5 pressed (or Start Debugging clicked)
- [ ] Build completed successfully
- [ ] Outlook launched
- [ ] Outlook fully loaded
- [ ] Home tab is active
- [ ] Looking for button...
- [ ] Button is VISIBLE
- [ ] Button is CLICKABLE
- [ ] Selected test email
- [ ] Clicked button
- [ ] Dialog appeared
- [ ] Clicked OK to report
- [ ] Email was moved to trash
- [ ] ✓ SUCCESS!

---

## 🆘 GETTING HELP

If something goes wrong:

**Tell me:**
1. What step you're on
2. What error message appears
3. A screenshot (if possible)
4. Any other details

**I can help with:**
- Build errors
- Add-in not loading
- Button not appearing
- Outlook not launching
- Any other issues

---

## 🚀 START DEBUG SESSION NOW

**Follow these exact steps:**

```
1. Close Outlook (very important)

2. Open Visual Studio

3. Open project: OutlookPhishingReporter.csproj

4. Select: Debug (from dropdown)

5. Press: F5 (or Debug → Start Debugging)

6. Wait for build to complete

7. Outlook opens automatically

8. Look for button in Home tab

9. Test it!

10. Report back with results
```

---

**When you're ready, start with Step 1 above and report back what happens!** 🎉

**I'm here to help debug any issues!** 💪
