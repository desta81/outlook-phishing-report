# 🚀 STEP-BY-STEP EXECUTION GUIDE - RUN NOW

## ✅ PREREQUISITES CHECK

Before we start, verify you have:
- ✓ Admin access
- ✓ Outlook closed (very important)
- ✓ ~20 minutes available
- ✓ ~1 GB free disk space

---

## 🎯 STEP 1: CLOSE OUTLOOK (CRITICAL)

**Action:**
1. Look at your taskbar (bottom of screen)
2. Find Outlook (mail icon)
3. Click it to bring window to front
4. Close ALL Outlook windows
   - Click X button on each window
   - Wait 5 seconds

**Verify:**
- No Outlook windows visible
- Outlook not in taskbar
- (Use Task Manager if needed: Ctrl + Shift + Esc → find Outlook.exe → End Task)

---

## 🎯 STEP 2: OPEN COMMAND PROMPT AS ADMINISTRATOR

**Action:**

**Method A (Fastest):**
```
Press: Win + R
Type: cmd
Press: Ctrl + Shift + Enter
(NOT just Enter - must use Ctrl + Shift + Enter)
```

**Method B (Alternative):**
```
Click: Start Menu
Type: Command Prompt
Right-click: Command Prompt
Select: Run as Administrator
```

**Verify:**
- Window title shows: "Administrator: Command Prompt"
- You're running as Administrator

---

## 🎯 STEP 3: NAVIGATE TO PROJECT DIRECTORY

**Action:**

Copy and paste this command:
```cmd
cd /d C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter
```

**What it does:**
- Changes directory to your project folder
- `/d` allows changing drive letters if needed

**Verify:**
- Command prompt shows: `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter>`
- You're in the correct directory

---

## 🎯 STEP 4: RUN THE BUILD SCRIPT

**Action:**

Type:
```cmd
BUILD-AND-RUN-MSI.bat
```

Press: **Enter**

**What happens:**
- Script starts immediately
- Black window shows build status
- Watch the messages

---

## 📊 STEP 5: WAIT FOR DLL BUILD (1-2 MINUTES)

**You'll see:**
```
╔════════════════════════════════════════════════════════════════════╗
║ STEP 1: BUILD PLUGIN DLL                                          ║
╚════════════════════════════════════════════════════════════════════╝

[*] Building DLL...
... (compilation messages)
[+] DLL build completed
[+] DLL verified
```

**What's happening:**
- Plugin code is being compiled
- Creating OutlookPhishingReporter.dll

**If errors appear:**
- STOP and tell me what the error says
- Don't continue

**If successful:**
- Continue to next step

---

## 📊 STEP 6: WAIT FOR MSI BUILD (5-10 MINUTES)

**You'll see:**
```
╔════════════════════════════════════════════════════════════════════╗
║ STEP 2: BUILD MSI INSTALLER                                       ║
╚════════════════════════════════════════════════════════════════════╝

[*] Building MSI...
... (compilation messages)
[+] MSI build completed
[+] MSI verified
Location: ...Distribution\OutlookPhishingReporter.msi
Size: 5-10 MB bytes
```

**What's happening:**
- MSI installer package is being created
- This is the largest part of the build

**If WiX error appears:**
- You need WiX Toolset installed
- I'll help you fix this

**If successful:**
- Continue to next step

---

## 🎯 STEP 7: CHOOSE INSTALLATION TYPE

**You'll see:**
```
╔════════════════════════════════════════════════════════════════════╗
║ STEP 3: RUN INSTALLER                                             ║
╚════════════════════════════════════════════════════════════════════╝

Ready to run the MSI installer!

Choose your installation type:

  1 = Normal Installation (with wizard)
  2 = Silent Installation (automatic, no interaction)
  3 = Skip Installation (just build MSI)
  4 = Exit

Select option (1-4): _
```

**Choose one:**

**Option 1 (RECOMMENDED for first time):**
```
Type: 1
Press: Enter
```
- Installation wizard opens
- Follow on-screen prompts
- Takes 2-3 minutes
- Most user-friendly

**Option 2 (for enterprise/testing):**
```
Type: 2
Press: Enter
```
- Automatic installation
- No wizard
- Takes 30-60 seconds
- Perfect for scripts

**Option 3 (just build, don't install yet):**
```
Type: 3
Press: Enter
```
- Creates MSI file only
- No installation
- For distribution/testing

---

## 🎯 STEP 8A: IF YOU CHOSE OPTION 1 (Normal Install)

**Installation wizard opens:**

1. **Welcome screen:**
   - Click: Next

2. **Select Installation Folder:**
   - Keep default (don't change)
   - Click: Next

3. **Ready to Install:**
   - Review settings
   - Click: Install

4. **Installation Progress:**
   - Wait for completion
   - Should be 2-3 minutes

5. **Installation Complete:**
   - Click: Finish
   - Close the wizard

---

## 🎯 STEP 8B: IF YOU CHOSE OPTION 2 (Silent Install)

**Silent installation runs:**
- No wizard appears
- Installation happens automatically
- Takes 30-60 seconds
- Wait for completion

---

## 🎯 STEP 8C: IF YOU CHOSE OPTION 3 (Build Only)

**No installation occurs:**
- Script finishes
- MSI file is ready
- Can install manually later
- Jump to Step 9

---

## 🎯 STEP 9: SCRIPT COMPLETES

**You'll see:**
```
╔════════════════════════════════════════════════════════════════════╗
║                    BUILD COMPLETE!                                ║
╚════════════════════════════════════════════════════════════════════╝

[+] DLL built successfully
[+] MSI created successfully

Location: C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\
          OutlookPhishingReporterSetup\Distribution\OutlookPhishingReporter.msi

Next steps:
  1. If not installed, run: OutlookPhishingReporter.msi
  2. Restart Outlook
  3. Verify "Report phishing" button appears
  4. Test report functionality

For silent installation, use:
  msiexec /i OutlookPhishingReporter.msi /quiet /norestart
```

**Action:**
- Close the command prompt window
- Or press any key when prompted

---

## 🎯 STEP 10: RESTART OUTLOOK (CRITICAL)

**Action:**

1. **Open Outlook:**
   - Click: Start Menu
   - Type: Outlook
   - Press: Enter

2. **Wait for it to load:**
   - Outlook window opens
   - Wait 5-10 seconds for full load

3. **Look for the button:**
   - Find: Home tab (should be selected)
   - Look at ribbon
   - Find: "Report phishing" button

**Expected result:**
- You should see a new button labeled "Report phishing"
- It should have an icon
- It should be in the ribbon

---

## 🎯 STEP 11: VERIFY INSTALLATION

**Action:**

1. **Visual Check:**
   ```
   Look for: "Report phishing" button
   Location: Home tab in Outlook ribbon
   Status: Should be clickable
   ```

2. **Test It:**
   ```
   1. Select any email in inbox
   2. Click: "Report phishing" button
   3. Confirmation dialog appears
   4. Click: OK
   5. Email forwarded to security team
   6. Email moved to Deleted Items
   ```

**Success Indicators:**
- ✓ Button appears in ribbon
- ✓ Button is clickable
- ✓ Dialog appears when clicked
- ✓ Email is forwarded
- ✓ Email moved to trash
- ✓ No error messages

**If button doesn't appear:**
- Close Outlook completely
- Wait 5 seconds
- Reopen Outlook
- Try again

**If still missing:**
- Run this command:
  ```cmd
  OutlookPhishingReporterSetup\Release\FIX-ERROR-NOW.bat
  ```
- Restart Outlook

---

## 🆘 TROUBLESHOOTING

### **Problem: "Admin privileges" error immediately**
```
Solution:
  Make sure you used: Ctrl + Shift + Enter
  (Not just Enter)
  
  Open Command Prompt again as Administrator
```

### **Problem: DLL build fails**
```
Solution:
  1. Tell me the error message
  2. Usually means Visual Studio not installed
  3. Or .NET Framework 4.7.2 missing
```

### **Problem: "WiX Toolset not found" error**
```
Solution:
  1. Download: https://wixtoolset.org/
  2. Install: WiX v3.14
  3. Restart computer
  4. Try running script again
```

### **Problem: Installation wizard doesn't appear (Option 1)**
```
Solution:
  1. Check if it opened in background
  2. Look in taskbar for window
  3. Or choose Option 2 (silent) instead
```

### **Problem: Button doesn't appear after install**
```
Solution:
  1. Completely close Outlook
  2. Wait 5 seconds
  3. Reopen Outlook
  4. If still missing, run FIX-ERROR-NOW.bat
```

### **Problem: "Script hangs" or takes too long**
```
Solution:
  1. Wait 5 more minutes (first build is slower)
  2. Or press Ctrl + C to cancel
  3. Close and try again
```

---

## 📊 EXPECTED TIMELINE

```
Time          Event
-----         -----
0 sec         You run: BUILD-AND-RUN-MSI.bat
0-2 min       DLL builds
2-10 min      MSI builds  
10 min        Script asks what to do
11-15 min     Installation runs (if chosen)
15 min        Complete! Script finishes
15-20 min     You restart Outlook
20 min        Button appears - Success! ✓
```

---

## ✅ SUCCESS CHECKLIST

Track your progress:

- [ ] Outlook closed
- [ ] Command Prompt opened as Administrator
- [ ] Navigated to project directory
- [ ] Ran: BUILD-AND-RUN-MSI.bat
- [ ] DLL built (no errors)
- [ ] MSI built (no errors)
- [ ] Chose installation type (1, 2, or 3)
- [ ] Installation completed
- [ ] Command window closed
- [ ] Restarted Outlook
- [ ] "Report phishing" button visible
- [ ] Button is clickable
- [ ] ✓ SUCCESS!

---

## 🚀 START NOW!

Ready to run it?

**Follow these exact steps:**

1. Close Outlook (close all windows)
2. Press: Win + R
3. Type: cmd
4. Press: Ctrl + Shift + Enter
5. Type: `cd /d C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter`
6. Press: Enter
7. Type: `BUILD-AND-RUN-MSI.bat`
8. Press: Enter
9. When asked, type: 1 (or 2)
10. Press: Enter
11. Wait 15 minutes
12. Restart Outlook
13. Done! ✓

**Tell me any errors you see, and I'll help you fix them!**

---

**Let's run it!** 🎉
