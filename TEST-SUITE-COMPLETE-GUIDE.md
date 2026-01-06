# 🧪 TEST SUITE - COMPLETE GUIDE

## Overview

The Outlook Phishing Reporter includes a comprehensive test suite to verify:
- ✅ Plugin installation
- ✅ Registry configuration
- ✅ .NET Framework compatibility
- ✅ Outlook integration
- ✅ Email configuration
- ✅ System readiness

---

## 🚀 RUN TESTS NOW

### **Option 1: Batch File (Easy)**

```
Navigate to: OutlookPhishingReporterSetup\Release\
Run: RUN-TEST-SUITE.bat
Right-click → Run as Administrator
```

### **Option 2: PowerShell (Advanced)**

```powershell
# Open PowerShell as Administrator
cd OutlookPhishingReporterSetup\Release
.\RUN-TEST-SUITE.ps1
```

**With verbose output:**
```powershell
.\RUN-TEST-SUITE.ps1 -Verbose
```

**Quick test (skip slow checks):**
```powershell
.\RUN-TEST-SUITE.ps1 -Quick
```

---

## 📋 TESTS PERFORMED

### **Test 1: Administrator Privileges**
- **Purpose:** Verify you have admin rights
- **Why it matters:** Registry access requires admin
- **Pass Condition:** Admin rights confirmed
- **Fail Action:** Run as Administrator

### **Test 2: .NET Framework 4.7.2**
- **Purpose:** Check .NET Framework version
- **Why it matters:** Plugin requires 4.7.2 or later
- **Pass Condition:** Framework 4.7.2+ detected
- **Fail Action:** Install .NET Framework 4.7.2+

### **Test 3: Outlook Installation**
- **Purpose:** Verify Outlook is installed
- **Why it matters:** Plugin extends Outlook
- **Pass Condition:** Microsoft Office found
- **Fail Action:** Install/repair Outlook

### **Test 4: Plugin Registry Configuration**
- **Purpose:** Check registry entries exist
- **Why it matters:** Registry stores plugin settings
- **Pass Condition:** Registry key and values found
- **Fail Action:** Run FIX-ERROR-NOW.bat

### **Test 5: LoadBehavior Setting**
- **Purpose:** Verify auto-load is enabled
- **Why it matters:** Controls plugin startup
- **Pass Condition:** LoadBehavior = 3
- **Fail Action:** Run FIX-ERROR-NOW.bat

### **Test 6: Outlook Process**
- **Purpose:** Check if Outlook is running
- **Why it matters:** Some tests require Outlook closed
- **Pass Condition:** Outlook not running
- **Fail Action:** Close Outlook

### **Test 7: Plugin DLL**
- **Purpose:** Locate plugin DLL file
- **Why it matters:** DLL is core plugin code
- **Pass Condition:** DLL file exists
- **Fail Action:** Reinstall plugin

### **Test 8: Log Directory**
- **Purpose:** Check log folder exists
- **Why it matters:** Logs help with troubleshooting
- **Pass Condition:** Log directory found
- **Fail Action:** Will be created on first use

### **Test 9: Email Configuration**
- **Purpose:** Verify admin email is set
- **Why it matters:** Email is where reports go
- **Pass Condition:** Valid email configured
- **Fail Action:** Run FIX-ERROR-NOW.bat

### **Test 10: Plugin Manifest**
- **Purpose:** Check manifest file
- **Why it matters:** Manifest defines plugin structure
- **Pass Condition:** Manifest file found
- **Fail Action:** Reinstall plugin

---

## ✅ UNDERSTANDING TEST RESULTS

### **[PASS]** - Green
```
Test succeeded, no action needed
```

### **[FAIL]** - Red
```
Test failed, action required
Recommended fix provided in output
```

### **[WARN]** - Yellow
```
Test passed but with warnings
May need attention but not critical
```

### **[INFO]** - Cyan
```
Informational message
No action required, just reporting status
```

---

## 🎯 TEST OUTPUT EXAMPLE

```
╔════════════════════════════════════════════════════════════════════╗
║   OUTLOOK PHISHING REPORTER - COMPLETE TEST SUITE                 ║
╚════════════════════════════════════════════════════════════════════╝

[TEST 1/10] Checking Administrator Privileges...
[PASS] Administrator privileges confirmed

[TEST 2/10] Checking .NET Framework 4.7.2...
[PASS] .NET Framework 4.7.2 or later detected

[TEST 3/10] Checking Microsoft Outlook Installation...
[PASS] Microsoft Office detected

[TEST 4/10] Checking Plugin Registry Configuration...
[PASS] Plugin registry key found
  [SUBTEST] Checking AdminEmail value...
    [PASS] AdminEmail: yosi.desta@outlook.co.il

[TEST 5/10] Checking LoadBehavior Setting...
[PASS] LoadBehavior = 3 (Auto-load enabled)

[TEST 6/10] Checking Outlook Process...
[PASS] Outlook is not running

[TEST 7/10] Checking Plugin DLL...
[PASS] Plugin DLL found

[TEST 8/10] Checking Log Directory...
[PASS] Log directory exists
        Log files: 5

[TEST 9/10] Checking Email Configuration...
[PASS] Admin email configured: yosi.desta@outlook.co.il
[PASS] Email format is valid

[TEST 10/10] Checking Plugin Manifest...
[PASS] Plugin manifest configured
[PASS] Manifest file exists

╔════════════════════════════════════════════════════════════════════╗
║                          TEST SUMMARY                             ║
╚════════════════════════════════════════════════════════════════════╝

[PASS] Tests Passed: 10
[FAIL] Tests Failed: 0
[INFO] Tests Total:  10

✓ ALL TESTS PASSED - Plugin appears to be correctly installed

Recommendations:
  1. Open Outlook
  2. Look for "Report phishing" button in ribbon
  3. Select a test email
  4. Click the button to verify it works
  5. Check that email is reported successfully
```

---

## 🔧 TROUBLESHOOTING BY TEST

### **If Test 1 Fails: No Admin Rights**
```
Solution:
1. Right-click on test suite file
2. Select "Run as Administrator"
3. Run test again
```

### **If Test 2 Fails: .NET Framework**
```
Solution:
1. Go to https://dotnet.microsoft.com/
2. Download .NET Framework 4.7.2
3. Install it
4. Restart computer
5. Run tests again
```

### **If Test 3 Fails: Outlook Not Found**
```
Solution:
1. Install Microsoft Outlook
2. Or repair Office installation
3. Restart computer
4. Run tests again
```

### **If Test 4 Fails: Registry Key Missing**
```
Solution:
1. Run: FIX-ERROR-NOW.bat
2. Or reinstall the plugin
3. Run tests again
```

### **If Test 5 Fails: LoadBehavior Not Set**
```
Solution:
1. Run: FIX-ERROR-NOW.bat
2. Restart Outlook
3. Run tests again
```

### **If Test 6 Fails: Outlook Running**
```
Solution:
1. Close all Outlook windows
2. Check Task Manager for OUTLOOK.EXE
3. End process if still running
4. Run tests again
```

### **If Test 7 Fails: DLL Not Found**
```
Solution:
1. Reinstall the plugin
2. Run tests again
3. If still failing, check installation location
```

### **If Test 9 Fails: Email Not Configured**
```
Solution:
1. Run: FIX-ERROR-NOW.bat
2. Provide email: yosi.desta@outlook.co.il
3. Restart Outlook
4. Run tests again
```

---

## 🎯 WHEN TO RUN TESTS

### **After Installation**
```
Run tests to verify installation succeeded
```

### **When Plugin Isn't Working**
```
Run tests to identify the problem
```

### **After Updating**
```
Run tests to verify update succeeded
```

### **After System Changes**
```
Run tests after Windows updates or changes
```

### **Troubleshooting Configuration Error**
```
Run tests to find the exact issue
```

---

## 📊 TEST PHASES

### **Phase 1: System Requirements** (Tests 1-3)
- Admin rights
- .NET Framework
- Outlook installation

### **Phase 2: Plugin Configuration** (Tests 4-5)
- Registry entries
- Auto-load setting

### **Phase 3: Plugin Files** (Tests 6-7)
- DLL file
- Plugin process

### **Phase 4: Operational Readiness** (Tests 8-10)
- Logging capability
- Email configuration
- Manifest files

---

## ⏱️ TEST DURATION

| Scope | Time | Details |
|-------|------|---------|
| **Quick** | 1-2 min | Basic checks only |
| **Standard** | 3-5 min | All tests |
| **Full** | 5-10 min | With verbose output |

---

## 🚀 NEXT STEPS

### **If All Tests Pass**
```
1. Open Outlook
2. Look for "Report phishing" button
3. Select a test email
4. Click button to verify
5. Check email was reported
```

### **If Tests Fail**
```
1. Follow troubleshooting steps for failed test
2. Run FIX-ERROR-NOW.bat if needed
3. Restart Outlook
4. Run tests again
```

---

## 💡 PRO TIPS

✅ **Run as Administrator**
- Required for registry access

✅ **Close Outlook**
- Better test coverage

✅ **Run Regularly**
- Check system health
- Verify configuration

✅ **Keep Logs**
- Run tests periodically
- Track issues over time

---

## 📝 TEST FILES

| File | Type | Usage |
|------|------|-------|
| **RUN-TEST-SUITE.bat** | Batch | Easy one-click testing |
| **RUN-TEST-SUITE.ps1** | PowerShell | Advanced testing with options |

---

## 🎉 SUMMARY

**Test Suite Features:**
- ✅ 10 comprehensive tests
- ✅ Batch and PowerShell versions
- ✅ Clear pass/fail reporting
- ✅ Specific troubleshooting steps
- ✅ Takes only 3-5 minutes
- ✅ No installation required

**Ready to test?**
```
Navigate to: OutlookPhishingReporterSetup\Release\
Run: RUN-TEST-SUITE.bat (as Administrator)
```

---

**Run tests to verify your plugin is working correctly!** 🧪✅
