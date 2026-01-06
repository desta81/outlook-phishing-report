# 📧 SETUP GUIDE - report@phishcage.cybeready.net

## Overview

This guide explains how to configure the Outlook Phishing Reporter plugin to use **report@phishcage.cybeready.net** as the security email address for phishing reports.

---

## ⚡ QUICK SETUP (1 MINUTE)

### **Option 1: Batch File (Easiest)**

**Steps:**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Right-click: Setup-CybeReady-Email.bat
3. Select: Run as Administrator
4. Wait for completion
5. Restart Outlook
Done!
```

**Time:** 1 minute  
**Difficulty:** Very Easy  

---

### **Option 2: PowerShell**

**Steps:**
```powershell
# Open PowerShell as Administrator
cd OutlookPhishingReporterSetup\Release
.\Setup-CybeReady-Email.ps1
```

**Verify Configuration:**
```powershell
.\Setup-CybeReady-Email.ps1 -Verify
```

**Time:** 1 minute  
**Difficulty:** Easy  

---

### **Option 3: Manual Registry Edit**

**Steps:**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail
5. Value: report@phishcage.cybeready.net
6. Click: OK
7. Restart Outlook
Done!
```

**Time:** 2-3 minutes  
**Difficulty:** Intermediate  

---

## 📋 EMAIL ADDRESS DETAILS

**Email:** `report@phishcage.cybeready.net`

**Format:** ✅ Valid
- Username: report
- Domain: phishcage.cybeready.net
- Extension: .net

**What This Means:**
- All phishing reports sent to this address
- Reports include full original email with headers
- Can monitor incoming phishing alerts
- Part of CybeReady phishing cage system

---

## 🎯 STEP-BY-STEP SETUP

### **Using Setup-CybeReady-Email.bat**

**Step 1: Find the file**
```
Open: OutlookPhishingReporterSetup\Release\
Look for: Setup-CybeReady-Email.bat
```

**Step 2: Run as admin**
```
Right-click on Setup-CybeReady-Email.bat
Select: Run as Administrator
```

**Step 3: Wait for script**
```
You will see:
  [*] Closing Outlook...
  [+] Outlook closed
  [*] Configuring registry...
  [+] Configuration saved
  [*] Verifying configuration...
  [+] Email verified: report@phishcage.cybeready.net
  
  CONFIGURATION COMPLETE!
```

**Step 4: Restart Outlook**
```
Close all Outlook windows
Wait 5 seconds
Reopen Outlook
Check for "Report phishing" button
```

**Step 5: Verify it works**
```
1. Select any email
2. Click "Report phishing" button
3. Confirm dialog
4. Email should be reported
5. No error! ✓
```

---

## ✔️ HOW TO VERIFY THE SETUP

### **Check 1: Visual Inspection**
```
1. Open Outlook
2. No error dialog should appear
3. "Report phishing" button should be visible
4. Button should be clickable
```

### **Check 2: Test Report**
```
1. Select a test email
2. Click "Report phishing" button
3. Dialog appears (normal)
4. Click OK to confirm
5. Email is reported to: report@phishcage.cybeready.net
6. No error! ✓
```

### **Check 3: Registry Verification**
```powershell
# Run in PowerShell as Administrator:
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -Name AdminEmail
```

Should return:
```
AdminEmail : report@phishcage.cybeready.net
```

---

## 📊 REGISTRY CONFIGURATION

### **Registry Location**
```
HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
```

### **Values Set**

| Name | Type | Value | Purpose |
|------|------|-------|---------|
| **AdminEmail** | String | report@phishcage.cybeready.net | Reports destination |
| **LoadBehavior** | DWORD | 3 | Auto-load on Outlook start |
| **Description** | String | Outlook Phishing Reporter | Display name |

---

## 🔧 TROUBLESHOOTING

### **Button Not Appearing**

**Problem:** "Report phishing" button missing from Outlook ribbon

**Solution:**
```
1. Restart Outlook completely
2. Close all Outlook windows
3. Run Setup-CybeReady-Email.bat or .ps1 again
4. Reopen Outlook
5. Button should appear
```

### **Reports Not Arriving**

**Problem:** Sending reports but they don't arrive at report@phishcage.cybeready.net

**Solution:**
```
1. Verify email address in registry is correct
2. Check email account is accessible
3. Check inbox for reports (might be in spam/junk)
4. Verify email forwarding rules
5. Test sending manual email to address
```

### **"Configuration Error" Message**

**Problem:** "Invalid admin email address" error when clicking button

**Solution:**
```
1. Run: Setup-CybeReady-Email.bat as Administrator
2. Or manually set registry value again
3. Restart Outlook
4. Try again
```

### **Registry Key Not Found**

**Problem:** Registry key doesn't exist or is invalid

**Solution:**
```
1. Run: Setup-CybeReady-Email.bat
2. This will create key and set all values
3. Restart Outlook
```

---

## 📧 CYBEREADY SETUP

### **Preparing report@phishcage.cybeready.net to Receive Reports**

**CybeReady Phishing Cage Integration:**
1. Account is monitored by CybeReady system
2. Reports analyzed automatically
3. Phishing indicators extracted
4. Intelligence shared with organization
5. Dashboard available for monitoring

**Recommended Setup:**
1. Verify account credentials
2. Check CybeReady dashboard access
3. Configure alert rules
4. Enable notifications
5. Monitor report trends

---

## 🔐 SECURITY NOTES

### **Email Account Security**

✅ **Recommended:**
- Dedicated email for phishing reports only
- Not personal email address
- Strong password protection
- Multi-factor authentication enabled
- Access monitored by CybeReady

⚠️ **Important:**
- CybeReady email has enhanced monitoring
- Reports automatically analyzed
- Intelligence extracted automatically
- Organization dashboard provided
- Regular security reviews

---

## 📝 NEXT STEPS

After setup is complete:

1. ✅ **Restart Outlook**
   - Close all windows
   - Reopen Outlook
   - Verify button appears

2. ✅ **Test Configuration**
   - Select a test email
   - Click "Report phishing"
   - Confirm dialog
   - Check for arrival at report@phishcage.cybeready.net

3. ✅ **Access CybeReady Dashboard**
   - Log in to CybeReady
   - View phishing intelligence
   - Monitor report trends
   - Configure alerts

4. ✅ **Notify Users**
   - If shared configuration
   - Explain how to use
   - Provide CybeReady resources

---

## 📋 COMPLETE SETUP COMMAND

**Copy and paste in Command Prompt (as Administrator):**
```cmd
cd OutlookPhishingReporterSetup\Release && Setup-CybeReady-Email.bat
```

**Or PowerShell (as Administrator):**
```powershell
cd 'OutlookPhishingReporterSetup\Release'; .\Setup-CybeReady-Email.ps1
```

---

## ✨ SUMMARY

**Setup Method:**
- Run: Setup-CybeReady-Email.bat (easiest)
- Or: Setup-CybeReady-Email.ps1
- Or: Manual registry edit

**Email Address:** report@phishcage.cybeready.net

**Reports Include:**
- Original email with all headers
- Attachments (if any)
- Routing information
- Complete forensic data

**CybeReady Integration:**
- Automatic analysis
- Intelligence extraction
- Dashboard monitoring
- Alert configuration
- Security intelligence

---

**Setup is complete! Your Outlook Phishing Reporter is now configured to send all reports to report@phishcage.cybeready.net** 📧
