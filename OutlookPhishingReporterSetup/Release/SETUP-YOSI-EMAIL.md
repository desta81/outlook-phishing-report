# 📧 SETUP GUIDE - yosi.desta@outlook.co.il

## Overview

This guide explains how to set up the Outlook Phishing Reporter plugin to use **yosi.desta@outlook.co.il** as the security email address for phishing reports.

---

## ⚡ QUICK SETUP (1 MINUTE)

### Option 1: Batch File (Easiest)

**Steps:**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Right-click: Setup-Default-Email.bat
3. Select: Run as Administrator
4. Wait for completion
5. Restart Outlook
Done!
```

**Time:** 1 minute  
**Difficulty:** Very Easy  

---

### Option 2: PowerShell

**Steps:**
```
1. Open PowerShell as Administrator
2. Navigate to: OutlookPhishingReporterSetup\Release\
3. Run: .\Setup-Default-Email.ps1
4. Wait for completion
5. Restart Outlook
Done!
```

**Verify Configuration:**
```powershell
.\Setup-Default-Email.ps1 -Verify
```

**Time:** 1 minute  
**Difficulty:** Easy  

---

### Option 3: Manual Registry Edit

**Steps:**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail
5. Value: yosi.desta@outlook.co.il
6. Click: OK
7. Close Registry Editor
8. Restart Outlook
Done!
```

**Time:** 2 minutes  
**Difficulty:** Intermediate  

---

## 📋 EMAIL ADDRESS DETAILS

**Email:** `yosi.desta@outlook.co.il`

**What This Means:**
- All phishing reports will be sent to this email address
- Reports include full original email with headers
- Automatic audit trail maintained
- Can monitor incoming phishing alerts

**Format Verification:**
- ✅ Valid format (username@domain.extension)
- ✅ Proper domain structure (.co.il = Israel domain)
- ✅ Recognized email provider (Outlook)

---

## 🔍 VERIFY SETUP

After setup, verify everything is configured correctly:

### Check 1: Registry Configuration

**Using Registry Editor:**
```
1. Win + R → regedit
2. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
3. Check these values:
   - AdminEmail = yosi.desta@outlook.co.il
   - LoadBehavior = 3
   - Description = Outlook Phishing Reporter
```

**Using PowerShell:**
```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
```

**Using Command Prompt:**
```cmd
reg query "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
```

### Check 2: Outlook Button

```
1. Open Outlook
2. Look for "Report phishing" button in ribbon
3. Button should be visible and clickable
4. Test on a sample email (don't send for real)
```

### Check 3: Test Report

```
1. Select a test email
2. Click "Report phishing" button
3. Click OK to confirm
4. Check your inbox at yosi.desta@outlook.co.il
5. Should receive test report email
```

---

## 🎯 WHAT HAPPENS WHEN REPORTING

### User Action
```
1. User selects suspicious email
2. User clicks "Report phishing" button
3. User confirms dialog
4. Email is reported
```

### Behind the Scenes
```
1. Original email is saved to temporary file
2. New report email is created
3. Original email attached to report
4. Report sent to: yosi.desta@outlook.co.il
5. Original email moved to Deleted Items
6. Temporary files cleaned up
7. Success message shown to user
```

### What's Sent
```
To: yosi.desta@outlook.co.il
Subject: Phish Report
Body: "Please review the attached phish email for investigation."
Attachment: Original suspicious email with full headers
```

---

## 📊 CONFIGURATION REFERENCE

### Registry Location
```
HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
```

### Values Set

| Name | Type | Value | Purpose |
|------|------|-------|---------|
| **AdminEmail** | String | yosi.desta@outlook.co.il | Reports destination |
| **LoadBehavior** | DWORD | 3 | Auto-load on Outlook start |
| **Description** | String | Outlook Phishing Reporter | Display name |

---

## 🔧 TROUBLESHOOTING

### Button Not Appearing

**Problem:** "Report phishing" button missing from Outlook ribbon

**Solution:**
```
1. Restart Outlook completely
2. Close all Outlook windows
3. Run Setup-Default-Email.bat or .ps1 again
4. Reopen Outlook
5. Button should appear
```

### Reports Not Arriving

**Problem:** Sending reports but they don't arrive at yosi.desta@outlook.co.il

**Solution:**
```
1. Verify email address in registry is correct
2. Check email account is accessible
3. Check inbox for reports (might be in Junk)
4. Verify email forwarding rules
5. Test sending manual email to address
```

### "Configuration Error" Message

**Problem:** "Invalid admin email address" error when clicking button

**Solution:**
```
1. Run: Setup-Default-Email.bat as Administrator
2. Or manually set registry value again
3. Restart Outlook
4. Try again
```

### Registry Key Not Found

**Problem:** Registry key doesn't exist or is invalid

**Solution:**
```
1. Run: Setup-Default-Email.bat
2. This will create key and set all values
3. Restart Outlook
```

---

## 📧 EMAIL ACCOUNT SETUP

### Preparing yosi.desta@outlook.co.il to Receive Reports

**Recommended Setup:**

1. **Create Email Rule**
   - Outlook → File → Rules → Manage Rules
   - Create rule for "Phish Report" subject
   - Auto-file to specific folder
   - Mark as important

2. **Set Up Alert**
   - Enable notifications for new reports
   - Or configure auto-forward to team
   - Set up auto-response if needed

3. **Archive Configuration**
   - Keep reports for audit trail
   - Archive old reports monthly
   - Maintain for compliance

4. **Team Access** (if shared)
   - Share mailbox with security team
   - Delegate access to monitor
   - Set up team mailbox rules

---

## 🔐 SECURITY NOTES

### Email Account Security

✅ **Recommended:**
- Dedicated email for phishing reports only
- Not personal email address
- Strong password protection
- Multi-factor authentication enabled
- Access logged and monitored

⚠️ **Important:**
- Only authorized personnel should have access
- Monitor for unauthorized access
- Review reports regularly
- Maintain audit trail
- Archive for compliance

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
   - Check for arrival at yosi.desta@outlook.co.il

3. ✅ **Notify Users**
   - If shared configuration
   - Explain how to use
   - Provide training if needed

4. ✅ **Monitor Reports**
   - Check regularly for reports
   - Investigate suspicious patterns
   - Take action on confirmed threats
   - Document findings

---

## ⚙️ AUTOMATED SETUP SCRIPTS

### Batch File Option

**File:** `Setup-Default-Email.bat`

**Features:**
- Closes Outlook automatically
- Sets all registry values
- Displays confirmation
- No user interaction needed

**Usage:**
```
Right-click → Run as Administrator
```

### PowerShell Option

**File:** `Setup-Default-Email.ps1`

**Features:**
- Same functionality as batch
- Verify option with -Verify flag
- Detailed output
- Better error handling

**Usage:**
```
Run as Administrator:
.\Setup-Default-Email.ps1

Or with verification:
.\Setup-Default-Email.ps1 -Verify
```

---

## 📞 SUPPORT

### If Setup Fails

1. **Try Alternative Method**
   - If batch fails, try PowerShell
   - If PowerShell fails, try manual registry

2. **Check Prerequisites**
   - Admin rights needed
   - Outlook should be closed
   - Registry access available

3. **Verify Configuration**
   - Use verification command
   - Check registry directly
   - Test with sample email

4. **Contact Support**
   - Share setup log if available
   - Include registry screenshot
   - Describe specific error

---

## ✨ SUMMARY

**Quick Setup:**
- Run: `Setup-Default-Email.bat` as Administrator
- Or: `.\Setup-Default-Email.ps1` as Administrator
- Restart Outlook
- Done!

**Email Address:** yosi.desta@outlook.co.il

**Reports Include:**
- Original email with all headers
- Attachments (if any)
- Routing information
- Complete forensic data

**Audit Trail:**
- All reports logged
- Timestamps recorded
- User information tracked
- Full history maintained

---

**Setup is complete! Your Outlook Phishing Reporter is now configured to send all reports to yosi.desta@outlook.co.il** 📧

For detailed information, see: `EMAIL_ADDRESS_MANAGEMENT_GUIDE.md`
