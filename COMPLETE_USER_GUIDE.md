# 📋 OUTLOOK PHISHING REPORTER - COMPLETE USER GUIDE

## Table of Contents
1. [Introduction](#introduction)
2. [Installation Guide](#installation-guide)
3. [How to Use the Plugin](#how-to-use-the-plugin)
4. [Troubleshooting](#troubleshooting)
5. [Uninstallation Guide](#uninstallation-guide)
6. [Advanced Features](#advanced-features)
7. [Log Management](#log-management)
8. [Frequently Asked Questions](#frequently-asked-questions)

---

## Introduction

The **Outlook Email Phishing Reporter Plugin** is a powerful security tool designed to help users quickly report suspicious emails directly to your organization's security team. By integrating seamlessly with Microsoft Outlook, this plugin streamlines the process of flagging potential threats, including phishing attempts, fraudulent messages, and suspicious content.

### Key Features
- ✅ One-click reporting directly from Outlook
- ✅ Works in both mail list and message views
- ✅ Automatic email archival to Deleted Items
- ✅ Configurable security team email address
- ✅ Custom icon and button label support
- ✅ Detailed logging for troubleshooting
- ✅ Compatible with Outlook 2016 and later

### System Requirements
- **Windows:** Windows 7 SP1 or later
- **Outlook:** Outlook 2016, 2019, or Microsoft 365
- **.NET Framework:** 4.7.2 or later
- **Admin Rights:** Required for installation
- **Email Server:** Connectivity required for reports

---

## Installation Guide

### Prerequisites
Before installing, ensure you have:
- Administrator access to your computer
- Microsoft Outlook 2016 or later installed and running
- .NET Framework 4.7.2 installed
- Sufficient disk space (200 MB)

### Installation Methods

#### Method 1: VBScript Installer (Recommended)

**Step 1: Locate the installer**
```
Navigate to: OutlookPhishingReporterSetup\Release\
Find: Install.vbs
```

**Step 2: Run the installer**
```
Double-click: Install.vbs
This will start the installation process
```

**Step 3: Configure the security email**
```
When prompted, enter your security team's email address:
Example: security@company.com
Format: username@domain.extension
```

**Step 4: Optional customization**
```
Custom Icon (optional):
- Provide path to PNG or ICO file
- Or leave blank for default icon

Custom Button Label (optional):
- Enter custom text for button
- Or leave blank for default "Report phishing"
```

**Step 5: Complete installation**
```
Click "Install" to finish
Installation takes 1-2 minutes
```

**Step 6: Restart Outlook**
```
Close all Outlook windows
Reopen Outlook
You will now see the "Report phishing" button
```

#### Method 2: Batch File Installer

**Step 1: Locate the installer**
```
Navigate to: OutlookPhishingReporterSetup\Release\
Find: setup.bat
```

**Step 2: Run as Administrator**
```
Right-click: setup.bat
Select: Run as Administrator
Follow the prompts
```

**Step 3: Complete installation**
```
Enter email address when prompted
Close and restart Outlook
```

#### Method 3: MSI Installer (Enterprise)

**Step 1: Build the MSI**
```
Navigate to: OutlookPhishingReporterSetup\MSISource\
Run: build-msi.bat
Requires WiX Toolset installed
```

**Step 2: Deploy the MSI**
```
Option A: Manual Installation
- Double-click: OutlookPhishingReporter.msi
- Follow wizard

Option B: Group Policy Deployment
- Deploy via SCCM or GPO
- Automatic for all users

Option C: Network Share
- Copy MSI to shared location
- Users run from network
```

---

## How to Use the Plugin

### Reporting from Mail List View

**Step 1: Select the suspicious email**
```
1. Open Outlook and go to your Inbox
2. Click on the email you suspect to be phishing
3. Note: You can only select ONE email at a time
```

**Step 2: Click the Report button**
```
1. Look at the Outlook ribbon
2. Find the "Report phishing" button (large icon)
3. Click the button
```

**Step 3: Confirm your report**
```
A dialog will appear with the message:
"You are about to report this email as phishing. 
The email will be forwarded to your security team 
for review. If you believe this email is legitimate 
or was selected by mistake, click Cancel."

Click "OK" to report
Click "Cancel" to abort
```

**Step 4: Verify success**
```
You will see:
"Thank you for reporting this email as phishing. 
It has been moved to the Deleted Items folder."

The email is now:
- Forwarded to your security team
- Moved to Deleted Items
- Logged for audit trail
```

### Reporting from Message View

**Step 1: Open the suspicious email**
```
1. Double-click the email to open it
2. The message reading window opens
```

**Step 2: Click the Report button**
```
1. Look at the message toolbar
2. Find "Report Phishing" button
3. Click the button
```

**Step 3: Confirm your report**
```
Same confirmation dialog as mail list view
Click "OK" to report
Click "Cancel" to abort
```

**Step 4: Verify success**
```
Success message appears
Message window closes
Email moved to Deleted Items
```

### What Happens to the Reported Email

When you report an email, the following occurs:

1. **Captured Data:**
   - Full email body
   - All attachments
   - Email headers (sender, recipient, subject, routing info)
   - Original email structure preserved

2. **Actions Taken:**
   - Email forwarded to security team
   - Original email moved to Deleted Items
   - Action logged for audit trail
   - Confirmation displayed to user

3. **Security Team Receives:**
   - Complete original email (.msg file)
   - Comprehensive metadata
   - User identification
   - Timestamp of report

---

## Troubleshooting

### Common Issues and Solutions

#### Error: "Please select at least one email"

**Problem:** No email is selected

**Solution:**
1. Click on an email in your inbox
2. Make sure it's highlighted/selected
3. Click the Report button again

---

#### Error: "Only one email at a time is allowed"

**Problem:** Multiple emails are selected

**Solution:**
1. Click on just ONE email
2. Deselect others by clicking elsewhere
3. Try reporting again

---

#### Error: "Configuration error: Invalid admin email address"

**Problem:** Security email not configured properly

**Solution - Quick Fix (2 minutes):**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Right-click: Diagnose-and-Repair.bat
3. Select: Run as Administrator
4. Enter email when prompted
5. Restart Outlook
```

**Solution - Manual Fix:**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail value
5. Enter: security@company.com
6. Restart Outlook
```

---

#### Error: "No Email Open"

**Problem:** Trying to report from message view without opening email

**Solution:**
1. Double-click an email to open it
2. Then click the Report button
3. Or use the mail list view instead

---

#### Error: "Failed to report the email"

**Problem:** Network or connectivity issue

**Solution:**
1. Check your internet connection
2. Verify security email is correct
3. Try again
4. Contact IT if problem persists

---

#### Button Doesn't Appear in Outlook

**Problem:** Button not showing after installation

**Solution:**
```
Step 1: Restart Outlook
- Close all Outlook windows
- Reopen Outlook

Step 2: Check Installation
- Verify plugin installed correctly
- Run diagnostic tool

Step 3: Repair Installation
- Run Diagnose-and-Repair.bat
- Or reinstall using Install.vbs

Step 4: Contact IT
- If button still missing
- Share diagnostic logs
```

---

### Diagnostic Tools

We provide two professional diagnostic tools:

#### Diagnose-and-Repair.bat

**Location:** `OutlookPhishingReporterSetup\Release\`

**How to use:**
```
1. Right-click: Diagnose-and-Repair.bat
2. Select: Run as Administrator
3. Tool will:
   - Check configuration
   - Identify issues
   - Offer to repair
   - Validate fix
```

**What it checks:**
- Registry key exists
- AdminEmail configured
- Email format valid
- LoadBehavior setting
- Description and Manifest

---

#### Diagnose-and-Repair.ps1

**Location:** `OutlookPhishingReporterSetup\Release\`

**Diagnostic mode (no repair):**
```
.\Diagnose-and-Repair.ps1
```

**Repair mode:**
```
.\Diagnose-and-Repair.ps1 -Repair
```

---

## Uninstallation Guide

### Method 1: Using Control Panel

**Step 1: Open Control Panel**
```
1. Right-click Start menu
2. Select: Settings
3. Go to: Apps → Apps & features
```

**Step 2: Find the plugin**
```
1. Look for "Outlook Phishing Reporter"
2. Or search for it by name
```

**Step 3: Uninstall**
```
1. Click on the plugin
2. Click: Uninstall
3. Confirm when prompted
```

---

### Method 2: Using Uninstall Script

**VBScript Method (Easiest):**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Double-click: Uninstall.vbs
3. Click: YES to confirm
4. Done!
```

**Batch File Method:**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Right-click: Uninstall.bat
3. Select: Run as Administrator
4. Confirm when prompted
```

**PowerShell Method:**
```
1. Open PowerShell as Administrator
2. Navigate to Release folder
3. Run: .\Uninstall.ps1
4. Confirm when prompted
```

---

### Method 3: MSI Uninstall

```
1. Navigate to: OutlookPhishingReporterSetup\Distribution\
2. Right-click: Uninstall-MSI.bat
3. Select: Run as Administrator
4. Choose uninstall method
5. Confirm
```

---

### After Uninstallation

✓ Plugin is completely removed  
✓ Registry entries deleted  
✓ Configuration cleared  
✓ Can reinstall anytime  

**Note:** Installation files can be safely deleted if no longer needed.

---

## Advanced Features

### Custom Icon Configuration

**Why use custom icons?**
- Brand your security reporting
- Use company logo
- Make button more recognizable

**How to configure:**

**During Installation:**
```
When installer asks:
"Do you want to set a custom icon? (Y/N): Y"

Enter path: C:\Icons\security-logo.png
```

**After Installation (Registry Edit):**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Create/Edit: CustomIconPath
5. Value: C:\path\to\icon.png
6. Restart Outlook
```

**Supported Formats:**
- PNG (.png)
- ICO (.ico)
- Recommended size: 32x32 or larger

---

### Custom Button Label

**How to configure:**

**During Installation:**
```
When installer asks:
"Do you want to customize the button label?"

Enter: Report Email as Unsafe
```

**After Installation (Registry Edit):**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Create/Edit: ButtonLabel
5. Value: Your Custom Label
6. Restart Outlook
```

**Examples:**
- "Report Phishing"
- "Report Email as Unsafe"
- "Security Alert"
- "Flag for Review"

---

### Changing the Security Email Address

**Method 1: Quick Fix Tool**
```
1. Right-click: QuickFix.bat
2. Select: Run as Administrator
3. Enter new email address
4. Restart Outlook
```

**Method 2: Registry Edit**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail
5. Enter new security email
6. Restart Outlook
```

**Method 3: Reinstall**
```
1. Run: Uninstall.vbs
2. Run: Install.vbs
3. Enter new email address
4. Restart Outlook
```

---

## Log Management

### Location of Log Files

```
C:\Users\[YourUsername]\AppData\Local\OutlookPhishingReporter\Logs\
```

**Log File Name:** `add-in.log`

### Viewing Logs

**Using Notepad:**
```
1. Open File Explorer
2. Navigate to log directory (see above)
3. Right-click: add-in.log
4. Select: Open with → Notepad
5. View log entries
```

**Using PowerShell:**
```
Get-Content "C:\Users\[YourUsername]\AppData\Local\OutlookPhishingReporter\Logs\add-in.log" -Tail 50
```

**Using Command Prompt:**
```
type "C:\Users\[YourUsername]\AppData\Local\OutlookPhishingReporter\Logs\add-in.log" | more
```

### Understanding Log Entries

**Log Format:**
```
YYYY-MM-DD HH:MM:SS.fff [LEVEL] Message
```

**Log Levels:**
- `[INFO]` - Normal operations
- `[ERROR]` - Errors and issues

**Example Log Entries:**
```
2024-01-05 14:23:45.123 [INFO] Report phishing clicked
2024-01-05 14:23:46.456 [INFO] Attached message: C:\Temp\abc123.msg
2024-01-05 14:23:47.789 [INFO] Phish report email sent successfully
2024-01-05 14:23:48.012 [INFO] Moved email to Deleted Items folder
2024-01-05 14:23:49.345 [INFO] Deleted temporary file: C:\Temp\abc123.msg
```

### Clearing Old Logs

**Safe to Delete:**
```
1. Open log folder
2. Delete old .log files
3. New logs generate automatically
4. No impact on plugin
```

**Recommended:**
- Keep logs for 30 days
- Archive important logs
- Regularly clear old logs

### Sharing Logs with IT Support

**When to share logs:**
- Plugin not working correctly
- Configuration errors
- Installation issues
- Reporting problems

**How to share:**
```
1. Open log file with Notepad
2. Copy contents
3. Paste in email to IT
4. Or send entire log file
5. Include description of issue
```

---

## Frequently Asked Questions

### Q: What email formats are accepted?

**A:** Standard email format:
```
Valid:
- security@company.com
- admin@yourdomain.org
- phishing@example.co.uk

Invalid:
- security (no @domain)
- @company.com (no username)
- admin @ company.com (spaces)
```

---

### Q: Can I report multiple emails at once?

**A:** No. The plugin allows reporting one email at a time. If you try to select multiple emails, you'll get an error message. Select one email and report it, then repeat for other emails.

---

### Q: Where does the reported email go?

**A:** After reporting:
1. Email is forwarded to security team
2. Original email moves to Deleted Items
3. Email is no longer in your Inbox
4. You can recover it from Deleted Items if needed

---

### Q: Is my email data secure?

**A:** Yes. The plugin:
- Only captures the selected email
- Includes full headers for investigation
- Sends to your organization's security team
- Doesn't access other emails in your mailbox
- Creates audit trail of all reports

---

### Q: What if I accidentally report the wrong email?

**A:** 
1. The email moves to Deleted Items
2. Security team still receives the report
3. You can recover email from Deleted Items
4. Contact security team to cancel investigation if needed

---

### Q: Can I customize the button appearance?

**A:** Yes. You can customize:
- **Custom Icon:** PNG or ICO file path
- **Custom Label:** Text for button
- **Custom Email:** Security team email address

See [Advanced Features](#advanced-features) for details.

---

### Q: What Outlook versions are supported?

**A:** 
- Outlook 2016
- Outlook 2019
- Microsoft 365 / Outlook Desktop App

**Not supported:**
- Outlook 2013 or earlier
- Outlook Web App (browser version)
- Outlook mobile apps

---

### Q: Do I need internet to use the plugin?

**A:** Yes. You need internet connectivity to:
- Forward reported emails to security team
- Communicate with organization's mail server
- Otherwise, plugin works offline

---

### Q: Can I report emails on behalf of others?

**A:** No. The plugin reports emails from your own mailbox only. Others must use the plugin on their own computers to report emails from their accounts.

---

### Q: What happens if reporting fails?

**A:** 
1. You'll see error message
2. Email is NOT moved to Deleted Items
3. Check internet connection
4. Verify security email address
5. Try again
6. Contact IT if persists

---

### Q: How often should I clear logs?

**A:** 
- Recommended: Clear monthly
- Or when folder gets large (>10MB)
- Safe to delete anytime
- New logs generate automatically

---

### Q: Can the button be removed from Outlook?

**A:** Yes, through uninstallation:
- Use Uninstall.vbs
- Use Control Panel → Add/Remove Programs
- Use Group Policy (enterprise)

---

### Q: What if I have issues after installation?

**A:** 
1. Try: Diagnose-and-Repair.bat
2. Review: Log files
3. Contact: Your IT department
4. Include: Diagnostic output and logs

---

## Support and Additional Resources

### Documentation Files
- `README_FIRST.md` - Quick start
- `SETUP_INSTRUCTIONS.md` - Detailed installation
- `DIAGNOSTIC_AND_REPAIR.md` - Troubleshooting
- `QUICK_FIX_GUIDE.md` - Common fixes

### Tools Provided
- `Install.vbs` - Easy installation
- `Uninstall.vbs` - Easy uninstallation
- `Diagnose-and-Repair.bat` - Automatic troubleshooting
- `QuickFix.bat` - Configuration repair

### Contact IT Support
```
If you need help:
1. Check this guide's troubleshooting section
2. Review log files (if applicable)
3. Run Diagnose-and-Repair.bat
4. Contact your IT department with:
   - Error message (if any)
   - Log file contents
   - Steps you've already tried
```

---

## Version Information

- **Product Name:** Outlook Phishing Reporter
- **Current Version:** 1.0
- **Release Date:** January 2026
- **Target Outlook:** 2016 and later
- **.NET Framework:** 4.7.2
- **Status:** Production Ready

---

**Thank you for helping protect your organization from phishing threats!** 🛡️

Last Updated: January 5, 2026
