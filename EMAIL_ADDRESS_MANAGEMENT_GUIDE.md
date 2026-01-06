# 📧 DEFAULT EMAIL ADDRESS MANAGEMENT GUIDE

## Overview

This guide explains how to configure, manage, and update the default security email address for the Outlook Phishing Reporter plugin.

---

## 🎯 CONFIGURATION METHODS

### Method 1: During Installation (Recommended)

**Using Install.vbs:**
```
1. Run: Install.vbs
2. When prompted: "Enter the admin email address for phishing reports"
3. Enter: security@company.com (your security team email)
4. Installation continues
5. Email is saved to registry
```

**Using setup.bat:**
```
1. Run: setup.bat as Administrator
2. When prompted for email address
3. Enter: security@company.com
4. Installation complete
```

---

### Method 2: After Installation - Quick Fix

**Using QuickFix.bat:**
```
1. Navigate to: OutlookPhishingReporterSetup\Release\
2. Right-click: QuickFix.bat
3. Select: Run as Administrator
4. Enter new email address
5. Restart Outlook
```

---

### Method 3: Manual Registry Edit

**Using Registry Editor:**
```
1. Press: Win + R
2. Type: regedit
3. Navigate to:
   HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
4. Edit: AdminEmail value
5. Enter: security@company.com
6. Click: OK
7. Restart Outlook
```

**Using PowerShell:**
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
$email = "security@company.com"
New-ItemProperty -Path $regPath -Name "AdminEmail" -Value $email -PropertyType String -Force
```

**Using Command Prompt:**
```cmd
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" /v "AdminEmail" /t REG_SZ /d "security@company.com" /f
```

---

## 📋 VALID EMAIL ADDRESS FORMAT

### Requirements
- Must contain exactly one `@` symbol
- Username before `@` (cannot be empty)
- Domain after `@` (cannot be empty)
- Domain extension (e.g., .com, .org, .co.uk)
- No spaces or special characters except: `.`, `_`, `-`, `+`

### Valid Examples
```
security@company.com
admin@yourdomain.org
phishing.report@example.co.uk
security-team@organization.net
it+phishing@domain.com
first.last@company.com
```

### Invalid Examples
```
security               (missing @domain)
@company.com           (missing username)
security @company.com  (space in email)
security@              (missing domain extension)
security@.com          (missing domain name)
security@@company.com  (double @@)
```

---

## 🔧 DEFAULT EMAIL CONFIGURATION

### Built-in Default
The plugin includes a default email address constant:
```csharp
private const string DEFAULT_EMAIL = "security@company.com";
```

### How Defaults Work

1. **Installation Time:**
   - User is prompted for email address
   - Email is saved to registry
   - No default is used if user provides valid email

2. **Runtime Check:**
   - Plugin loads email from registry
   - If email is missing or invalid:
     - Error logged
     - Error message shown to user
     - Button click is prevented

3. **Fallback:**
   - No automatic fallback to default
   - User must configure valid email
   - This ensures intentional configuration

---

## 🚨 TROUBLESHOOTING

### Error: "Invalid admin email address"

**Problem:** Email not configured or invalid format

**Solution:**
```
1. Run: QuickFix.bat as Administrator
2. Or: Edit registry manually
3. Enter valid email address
4. Restart Outlook
```

### Email Address Not Taking Effect

**Solution:**
```
1. Restart Outlook completely
2. Close all Outlook windows
3. Verify registry has correct value
4. Reopen Outlook
```

### Need to Change Email Address

**Option 1 - Quick:**
```
Run: QuickFix.bat
Enter new email
Restart Outlook
```

**Option 2 - Manual:**
```
Regedit → Edit AdminEmail value → Save → Restart Outlook
```

**Option 3 - PowerShell:**
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
Set-ItemProperty -Path $regPath -Name "AdminEmail" -Value "newemail@company.com"
```

---

## 🔄 EMAIL ADDRESS CHANGE PROCEDURES

### Scenario 1: Department Email Changes
```
Old email: phishing@olddept.com
New email: phishing@newdept.com

Steps:
1. QuickFix.bat → Enter new email
2. Or: Registry edit
3. Verify with test report
```

### Scenario 2: Distribute to Multiple Users
```
Option A - Individual:
- Each user runs Install.vbs
- Enters their organization's email

Option B - Enterprise:
- Use MSI with configured email
- Deploy via Group Policy
- Email pre-configured for all

Option C - Batch Script:
- Create script with email pre-filled
- Deploy to users
- Minimal user interaction
```

### Scenario 3: Transition to New Security Team
```
1. Configure new email address
2. Test reporting
3. Verify emails arrive at new address
4. Notify users of change
5. Remove old email from distribution list
```

---

## 📊 REGISTRY REFERENCE

**Location:**
```
HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
```

**AdminEmail Entry:**
```
Name:     AdminEmail
Type:     String (REG_SZ)
Example:  security@company.com
Required: Yes
```

**Related Entries:**
```
LoadBehavior: 3 (auto-load)
Description: Outlook Phishing Reporter
Manifest: Path to manifest file
CustomIconPath: (optional) custom icon
ButtonLabel: (optional) custom label
```

---

## ✅ VERIFICATION CHECKLIST

After configuring email address, verify:

- [ ] Email is valid format (has @ and domain extension)
- [ ] Email is correctly entered in registry
- [ ] Outlook has been restarted
- [ ] Button appears in ribbon
- [ ] Button is clickable
- [ ] No error when clicking button
- [ ] Test email received by security team
- [ ] Reported email moved to Deleted Items

---

## 📝 BATCH DEPLOYMENT SCRIPT

**Deploy email to multiple users:**

```batch
@echo off
REM Batch deploy phishing reporter email config

set REGPATH=HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
set EMAIL=security@company.com

REM Create registry key
reg add "%REGPATH%" /f

REM Set AdminEmail
reg add "%REGPATH%" /v "AdminEmail" /t REG_SZ /d "%EMAIL%" /f

REM Set other required values
reg add "%REGPATH%" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "%REGPATH%" /v "Description" /t REG_SZ /d "Outlook Phishing Reporter" /f

echo Configuration complete: %EMAIL%
```

---

## 🔐 SECURITY CONSIDERATIONS

### Best Practices

1. **Dedicated Security Email:**
   - Use mailbox specifically for phishing reports
   - Not personal email
   - Monitored by security team

2. **Access Control:**
   - Limit access to phishing reports mailbox
   - Audit access logs
   - Monitor for abuse

3. **Email Monitoring:**
   - Set up filters for phishing reports
   - Configure auto-responses
   - Implement mail rules

4. **Validation:**
   - Always verify email format
   - Test with sample report
   - Confirm receipt by security team

---

## 🎓 ADVANCED CONFIGURATION

### PowerShell Module Example

```powershell
function Set-PhishingReporterEmail {
    param(
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress,
        
        [switch]$Force
    )
    
    # Validate email
    if ($EmailAddress -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        throw "Invalid email format: $EmailAddress"
    }
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    
    if (-not $Force) {
        $confirm = Read-Host "Set email to $EmailAddress ? (Y/N)"
        if ($confirm -ne 'Y') { return }
    }
    
    Set-ItemProperty -Path $regPath -Name "AdminEmail" -Value $EmailAddress -Force
    Write-Host "Email configured: $EmailAddress"
}
```

---

## 📞 COMMON EMAIL SCENARIOS

| Scenario | Email Address | Notes |
|----------|---------------|-------|
| **Small Org** | security@company.com | Single email |
| **Department** | phishing@dept.company.com | Department specific |
| **Ticketing System** | phishing-reports@tickets.company.com | Auto-creates tickets |
| **Shared Mailbox** | phishing-team@company.com | Multiple monitors |
| **Individual** | john.doe@company.com | Personal responsibility |
| **Distribution List** | phishing-team@company.com | Routed to group |

---

## 🚀 DEPLOYMENT EXAMPLES

### Example 1: Single User Installation
```
User runs: Install.vbs
Enters: security@company.com
Result: Email configured for user
```

### Example 2: Group Policy Deployment
```
1. Build MSI with email configured
2. Deploy via Group Policy
3. All users get same email
4. No user interaction needed
```

### Example 3: Script-Based Deployment
```
1. Create batch script with email
2. Distribute to users
3. Users run script
4. Email auto-configured
5. Minimal support needed
```

---

## ✨ SUMMARY

**Email Address Configuration:**
- ✅ Set during installation (recommended)
- ✅ Changed via QuickFix.bat (easiest after install)
- ✅ Edited via Registry (manual control)
- ✅ Configured via PowerShell (automation)
- ✅ Deployed via Group Policy (enterprise)

**Key Points:**
- Must be valid email format
- Must be monitored security mailbox
- Can be changed anytime
- Restart Outlook to apply changes
- Test with sample report

**Tools Provided:**
- `Install.vbs` - Set during install
- `QuickFix.bat` - Change anytime
- `Diagnose-and-Repair.bat` - Fix issues
- PowerShell scripts - Automation

---

**Everything you need to manage your phishing reporter email address!** 📧
