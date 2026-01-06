# 📄 QUICK REFERENCE CARD

## INSTALLATION (Choose One)

### Option 1: VBScript (Easiest)
```
1. Double-click: Install.vbs
2. Enter email: security@company.com
3. Restart Outlook
✓ Done in 2 minutes!
```

### Option 2: Batch File
```
1. Right-click: setup.bat → Run as Administrator
2. Follow prompts
3. Enter email
4. Restart Outlook
```

### Option 3: MSI (Enterprise)
```
1. Double-click: OutlookPhishingReporter.msi
2. Follow wizard
3. Restart Outlook
```

---

## REPORTING EMAIL

### From Mail List
```
1. Click email to select
2. Click "Report phishing" button (ribbon)
3. Click OK to confirm
4. Email moved to Deleted Items
```

### From Message View
```
1. Double-click email to open
2. Click "Report Phishing" button
3. Click OK to confirm
4. Window closes
5. Email moved to Deleted Items
```

---

## TROUBLESHOOTING

### Error: "Invalid admin email"
```
Run: Diagnose-and-Repair.bat
Or: Win + R → regedit → Edit AdminEmail
```

### Button not showing
```
1. Restart Outlook
2. Run: Diagnose-and-Repair.bat
3. Reinstall if needed
```

### Can't report email
```
1. Check internet connection
2. Verify email address format
3. Try again
4. Contact IT
```

---

## CONFIGURATION

### Change Email Address
```
Registry: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
Edit: AdminEmail
Enter: security@company.com
Restart Outlook
```

### Custom Icon
```
Set Registry: CustomIconPath
Value: C:\path\to\icon.png
Restart Outlook
```

### Custom Button Label
```
Set Registry: ButtonLabel
Value: Your Custom Text
Restart Outlook
```

---

## UNINSTALL

### Option 1: VBScript (Easiest)
```
Double-click: Uninstall.vbs
Click YES
Done!
```

### Option 2: Batch File
```
Right-click: Uninstall.bat → Run as Administrator
Confirm uninstall
```

### Option 3: Control Panel
```
Settings → Apps → Apps & features
Find: Outlook Phishing Reporter
Click: Uninstall
```

---

## FILES & LOCATIONS

| File | Purpose |
|------|---------|
| `Install.vbs` | Install plugin |
| `Uninstall.vbs` | Remove plugin |
| `Diagnose-and-Repair.bat` | Fix errors |
| `QuickFix.bat` | Fix email config |
| `add-in.log` | View logs |

**Location:** `OutlookPhishingReporterSetup\Release\`

**Logs:** `C:\Users\[You]\AppData\Local\OutlookPhishingReporter\Logs\`

---

## VALID EMAIL FORMAT

```
✓ VALID:
- security@company.com
- admin@domain.org
- phishing@example.co.uk

✗ INVALID:
- security (no @)
- @company.com (no user)
- admin@company (missing extension)
```

---

## KEYBOARD SHORTCUTS

| Action | Keys |
|--------|------|
| Open Registry Editor | Win + R, regedit |
| Open Run Dialog | Win + R |
| Task Manager | Ctrl + Shift + Esc |

---

## CONTACT IT

**When to contact IT:**
- Installation fails
- Button doesn't appear
- Can't report email
- Error messages persist

**Provide:**
- Error message (exact text)
- Log file contents
- Steps already tried

---

## QUICK LINKS

- **Full Guide:** `COMPLETE_USER_GUIDE.md`
- **Troubleshooting:** `DIAGNOSTIC_AND_REPAIR.md`
- **Setup Help:** `SETUP_INSTRUCTIONS.md`
- **Error Fixes:** `CONFIG_ERROR_COMPLETE_SOLUTION.md`

---

**Print this card for quick reference!** 📌
