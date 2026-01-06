# Outlook Phishing Reporter - Build Instructions

## Current Status
✅ All code complete with 5 features:
1. Large button icons
2. Message toolbar button (Inspector view)
3. Auto-move to Deleted Items folder
4. Custom success popup message
5. Custom icon and label support

## Icon File
**Custom Icon:** `phishing.png` (shield with fishing hook)
- Location: `C:\Users\YosiD\OneDrive\Desktop\etio\Outlook-Phishing-Reporter\phishing.png`
- This icon will be embedded in the add-in

## How to Build

### Option 1: Using Visual Studio (Recommended)

1. **Install Office Development Tools:**
   - Open Visual Studio Insiders (already running)
   - Go to **Tools** → **Get Tools and Features**
   - Check **"Office/SharePoint development"**
   - Click **"Modify"** and wait 5-10 minutes

2. **Update Icon in Visual Studio:**
   - In Solution Explorer, expand **OutlookPhishingReporter** project
   - Double-click **OutlookPhishingReporter.cs [Design]**
   - Click on the **btnReportPhishing** button
   - In Properties window, find **Image** property
   - Click the "..." button
   - Click **Import** and select `phishing.png`
   - Repeat for **btnReportPhishingRead** button
   - Save all (Ctrl+Shift+S)

3. **Build the Solution:**
   - Press **Ctrl+Shift+B** (Build Solution)
   - Wait for build to complete

4. **Build the MSI Installer:**
   - In Solution Explorer, right-click **OutlookPhishingReporterSetup**
   - Click **Build**

5. **Get Your Files:**
   - **MSI:** `OutlookPhishingReporterSetup\Release\OutlookPhishingReporterSetup.msi`
   - **Setup.exe:** `OutlookPhishingReporterSetup\Release\setup.exe`

### Option 2: Command Line (After Office Tools Installed)

```powershell
# Build the solution
& 'C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\MSBuild.exe' `
    'OutlookPhishingReporter.sln' `
    /p:Configuration=Release `
    /t:Rebuild
```

## Installation

After building, use the MSI to install:

```batch
OutlookPhishingReporterSetup\Release\setup.exe
```

Or run:
```batch
install.bat
```

## Configuration

After installation, configure the admin email:

```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
New-ItemProperty -Path $regPath -Name "AdminEmail" -Value "yosi.desta@outlook.co.il" -PropertyType String -Force
```

## Custom Branding (Optional)

To use a different icon or label per installation:

```powershell
# Custom icon
New-ItemProperty -Path $regPath -Name "CustomIconPath" -Value "C:\Company\custom-icon.png" -PropertyType String -Force

# Custom button label
New-ItemProperty -Path $regPath -Name "ButtonLabel" -Value "Report Threat" -PropertyType String -Force
```

## Files Overview

| File | Purpose |
|------|---------|
| `OutlookPhishingReporter.cs` | Main business logic (338 lines) |
| `OutlookPhishingReporter.Designer.cs` | Ribbon UI configuration |
| `OutlookPhishingReporter.resx` | Resources (icons, strings) |
| `phishing.png` | Default button icon |
| `install.bat` | Installation launcher |
| `Install.ps1` | PowerShell installer |

## Troubleshooting

**Q: Build fails with "OfficeTools not found"**
A: Install Office/SharePoint development workload in Visual Studio

**Q: How do I update the icon?**
A: Open the Designer view and update the Image property, or place a custom icon and set CustomIconPath in registry

**Q: Where are the output files?**
A: Check `OutlookPhishingReporter\bin\Release\` for DLL and `OutlookPhishingReporterSetup\Release\` for MSI

## Support

Admin Email: yosi.desta@outlook.co.il
