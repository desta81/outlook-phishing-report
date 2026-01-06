# Outlook Phishing Reporter

A VSTO (Visual Studio Tools for Office) add-in for Microsoft Outlook that enables users to report phishing emails directly from their inbox.

## Overview

The Outlook Phishing Reporter plugin adds a "Report phishing" button to the Outlook ribbon, allowing users to quickly report suspicious emails to their security team. The plugin automatically forwards the suspected phishing email to a configured security inbox and moves the reported email to the Deleted Items folder.

## Features

- **Easy Reporting**: One-click reporting of phishing emails from the Outlook ribbon
- **Automatic Forwarding**: Sends reported emails to configured security team email address
- **Email Validation**: Validates email addresses for proper configuration
- **Logging**: Comprehensive logging system for troubleshooting
- **Registry Configuration**: Stores settings in Windows Registry for easy customization
- **Multi-View Support**: Works in both Mail Explorer and Mail Read inspector views
- **Automatic Cleanup**: Removes temporary files after sending reports
- **Error Handling**: Robust error handling with user-friendly messages

## System Requirements

- Microsoft Outlook 2016 or later
- Windows 7 or later
- .NET Framework 4.7.2
- Administrator privileges for installation

## Installation

### Using MSI Installer (Recommended)

1. Download the `OutlookPhishingReporter.msi` file
2. Run the installer with administrator privileges
3. Follow the installation wizard
4. Restart Outlook
5. Configure the security team email address (see Configuration section)

### Manual Installation

1. Build the project from source
2. Register the add-in in Outlook
3. Configure the email address via registry

## Configuration

### Default Email (Debug Mode)

In debug mode, the plugin uses `yosi.desta@outlook.co.il` as the default email address.

### Setting Custom Email Address

Edit the Windows Registry:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
```

Add or modify:
- **Value Name**: `AdminEmail`
- **Value Data**: Your security team email address (e.g., security@company.com)

### Optional Registry Settings

- **CustomIconPath**: Path to custom button icon (e.g., C:\icons\phishing-icon.png)
- **ButtonLabel**: Custom button label text (default: "Report phishing")

## Usage

### Reporting an Email

1. Select an email in your Outlook inbox
2. Click the "Report phishing" button in the Home tab of the ribbon
3. Confirm the action in the dialog box
4. The email will be forwarded to the configured security email and moved to Deleted Items

### Keyboard Shortcut

Currently, the button must be clicked via the ribbon. Keyboard shortcuts can be added in future versions.

## Architecture

### Project Structure

```
OutlookPhishingReporter/
├── OutlookPhishingReporter.cs          # Main ribbon and event handlers
├── OutlookPhishingReporter.Designer.cs # UI designer code
├── ThisAddIn.cs                        # VSTO add-in lifecycle
├── Utilities/
│   └── FileLogger.cs                   # Logging functionality
└── Properties/                         # Assembly metadata
```

### Key Components

- **PhishingReporterRibbon**: Main ribbon UI component with button click handlers
- **FileLogger**: Static logger for debugging and troubleshooting
- **Registry Integration**: Reads configuration from Windows Registry
- **Email Validation**: Validates recipient email addresses

## Logging

The plugin creates logs in:
```
C:\Users\<Username>\AppData\Local\OutlookPhishingReporter\Logs\add-in.log
```

Check this file for troubleshooting information.

## Development

### Building from Source

1. Clone the repository:
   ```bash
   git clone -b develop https://github.com/samicsc0/Outlook-Spam-Reporter.git
   ```

2. Open in Visual Studio:
   ```
   File → Open Project/Solution → OutlookPhishingReporter.csproj
   ```

3. Build the solution:
   ```
   Build → Build Solution
   ```

4. Debug:
   ```
   Press F5 to start debugging (Outlook will launch with the plugin)
   ```

### Building MSI Installer

```bash
cd OutlookPhishingReporterSetup\MSISource
build-msi.bat
```

The MSI will be created in: `OutlookPhishingReporterSetup\Distribution\`

## Git Branches

- **main**: Production-ready code
- **develop**: Active development branch

Create feature branches from `develop` for new features.

## Troubleshooting

### Button Not Appearing

1. Close Outlook completely
2. Restart Visual Studio debug session (F5)
3. If still missing, check registry configuration

### "Configuration error: Invalid admin email address"

Ensure the email address is configured in the registry:
```
HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter
Value: AdminEmail
```

### Plugin Not Loading

1. Check Windows Event Viewer for VSTO errors
2. Verify .NET Framework 4.7.2 is installed
3. Run Outlook in safe mode to check for conflicts

## License

[Specify your license here]

## Support

For issues and questions:
- Check the logs in `%LocalAppData%\OutlookPhishingReporter\Logs\`
- Review the GitHub issues page
- Contact your system administrator

## Version History

See `RELEASE_NOTES.md` for detailed version information.

## Technologies Used

- **Language**: C# 7.3
- **.NET Framework**: 4.7.2
- **UI Framework**: VSTO (Visual Studio Tools for Office)
- **Build Tool**: WiX Toolset (for MSI installer)
- **Office Version**: Outlook 2016+

## Contributing

1. Create a feature branch from `develop`
2. Make your changes
3. Create a pull request for review
4. Merge after approval

## Author

Yosi Desta

## Changelog

### v1.0.0
- Initial release
- Phishing report functionality
- Registry-based configuration
- MSI installer support
- Comprehensive logging

---

**Repository**: https://github.com/samicsc0/Outlook-Spam-Reporter
