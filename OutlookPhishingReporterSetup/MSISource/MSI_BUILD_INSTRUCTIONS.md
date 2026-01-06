# Building MSI Installer for Outlook Phishing Reporter

## Prerequisites
1. Install WiX Toolset v3.11 or later
   - Download: https://github.com/wixtoolset/wix3/releases

2. Install Visual Studio 2019 or later (optional)

## Build Steps

1. Open Command Prompt in OutlookPhishingReporterSetup\MSISource
2. Run these commands:

`atch
REM Compile WiX source file
"C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" Product.wxs

REM Link to create MSI
"C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe" -out ..\Distribution\OutlookPhishingReporter.msi Product.wixobj
`

## Deployment

Silent installation:
`atch
msiexec /i OutlookPhishingReporter.msi /quiet /norestart
`

With log file:
`atch
msiexec /i OutlookPhishingReporter.msi /l*v install.log
`

## Group Policy Deployment
1. Place MSI on network share: \\SERVER\Software\OutlookPhishingReporter.msi
2. Create GPO for deployment
3. Configure registry preferences for AdminEmail settings

## Support
For issues, contact your IT department
