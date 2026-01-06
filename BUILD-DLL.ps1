# ========================================================================
# Outlook Phishing Reporter - Build DLL File (PowerShell)
# ========================================================================
# This script builds the plugin DLL using MSBuild

param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$Clean = $false,
    [switch]$Publish = $false
)

$ErrorActionPreference = "Stop"

function Show-Header {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   OUTLOOK PHISHING REPORTER - BUILD SCRIPT                        ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Test-ProjectFile {
    $projectPath = "OutlookPhishingReporter\OutlookPhishingReporter.csproj"
    
    if (-not (Test-Path $projectPath)) {
        Write-Host "ERROR: Project file not found!" -ForegroundColor Red
        Write-Host "Expected: $projectPath" -ForegroundColor Yellow
        exit 1
    }
    
    Write-Host "[+] Project file found: $projectPath" -ForegroundColor Green
}

function Build-Project {
    param(
        [string]$Configuration,
        [string]$Platform = "AnyCPU"
    )
    
    $projectPath = "OutlookPhishingReporter\OutlookPhishingReporter.csproj"
    
    Write-Host "[*] Building $Configuration configuration..." -ForegroundColor Cyan
    Write-Host ""
    
    try {
        msbuild $projectPath /p:Configuration=$Configuration /p:Platform=$Platform
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "ERROR: Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
            Write-Host ""
            Write-Host "Troubleshooting:" -ForegroundColor Yellow
            Write-Host "  1. Ensure Visual Studio is installed" -ForegroundColor White
            Write-Host "  2. Check for missing references" -ForegroundColor White
            Write-Host "  3. Verify .NET Framework 4.7.2 is installed" -ForegroundColor White
            Write-Host "  4. Run from Visual Studio Developer Command Prompt" -ForegroundColor White
            Write-Host ""
            exit 1
        }
    } catch {
        Write-Host "ERROR: $($_)" -ForegroundColor Red
        exit 1
    }
}

function Verify-Build {
    param(
        [string]$Configuration
    )
    
    $dllPath = "OutlookPhishingReporter\bin\$Configuration\OutlookPhishingReporter.dll"
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                      BUILD VERIFICATION                           ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    
    if (Test-Path $dllPath) {
        $fileInfo = Get-Item $dllPath
        $sizeMB = "{0:N2}" -f ($fileInfo.Length / 1MB)
        
        Write-Host "[+] DLL File Created Successfully" -ForegroundColor Green
        Write-Host ""
        Write-Host "Location: $dllPath" -ForegroundColor White
        Write-Host "Size:     $sizeMB MB" -ForegroundColor White
        Write-Host "Date:     $($fileInfo.LastWriteTime)" -ForegroundColor White
        Write-Host ""
        
        return $true
    } else {
        Write-Host "[ERROR] DLL file not found!" -ForegroundColor Red
        Write-Host "Expected: $dllPath" -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
}

function Show-NextSteps {
    param(
        [string]$Configuration
    )
    
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                          NEXT STEPS                               ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "1. Build MSI Installer (Optional):" -ForegroundColor Yellow
    Write-Host "   Navigate to: OutlookPhishingReporterSetup\MSISource\" -ForegroundColor White
    Write-Host "   Run: build-msi.bat" -ForegroundColor White
    Write-Host ""
    
    Write-Host "2. Test the Plugin:" -ForegroundColor Yellow
    Write-Host "   Run: RUN-TEST-SUITE.bat" -ForegroundColor White
    Write-Host ""
    
    Write-Host "3. Install Directly:" -ForegroundColor Yellow
    Write-Host "   Copy DLL to: `$env:APPDATA\Microsoft\Addins\" -ForegroundColor White
    Write-Host "   Restart Outlook" -ForegroundColor White
    Write-Host ""
    
    Write-Host "DLL Ready for:" -ForegroundColor Cyan
    Write-Host "  ✓ MSI packaging" -ForegroundColor White
    Write-Host "  ✓ Direct installation" -ForegroundColor White
    Write-Host "  ✓ Enterprise deployment" -ForegroundColor White
    Write-Host "  ✓ Testing and verification" -ForegroundColor White
    Write-Host ""
}

# Main execution
Show-Header

Write-Host "Configuration: $Configuration" -ForegroundColor White
Write-Host "Platform: AnyCPU" -ForegroundColor White
Write-Host ""

Test-ProjectFile

if ($Clean) {
    Write-Host "[*] Cleaning previous builds..." -ForegroundColor Cyan
    msbuild "OutlookPhishingReporter\OutlookPhishingReporter.csproj" /t:Clean /p:Configuration=$Configuration | Out-Null
    Write-Host "[+] Clean complete" -ForegroundColor Green
    Write-Host ""
}

Build-Project -Configuration $Configuration

if (Verify-Build -Configuration $Configuration) {
    Show-NextSteps -Configuration $Configuration
    
    if ($Publish) {
        Write-Host "[*] Publishing application..." -ForegroundColor Cyan
        msbuild "OutlookPhishingReporter\OutlookPhishingReporter.csproj" /p:Configuration=$Configuration /t:Publish
        Write-Host "[+] Publish complete" -ForegroundColor Green
    }
} else {
    exit 1
}

Write-Host "Build script completed." -ForegroundColor Green
