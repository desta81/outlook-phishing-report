param(
    [switch]$Verify = $false
)

# Outlook Phishing Reporter - Setup for yosi.desta@outlook.co.il
# Usage: .\Setup-Default-Email.ps1 or .\Setup-Default-Email.ps1 -Verify

$AdminEmail = "yosi.desta@outlook.co.il"
$RegPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"

function Show-Header {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Outlook Phishing Reporter - Setup" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Test-AdminRights {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "ERROR: This script requires administrator privileges." -ForegroundColor Red
        Write-Host "Please run PowerShell as Administrator." -ForegroundColor Yellow
        exit 1
    }
}

function Configure-Email {
    Write-Host "[*] Closing Outlook..." -ForegroundColor Cyan
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
    
    Write-Host "[*] Configuring registry..." -ForegroundColor Cyan
    
    # Create registry key
    New-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
    
    # Set values
    New-ItemProperty -Path $RegPath -Name "AdminEmail" -Value $AdminEmail -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force | Out-Null
    
    Write-Host "[+] Configuration saved" -ForegroundColor Green
}

function Verify-Configuration {
    Write-Host ""
    Write-Host "[*] Verifying configuration..." -ForegroundColor Cyan
    Write-Host ""
    
    if (Test-Path $RegPath) {
        Write-Host "[+] Registry key found" -ForegroundColor Green
        
        $email = (Get-ItemProperty -Path $RegPath -Name "AdminEmail" -ErrorAction SilentlyContinue).AdminEmail
        if ($email -eq $AdminEmail) {
            Write-Host "[+] AdminEmail correctly configured: $email" -ForegroundColor Green
        } else {
            Write-Host "[!] AdminEmail mismatch. Found: $email" -ForegroundColor Yellow
        }
        
        $behavior = (Get-ItemProperty -Path $RegPath -Name "LoadBehavior" -ErrorAction SilentlyContinue).LoadBehavior
        Write-Host "[+] LoadBehavior: $behavior" -ForegroundColor Green
        
    } else {
        Write-Host "[!] Registry key not found" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Main execution
Show-Header
Test-AdminRights

Configure-Email

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Configuration Complete!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[+] Admin Email: $AdminEmail" -ForegroundColor Green
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Restart Microsoft Outlook" -ForegroundColor White
Write-Host "  2. You will see 'Report phishing' button" -ForegroundColor White
Write-Host "  3. Click to report suspicious emails" -ForegroundColor White
Write-Host ""
Write-Host "All phishing reports will be sent to: $AdminEmail" -ForegroundColor Cyan
Write-Host ""

if ($Verify) {
    Verify-Configuration
}
