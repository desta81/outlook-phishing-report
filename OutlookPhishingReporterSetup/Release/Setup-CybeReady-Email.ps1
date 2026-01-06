param(
    [switch]$Verify = $false
)

# Outlook Phishing Reporter - Setup for report@phishcage.cybeready.net
# Usage: .\Setup-CybeReady-Email.ps1 or .\Setup-CybeReady-Email.ps1 -Verify

$AdminEmail = "report@phishcage.cybeready.net"
$RegPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"

function Show-Header {
    Write-Host ""
    Write-Host "╔═════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  Outlook Phishing Reporter - Setup Configuration                   ║" -ForegroundColor Cyan
    Write-Host "║  Email: report@phishcage.cybeready.net                             ║" -ForegroundColor Cyan
    Write-Host "╚═════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
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
    Write-Host "[+] Outlook closed" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "[*] Configuring registry..." -ForegroundColor Cyan
    
    # Create registry key
    New-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
    
    # Set values
    New-ItemProperty -Path $RegPath -Name "AdminEmail" -Value $AdminEmail -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force | Out-Null
    
    Write-Host "[+] Configuration saved" -ForegroundColor Green
    Write-Host ""
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

Write-Host "╔═════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                    CONFIGURATION COMPLETE!                         ║" -ForegroundColor Green
Write-Host "╚═════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "[+] Email configured: $AdminEmail" -ForegroundColor Green
Write-Host ""
Write-Host "IMPORTANT - NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Close this window" -ForegroundColor White
Write-Host "  2. Restart Microsoft Outlook" -ForegroundColor White
Write-Host "  3. The configuration will be applied" -ForegroundColor White
Write-Host "  4. 'Report phishing' button will work" -ForegroundColor White
Write-Host ""
Write-Host "All reports will be sent to: $AdminEmail" -ForegroundColor Cyan
Write-Host ""
Write-Host "To verify configuration later, run:" -ForegroundColor Gray
Write-Host "  .\Setup-CybeReady-Email.ps1 -Verify" -ForegroundColor Gray
Write-Host ""

if ($Verify) {
    Verify-Configuration
}
