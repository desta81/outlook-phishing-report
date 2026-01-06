param(
    [switch]$Force = $false
)

# Outlook Phishing Reporter - PowerShell Uninstaller
# Usage: .\Uninstall.ps1 or .\Uninstall.ps1 -Force

function Show-Header {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Outlook Phishing Reporter - Uninstaller" -ForegroundColor Cyan
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

function Confirm-Uninstall {
    if ($Force) {
        return $true
    }
    
    Write-Host "WARNING: This will uninstall the Outlook Phishing Reporter add-in." -ForegroundColor Yellow
    Write-Host ""
    
    $response = Read-Host "Are you sure you want to uninstall? (Y/N)"
    
    return $response -eq "Y" -or $response -eq "y"
}

function Close-Outlook {
    Write-Host "[*] Checking for running Outlook instances..." -ForegroundColor Cyan
    
    $outlookProcess = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    
    if ($outlookProcess) {
        Write-Host "[!] Outlook is running, closing it..." -ForegroundColor Yellow
        Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Write-Host "[+] Outlook closed" -ForegroundColor Green
    } else {
        Write-Host "[+] Outlook is not running" -ForegroundColor Green
    }
}

function Remove-RegistryEntry {
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    
    Write-Host "[*] Removing registry entries..." -ForegroundColor Cyan
    
    if (Test-Path $regPath) {
        try {
            Remove-Item -Path $regPath -Force -ErrorAction Stop
            Write-Host "[+] Registry entries removed successfully" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "[ERROR] Could not remove registry entries: $_" -ForegroundColor Red
            return $false
        }
    } else {
        Write-Host "[!] Registry entry not found (add-in may not be installed)" -ForegroundColor Yellow
        return $true
    }
}

# Main execution
Show-Header

Test-AdminRights

if (-not (Confirm-Uninstall)) {
    Write-Host "Uninstallation cancelled." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

Write-Host ""

Close-Outlook

$success = Remove-RegistryEntry

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan

if ($success) {
    Write-Host "Uninstallation Complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "[+] The Outlook Phishing Reporter add-in has been uninstalled." -ForegroundColor Green
    Write-Host "[+] You may restart Outlook at any time." -ForegroundColor Green
    Write-Host ""
    Write-Host "To reinstall, run: Install.vbs" -ForegroundColor Cyan
    Write-Host ""
    exit 0
} else {
    Write-Host "Uninstallation Failed!" -ForegroundColor Red
    Write-Host ""
    Write-Host "[ERROR] Could not complete uninstallation." -ForegroundColor Red
    Write-Host "[HELP] Try running as Administrator or manually deleting the registry key:" -ForegroundColor Yellow
    Write-Host "       HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
