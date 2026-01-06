param(
    [switch]$Repair = $false
)

# Outlook Phishing Reporter - Diagnostic Tool (PowerShell)
# Usage: .\Diagnose-and-Repair.ps1 or .\Diagnose-and-Repair.ps1 -Repair

function Show-Header {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Outlook Phishing Reporter - Diagnostic" -ForegroundColor Cyan
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

function Get-RegistryValue {
    param($Path, $Name)
    try {
        $key = Get-Item -Path "Registry::$Path" -ErrorAction SilentlyContinue
        if ($key) {
            return $key.GetValue($Name)
        }
    } catch {
        return $null
    }
    return $null
}

function Test-RegistryKey {
    param($Path)
    try {
        $key = Get-Item -Path "Registry::$Path" -ErrorAction SilentlyContinue
        return $null -ne $key
    } catch {
        return $false
    }
}

function Show-Diagnostics {
    Write-Host "[*] Running diagnostics..." -ForegroundColor Cyan
    Write-Host ""
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    
    # Step 1: Check registry key
    Write-Host "[STEP 1] Checking registry key..." -ForegroundColor White
    if (Test-RegistryKey $regPath) {
        Write-Host "[+] Registry key found" -ForegroundColor Green
    } else {
        Write-Host "[!] Registry key NOT FOUND" -ForegroundColor Yellow
        return $false
    }
    
    # Step 2: Check AdminEmail
    Write-Host ""
    Write-Host "[STEP 2] Checking AdminEmail configuration..." -ForegroundColor White
    $adminEmail = Get-RegistryValue $regPath "AdminEmail"
    if ($adminEmail) {
        Write-Host "[+] AdminEmail found: $adminEmail" -ForegroundColor Green
    } else {
        Write-Host "[!] AdminEmail NOT SET" -ForegroundColor Yellow
        return $false
    }
    
    # Step 3: Check LoadBehavior
    Write-Host ""
    Write-Host "[STEP 3] Checking LoadBehavior..." -ForegroundColor White
    $loadBehavior = Get-RegistryValue $regPath "LoadBehavior"
    if ($loadBehavior) {
        Write-Host "[+] LoadBehavior: $loadBehavior" -ForegroundColor Green
    } else {
        Write-Host "[!] LoadBehavior NOT SET" -ForegroundColor Yellow
    }
    
    # Step 4: Check Description
    Write-Host ""
    Write-Host "[STEP 4] Checking Description..." -ForegroundColor White
    $description = Get-RegistryValue $regPath "Description"
    if ($description) {
        Write-Host "[+] Description: $description" -ForegroundColor Green
    } else {
        Write-Host "[!] Description NOT SET" -ForegroundColor Yellow
    }
    
    # Step 5: Check Manifest
    Write-Host ""
    Write-Host "[STEP 5] Checking Manifest path..." -ForegroundColor White
    $manifest = Get-RegistryValue $regPath "Manifest"
    if ($manifest) {
        Write-Host "[+] Manifest found" -ForegroundColor Green
        if (Test-Path $manifest) {
            Write-Host "[+] Manifest file exists" -ForegroundColor Green
        } else {
            Write-Host "[!] Manifest file not found: $manifest" -ForegroundColor Yellow
            return $false
        }
    } else {
        Write-Host "[!] Manifest NOT SET" -ForegroundColor Yellow
    }
    
    return $true
}

function Invoke-Repair {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Repair Wizard" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Close Outlook
    Write-Host "[*] Closing Outlook..." -ForegroundColor Cyan
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
    
    # Remove old registry entry
    Write-Host "[*] Removing old configuration..." -ForegroundColor Cyan
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Force -ErrorAction SilentlyContinue
    }
    
    # Prompt for email
    Write-Host ""
    Write-Host "[REQUIRED] Configure Admin Email" -ForegroundColor Yellow
    Write-Host ""
    
    do {
        $adminEmail = Read-Host "Enter admin email address (e.g., security@company.com)"
        
        if ([string]::IsNullOrWhiteSpace($adminEmail)) {
            Write-Host "ERROR: Email is required" -ForegroundColor Red
            continue
        }
        
        # Validate email format
        if ($adminEmail -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
            break
        } else {
            Write-Host "ERROR: Invalid email format. Use: user@domain.com" -ForegroundColor Red
        }
    } while ($true)
    
    # Create registry entries
    Write-Host ""
    Write-Host "[*] Creating registry configuration..." -ForegroundColor Cyan
    
    New-Item -Path $regPath -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "AdminEmail" -Value $adminEmail -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force | Out-Null
    
    Write-Host "[+] Configuration saved" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Repair Complete!" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[+] Admin Email: $adminEmail" -ForegroundColor Green
    Write-Host ""
    Write-Host "IMPORTANT: Please restart Microsoft Outlook now" -ForegroundColor Yellow
    Write-Host ""
}

# Main execution
Show-Header
Test-AdminRights

Write-Host "Starting diagnostics..." -ForegroundColor Cyan
Write-Host ""

$diagnosticsPass = Show-Diagnostics

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan

if ($diagnosticsPass) {
    Write-Host "Diagnostics Complete - All Configuration Valid!" -ForegroundColor Green
    Write-Host ""
    Write-Host "All required registry values are configured correctly." -ForegroundColor Green
} else {
    Write-Host "Diagnostics Failed - Configuration Issues Detected!" -ForegroundColor Yellow
    Write-Host ""
    
    if ($Repair) {
        Invoke-Repair
    } else {
        Write-Host "Issues found. Run with -Repair flag to fix automatically:" -ForegroundColor Yellow
        Write-Host "  .\Diagnose-and-Repair.ps1 -Repair" -ForegroundColor Cyan
    }
}

Write-Host ""
