param(
    [switch]$Force = $false,
    [switch]$Verify = $false
)

# IMMEDIATE FIX - Configuration Error: Invalid Admin Email
# Usage: .\FIX-ERROR-NOW.ps1 or .\FIX-ERROR-NOW.ps1 -Verify

$AdminEmail = "yosi.desta@outlook.co.il"
$RegPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"

function Show-Header {
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║     CONFIGURATION ERROR - IMMEDIATE FIX                          ║" -ForegroundColor Red
    Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
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

function Fix-Configuration {
    Write-Host "[STEP 1] Closing Outlook..." -ForegroundColor Cyan
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "[+] Outlook closed" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "[STEP 2] Removing old configuration..." -ForegroundColor Cyan
    if (Test-Path $RegPath) {
        Remove-Item -Path $RegPath -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Old configuration removed" -ForegroundColor Green
    } else {
        Write-Host "[*] No existing configuration found" -ForegroundColor Yellow
    }
    Write-Host ""
    
    Write-Host "[STEP 3] Creating new configuration..." -ForegroundColor Cyan
    New-Item -Path $RegPath -Force -ErrorAction Stop | Out-Null
    New-ItemProperty -Path $RegPath -Name "AdminEmail" -Value $AdminEmail -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "Description" -Value "Outlook Phishing Reporter" -PropertyType String -Force | Out-Null
    Write-Host "[+] New configuration created" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "[STEP 4] Verifying configuration..." -ForegroundColor Cyan
    $savedEmail = (Get-ItemProperty -Path $RegPath -Name "AdminEmail" -ErrorAction SilentlyContinue).AdminEmail
    
    if ($savedEmail -eq $AdminEmail) {
        Write-Host "[+] Email verified: $savedEmail" -ForegroundColor Green
    } else {
        Write-Host "[!] WARNING: Email mismatch!" -ForegroundColor Yellow
        Write-Host "    Expected: $AdminEmail" -ForegroundColor Yellow
        Write-Host "    Got: $savedEmail" -ForegroundColor Yellow
    }
    Write-Host ""
}

function Verify-Only {
    Write-Host "[*] Verifying current configuration..." -ForegroundColor Cyan
    Write-Host ""
    
    if (Test-Path $RegPath) {
        Write-Host "[+] Registry key found" -ForegroundColor Green
        
        $email = (Get-ItemProperty -Path $RegPath -Name "AdminEmail" -ErrorAction SilentlyContinue).AdminEmail
        $behavior = (Get-ItemProperty -Path $RegPath -Name "LoadBehavior" -ErrorAction SilentlyContinue).LoadBehavior
        $description = (Get-ItemProperty -Path $RegPath -Name "Description" -ErrorAction SilentlyContinue).Description
        
        Write-Host "  AdminEmail:   $email" -ForegroundColor White
        Write-Host "  LoadBehavior: $behavior" -ForegroundColor White
        Write-Host "  Description:  $description" -ForegroundColor White
        
        if ($email -eq $AdminEmail) {
            Write-Host "[+] Configuration is correct!" -ForegroundColor Green
        } else {
            Write-Host "[!] Configuration needs fixing!" -ForegroundColor Red
        }
    } else {
        Write-Host "[!] Registry key not found!" -ForegroundColor Red
        Write-Host "    Need to run fix first" -ForegroundColor Red
    }
    Write-Host ""
}

# Main execution
Show-Header
Test-AdminRights

if ($Verify) {
    Verify-Only
} else {
    Fix-Configuration
}

Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                      FIX COMPLETE!                              ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "[+] Configuration error has been fixed" -ForegroundColor Green
Write-Host ""
Write-Host "IMPORTANT - NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Close this window" -ForegroundColor White
Write-Host "  2. Restart Microsoft Outlook completely" -ForegroundColor White
Write-Host "  3. The error should be gone" -ForegroundColor White
Write-Host "  4. Click 'Report phishing' button - it will work!" -ForegroundColor White
Write-Host ""
Write-Host "Email configured: $AdminEmail" -ForegroundColor Cyan
Write-Host ""
Write-Host "To verify configuration later, run:" -ForegroundColor Gray
Write-Host "  .\FIX-ERROR-NOW.ps1 -Verify" -ForegroundColor Gray
Write-Host ""
