param(
    [string]$EmailAddress = "yosi.desta@outlook.co.il"
)

Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║         SETUP EMAIL ADDRESS - YOSI.DESTA@OUTLOOK.CO.IL          ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""

# Check if running as Administrator
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must run as Administrator!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please:" -ForegroundColor Yellow
    Write-Host "  1. Right-click this script" -ForegroundColor White
    Write-Host "  2. Select: Run with PowerShell" -ForegroundColor White
    Write-Host "  3. If prompted, click: Run" -ForegroundColor White
    Write-Host ""
    pause
    exit 1
}

Write-Host "[*] Running as Administrator: ✓" -ForegroundColor Green
Write-Host ""

# Registry path
$RegistryPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
$ValueName = "AdminEmail"

Write-Host "[*] Setting up email address..." -ForegroundColor Cyan
Write-Host "    Email: $EmailAddress" -ForegroundColor White
Write-Host ""

try {
    # Create registry path if it doesn't exist
    if (-not (Test-Path $RegistryPath)) {
        Write-Host "[*] Creating registry path..." -ForegroundColor Yellow
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    # Set the email value
    Write-Host "[*] Writing to registry..." -ForegroundColor Cyan
    Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $EmailAddress -Type String

    Write-Host "[+] Email address set successfully!" -ForegroundColor Green
    Write-Host ""

    # Verify it was set
    Write-Host "[*] Verifying..." -ForegroundColor Cyan
    $VerifyValue = Get-ItemProperty -Path $RegistryPath -Name $ValueName -ErrorAction Stop
    
    if ($VerifyValue.$ValueName -eq $EmailAddress) {
        Write-Host "[+] Verified: Email is correctly set" -ForegroundColor Green
        Write-Host ""
        Write-Host "Registry Path: $RegistryPath" -ForegroundColor Gray
        Write-Host "Value Name:   $ValueName" -ForegroundColor Gray
        Write-Host "Value Data:   $($VerifyValue.$ValueName)" -ForegroundColor Gray
        Write-Host ""
    } else {
        Write-Host "[ERROR] Verification failed!" -ForegroundColor Red
        exit 1
    }

    Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "✅ EMAIL SETUP COMPLETE!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "  1. Close Outlook (if open)" -ForegroundColor White
    Write-Host "  2. Close Visual Studio debug session" -ForegroundColor White
    Write-Host "  3. Restart debug in Visual Studio (F5)" -ForegroundColor White
    Write-Host "  4. Select an email" -ForegroundColor White
    Write-Host "  5. Click 'Report phishing' button" -ForegroundColor White
    Write-Host "  6. Email will be sent to: $EmailAddress" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""

} catch {
    Write-Host "[ERROR] Failed to set email address!" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    pause
    exit 1
}

Write-Host "Press any key to close..." -ForegroundColor Yellow
[System.Console]::ReadKey() | Out-Null
