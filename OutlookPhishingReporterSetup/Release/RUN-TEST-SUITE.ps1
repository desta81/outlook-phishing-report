# ========================================================================
# Outlook Phishing Reporter - Complete Test Suite (PowerShell)
# ========================================================================
# This script tests all aspects of the plugin installation and functionality

param(
    [switch]$Verbose = $false,
    [switch]$Quick = $false
)

$Script:TestsPassed = 0
$Script:TestsFailed = 0
$Script:TestsTotal = 0

function Show-Header {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   OUTLOOK PHISHING REPORTER - COMPLETE TEST SUITE                 ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Test-AdminRights {
    Write-Host "[TEST 1/10] Checking Administrator Privileges..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "[PASS] Administrator privileges confirmed" -ForegroundColor Green
        $Script:TestsPassed++
        return $true
    } else {
        Write-Host "[FAIL] Administrator privileges required" -ForegroundColor Red
        Write-Host "       Please run as Administrator" -ForegroundColor Yellow
        $Script:TestsFailed++
        return $false
    }
}

function Test-DotNetFramework {
    Write-Host "[TEST 2/10] Checking .NET Framework 4.7.2..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    try {
        $key = Get-ItemProperty 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue
        if ($key) {
            $release = $key.Release
            # 4.7.2 = 461808
            if ($release -ge 461808) {
                Write-Host "[PASS] .NET Framework 4.7.2 or later detected" -ForegroundColor Green
                if ($Verbose) { Write-Host "       Release: $release" -ForegroundColor Gray }
                $Script:TestsPassed++
                return $true
            } else {
                Write-Host "[WARN] .NET Framework older than 4.7.2" -ForegroundColor Yellow
                Write-Host "       Recommend upgrading to .NET Framework 4.7.2+" -ForegroundColor Yellow
                return $false
            }
        } else {
            Write-Host "[FAIL] .NET Framework not found" -ForegroundColor Red
            $Script:TestsFailed++
            return $false
        }
    } catch {
        Write-Host "[FAIL] Error checking .NET Framework: $_" -ForegroundColor Red
        $Script:TestsFailed++
        return $false
    }
}

function Test-OutlookInstallation {
    Write-Host "[TEST 3/10] Checking Microsoft Outlook Installation..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    try {
        $key = Get-Item 'HKLM:\Software\Microsoft\Office' -ErrorAction SilentlyContinue
        if ($key) {
            Write-Host "[PASS] Microsoft Office detected" -ForegroundColor Green
            $Script:TestsPassed++
            return $true
        } else {
            Write-Host "[FAIL] Microsoft Office not found" -ForegroundColor Red
            $Script:TestsFailed++
            return $false
        }
    } catch {
        Write-Host "[FAIL] Error checking Outlook: $_" -ForegroundColor Red
        $Script:TestsFailed++
        return $false
    }
}

function Test-PluginRegistry {
    Write-Host "[TEST 4/10] Checking Plugin Registry Configuration..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    
    try {
        $key = Get-Item $regPath -ErrorAction SilentlyContinue
        if ($key) {
            Write-Host "[PASS] Plugin registry key found" -ForegroundColor Green
            $Script:TestsPassed++
            
            # Check AdminEmail
            Write-Host "  [SUBTEST] Checking AdminEmail value..." -ForegroundColor Gray
            $email = (Get-ItemProperty -Path $regPath -Name AdminEmail -ErrorAction SilentlyContinue).AdminEmail
            if ($email) {
                Write-Host "    [PASS] AdminEmail: $email" -ForegroundColor Green
                return $email
            } else {
                Write-Host "    [WARN] AdminEmail not configured" -ForegroundColor Yellow
                return $null
            }
        } else {
            Write-Host "[FAIL] Plugin registry key not found" -ForegroundColor Red
            Write-Host "       Plugin may not be installed" -ForegroundColor Yellow
            $Script:TestsFailed++
            return $null
        }
    } catch {
        Write-Host "[FAIL] Error checking registry: $_" -ForegroundColor Red
        $Script:TestsFailed++
        return $null
    }
}

function Test-LoadBehavior {
    Write-Host "[TEST 5/10] Checking LoadBehavior Setting..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    
    try {
        $behavior = (Get-ItemProperty -Path $regPath -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
        if ($behavior -eq 3) {
            Write-Host "[PASS] LoadBehavior = 3 (Auto-load enabled)" -ForegroundColor Green
            $Script:TestsPassed++
            return $true
        } elseif ($null -eq $behavior) {
            Write-Host "[FAIL] LoadBehavior not set" -ForegroundColor Red
            $Script:TestsFailed++
            return $false
        } else {
            Write-Host "[WARN] LoadBehavior = $behavior (expected 3)" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "[FAIL] Error checking LoadBehavior: $_" -ForegroundColor Red
        $Script:TestsFailed++
        return $false
    }
}

function Test-OutlookProcess {
    Write-Host "[TEST 6/10] Checking Outlook Process..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $outlookProcess = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Host "[INFO] Outlook is currently running" -ForegroundColor Cyan
        Write-Host "       Some tests cannot be performed while Outlook is running" -ForegroundColor Gray
        return $false
    } else {
        Write-Host "[PASS] Outlook is not running" -ForegroundColor Green
        Write-Host "       Good for full configuration testing" -ForegroundColor Gray
        $Script:TestsPassed++
        return $true
    }
}

function Test-PluginDLL {
    Write-Host "[TEST 7/10] Checking Plugin DLL..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $dllPath = "$env:APPDATA\Microsoft\Addins\OutlookPhishingReporter.dll"
    if (Test-Path $dllPath) {
        Write-Host "[PASS] Plugin DLL found" -ForegroundColor Green
        if ($Verbose) { Write-Host "       Path: $dllPath" -ForegroundColor Gray }
        $Script:TestsPassed++
        return $true
    } else {
        Write-Host "[INFO] DLL not in standard location" -ForegroundColor Cyan
        Write-Host "       Plugin may be installed elsewhere or not installed" -ForegroundColor Gray
        return $false
    }
}

function Test-LogDirectory {
    Write-Host "[TEST 8/10] Checking Log Directory..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $logDir = "$env:LOCALAPPDATA\OutlookPhishingReporter\Logs"
    if (Test-Path $logDir) {
        Write-Host "[PASS] Log directory exists" -ForegroundColor Green
        Write-Host "       Location: $logDir" -ForegroundColor Gray
        
        # Count log files
        $logCount = (Get-ChildItem "$logDir\*.log" -ErrorAction SilentlyContinue | Measure-Object).Count
        Write-Host "       Log files: $logCount" -ForegroundColor Gray
        
        $Script:TestsPassed++
        return $true
    } else {
        Write-Host "[INFO] Log directory not yet created" -ForegroundColor Cyan
        Write-Host "       Will be created on first use" -ForegroundColor Gray
        return $false
    }
}

function Test-EmailConfiguration {
    Write-Host "[TEST 9/10] Checking Email Configuration..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    $email = (Get-ItemProperty -Path $regPath -Name AdminEmail -ErrorAction SilentlyContinue).AdminEmail
    
    if ($email) {
        Write-Host "[PASS] Admin email configured: $email" -ForegroundColor Green
        $Script:TestsPassed++
        
        # Validate email format
        if ($email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
            Write-Host "[PASS] Email format is valid" -ForegroundColor Green
            return $true
        } else {
            Write-Host "[WARN] Email format may be invalid" -ForegroundColor Yellow
            return $false
        }
    } else {
        Write-Host "[FAIL] Admin email not configured" -ForegroundColor Red
        Write-Host "       Need to run: FIX-ERROR-NOW.bat" -ForegroundColor Yellow
        $Script:TestsFailed++
        return $false
    }
}

function Test-PluginManifest {
    Write-Host "[TEST 10/10] Checking Plugin Manifest..." -ForegroundColor Yellow
    $Script:TestsTotal++
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"
    $manifest = (Get-ItemProperty -Path $regPath -Name Manifest -ErrorAction SilentlyContinue).Manifest
    
    if ($manifest) {
        Write-Host "[PASS] Plugin manifest configured" -ForegroundColor Green
        Write-Host "       Path: $manifest" -ForegroundColor Gray
        
        if (Test-Path $manifest) {
            Write-Host "[PASS] Manifest file exists" -ForegroundColor Green
            $Script:TestsPassed++
            return $true
        } else {
            Write-Host "[WARN] Manifest file not found at configured path" -ForegroundColor Yellow
            return $false
        }
    } else {
        Write-Host "[INFO] Manifest not configured in registry" -ForegroundColor Cyan
        return $false
    }
}

function Show-Summary {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                          TEST SUMMARY                             ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Script:TestsPassed -gt 0) {
        Write-Host "[PASS] Tests Passed: $($Script:TestsPassed)" -ForegroundColor Green
    }
    if ($Script:TestsFailed -gt 0) {
        Write-Host "[FAIL] Tests Failed: $($Script:TestsFailed)" -ForegroundColor Red
    }
    Write-Host "[INFO] Tests Total:  $($Script:TestsTotal)" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Script:TestsFailed -eq 0) {
        Write-Host "✓ ALL TESTS PASSED - Plugin appears to be correctly installed" -ForegroundColor Green
        Write-Host ""
        Write-Host "Recommendations:" -ForegroundColor White
        Write-Host "  1. Open Outlook" -ForegroundColor Gray
        Write-Host "  2. Look for 'Report phishing' button in ribbon" -ForegroundColor Gray
        Write-Host "  3. Select a test email" -ForegroundColor Gray
        Write-Host "  4. Click the button to verify it works" -ForegroundColor Gray
        Write-Host "  5. Check that email is reported successfully" -ForegroundColor Gray
    } else {
        Write-Host "✗ SOME TESTS FAILED - Please address the issues above" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting Steps:" -ForegroundColor White
        Write-Host "  1. Run FIX-ERROR-NOW.bat to fix configuration" -ForegroundColor Gray
        Write-Host "  2. Verify .NET Framework 4.7.2 is installed" -ForegroundColor Gray
        Write-Host "  3. Verify Outlook is properly installed" -ForegroundColor Gray
        Write-Host "  4. Restart your computer" -ForegroundColor Gray
        Write-Host "  5. Try reinstalling the plugin" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    END OF TEST SUITE                              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# ========================================================================
# Main Execution
# ========================================================================

Show-Header

if (-not (Test-AdminRights)) {
    Write-Host "Error: Administrator rights required" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Run tests
if (-not $Quick) {
    Test-DotNetFramework | Out-Null
    Write-Host ""
    Test-OutlookInstallation | Out-Null
    Write-Host ""
}

$email = Test-PluginRegistry
Write-Host ""
Test-LoadBehavior | Out-Null
Write-Host ""
Test-OutlookProcess | Out-Null
Write-Host ""
Test-PluginDLL | Out-Null
Write-Host ""
Test-LogDirectory | Out-Null
Write-Host ""
Test-EmailConfiguration | Out-Null
Write-Host ""
Test-PluginManifest | Out-Null
Write-Host ""

# Show summary
Show-Summary
