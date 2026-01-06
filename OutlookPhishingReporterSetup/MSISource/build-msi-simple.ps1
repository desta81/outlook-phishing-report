#!/usr/bin/env pwsh

# Simple MSI Build Script for Outlook Phishing Reporter

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$outputDir = Join-Path (Split-Path -Parent $sourceDir) "Distribution"
$productWxs = Join-Path $sourceDir "Product.wxs"

Write-Host "=========================================="
Write-Host "Building Outlook Phishing Reporter MSI"
Write-Host "=========================================="
Write-Host "Source: $sourceDir"
Write-Host "Output: $outputDir"
Write-Host ""

# Create output directory
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "[+] Created output directory"
}

# Check for WiX Toolset
$candle = Get-Command candle.exe -ErrorAction SilentlyContinue
if (!$candle) {
    Write-Host "[ERROR] WiX Toolset not found in PATH"
    exit 1
}
Write-Host "[+] WiX Toolset found"

# Step 1: Compile
Write-Host ""
Write-Host "[1/2] Compiling WiX source..."
Push-Location $sourceDir
candle.exe Product.wxs -o Product.wixobj
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Compilation failed!"
    Pop-Location
    exit 1
}
Write-Host "[+] Compilation successful"

# Step 2: Link
Write-Host "[2/2] Linking MSI..."
light.exe -out "$outputDir\OutlookPhishingReporter.msi" Product.wixobj
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Linking failed!"
    Pop-Location
    exit 1
}
Write-Host "[+] Linking successful"
Pop-Location

# Verify
if (Test-Path "$outputDir\OutlookPhishingReporter.msi") {
    $size = (Get-Item "$outputDir\OutlookPhishingReporter.msi").Length
    Write-Host ""
    Write-Host "=========================================="
    Write-Host "Build Complete!"
    Write-Host "=========================================="
    Write-Host "[+] MSI created: $outputDir\OutlookPhishingReporter.msi"
    Write-Host "[+] Size: $([math]::Round($size/1KB, 2)) KB"
    Write-Host ""
} else {
    Write-Host "[ERROR] MSI was not created!"
    exit 1
}
