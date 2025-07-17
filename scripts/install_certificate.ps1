# Certificate Installation Script for Modbus Data Exporter
# This script installs the code signing certificate to make the application trusted

param(
    [string]$CertPath = "modbus_exporter_certificate.cer",
    [switch]$Force
)

Write-Host "=== Modbus Data Exporter Certificate Installer ===" -ForegroundColor Cyan
Write-Host "This script will install the code signing certificate to make the application trusted by your system." -ForegroundColor Yellow
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "⚠️  WARNING: Not running as Administrator" -ForegroundColor Red
    Write-Host "For system-wide trust, please run this script as Administrator." -ForegroundColor Yellow
    Write-Host "Continuing with user-level installation..." -ForegroundColor Yellow
    Write-Host ""
}

# Check if certificate file exists
if (-not (Test-Path $CertPath)) {
    Write-Host "❌ Certificate file not found: $CertPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please ensure the certificate file is in the same directory as this script." -ForegroundColor Yellow
    Write-Host "You can download it from the GitHub releases page." -ForegroundColor Yellow
    exit 1
}

# Display certificate information
try {
    Write-Host "📋 Certificate Information:" -ForegroundColor Green
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
    Write-Host "  Subject: $($cert.Subject)" -ForegroundColor White
    Write-Host "  Issuer: $($cert.Issuer)" -ForegroundColor White
    Write-Host "  Valid From: $($cert.NotBefore)" -ForegroundColor White
    Write-Host "  Valid Until: $($cert.NotAfter)" -ForegroundColor White
    Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
    Write-Host ""
    
    # Check if certificate is expired
    if ($cert.NotAfter -lt (Get-Date)) {
        Write-Host "⚠️  WARNING: Certificate has expired!" -ForegroundColor Red
        if (-not $Force) {
            $response = Read-Host "Continue anyway? (y/N)"
            if ($response -ne "y" -and $response -ne "Y") {
                Write-Host "Installation cancelled." -ForegroundColor Yellow
                exit 1
            }
        }
    }
} catch {
    Write-Host "❌ Error reading certificate: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Determine installation location
$storeLocation = if ($isAdmin) { "LocalMachine" } else { "CurrentUser" }
Write-Host "📁 Installing to: $storeLocation\TrustedPublisher" -ForegroundColor Cyan

# Install the certificate
try {
    Write-Host "🔄 Installing certificate..." -ForegroundColor Yellow
    
    Import-Certificate -FilePath $CertPath -CertStoreLocation "Cert:\$storeLocation\TrustedPublisher" -ErrorAction Stop | Out-Null
    
    Write-Host "✅ Certificate installed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "🎉 The Modbus Data Exporter is now trusted on this machine!" -ForegroundColor Green
    Write-Host "You can now run the application without security warnings." -ForegroundColor Green
    
} catch {
    Write-Host "❌ Failed to install certificate: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible solutions:" -ForegroundColor Yellow
    Write-Host "1. Run this script as Administrator" -ForegroundColor Yellow
    Write-Host "2. Check that the certificate file is not corrupted" -ForegroundColor Yellow
    Write-Host "3. Ensure PowerShell execution policy allows script execution" -ForegroundColor Yellow
    exit 1
}

# Verify installation
try {
    Write-Host "🔍 Verifying installation..." -ForegroundColor Yellow
    $installedCert = Get-ChildItem -Path "Cert:\$storeLocation\TrustedPublisher" | Where-Object {$_.Thumbprint -eq $cert.Thumbprint}
    
    if ($installedCert) {
        Write-Host "✅ Verification successful - Certificate is properly installed!" -ForegroundColor Green
    } else {
        Write-Host "⚠️  Verification failed - Certificate may not be properly installed" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️  Could not verify installation: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Installation Complete ===" -ForegroundColor Cyan
Write-Host "You can now run the Modbus Data Exporter without security warnings." -ForegroundColor Green
Write-Host ""
Write-Host "Need help? Check the documentation at:" -ForegroundColor Yellow
Write-Host "https://github.com/sweidinger/modbus-exporter-gui-test/blob/main/docs/CODE_SIGNING_GUIDE.md" -ForegroundColor Cyan
