# PowerShell script to sign executable files
# This script signs Windows executables with a code signing certificate

param(
    [Parameter(Mandatory=$true)]
    [string]$ExecutablePath,
    [string]$CertificatePath = "certificate.pfx",
    [string]$CertificatePassword = "ModbusExporter2024!",
    [string]$TimestampServer = "http://timestamp.sectigo.com",
    [string]$Description = "Modbus Data Exporter - Industrial IoT Data Collection Tool"
)

Write-Host "Starting code signing process..." -ForegroundColor Green

# Check if the executable exists
if (-not (Test-Path $ExecutablePath)) {
    Write-Error "Executable file not found: $ExecutablePath"
    exit 1
}

# Check if the certificate exists
if (-not (Test-Path $CertificatePath)) {
    Write-Error "Certificate file not found: $CertificatePath"
    exit 1
}

try {
    # Convert password to secure string
    $securePassword = ConvertTo-SecureString -String $CertificatePassword -Force -AsPlainText
    
    # Sign the executable
    Write-Host "Signing executable: $ExecutablePath" -ForegroundColor Yellow
    
    $result = Set-AuthenticodeSignature -FilePath $ExecutablePath -Certificate (Get-PfxCertificate -FilePath $CertificatePath -Password $securePassword) -TimestampServer $TimestampServer -HashAlgorithm SHA256
    
    if ($result.Status -eq "Valid") {
        Write-Host "✓ Successfully signed executable!" -ForegroundColor Green
        Write-Host "  Status: $($result.Status)" -ForegroundColor Green
        Write-Host "  Signature Type: $($result.SignatureType)" -ForegroundColor Green
        Write-Host "  Certificate: $($result.SignerCertificate.Subject)" -ForegroundColor Green
        
        # Verify the signature
        $verification = Get-AuthenticodeSignature -FilePath $ExecutablePath
        Write-Host "  Verification Status: $($verification.Status)" -ForegroundColor Green
        
        # Display file information
        $fileInfo = Get-Item $ExecutablePath
        Write-Host "  File Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB" -ForegroundColor Cyan
        Write-Host "  Last Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Cyan
        
        return $true
    } else {
        Write-Error "Failed to sign executable. Status: $($result.Status)"
        Write-Error "Status Message: $($result.StatusMessage)"
        return $false
    }
} catch {
    Write-Error "Error during signing process: $($_.Exception.Message)"
    return $false
}
