# Code Signing Guide for Modbus Data Exporter

This guide explains how to set up code signing for the Modbus Data Exporter to make it trusted by enterprise security software like SentinelOne.

## Overview

Code signing provides digital authentication to verify the integrity and origin of software. This is essential for enterprise environments where security software may block unsigned executables.

## Quick Start

### For End Users (Installing Trust)

1. **Download the certificate** from the releases page (`modbus_exporter_certificate.cer`)
2. **Install the certificate** to the Trusted Publishers store:
   ```powershell
   # Run as Administrator
   certlm.msc
   ```
   - Navigate to: `Trusted Publishers > Certificates`
   - Right-click → `All Tasks` → `Import`
   - Select the `.cer` file and complete the import

3. **Verify the installation**:
   ```powershell
   Get-ChildItem -Path "Cert:\LocalMachine\TrustedPublisher" | Where-Object {$_.Subject -like "*Modbus Data Exporter*"}
   ```

### For Developers (Code Signing)

1. **Create a certificate** (one-time setup):
   ```powershell
   .\scripts\create_certificate.ps1
   ```

2. **Sign the executable**:
   ```powershell
   .\scripts\sign_executable.ps1 -ExecutablePath "dist\modbus_exporter.exe"
   ```

## Detailed Setup

### Self-Signed Certificate (Development/Internal Use)

For internal company use or development environments:

1. **Generate the certificate**:
   ```powershell
   .\scripts\create_certificate.ps1 -CertName "Your Company - Modbus Exporter" -Publisher "Your Company Name"
   ```

2. **Distribute the public certificate** (`.cer` file) to all target machines
3. **Install on each machine** using the steps above

### Commercial Certificate (Production Use)

For broader distribution or stricter security requirements:

1. **Purchase a code signing certificate** from a trusted CA like:
   - DigiCert
   - Sectigo (formerly Comodo)
   - GlobalSign
   - Entrust

2. **Replace the self-signed certificate** with the commercial one in the signing script

3. **Update the GitHub Actions workflow** with the commercial certificate (stored as secrets)

## GitHub Actions Integration

The build workflow automatically signs executables when certificates are available:

### Setup GitHub Secrets

1. **Add certificate secrets** to your GitHub repository:
   - `CODE_SIGNING_CERTIFICATE_BASE64`: Base64-encoded PFX certificate
   - `CODE_SIGNING_PASSWORD`: Certificate password

2. **Encode your certificate**:
   ```powershell
   $bytes = [System.IO.File]::ReadAllBytes("certificate.pfx")
   $base64 = [System.Convert]::ToBase64String($bytes)
   Write-Output $base64
   ```

### Workflow Configuration

The workflow will automatically:
- Decode the certificate from GitHub secrets
- Sign the Windows executable
- Verify the signature
- Include the public certificate in releases

## Enterprise Deployment

### Group Policy (Domain Environment)

For domain-joined machines, deploy the certificate via Group Policy:

1. **Copy the certificate** to a network share
2. **Create a GPO**:
   - Computer Configuration → Windows Settings → Security Settings → Public Key Policies → Trusted Publishers
   - Right-click → Import → Select the `.cer` file

3. **Apply the GPO** to target OUs

### Manual Installation Script

For non-domain machines, use this PowerShell script:

```powershell
# install_certificate.ps1
param([string]$CertPath = "modbus_exporter_certificate.cer")

if (-not (Test-Path $CertPath)) {
    Write-Error "Certificate file not found: $CertPath"
    exit 1
}

try {
    Import-Certificate -FilePath $CertPath -CertStoreLocation "Cert:\LocalMachine\TrustedPublisher"
    Write-Host "✓ Certificate installed successfully!" -ForegroundColor Green
    Write-Host "The Modbus Data Exporter is now trusted on this machine." -ForegroundColor Green
} catch {
    Write-Error "Failed to install certificate: $($_.Exception.Message)"
    exit 1
}
```

## Security Considerations

### Certificate Validation

Always verify the certificate before installation:

```powershell
# Check certificate details
$cert = Get-PfxCertificate -FilePath "certificate.pfx"
Write-Host "Subject: $($cert.Subject)"
Write-Host "Issuer: $($cert.Issuer)"
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Valid From: $($cert.NotBefore)"
Write-Host "Valid Until: $($cert.NotAfter)"
```

### Best Practices

1. **Use commercial certificates** for production environments
2. **Protect private keys** - never commit PFX files to version control
3. **Use timestamping** to ensure signatures remain valid after certificate expiry
4. **Monitor certificate expiry** and renew before expiration
5. **Test on target environments** before widespread deployment

## Troubleshooting

### Common Issues

1. **"Publisher could not be verified"**:
   - Certificate not installed in Trusted Publishers store
   - Certificate expired or invalid

2. **"This app has been blocked by your administrator"**:
   - SmartScreen or security software blocking
   - Install certificate as described above

3. **Windows SmartScreen warning appears**:
   - SmartScreen doesn't recognize self-signed certificates
   - Click "More info" → "Run anyway" to bypass
   - For permanent solution, use a commercial certificate

4. **SentinelOne or other security software blocks the application**:
   - Install certificate to Trusted Publishers store
   - Contact IT department for domain-wide deployment
   - Request application whitelisting as alternative

5. **Signature verification fails**:
   - Check if certificate is valid and not expired
   - Verify timestamping server is accessible

### Verification Commands

```powershell
# Check if executable is signed
Get-AuthenticodeSignature -FilePath "modbus_exporter.exe"

# List trusted publishers
Get-ChildItem -Path "Cert:\LocalMachine\TrustedPublisher"

# Check certificate details
Get-AuthenticodeSignature -FilePath "modbus_exporter.exe" | Select-Object -ExpandProperty SignerCertificate
```

## Support

For issues with code signing or certificate installation:

1. Check the troubleshooting section above
2. Verify your PowerShell execution policy allows script execution
3. Ensure you're running as Administrator when installing certificates
4. Contact your IT administrator for enterprise deployment assistance

---

**Note**: This guide assumes Windows environments. For cross-platform scenarios, consider additional signing methods for macOS (notarization) and Linux (GPG signing).
