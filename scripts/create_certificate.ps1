# PowerShell script to create a self-signed certificate for code signing
# This creates a certificate that can be used to sign executables for enterprise environments

param(
    [string]$CertName = "Modbus Data Exporter",
    [string]$Publisher = "Stefan Weidinger",
    [string]$OutputPath = "certificate.pfx",
    [string]$Password = "ModbusExporter2024!"
)

Write-Host "Creating self-signed certificate for code signing..." -ForegroundColor Green

# Create the certificate
$cert = New-SelfSignedCertificate -Type CodeSigning -Subject "CN=$CertName, O=$Publisher" -KeyAlgorithm RSA -KeyLength 2048 -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -KeyExportPolicy Exportable -KeyUsage DigitalSignature -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)

Write-Host "Certificate created with thumbprint: $($cert.Thumbprint)" -ForegroundColor Yellow

# Export the certificate to PFX format
$certPath = "cert:\CurrentUser\My\$($cert.Thumbprint)"
$pfxPassword = ConvertTo-SecureString -String $Password -Force -AsPlainText

Export-PfxCertificate -Cert $certPath -FilePath $OutputPath -Password $pfxPassword
Write-Host "Certificate exported to: $OutputPath" -ForegroundColor Green

# Export the public key for distribution
$cerPath = $OutputPath -replace '\.pfx$', '.cer'
Export-Certificate -Cert $certPath -FilePath $cerPath
Write-Host "Public certificate exported to: $cerPath" -ForegroundColor Green

Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Cyan
Write-Host "1. This is a self-signed certificate. For production use, consider purchasing a certificate from a trusted CA." -ForegroundColor Yellow
Write-Host "2. To install the certificate on target machines, run: certlm.msc and import the .cer file to 'Trusted Publishers' store." -ForegroundColor Yellow
Write-Host "3. The PFX password is: $Password" -ForegroundColor Yellow
Write-Host "4. Store the PFX file securely and never commit it to version control." -ForegroundColor Red

return @{
    Thumbprint = $cert.Thumbprint
    PfxPath = $OutputPath
    CerPath = $cerPath
    Password = $Password
}
