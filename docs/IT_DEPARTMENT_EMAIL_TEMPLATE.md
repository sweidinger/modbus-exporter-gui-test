# IT Department Request Template

Copy and customize this email template when requesting certificate deployment from your IT department.

---

**Subject:** Request for Code Signing Certificate Deployment - Modbus Data Exporter

**To:** IT Support / IT Security Team

**Body:**

Dear IT Team,

I am requesting the deployment of a code signing certificate for a custom business application I have developed called "Modbus Data Exporter."

## Application Details
- **Name:** Modbus Data Exporter
- **Purpose:** Industrial IoT data collection and analysis tool for Modbus devices
- **Developer:** [Your Name]
- **Department:** [Your Department]
- **Business Justification:** This tool is used for collecting and analyzing data from industrial Modbus devices, improving operational efficiency and data visibility.

## Certificate Details
- **Certificate File:** `modbus_exporter_certificate.cer` (attached)
- **Certificate Type:** Code Signing Certificate (Self-Signed)
- **Publisher:** [Your Name/Company]
- **Validity Period:** 5 years
- **Intended Use:** Sign Windows executables to prevent security warnings

## Technical Information
- **Certificate Store:** Trusted Publishers
- **Scope:** Domain-wide deployment via Group Policy
- **Security Impact:** This certificate will allow the application to run without security warnings from Windows SmartScreen and enterprise security software like SentinelOne.

## Certificate Verification
You can verify the certificate details using:
```powershell
# View certificate information
Get-PfxCertificate -FilePath "modbus_exporter_certificate.cer"

# After installation, verify it's properly installed
Get-ChildItem -Path "Cert:\LocalMachine\TrustedPublisher" | Where-Object {$_.Subject -like "*Modbus Data Exporter*"}
```

## Deployment Request
Please deploy this certificate to:
- [ ] All domain-joined Windows machines
- [ ] Specific machines/OUs: [List specific requirements]
- [ ] My machine only: [Machine name/IP]

## Security Considerations
- This is a self-signed certificate created specifically for internal use
- The certificate only allows the signed application to run without warnings
- No other applications or processes will be affected
- The certificate can be revoked at any time if needed

## Documentation
Full documentation is available at:
https://github.com/sweidinger/modbus-exporter-gui-test/blob/main/docs/CODE_SIGNING_GUIDE.md

Please let me know if you need any additional information or have questions about this request.

Best regards,
[Your Name]
[Your Title]
[Your Contact Information]

---

**Attachments:**
- `modbus_exporter_certificate.cer`
- Application documentation (if available)

## Follow-up Actions After Deployment

Once the certificate is deployed, you can verify it's working by:

1. **Check certificate installation:**
   ```powershell
   Get-ChildItem -Path "Cert:\LocalMachine\TrustedPublisher" | Where-Object {$_.Subject -like "*Modbus Data Exporter*"}
   ```

2. **Test the application:**
   - Download and run the signed executable
   - Verify no security warnings appear
   - Confirm SentinelOne doesn't block the application

3. **Report success:**
   - Inform IT that the deployment was successful
   - Provide feedback on any remaining issues
