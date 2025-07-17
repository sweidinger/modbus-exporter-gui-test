# Windows SmartScreen Bypass Guide

This guide helps users run the Modbus Data Exporter when Windows SmartScreen shows security warnings.

## What is Windows SmartScreen?

Windows SmartScreen is a security feature that helps protect against malicious software by:
- Checking file reputation
- Verifying publisher certificates
- Analyzing download patterns

## Why Does SmartScreen Block the Modbus Data Exporter?

SmartScreen shows warnings for the Modbus Data Exporter because:
1. **Self-signed certificate**: The app uses a self-signed certificate, not one from a well-known Certificate Authority
2. **New application**: The app doesn't have established reputation in Microsoft's database
3. **Low download frequency**: The app hasn't been downloaded enough times to build reputation

## How to Bypass SmartScreen Warnings

### Step 1: Download and Run the Application

1. Download the latest release from GitHub
2. Extract the ZIP file
3. Double-click the executable

### Step 2: Handle SmartScreen Warning

When SmartScreen appears:

1. **Click "More info"** (don't click "Don't run")
   ![SmartScreen More Info](https://via.placeholder.com/400x200/0078d4/ffffff?text=Click+More+Info)

2. **Click "Run anyway"**
   ![SmartScreen Run Anyway](https://via.placeholder.com/400x200/0078d4/ffffff?text=Click+Run+Anyway)

3. The application will start normally

### Step 3: Future Runs

After the first bypass:
- Windows may remember your choice
- The warning might not appear again
- If it does, repeat the bypass steps

## Alternative Solutions

### Option 1: Install the Certificate (Recommended)

Installing the code signing certificate will reduce security warnings:

1. **Download the certificate installation script**:
   - `install_certificate.ps1` from the repository
   - `modbus_exporter_certificate.cer` file

2. **Run the installation script**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File install_certificate.ps1
   ```

3. **Follow the prompts** - works without admin rights

### Option 2: Add to Windows Defender Exclusions

If you have admin rights:

1. **Open Windows Security**
2. **Go to Virus & threat protection**
3. **Click "Manage settings" under Virus & threat protection settings**
4. **Add an exclusion** for the application folder

### Option 3: Contact IT Department

For company computers:
1. **Request certificate deployment** via Group Policy
2. **Ask for application whitelisting** in security software
3. **Use the IT department email template** provided in the repository

## Technical Details

### Certificate Information
- **Publisher**: Stefan Weidinger
- **Certificate Type**: Self-signed code signing certificate
- **Validity**: 5 years
- **Purpose**: Authenticate the Modbus Data Exporter application

### Security Assurance
- The application is **digitally signed** with a verifiable certificate
- The signature ensures the **file hasn't been tampered with**
- The certificate can be **verified and traced** back to the developer
- Source code is **publicly available** on GitHub

## Frequently Asked Questions

### Q: Is it safe to bypass SmartScreen for this application?
**A**: Yes, the application is digitally signed and the source code is available on GitHub. The SmartScreen warning appears only because of the self-signed certificate, not because of any security risk.

### Q: Will this affect my computer's security?
**A**: No, bypassing SmartScreen for this specific application doesn't reduce your overall security. SmartScreen will continue to protect against other potentially harmful downloads.

### Q: Can I permanently trust this publisher?
**A**: Yes, by installing the certificate as described in Option 1 above, you can establish permanent trust for applications signed with this certificate.

### Q: What if my antivirus software also blocks it?
**A**: Some antivirus software may also show warnings. The same principle applies - the application is safe to run. You can add it to your antivirus exclusions or install the certificate.

### Q: Why not use a commercial certificate?
**A**: Commercial certificates cost $200-500 per year. For internal/development use, self-signed certificates provide the same security benefits at no cost.

## Need Help?

If you're still having issues:

1. **Check the main documentation**: [CODE_SIGNING_GUIDE.md](CODE_SIGNING_GUIDE.md)
2. **Contact your IT department** for company-wide deployment
3. **Report issues** on the GitHub repository
4. **Email the developer** with specific error messages

---

**Remember**: These warnings are designed to protect you, but in this case, the application is safe to run. The bypass steps above will allow you to use the Modbus Data Exporter while maintaining your system's security.
