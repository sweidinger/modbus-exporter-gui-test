#!/bin/bash

# Certificate verification script
# This script verifies and displays information about generated certificates

# Default certificate path
CERT_PATH="certificate"

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -c|--cert)
            CERT_PATH="$2"
            shift 2
            ;;
        -h|--help)
            echo "Usage: $0 [OPTIONS]"
            echo "Verify and display certificate information"
            echo ""
            echo "Options:"
            echo "  -c, --cert       Certificate path prefix (default: $CERT_PATH)"
            echo "  -h, --help       Show this help message"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            exit 1
            ;;
    esac
done

echo "🔍 Certificate Verification Tool"
echo "================================="
echo ""

# Check if certificate files exist
if [[ ! -f "${CERT_PATH}.crt" ]]; then
    echo "❌ Certificate file not found: ${CERT_PATH}.crt"
    echo "Please run ./scripts/create_certificate.sh first"
    exit 1
fi

echo "📋 Certificate Details:"
echo "----------------------"
openssl x509 -in "${CERT_PATH}.crt" -noout -text | grep -A 10 "Subject:"
echo ""

echo "📅 Validity Period:"
echo "-------------------"
openssl x509 -in "${CERT_PATH}.crt" -noout -text | grep -A 3 "Validity"
echo ""

echo "🔍 Certificate Fingerprints:"
echo "----------------------------"
echo "SHA256: $(openssl x509 -in "${CERT_PATH}.crt" -noout -fingerprint -sha256 | cut -d'=' -f2)"
echo "SHA1:   $(openssl x509 -in "${CERT_PATH}.crt" -noout -fingerprint -sha1 | cut -d'=' -f2)"
echo ""

echo "🔑 Key Information:"
echo "-------------------"
if [[ -f "${CERT_PATH}.key" ]]; then
    echo "Private Key: ✅ Present"
    echo "Key Size: $(openssl rsa -in "${CERT_PATH}.key" -noout -text | grep "Private-Key" | cut -d'(' -f2 | cut -d' ' -f1)"
else
    echo "Private Key: ❌ Missing"
fi
echo ""

echo "📦 Certificate Formats:"
echo "----------------------"
if [[ -f "${CERT_PATH}.crt" ]]; then
    echo "PEM Certificate (.crt): ✅ Present"
fi
if [[ -f "${CERT_PATH}.cer" ]]; then
    echo "Public Certificate (.cer): ✅ Present"
fi
if [[ -f "${CERT_PATH}.pfx" ]]; then
    echo "PKCS#12 Archive (.pfx): ✅ Present"
    echo "PFX Contents:"
    openssl pkcs12 -in "${CERT_PATH}.pfx" -noout -info -passin pass:ModbusExporter2024! 2>/dev/null || echo "  Unable to read PFX (password required)"
fi
echo ""

echo "⏰ Certificate Status:"
echo "----------------------"
# Check if certificate is valid (not expired)
if openssl x509 -in "${CERT_PATH}.crt" -noout -checkend 0 >/dev/null 2>&1; then
    echo "Status: ✅ Valid"
else
    echo "Status: ❌ Expired"
fi

# Check expiration
EXPIRY_DATE=$(openssl x509 -in "${CERT_PATH}.crt" -noout -enddate | cut -d'=' -f2)
echo "Expires: $EXPIRY_DATE"

# Days until expiry
DAYS_UNTIL_EXPIRY=$(openssl x509 -in "${CERT_PATH}.crt" -noout -checkend 0 | grep -o '[0-9]*' | head -1)
if [[ -n "$DAYS_UNTIL_EXPIRY" ]]; then
    echo "Days until expiry: $DAYS_UNTIL_EXPIRY"
fi
echo ""

echo "🔒 Security Information:"
echo "------------------------"
echo "Signature Algorithm: $(openssl x509 -in "${CERT_PATH}.crt" -noout -text | grep "Signature Algorithm" | head -1 | cut -d':' -f2 | xargs)"
echo "Key Usage: $(openssl x509 -in "${CERT_PATH}.crt" -noout -text | grep -A 1 "Key Usage" | tail -1 | xargs)"
echo "Extended Key Usage: $(openssl x509 -in "${CERT_PATH}.crt" -noout -text | grep -A 1 "Extended Key Usage" | tail -1 | xargs)"
echo ""

echo "🌐 Base64 Encoded Certificate (for GitHub Actions):"
echo "===================================================="
echo "Copy the following to your GitHub repository secrets as CODE_SIGNING_CERTIFICATE_BASE64:"
echo ""
base64 -i "${CERT_PATH}.pfx" | fold -w 64
echo ""
echo "🔑 Don't forget to also add CODE_SIGNING_PASSWORD: ModbusExporter2024!"
echo ""

echo "✅ Certificate verification complete!"
