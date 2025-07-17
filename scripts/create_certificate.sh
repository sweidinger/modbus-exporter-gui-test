#!/bin/bash

# Cross-platform certificate generation script using OpenSSL
# This script creates a self-signed code signing certificate on macOS/Linux

# Default values
CERT_NAME="Modbus Data Exporter"
PUBLISHER="Stefan Weidinger"
OUTPUT_PATH="certificate"
PASSWORD="ModbusExporter2024!"
VALIDITY_DAYS=1825  # 5 years

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -n|--name)
            CERT_NAME="$2"
            shift 2
            ;;
        -p|--publisher)
            PUBLISHER="$2"
            shift 2
            ;;
        -o|--output)
            OUTPUT_PATH="$2"
            shift 2
            ;;
        -w|--password)
            PASSWORD="$2"
            shift 2
            ;;
        -d|--days)
            VALIDITY_DAYS="$2"
            shift 2
            ;;
        -h|--help)
            echo "Usage: $0 [OPTIONS]"
            echo "Generate a self-signed code signing certificate"
            echo ""
            echo "Options:"
            echo "  -n, --name       Certificate name (default: $CERT_NAME)"
            echo "  -p, --publisher  Publisher name (default: $PUBLISHER)"
            echo "  -o, --output     Output file prefix (default: $OUTPUT_PATH)"
            echo "  -w, --password   Certificate password (default: $PASSWORD)"
            echo "  -d, --days       Validity in days (default: $VALIDITY_DAYS)"
            echo "  -h, --help       Show this help message"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            exit 1
            ;;
    esac
done

echo "🔐 Creating self-signed code signing certificate..."
echo "📋 Certificate Name: $CERT_NAME"
echo "🏢 Publisher: $PUBLISHER"
echo "📅 Validity: $VALIDITY_DAYS days"
echo ""

# Check if OpenSSL is available
if ! command -v openssl &> /dev/null; then
    echo "❌ OpenSSL is not installed. Please install it first:"
    echo "   macOS: brew install openssl"
    echo "   Linux: sudo apt-get install openssl (Ubuntu/Debian)"
    echo "          sudo yum install openssl (RHEL/CentOS)"
    exit 1
fi

# Create a temporary config file for the certificate
CONFIG_FILE=$(mktemp)
cat > "$CONFIG_FILE" << EOF
[req]
distinguished_name = req_distinguished_name
x509_extensions = v3_req
prompt = no

[req_distinguished_name]
C = US
ST = State
L = City
O = $PUBLISHER
CN = $CERT_NAME

[v3_req]
keyUsage = digitalSignature
extendedKeyUsage = codeSigning
basicConstraints = CA:FALSE
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid:always,issuer:always
EOF

# Generate private key
echo "🔑 Generating private key..."
openssl genrsa -out "${OUTPUT_PATH}.key" 2048

# Generate certificate
echo "📜 Generating certificate..."
openssl req -new -x509 -key "${OUTPUT_PATH}.key" -out "${OUTPUT_PATH}.crt" -days $VALIDITY_DAYS -config "$CONFIG_FILE"

# Create PKCS#12 file (PFX format for Windows)
echo "📦 Creating PKCS#12 file..."
openssl pkcs12 -export -out "${OUTPUT_PATH}.pfx" -inkey "${OUTPUT_PATH}.key" -in "${OUTPUT_PATH}.crt" -password "pass:$PASSWORD"

# Create public certificate for distribution (same as .crt but with .cer extension for Windows)
echo "🔓 Creating public certificate..."
cp "${OUTPUT_PATH}.crt" "${OUTPUT_PATH}.cer"

# Display certificate information
echo ""
echo "📊 Certificate Information:"
openssl x509 -in "${OUTPUT_PATH}.crt" -noout -text | grep -A 5 "Subject:"
openssl x509 -in "${OUTPUT_PATH}.crt" -noout -text | grep -A 5 "Validity"
echo ""
echo "🔍 Certificate Fingerprint:"
openssl x509 -in "${OUTPUT_PATH}.crt" -noout -fingerprint -sha256

# Clean up
rm "$CONFIG_FILE"

echo ""
echo "✅ Certificate generation completed!"
echo ""
echo "📁 Generated files:"
echo "  • ${OUTPUT_PATH}.key - Private key (keep secure!)"
echo "  • ${OUTPUT_PATH}.crt - Certificate (PEM format)"
echo "  • ${OUTPUT_PATH}.pfx - PKCS#12 file for Windows code signing"
echo "  • ${OUTPUT_PATH}.cer - Public certificate for distribution"
echo ""
echo "🔒 Certificate password: $PASSWORD"
echo ""
echo "📋 To use with GitHub Actions:"
echo "1. Encode the PFX file to base64:"
echo "   base64 -i ${OUTPUT_PATH}.pfx | pbcopy"
echo "2. Add to GitHub Secrets:"
echo "   - CODE_SIGNING_CERTIFICATE_BASE64: [paste base64 output]"
echo "   - CODE_SIGNING_PASSWORD: $PASSWORD"
echo ""
echo "🏢 To install on Windows machines:"
echo "1. Copy ${OUTPUT_PATH}.cer to the target machine"
echo "2. Run: certlm.msc (as Administrator)"
echo "3. Import to 'Trusted Publishers' store"
echo "Or use the install_certificate.ps1 script"
echo ""
echo "⚠️  IMPORTANT NOTES:"
echo "• This is a self-signed certificate - for production, consider a commercial certificate"
echo "• Keep the .key and .pfx files secure and never commit them to version control"
echo "• The .cer file can be distributed publicly for certificate installation"
echo "• Certificate is valid for $(($VALIDITY_DAYS / 365)) years"

# Set appropriate permissions
chmod 600 "${OUTPUT_PATH}.key" "${OUTPUT_PATH}.pfx"
chmod 644 "${OUTPUT_PATH}.crt" "${OUTPUT_PATH}.cer"

echo ""
echo "🎉 Ready to sign Windows executables!"
