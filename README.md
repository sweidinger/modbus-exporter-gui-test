# Modbus Data Exporter

A GUI application for exporting Modbus device data to CSV and Excel formats.

## Features

- Connect to Modbus TCP devices
- Export device data to CSV and Excel formats
- Real-time connection testing
- Cross-platform support (Windows, macOS, Linux)

## Requirements

- Python 3.11 or higher
- See `requirements.txt` for Python dependencies

## Quick Start (Pre-built Executables)

1. **Download** the latest release from the [Releases](https://github.com/sweidinger/modbus-exporter-gui-test/releases) page
2. **Extract** the downloaded ZIP file
3. **For Windows**: If you get security warnings, install the code signing certificate first:
   ```powershell
   # Download install_certificate.ps1 from the repository
   powershell -ExecutionPolicy Bypass -File install_certificate.ps1
   ```
4. **Run** the executable:
   - **Windows**: `modbus_exporter_v[VERSION]_windows.exe`
   - **macOS**: `modbus_exporter_v[VERSION]_macos` (right-click → Open first time)

### Enterprise Security & Code Signing

For corporate environments with security software (SentinelOne, CrowdStrike, etc.):

- ✅ **Windows executables are digitally signed** with a code signing certificate
- 🔒 **Install the certificate** using the provided script to establish trust
- 🏢 **IT administrators** can deploy the certificate via Group Policy for domain-wide trust
- 📖 **See [Code Signing Guide](docs/CODE_SIGNING_GUIDE.md)** for detailed instructions

### Features

- 🔌 **Modbus TCP/RTU support** - Connect to industrial devices
- 📊 **Live diagnostics view** - Real-time monitoring with customizable columns
- 📈 **Enhanced diagnostics** - Signal quality analysis and RF communication
- 📄 **Sensor pairing sheets** - JSON-integrated Excel generation
- 💾 **Multiple export formats** - CSV and Excel with automatic naming
- 🎨 **Professional dark theme** - Modern, responsive GUI
- ⚡ **Asynchronous operations** - Non-blocking data collection
- 🔍 **Version system** - Integrated versioning with GitHub Actions

## Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

```bash
python main.py
```

### Building Windows Executable

#### Using GitHub Actions (Recommended)

1. Push your code to GitHub
2. The GitHub Actions workflow will automatically build executables for Windows, macOS, and Linux
3. Download the artifacts from the Actions tab

#### Building Locally

1. Install build dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the build script:
   ```bash
   python build_windows.py
   ```

3. Or use PyInstaller directly:
   ```bash
   pyinstaller --onefile --windowed --name="PAS Wireless Device Exporter" main.py
   ```

#### Using the Spec File

For advanced configuration, use the spec file:

```bash
pyinstaller modbus_exporter.spec
```

## Dependencies for Windows Build

The `requirements.txt` includes everything needed for Windows executable building:

- **pyinstaller**: Creates standalone executables
- **auto-py-to-exe**: GUI tool for PyInstaller (optional)
- **pywin32**: Windows-specific functionality
- **pymodbus**: Modbus protocol implementation
- **openpyxl**: Excel file support

## GitHub Actions Workflow

The `.github/workflows/build-windows.yml` file provides:

- Automatic building on push/PR
- Multi-platform support (Windows, macOS, Linux)
- Artifact upload
- Automatic release creation

## File Structure

```
├── main.py                 # Main application
├── requirements.txt        # Python dependencies
├── modbus_exporter.spec   # PyInstaller configuration
├── build_windows.py       # Local build script
├── .github/
│   └── workflows/
│       └── build-windows.yml  # GitHub Actions workflow
└── README.md              # This file
```

## Troubleshooting

### Common Issues

1. **Missing tkinter**: On some Linux systems, install `python3-tk`
2. **PyInstaller issues**: Try using the spec file for better control
3. **Modbus connection**: Ensure the device IP and port are correct

### Build Issues

1. **Windows**: Ensure Visual C++ Build Tools are installed
2. **Icon missing**: Remove `--icon=icon.ico` from build commands if no icon file
3. **Large executable**: Use `--exclude-module` to remove unused modules

## License

MIT License
