#!/usr/bin/env python3
"""
Modbus Data Exporter GUI Application
Exports Modbus device data to CSV or Excel format
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import time
import csv
import os
from datetime import datetime
import struct

# Try to import modbus library
try:
    from pymodbus.client import ModbusTcpClient as ModbusClient
    MODBUS_AVAILABLE = True
except ImportError:
    try:
        from pymodbus.client.sync import ModbusTcpClient as ModbusClient
        MODBUS_AVAILABLE = True
    except ImportError:
        MODBUS_AVAILABLE = False

# Try to import openpyxl for Excel export
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Original decode_ascii function
def decode_ascii(registers):
    return "".join(
        chr((reg >> 8) & 0xFF) + chr(reg & 0xFF) for reg in registers
    ).split("\x00")[0].strip()

def decode_float32(registers):
    """Decode Float32 value from two Modbus registers."""
    if registers and len(registers) == 2:
        combined = (registers[0] << 16) | registers[1]
        return struct.unpack('!f', struct.pack('!I', combined))[0]
    return None

# Original read_registers function
def read_registers(client, device_id, address, count, log_widget=None):
    try:
        # Try newer parameter style first
        try:
            result = client.read_holding_registers(address, count=count, slave=device_id)
        except TypeError:
            # Fall back to older parameter style
            result = client.read_holding_registers(address, count, unit=device_id)
        
        if result.isError():
            raise Exception(f"Modbus-Fehler: {result}")
        return result.registers
    except Exception as e:
        if log_widget:
            log_widget.log_message(f"⚠ Fehler beim Lesen der Register {address}: {e}")
        return None

def read_enhanced_diagnostics(client, device_id, device_type, log_widget=None):
    """Read enhanced diagnostics for TH110 and CL110 devices"""
    diagnostics = {}
    # Define additional registers and their processing
    enhanced_registers = [
        (31144, 1, "RF Communication Validity", "BITMAP"),
        (31145, 1, "Communication Status", "BITMAP"),
        (31151, 2, "Gateway PER", "Float32"),
        (31153, 2, "RSSI", "Float32"),
        (31155, 1, "LQI", "UINT16"),
        (31156, 2, "PER Max", "Float32"),
        (31158, 2, "RSSI Min", "Float32"),
        (31160, 1, "LQI Min", "UINT16")
    ]
    
    # Add CL110-specific Battery Voltage register
    if device_type == "CL110":
        enhanced_registers.append((3315, 2, "Battery Voltage", "Float32"))
    
    for addr, count, field_name, field_type in enhanced_registers:
        regs = read_registers(client, device_id, addr, count, log_widget)
        if regs:
            value = None
            if field_type == "Float32":
                value = round(decode_float32(regs), 2)
            elif field_type == "UINT16":
                value = regs[0] if regs else None
            elif field_type == "BITMAP":
                value = regs[0] if regs else None
            # Add more type handlers if needed
            
            diagnostics[field_name] = value
            if log_widget:
                log_widget.log_message(f"  ✓ {field_name}: {value}")
        else:
            diagnostics[field_name] = "N/A"
            if log_widget:
                log_widget.log_message(f"  ⚠ {field_name}: Error reading")
    
    return diagnostics

# Original get_device_ids function
def get_device_ids(client, log_widget=None):
    base = 504
    step = 5
    max_devices = 100
    device_ids = []

    if log_widget:
        log_widget.log_message("→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)")
    
    for i in range(max_devices):
        addr = base + (i * step)
        result = read_registers(client, 255, addr, 1, log_widget)
        if result and result[0] not in (0, 0xFFFF):
            device_id = result[0]
            if log_widget:
                log_widget.log_message(f"✓ Reg {addr}: DeviceID {device_id}")
            device_ids.append(device_id)
        else:
            if log_widget:
                log_widget.log_message(f"- Kein gültiger DeviceID-Wert in Register {addr}")
    return device_ids

# Original collect_data function
def collect_data(ip, log_widget=None):
    client = ModbusClient(ip)
    if not client.connect():
        if log_widget:
            log_widget.log_message("❌ Verbindung fehlgeschlagen.")
        return None

    if log_widget:
        log_widget.log_message("✓ Verbindung erfolgreich hergestellt.")
    
    device_ids = get_device_ids(client, log_widget)
    if not device_ids:
        if log_widget:
            log_widget.log_message("⚠ Keine gültigen DeviceIDs gefunden.")
        client.close()
        return None

    data = []
    for idx, device_id in enumerate(device_ids, start=1):
        if log_widget:
            log_widget.log_message(f"[{idx}/{len(device_ids)}] Verarbeite Device ID {device_id}")
        
        device_data = {
            "DeviceID": device_id,
            "DeviceType": "",
            "RFID": "",
            "SerialNumber": "",
        }

        # Commercial Reference → 31060
        ref_regs = read_registers(client, device_id, 31060, 16, log_widget)
        ref = decode_ascii(ref_regs) if ref_regs else ""
        if log_widget:
            log_widget.log_message(f"→ Device {device_id} hat Commercial Reference: {ref}")

        device_type = ""
        if ref == "EMS59443":
            device_type = "CL110"
        elif ref == "EMS59440":
            device_type = "TH110"
        else:
            device_type = "Unknown"
        device_data["DeviceType"] = device_type

        # RFID → 31026 (6 Register, hex)
        rfid_regs = read_registers(client, device_id, 31026, 6, log_widget)
        if rfid_regs:
            if log_widget:
                log_widget.log_message(f"  📦 RFID (Reg 31026, 6): {rfid_regs}")
            hex_str = "".join(f"{reg:04X}" for reg in rfid_regs if reg > 0)
            device_data["RFID"] = hex_str[:8]
        else:
            if log_widget:
                log_widget.log_message("  ⚠ RFID: Fehler beim Lesen")

        # Serial Number → 31088 (10 Register, ASCII)
        sn_regs = read_registers(client, device_id, 31088, 10, log_widget)
        if sn_regs:
            sn = decode_ascii(sn_regs)
            if log_widget:
                log_widget.log_message(f"  📦 SerialNumber (Reg 31088, 10): {sn_regs}")
                log_widget.log_message(f"  ✓ SerialNumber: {sn}")
            device_data["SerialNumber"] = sn
        else:
            if log_widget:
                log_widget.log_message("  ⚠ SerialNumber: Fehler beim Lesen")

        # Enhanced Diagnostics if enabled
        if hasattr(log_widget, 'enhanced_diagnostics_var') and log_widget.enhanced_diagnostics_var.get() and device_type in ["TH110", "CL110"]:
            enhanced_diagnostics = read_enhanced_diagnostics(client, device_id, device_type, log_widget)
            device_data["EnhancedDiagnostics"] = enhanced_diagnostics
            if log_widget:
                log_widget.log_message(f"→ Enhanced Diagnostics for {device_type}: {enhanced_diagnostics}")
        else:
            device_data["EnhancedDiagnostics"] = {}

        # Product Model (nur Debug) → 31106
        pm_regs = read_registers(client, device_id, 31106, 8, log_widget)
        if pm_regs:
            pm = decode_ascii(pm_regs)
            if log_widget:
                log_widget.log_message(f"  📦 ProductModel (Reg 31106, 8): {pm_regs}")
                log_widget.log_message(f"  ✓ ProductModel: {pm}")

        data.append(device_data)

    client.close()
    return data

class ModbusExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus Data Exporter")
        self.root.geometry("600x700")
        self.root.configure(bg='#333333')
        
        
        # Variables
        self.is_running = False
        self.export_thread = None
        self.csv_var = tk.BooleanVar(value=True)
        self.excel_var = tk.BooleanVar(value=EXCEL_AVAILABLE)
        self.enhanced_diagnostics_var = tk.BooleanVar(value=False)
        
        # Setup GUI
        self.setup_gui()
        
        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Bind ESC key to exit
        self.root.bind('<Escape>', lambda e: self.on_closing())
        
        # Make window focusable
        self.root.focus_force()

    def setup_gui(self):
        """Setup the main GUI elements"""
        
        # Title
        title_label = tk.Label(self.root, text="Modbus Data Exporter", 
                              font=("Arial", 16, "bold"), 
                              bg='#333333', fg='white')
        title_label.pack(pady=10)
        
        # IP Address Section
        ip_frame = tk.Frame(self.root, bg='#cccccc', relief='raised', bd=2)
        ip_frame.pack(fill='x', padx=20, pady=10)
        
        ip_label = tk.Label(ip_frame, text="Modbus Device IP Address:", 
                           font=("Arial", 12), bg='#cccccc', fg='#000000')
        ip_label.pack(pady=5)
        
        # IP Entry with clear visibility
        self.ip_entry = tk.Entry(ip_frame, width=20, font=("Arial", 12),
                                relief='sunken', bd=2, bg='#ffffff', fg='#000000',
                                highlightbackground='#cccccc',
                                highlightcolor='#0066cc', highlightthickness=1,
                                insertbackground='#000000')
        self.ip_entry.pack(pady=5)
        self.ip_entry.config(state='normal')  # Make sure the entry is enabled
        self.ip_entry.focus_set()  # Set focus to the IP entry
        self.ip_entry.insert(0, "10.0.1.110")  # Default value
        
        # Test IP Button Frame
        test_frame = tk.Frame(self.root, bg='#cccccc', relief='raised', bd=2)
        test_frame.pack(fill='x', padx=20, pady=10)
        
        # Test IP Button
        test_btn = tk.Button(test_frame, text="Test IP Connection", 
                            command=self.test_ip,
                            font=("Arial", 12),
                            bg='#cccccc',  # Match parent frame background
                            fg='black',
                            relief='flat',
                            bd=0,
                            highlightthickness=0,
                            activebackground='#bbbbbb',
                            activeforeground='black')
        test_btn.pack(fill='x', padx=10, pady=10)
        
        # Export Options Section
        export_frame = tk.Frame(self.root, bg='#cccccc', relief='raised', bd=2)
        export_frame.pack(fill='x', padx=20, pady=10)
        
        export_label = tk.Label(export_frame, text="Export Format:", 
                               font=("Arial", 12), bg='#cccccc', fg='#000000')
        export_label.pack(pady=5)
        
        # CSV Checkbox
        csv_cb = tk.Checkbutton(export_frame, text="Export to CSV",
                               variable=self.csv_var, font=("Arial", 10),
                               bg='#cccccc', fg='#000000', activeforeground='#000000',
                               activebackground='#aaaaaa')
        csv_cb.pack(pady=2)
        
        # Excel Checkbox
        excel_text = "Export to Excel" if EXCEL_AVAILABLE else "Export to Excel (openpyxl not installed)"
        excel_cb = tk.Checkbutton(export_frame, text=excel_text,
                                 variable=self.excel_var, font=("Arial", 10),
                                 bg='#cccccc', fg='#000000', activeforeground='#000000',
                                 activebackground='#aaaaaa', disabledforeground='#000000',
                                 state='normal' if EXCEL_AVAILABLE else 'disabled')
        excel_cb.pack(pady=2)
        
        # Enhanced Diagnostics Checkbox
        enhanced_diag_cb = tk.Checkbutton(export_frame, text="Enable Enhanced Diagnostics",
                                       variable=self.enhanced_diagnostics_var, font=("Arial", 10),
                                       bg='#cccccc', fg='#000000', activeforeground='#000000',
                                       activebackground='#aaaaaa')
        enhanced_diag_cb.pack(pady=5)

        # Control Buttons
        button_frame = tk.Frame(self.root, bg='#333333')
        button_frame.pack(pady=20)
        
        self.start_btn = tk.Button(button_frame, text="START EXPORT",
                                  command=self.start_export,
                                  bg='#2196F3', fg='black',
                                  font=("Arial", 14, "bold"),
                                  width=15, height=2)
        self.start_btn.pack(side='left', padx=10)
        
        self.stop_btn = tk.Button(button_frame, text="STOP EXPORT",
                                 command=self.stop_export,
                                 bg='#f44336', fg='black',
                                 font=("Arial", 14, "bold"),
                                 width=15, height=2,
                                 state='disabled')
        self.stop_btn.pack(side='left', padx=10)
        
        exit_btn = tk.Button(button_frame, text="EXIT",
                            command=self.on_closing,
                            bg='#757575', fg='black',
                            font=("Arial", 14, "bold"),
                            width=15, height=2)
        exit_btn.pack(side='left', padx=10)
        
        # Status Section
        status_frame = tk.Frame(self.root, bg='#cccccc', relief='raised', bd=2)
        status_frame.pack(fill='x', padx=20, pady=10)
        
        status_title = tk.Label(status_frame, text="Status:", 
                               font=("Arial", 12, "bold"), bg='#cccccc', fg='#000000')
        status_title.pack(pady=5)
        
        self.status_label = tk.Label(status_frame, text="Ready", 
                                    font=("Arial", 11), bg='#cccccc', fg='#000000')
        self.status_label.pack(pady=5)
        
        # Log Section
        log_frame = tk.Frame(self.root, bg='#cccccc', relief='raised', bd=2)
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        log_title = tk.Label(log_frame, text="Log Output:", 
                            font=("Arial", 12, "bold"), bg='#cccccc', fg='#000000')
        log_title.pack(pady=5)
        
        # Log text with scrollbar
        log_container = tk.Frame(log_frame, bg='#cccccc')
        log_container.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_container, height=8, width=70,
                               bg='#000000', fg='#00ff00', font=("Courier", 10),
                               relief='sunken', bd=2)
        
        log_scrollbar = tk.Scrollbar(log_container, orient='vertical',
                                    command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scrollbar.pack(side='right', fill='y')
        
        # Initial log message
        self.log_message("Application started. Ready to export Modbus data.")
        if not MODBUS_AVAILABLE:
            self.log_message("WARNING: pymodbus not installed. Using simulation mode.")

    def log_message(self, message):
        """Add a timestamped message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Print to console
        print(log_entry.strip())
        
        # Add to GUI log
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_status(self, message, color='#4CAF50'):
        """Update the status label"""
        self.status_label.config(text=message, fg=color)
        self.root.update_idletasks()

    def test_ip(self):
        """Test the IP address connectivity"""
        ip = self.ip_entry.get().strip()
        if not ip:
            messagebox.showerror("Error", "Please enter an IP address")
            return
        
        self.log_message(f"Testing IP address: {ip}")
        self.update_status("Testing connection...", '#FF9800')
        
        # Simulate connection test
        threading.Thread(target=self._test_ip_thread, args=(ip,), daemon=True).start()

    def _test_ip_thread(self, ip):
        """Test IP connection in a separate thread"""
        try:
            if MODBUS_AVAILABLE:
                client = ModbusClient(ip, port=502)
                if client.connect():
                    self.log_message(f"✓ Successfully connected to {ip}")
                    self.update_status("Connection successful", '#4CAF50')
                    client.close()
                else:
                    self.log_message(f"✗ Failed to connect to {ip}")
                    self.update_status("Connection failed", '#f44336')
            else:
                # Simulate test
                time.sleep(2)
                self.log_message(f"✓ IP test completed for {ip} (simulation mode)")
                self.update_status("Test completed (simulation)", '#4CAF50')
        except Exception as e:
            self.log_message(f"✗ Error testing IP {ip}: {str(e)}")
            self.update_status("Connection error", '#f44336')

    def start_export(self):
        """Start the data export process"""
        if self.is_running:
            return
        
        ip = self.ip_entry.get().strip()
        if not ip:
            messagebox.showerror("Error", "Please enter an IP address")
            return
        
        if not self.csv_var.get() and not self.excel_var.get():
            messagebox.showerror("Error", "Please select at least one export format")
            return
        
        self.is_running = True
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        
        self.log_message(f"Starting export from {ip}")
        self.update_status("Exporting data...", '#FF9800')
        
        # Start export in separate thread
        self.export_thread = threading.Thread(target=self._export_data, args=(ip,), daemon=True)
        self.export_thread.start()

    def stop_export(self):
        """Stop the data export process"""
        if not self.is_running:
            return
        
        self.is_running = False
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        
        self.log_message("Export stopped by user")
        self.update_status("Export stopped", '#FF9800')

    def _export_data(self, ip):
        """Export data using the original collect_data function"""
        try:
            if MODBUS_AVAILABLE:
                # Use original collect_data function
                data = collect_data(ip, self)
                if data:
                    self.log_message(f"Collected {len(data)} device records. Saving files...")
                    self._save_original_data(data)
                    self.log_message("Export completed successfully!")
                    self.update_status("Export completed", '#4CAF50')
                else:
                    self.log_message("No data collected or connection failed")
                    self.update_status("Export failed", '#f44336')
            else:
                # Simulation mode
                self.log_message("Running in simulation mode...")
                data = []
                for i in range(3):  # Simulate 3 devices
                    if not self.is_running:
                        break
                    
                    device_data = {
                        "DeviceID": 100 + i,
                        "DeviceType": "CL110" if i % 2 == 0 else "TH110",
                        "RFID": f"ABCD{i:04d}",
                        "SerialNumber": f"SN{i:06d}",
                    }
                    data.append(device_data)
                    self.log_message(f"Simulated device {i+1}: {device_data}")
                    time.sleep(1)
                
                if data and self.is_running:
                    self.log_message(f"Collected {len(data)} simulated device records. Saving files...")
                    self._save_original_data(data)
                    self.log_message("Export completed successfully!")
                    self.update_status("Export completed", '#4CAF50')
        
        except Exception as e:
            self.log_message(f"Export error: {str(e)}")
            self.update_status("Export error", '#f44336')
        finally:
            if self.is_running:
                self.is_running = False
                self.start_btn.config(state='normal')
                self.stop_btn.config(state='disabled')

    def flatten_diagnostics(self, data):
        """Flatten the enhanced diagnostics into individual fields for export"""
        headers = set()
        flattened_data = []
        
        for device in data:
            flat_device = {
                "DeviceID": device.get("DeviceID", ""),
                "DeviceType": device.get("DeviceType", ""),
                "RFID": device.get("RFID", ""),
                "SerialNumber": device.get("SerialNumber", "")
            }
            diagnostics = device.get("EnhancedDiagnostics", {})
            for key, value in diagnostics.items():
                flat_device[key] = value
                headers.add(key)
            flattened_data.append(flat_device)
        
        return list(sorted(headers)), flattened_data

    def _save_original_data(self, data):
        """Save data in the original format"""
        # Ask user for file location
        base_file = filedialog.asksaveasfilename(
            title="Save Export As",
            defaultextension="",
            filetypes=[("All Files", "*.*")]
        )
        
        if not base_file:
            self.log_message("Export cancelled by user")
            return
        
        # Save as CSV
        if self.csv_var.get():
            filename = base_file + ".csv"
            header_extras, flattened_data = self.flatten_diagnostics(data)
            fieldnames = ["DeviceID", "DeviceType", "RFID", "SerialNumber"] + header_extras
            with open(filename, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for row in flattened_data:
                    writer.writerow({k: row.get(k, "") for k in fieldnames})
            self.log_message(f"✓ CSV-Datei gespeichert: {filename}")
        
        # Save as Excel
        if self.excel_var.get() and EXCEL_AVAILABLE:
            filename = base_file + ".xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Modbus Export"
            header_extras, flattened_data = self.flatten_diagnostics(data)
            headers = ["DeviceID", "DeviceType", "RFID", "SerialNumber"] + header_extras
            ws.append(headers)
            for row in flattened_data:
                ws.append([row.get(h, "") for h in headers])
            wb.save(filename)
            self.log_message(f"✓ Excel-Datei gespeichert: {filename}")

    def on_closing(self):
        """Handle application closing"""
        if self.is_running:
            if messagebox.askokcancel("Quit", "Export is running. Do you want to stop and quit?"):
                self.stop_export()
                self.root.after(1000, self.root.destroy)  # Give time for cleanup
        else:
            self.root.destroy()

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = ModbusExporterGUI(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()

if __name__ == "__main__":
    main()

