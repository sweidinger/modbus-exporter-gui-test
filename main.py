import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pymodbus.client import ModbusTcpClient
from pymodbus.exceptions import ModbusIOException

def read_ascii_registers(client, unit, start_address, count):
    try:
        result = client.read_holding_registers(start_address, count, unit=unit)
        if not result or isinstance(result, ModbusIOException) or not hasattr(result, "registers"):
            return ""
        chars = [chr((reg >> 8) & 0xFF) + chr(reg & 0xFF) for reg in result.registers]
        return ''.join(chars).strip('\x00')
    except Exception:
        return ""

def read_registers(client, unit, start_address, count):
    try:
        result = client.read_holding_registers(start_address, count, unit=unit)
        if not result or isinstance(result, ModbusIOException) or not hasattr(result, "registers"):
            return []
        return result.registers
    except Exception:
        return []

def convert_registers_to_hex(registers):
    # Nutzt nur die ersten 4 Register für RFID
    if len(registers) < 4:
        return "ERROR"
    try:
        hex_str = ''.join(f'{reg:04X}' for reg in registers[:4])
        return hex_str
    except Exception:
        return "ERROR"

def scan_devices(ip_address, log_widget):
    client = ModbusTcpClient(ip_address)
    if not client.connect():
        log_widget.insert(tk.END, f"✗ Verbindung zu {ip_address} fehlgeschlagen.\n")
        return []

    log_widget.insert(tk.END, f"✓ Verbindung erfolgreich hergestellt.\n")
    log_widget.insert(tk.END, "→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)\n")
    log_widget.update()

    device_ids = []
    for reg in range(504, 1000, 5):
        result = read_registers(client, 255, reg, 1)
        if result and result[0] > 0:
            log_widget.insert(tk.END, f"✓ Reg {reg}: DeviceID {result[0]}\n")
            device_ids.append(result[0])
        else:
            log_widget.insert(tk.END, f"- Kein gültiger DeviceID-Wert in Register {reg}\n")
        log_widget.update()

    results = []
    for idx, device_id in enumerate(device_ids):
        log_widget.insert(tk.END, f"[{idx+1}/{len(device_ids)}] Verarbeite Device ID {device_id}\n")

        reference = read_ascii_registers(client, device_id, 31060, 16)
        log_widget.insert(tk.END, f"→ Device {device_id} hat Commercial Reference: {reference}\n")

        if reference == "EMS59440":
            device_type = "TH110"
        elif reference == "EMS59443":
            device_type = "CL110"
        else:
            device_type = "HeatTag"  # Platzhalter – ggf. später Referenz ergänzen

        rfid_regs = read_registers(client, device_id, 31026, 6)
        rfid_hex = convert_registers_to_hex(rfid_regs)
        log_widget.insert(tk.END, f"  📦 RFID (Reg 31026, 6): {rfid_regs}\n")

        serial_regs = read_registers(client, device_id, 31088, 10)
        serial_ascii = read_ascii_registers(client, device_id, 31088, 10)
        log_widget.insert(tk.END, f"  📦 SerialNumber (Reg 31088, 10): {serial_regs}\n")
        log_widget.insert(tk.END, f"  ✓ SerialNumber: {serial_ascii}\n")

        model_regs = read_registers(client, device_id, 31106, 8)
        model_ascii = read_ascii_registers(client, device_id, 31106, 8)
        log_widget.insert(tk.END, f"  📦 ProductModel (Reg 31106, 8): {model_regs}\n")
        log_widget.insert(tk.END, f"  ✓ ProductModel: {model_ascii}\n")

        results.append({
            "Device ID": device_id,
            "Device Type": device_type,
            "RFID": rfid_hex,
            "SerialNumber": serial_ascii
        })

        log_widget.update()

    client.close()
    return results

def save_to_csv(data, log_widget):
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not filepath:
        return

    with open(filepath, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=["Device ID", "Device Type", "RFID", "SerialNumber"])
        writer.writeheader()
        writer.writerows(data)

    log_widget.insert(tk.END, f"✓ CSV-Datei gespeichert: {filepath}\n")

def start_scan(entry, log_widget):
    log_widget.delete(1.0, tk.END)
    ip_address = entry.get()
    if not ip_address:
        messagebox.showerror("Fehler", "Bitte eine IP-Adresse eingeben.")
        return
    log_widget.insert(tk.END, f"Starte Verbindung zu {ip_address}...\n")
    log_widget.update()
    results = scan_devices(ip_address, log_widget)
    if results:
        save_to_csv(results, log_widget)
    else:
        log_widget.insert(tk.END, "✗ Keine Geräte gefunden oder Verbindungsfehler.\n")

def create_gui():
    root = tk.Tk()
    root.title("Modbus Export Tool")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="PanelServer IP-Adresse:").pack(anchor="w")
    ip_entry = tk.Entry(frame, width=40)
    ip_entry.pack(anchor="w")
    ip_entry.insert(0, "10.0.1.110")

    tk.Button(frame, text="Scan starten", command=lambda: start_scan(ip_entry, log_output)).pack(pady=10)

    log_output = scrolledtext.ScrolledText(frame, height=30, width=100)
    log_output.pack(fill=tk.BOTH, expand=True)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
