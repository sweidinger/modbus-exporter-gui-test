import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pymodbus.client.sync import ModbusTcpClient
import csv
import struct
from openpyxl import Workbook
import os

def log(msg, log_widget=None):
    print(msg)
    if log_widget:
        log_widget.insert(tk.END, msg + "\n")
        log_widget.see(tk.END)

def decode_ascii(registers):
    return "".join(
        chr((reg >> 8) & 0xFF) + chr(reg & 0xFF) for reg in registers
    ).split("\x00")[0].strip()

def read_registers(client, device_id, address, count, log_widget=None):
    try:
        result = client.read_holding_registers(address, count, unit=device_id)
        if result.isError():
            raise Exception(f"Modbus-Fehler: {result}")
        return result.registers
    except Exception as e:
        log(f"⚠ Fehler beim Lesen der Register {address}: {e}", log_widget)
        return None

def get_device_ids(client, log_widget=None):
    base = 504
    step = 5
    max_devices = 100
    device_ids = []

    log("→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)", log_widget)
    for i in range(max_devices):
        addr = base + (i * step)
        result = read_registers(client, 255, addr, 1, log_widget)
        if result and result[0] not in (0, 0xFFFF):
            device_id = result[0]
            log(f"✓ Reg {addr}: DeviceID {device_id}", log_widget)
            device_ids.append(device_id)
        else:
            log(f"- Kein gültiger DeviceID-Wert in Register {addr}", log_widget)
    return device_ids

def collect_data(ip, log_widget=None):
    client = ModbusTcpClient(ip)
    if not client.connect():
        log("❌ Verbindung fehlgeschlagen.", log_widget)
        return None

    log("✓ Verbindung erfolgreich hergestellt.", log_widget)
    device_ids = get_device_ids(client, log_widget)
    if not device_ids:
        log("⚠ Keine gültigen DeviceIDs gefunden.", log_widget)
        client.close()
        return None

    data = []
    for idx, device_id in enumerate(device_ids, start=1):
        log(f"[{idx}/{len(device_ids)}] Verarbeite Device ID {device_id}", log_widget)
        device_data = {
            "DeviceID": device_id,
            "DeviceType": "",
            "RFID": "",
            "SerialNumber": "",
        }

        # Commercial Reference → 31060
        ref_regs = read_registers(client, device_id, 31060, 16, log_widget)
        ref = decode_ascii(ref_regs) if ref_regs else ""
        log(f"→ Device {device_id} hat Commercial Reference: {ref}", log_widget)

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
            log(f"  📦 RFID (Reg 31026, 6): {rfid_regs}", log_widget)
            hex_str = "".join(f"{reg:04X}" for reg in rfid_regs if reg > 0)
            device_data["RFID"] = hex_str[:8]
        else:
            log("  ⚠ RFID: Fehler beim Lesen", log_widget)

        # Serial Number → 31088 (10 Register, ASCII)
        sn_regs = read_registers(client, device_id, 31088, 10, log_widget)
        if sn_regs:
            sn = decode_ascii(sn_regs)
            log(f"  📦 SerialNumber (Reg 31088, 10): {sn_regs}", log_widget)
            log(f"  ✓ SerialNumber: {sn}", log_widget)
            device_data["SerialNumber"] = sn
        else:
            log("  ⚠ SerialNumber: Fehler beim Lesen", log_widget)

        # Product Model (nur Debug) → 31106
        pm_regs = read_registers(client, device_id, 31106, 8, log_widget)
        if pm_regs:
            pm = decode_ascii(pm_regs)
            log(f"  📦 ProductModel (Reg 31106, 8): {pm_regs}", log_widget)
            log(f"  ✓ ProductModel: {pm}", log_widget)

        data.append(device_data)

    client.close()
    return data

def save_csv(data, filename, log_widget=None):
    fieldnames = ["DeviceID", "DeviceType", "RFID", "SerialNumber"]
    with open(filename, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow({k: row.get(k, "") for k in fieldnames})
    log(f"✓ CSV-Datei gespeichert: {filename}", log_widget)

def save_excel(data, filename, log_widget=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Modbus Export"
    headers = ["DeviceID", "DeviceType", "RFID", "SerialNumber"]
    ws.append(headers)
    for row in data:
        ws.append([row.get(h, "") for h in headers])
    wb.save(filename)
    log(f"✓ Excel-Datei gespeichert: {filename}", log_widget)

def on_start():
    ip = ip_entry.get()
    if not ip:
        messagebox.showerror("Fehler", "Bitte IP-Adresse eingeben.")
        return

    if not (export_csv_var.get() or export_xlsx_var.get()):
        messagebox.showerror("Fehler", "Bitte mindestens ein Exportformat auswählen.")
        return

    log_text.delete(1.0, tk.END)
    log(f"Starte Verbindung zu {ip}...", log_text)
    data = collect_data(ip, log_text)

    if data:
        base_file = filedialog.asksaveasfilename(
            title="Dateiname ohne Endung",
            defaultextension="",
            filetypes=[("Alle Dateien", "*.*")]
        )
        if not base_file:
            return

        if export_csv_var.get():
            save_csv(data, base_file + ".csv", log_text)
        if export_xlsx_var.get():
            save_excel(data, base_file + ".xlsx", log_text)
        messagebox.showinfo("Erfolg", "Export abgeschlossen.")

# GUI
root = tk.Tk()
root.title("Modbus Export Tool")

frame = ttk.Frame(root, padding=10)
frame.grid()

ttk.Label(frame, text="PanelServer IP-Adresse:").grid(row=0, column=0, sticky="w")
ip_entry = ttk.Entry(frame, width=20)
ip_entry.grid(row=0, column=1)
ip_entry.insert(0, "10.0.1.110")

start_btn = ttk.Button(frame, text="Start", command=on_start)
start_btn.grid(row=0, column=2, padx=5)

export_csv_var = tk.BooleanVar(value=True)
export_xlsx_var = tk.BooleanVar(value=False)

csv_check = ttk.Checkbutton(frame, text="CSV exportieren", variable=export_csv_var)
xlsx_check = ttk.Checkbutton(frame, text="Excel exportieren", variable=export_xlsx_var)
csv_check.grid(row=1, column=0, sticky="w", pady=5)
xlsx_check.grid(row=1, column=1, sticky="w", pady=5)

log_text = tk.Text(frame, width=100, height=30)
log_text.grid(row=2, column=0, columnspan=3, pady=10)

root.mainloop()
