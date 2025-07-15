from pymodbus.client.tcp import ModbusTcpClient
from pymodbus.exceptions import ModbusIOException
import struct
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import csv
import os

# ------------------------ GUI SETUP ------------------------

def log_debug(message):
    debug_text.configure(state='normal')
    debug_text.insert(tk.END, message + "\n")
    debug_text.see(tk.END)
    debug_text.configure(state='disabled')
    root.update()

def connect_to_modbus(ip):
    try:
        log_debug(f"Starte Verbindung zu {ip}...")
        client = ModbusTcpClient(ip, port=502)
        if client.connect():
            log_debug("✓ Verbindung erfolgreich hergestellt.")
            return client
        else:
            log_debug("✗ Verbindung fehlgeschlagen.")
            return None
    except Exception as e:
        log_debug(f"Fehler bei Verbindung: {e}")
        return None

def read_ascii(client, unit, start, length):
    try:
        rr = client.read_holding_registers(start, length, unit=unit)
        if isinstance(rr, ModbusIOException) or not rr or not rr.registers:
            return ""
        chars = []
        for reg in rr.registers:
            chars.append((reg >> 8) & 0xFF)
            chars.append(reg & 0xFF)
        return bytes(chars).decode("ascii", errors="ignore").strip('\x00')
    except Exception as e:
        log_debug(f"Fehler beim Lesen von ASCII-Register {start}: {e}")
        return ""

def read_uint64(client, unit, start):
    try:
        rr = client.read_holding_registers(start, 4, unit=unit)
        if not rr or not rr.registers:
            return ""
        value = (rr.registers[0] << 48) + (rr.registers[1] << 32) + (rr.registers[2] << 16) + rr.registers[3]
        return hex(value)
    except Exception as e:
        log_debug(f"Fehler beim Lesen von UINT64 {start}: {e}")
        return ""

def detect_devices(client):
    device_ids = []
    log_debug("→ Suche DeviceIDs in alternativen Registern (509, 514, 519, ...)")
    for i in range(509, 600, 5):
        rr = client.read_holding_registers(i, 1, unit=255)
        if rr and rr.registers and rr.registers[0] != 0xFFFF:
            device_ids.append(rr.registers[0])
            log_debug(f"✓ Reg {i}: DeviceID {rr.registers[0]}")
    if not device_ids:
        log_debug("⚠ Keine DeviceIDs gefunden in alternativen Registern.")
    return device_ids

def read_device_data(client, unit_id):
    log_debug(f"→ Lese Daten von Device {unit_id}")
    name = read_ascii(client, unit_id, 31001, 10)
    product_model = read_ascii(client, unit_id, 31107, 8)
    serial = read_ascii(client, unit_id, 31089, 10)
    rfid = read_uint64(client, unit_id, 31027)
    log_debug(f"  ✓ DeviceName: {name}")
    return {"ID": unit_id, "Name": name, "Typ": product_model, "Seriennummer": serial, "RFID": rfid}

def export_data():
    ip = ip_entry.get()
    if not ip:
        messagebox.showwarning("Warnung", "Bitte IP-Adresse eingeben.")
        return

    client = connect_to_modbus(ip)
    if not client:
        return

    device_ids = detect_devices(client)
    if not device_ids:
        log_debug("⚠ Keine gültigen DeviceIDs gefunden.")
        client.close()
        return

    data = []
    for i, unit in enumerate(device_ids, start=1):
        log_debug(f"[{i}/{len(device_ids)}] Verarbeite Device ID {unit}")
        info = read_device_data(client, unit)
        data.append(info)

    client.close()

    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV-Dateien", "*.csv")])
    if filepath:
        with open(filepath, "w", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=["ID", "Name", "Typ", "Seriennummer", "RFID"])
            writer.writeheader()
            writer.writerows(data)
        log_debug(f"✓ CSV-Datei gespeichert: {filepath}")
        messagebox.showinfo("Erfolg", f"CSV-Datei erfolgreich gespeichert:\n{filepath}")

# ------------------------ GUI ------------------------

root = tk.Tk()
root.title("Modbus Export Tool")
root.geometry("640x480")

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="PanelServer IP-Adresse:").pack(side=tk.LEFT, padx=(0, 5))
ip_entry = tk.Entry(frame, width=20)
ip_entry.pack(side=tk.LEFT)
ip_entry.insert(0, "10.0.1.110")

start_button = tk.Button(frame, text="Export starten", command=export_data)
start_button.pack(side=tk.LEFT, padx=10)

debug_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=25, state='disabled')
debug_text.pack(expand=True, fill='both', padx=10, pady=10)

root.mainloop()
