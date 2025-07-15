import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from pymodbus.client.sync import ModbusTcpClient
import csv

PANELSERVER_UNIT_ID = 255
DEVICE_ID_REGISTERS = list(range(504, 1000, 5))

# Gerätespezifische Register-Zuordnung
DEVICE_REGISTER_MAP = {
    "CL110": {
        "RFID": (31026, 4, "uint64"),
        "SerialNumber": (31088, 10, "ascii"),
        "ProductModel": (31106, 8, "ascii"),
    },
    "TH110": {
        "RFID": (31026, 4, "uint64"),
        "SerialNumber": (31088, 10, "ascii"),
        "ProductModel": (31106, 8, "ascii"),
    },
    "HeatTag": {
        "RFID": (31026, 4, "uint64"),
        "SerialNumber": (31088, 10, "ascii"),
        "ProductModel": (31106, 8, "ascii"),
    },
}

def decode_ascii(registers):
    return ''.join(chr((r >> 8) & 0xFF) + chr(r & 0xFF) for r in registers).strip('\x00')

def decode_uint64(registers):
    result = 0
    for r in registers:
        result = (result << 16) | r
    return str(result)

def read_register(client, device_id, address, count):
    try:
        res = client.read_holding_registers(address, count, unit=device_id)
        if not res.isError():
            return res.registers
    except:
        pass
    return None

def log(msg, log_widget=None):
    print(msg)
    if log_widget:
        log_widget.insert(tk.END, msg + '\n')
        log_widget.see(tk.END)
        log_widget.update()

def get_device_ids(client, log_widget=None):
    ids = []
    for reg in DEVICE_ID_REGISTERS:
        res = client.read_holding_registers(reg, 1, unit=PANELSERVER_UNIT_ID)
        if not res.isError():
            device_id = res.registers[0]
            if device_id not in (0, 0xFFFF):
                ids.append(device_id)
                log(f"✓ Reg {reg}: DeviceID {device_id}", log_widget)
            else:
                log(f"- Kein gültiger DeviceID-Wert in Register {reg}", log_widget)
        else:
            log(f"- Fehler beim Lesen von Register {reg}", log_widget)
    return list(set(ids))

def read_device_data(client, device_id, log_widget=None):
    data = {"DeviceID": device_id}

    # Commercial Reference statt DeviceName
    raw_cr = read_register(client, device_id, 31060, 16)
    if raw_cr:
        commercial_ref = decode_ascii(raw_cr)
        data["CommercialReference"] = commercial_ref
        log(f"→ Device {device_id} hat Commercial Reference: {commercial_ref}", log_widget)
        if "EMS59443" in commercial_ref:
            device_type = "CL110"
        elif "EMS59440" in commercial_ref:
            device_type = "TH110"
        else:
            device_type = "HeatTag"  # oder ggf. 'Unbekannt'
        data["DeviceType"] = device_type
    else:
        log(f"⚠ Device {device_id}: Commercial Reference konnte nicht gelesen werden", log_widget)
        return data

    reg_map = DEVICE_REGISTER_MAP.get(device_type)
    if not reg_map:
        log(f"⚠ Kein Register-Mapping für Device {device_id} ({device_type})", log_widget)
        return data

    for key, (reg, count, dtype) in reg_map.items():
        raw = read_register(client, device_id, reg, count)
        if raw:
            log(f"  📦 {key} (Reg {reg}, {count}): {raw}", log_widget)  # Zeige Rohwerte
            if dtype == "ascii":
                data[key] = decode_ascii(raw)
            elif dtype == "uint64":
                data[key] = decode_uint64(raw)
            else:
                data[key] = str(raw)
            log(f"  ✓ {key}: {data[key]}", log_widget)
        else:
            data[key] = "ERROR"
            log(f"  ⚠ {key}: Fehler beim Lesen", log_widget)

    return data

def run_query(ip, log_widget):
    log("Starte Verbindung zu " + ip + "...", log_widget)
    client = ModbusTcpClient(ip)
    if not client.connect():
        messagebox.showerror("Fehler", "Verbindung zum PanelServer fehlgeschlagen.")
        log("✗ Verbindung fehlgeschlagen", log_widget)
        return

    log("✓ Verbindung erfolgreich hergestellt.", log_widget)
    log("→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)", log_widget)

    device_ids = get_device_ids(client, log_widget)
    if not device_ids:
        log("⚠ Keine gültigen DeviceIDs gefunden.", log_widget)
        return

    data = []
    for i, device_id in enumerate(device_ids):
        log(f"[{i+1}/{len(device_ids)}] Verarbeite Device ID {device_id}", log_widget)
        result = read_device_data(client, device_id, log_widget)
        data.append(result)

    client.close()

    if not data:
        log("⚠ Keine Daten zum Exportieren vorhanden.", log_widget)
        return

    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV-Dateien", "*.csv")])
    if filename:
        fieldnames = sorted({k for d in data for k in d.keys()})
        with open(filename, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)
        log(f"✓ CSV-Datei gespeichert: {filename}", log_widget)
        messagebox.showinfo("Erfolg", f"Daten gespeichert in {filename}")

class App:
    def __init__(self, root):
        self.root = root
        root.title("Modbus Wireless Exporter")
        root.geometry("600x500")

        tk.Label(root, text="PanelServer IP-Adresse:").pack(pady=5)
        self.ip_entry = tk.Entry(root, width=30)
        self.ip_entry.pack()

        self.log_box = scrolledtext.ScrolledText(root, width=80, height=25)
        self.log_box.pack(padx=10, pady=10)

        tk.Button(root, text="Abfragen & Exportieren", command=self.start).pack(pady=10)

    def start(self):
        ip = self.ip_entry.get().strip()
        if not ip:
            messagebox.showwarning("Eingabe fehlt", "Bitte IP-Adresse eingeben.")
            return
        self.log_box.delete(1.0, tk.END)
        run_query(ip, self.log_box)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
