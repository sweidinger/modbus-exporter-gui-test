import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pymodbus.client.sync import ModbusTcpClient
import csv

PANELSERVER_UNIT_ID = 255
DEVICE_ID_REGISTERS = list(range(504, 1001, 5))

# Registerdefinitionen je Gerätetyp
REGISTER_SETS = {
    "CL110": {
        "DeviceName": (31001, 10, "ascii"),
        "RFID": (31027, 4, "uint64"),
        "SerialNumber": (31089, 10, "ascii"),
        "ProductModel": (31107, 8, "ascii"),
    },
    "TH110": {
        "DeviceName": (31001, 10, "ascii"),
        "RFID": (31027, 4, "uint64"),
        "SerialNumber": (31089, 10, "ascii"),
        "ProductModel": (31107, 8, "ascii"),
    },
    "HeatTag": {
        "DeviceName": (31001, 10, "ascii"),
        "RFID": (31027, 4, "uint64"),
        "SerialNumber": (31089, 10, "ascii"),
        "ProductModel": (31107, 8, "ascii"),
    }
}

def decode_ascii(registers):
    return ''.join(chr((r >> 8) & 0xFF) + chr(r & 0xFF) for r in registers).strip('\x00')

def decode_uint64(registers):
    result = 0
    for r in registers:
        result = (result << 16) | r
    return str(result)

def log(text, log_widget):
    log_widget.insert(tk.END, text + "\n")
    log_widget.see(tk.END)
    log_widget.update()

def get_device_ids(client, log_widget):
    ids = []
    log("→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)", log_widget)
    for reg in DEVICE_ID_REGISTERS:
        res = client.read_holding_registers(reg, 1, unit=PANELSERVER_UNIT_ID)
        if not res.isError():
            val = res.registers[0]
            if val not in (0, 0xFFFF):
                log(f"✓ Reg {reg}: DeviceID {val}", log_widget)
                ids.append(val)
            else:
                log(f"- Kein gültiger DeviceID-Wert in Register {reg}", log_widget)
        else:
            log(f"⚠ Fehler beim Lesen von Register {reg}", log_widget)
    return list(set(ids))

def read_register(client, device_id, address, count):
    res = client.read_holding_registers(address, count, unit=device_id)
    if res.isError():
        return None
    return res.registers

def read_device_data(client, device_id, log_widget):
    data = {"DeviceID": device_id, "DeviceType": "Unbekannt"}
    
    # Ersten Versuch: ProductModel auslesen
    raw_pm = read_register(client, device_id, 31107, 8)
    if raw_pm:
        product_model = decode_ascii(raw_pm)
        data["ProductModel"] = product_model
        log(f"→ Device {device_id} hat ProductModel: {product_model}", log_widget)
        if "CL110" in product_model:
            device_type = "CL110"
        elif "TH110" in product_model:
            device_type = "TH110"
        elif "HT" in product_model or "HeatTag" in product_model:
            device_type = "HeatTag"
        else:
            device_type = "Unbekannt"
        data["DeviceType"] = device_type
    else:
        log(f"⚠ Device {device_id}: ProductModel konnte nicht gelesen werden", log_widget)
        return data

    regset = REGISTER_SETS.get(device_type)
    if not regset:
        log(f"⚠ Device {device_id}: Kein Register-Set bekannt für {device_type}", log_widget)
        return data

    for key, (reg, count, dtype) in regset.items():
        raw = read_register(client, device_id, reg, count)
        if raw:
            if dtype == "ascii":
                data[key] = decode_ascii(raw)
            elif dtype == "uint64":
                data[key] = decode_uint64(raw)
        else:
            data[key] = "ERROR"
            log(f"⚠ Device {device_id}: Fehler beim Lesen von {key}", log_widget)
    return data

def run_query(ip, log_widget):
    log(f"Starte Verbindung zu {ip}...", log_widget)
    client = ModbusTcpClient(ip)
    if not client.connect():
        log("❌ Verbindung fehlgeschlagen.", log_widget)
        messagebox.showerror("Fehler", "Verbindung zum PanelServer fehlgeschlagen.")
        return

    log("✓ Verbindung erfolgreich hergestellt.", log_widget)
    device_ids = get_device_ids(client, log_widget)

    if not device_ids:
        log("⚠ Keine gültigen DeviceIDs gefunden.", log_widget)
        messagebox.showwarning("Keine Geräte", "Es konnten keine gültigen Geräte gefunden werden.")
        return

    data = []
    for idx, device_id in enumerate(device_ids, 1):
        log(f"[{idx}/{len(device_ids)}] Verarbeite Device ID {device_id}", log_widget)
        device_data = read_device_data(client, device_id, log_widget)
        if device_data:
            data.append(device_data)

    client.close()

    if data:
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV-Datei", "*.csv")])
        if filename:
            with open(filename, "w", newline="") as f:
                fieldnames = ["DeviceID", "DeviceType", "DeviceName", "RFID", "SerialNumber", "ProductModel"]
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(data)
            log(f"✓ CSV-Datei gespeichert: {filename}", log_widget)
            messagebox.showinfo("Fertig", "Daten erfolgreich exportiert.")
    else:
        log("⚠ Keine Daten zum Exportieren vorhanden.", log_widget)

class App:
    def __init__(self, root):
        self.root = root
        root.title("Modbus Wireless Exporter")
        root.geometry("600x500")

        tk.Label(root, text="PanelServer IP-Adresse:").pack(pady=5)
        self.ip_entry = tk.Entry(root, width=40)
        self.ip_entry.pack(pady=5)

        tk.Button(root, text="Abfragen & Exportieren", command=self.start).pack(pady=10)

        self.log_output = scrolledtext.ScrolledText(root, height=20, width=80, state=tk.NORMAL)
        self.log_output.pack(pady=10)

    def start(self):
        ip = self.ip_entry.get()
        if not ip:
            messagebox.showwarning("Eingabe fehlt", "Bitte IP-Adresse eingeben.")
            return
        self.log_output.delete("1.0", tk.END)
        run_query(ip, self.log_output)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
