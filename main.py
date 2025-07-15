import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from pymodbus.client.sync import ModbusTcpClient
import csv

# Modbus Register Info
PANELSERVER_UNIT_ID = 255
DEVICE_ID_REGISTERS = list(range(504, 1000, 5))  # 0-basiert!

# Wireless Device Register Map (alle Adressen um -1 korrigiert!)
REGISTERS = {
    "DeviceName": (31000, 10, "ascii"),
    "RFID": (31026, 4, "uint64"),
    "SerialNumber": (31088, 10, "ascii"),
    "ProductModel": (31106, 8, "ascii"),
}

def decode_ascii(registers):
    return ''.join(chr((r >> 8) & 0xFF) + chr(r & 0xFF) for r in registers).strip('\x00')

def decode_uint64(registers):
    result = 0
    for r in registers:
        result = (result << 16) | r
    return str(result)

def get_device_ids(client, log_callback):
    ids = []
    log_callback("→ Suche DeviceIDs in alternativen Registern (504, 509, 514, ...)")
    for reg in DEVICE_ID_REGISTERS:
        res = client.read_holding_registers(reg, 1, unit=PANELSERVER_UNIT_ID)
        if not res.isError():
            device_id = res.registers[0]
            if device_id not in (0, 0xFFFF):
                log_callback(f"✓ Reg {reg}: DeviceID {device_id}")
                ids.append(device_id)
            else:
                log_callback(f"- Kein gültiger DeviceID-Wert in Register {reg}")
        else:
            log_callback(f"⚠ Fehler beim Lesen von Register {reg}")
    return list(set(ids))

def read_device_data(client, device_id, log_callback):
    data = {"DeviceID": device_id}
    log_callback(f"→ Lese Daten von Device {device_id}")
    for key, (reg, count, dtype) in REGISTERS.items():
        try:
            res = client.read_holding_registers(reg, count, unit=device_id)
            if res.isError():
                log_callback(f"  ⚠ Fehler beim Lesen von {key} (Reg {reg})")
                data[key] = "ERROR"
            else:
                if dtype == "ascii":
                    data[key] = decode_ascii(res.registers)
                elif dtype == "uint64":
                    data[key] = decode_uint64(res.registers)
                log_callback(f"  ✓ {key}: {data[key]}")
        except Exception as e:
            log_callback(f"  ⚠ Ausnahme beim Lesen von {key}: {e}")
            data[key] = "ERROR"
    return data

def run_query(ip, log_callback):
    client = ModbusTcpClient(ip)
    if not client.connect():
        log_callback("✖ Verbindung zum PanelServer fehlgeschlagen.")
        messagebox.showerror("Fehler", "Verbindung zum PanelServer fehlgeschlagen.")
        return

    log_callback(f"✓ Verbindung erfolgreich hergestellt.")
    device_ids = get_device_ids(client, log_callback)

    if not device_ids:
        log_callback("⚠ Keine gültigen DeviceIDs gefunden.")
        messagebox.showwarning("Keine Geräte", "Es wurden keine gültigen DeviceIDs gefunden.")
        client.close()
        return

    data = []
    for i, device_id in enumerate(device_ids):
        log_callback(f"[{i+1}/{len(device_ids)}] Verarbeite Device ID {device_id}")
        data.append(read_device_data(client, device_id, log_callback))

    client.close()

    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        with open(filename, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["DeviceID"] + list(REGISTERS.keys()))
            writer.writeheader()
            writer.writerows(data)
        log_callback(f"✓ CSV-Datei gespeichert: {filename}")
        messagebox.showinfo("Fertig", f"Daten gespeichert in {filename}")
    else:
        log_callback("⚠ Export abgebrochen – keine Datei ausgewählt.")

class App:
    def __init__(self, root):
        self.root = root
        root.title("Modbus Wireless Exporter")
        root.geometry("600x400")

        tk.Label(root, text="PanelServer IP-Adresse:").pack(pady=5)
        self.ip_entry = tk.Entry(root, width=30)
        self.ip_entry.pack()

        tk.Button(root, text="Abfragen & Exportieren", command=self.start).pack(pady=10)

        self.log_window = scrolledtext.ScrolledText(root, height=15, width=80, state='disabled')
        self.log_window.pack(padx=10, pady=5)

    def log(self, msg):
        self.log_window.config(state='normal')
        self.log_window.insert(tk.END, msg + '\n')
        self.log_window.see(tk.END)
        self.log_window.config(state='disabled')

    def start(self):
        ip = self.ip_entry.get().strip()
        if not ip:
            messagebox.showwarning("Eingabe fehlt", "Bitte IP-Adresse eingeben.")
            return
        self.log_window.config(state='normal')
        self.log_window.delete(1.0, tk.END)
        self.log_window.config(state='disabled')
        run_query(ip, self.log)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
