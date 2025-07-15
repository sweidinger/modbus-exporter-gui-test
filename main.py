import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from pymodbus.client.sync import ModbusTcpClient
import csv

# Modbus Register Info
PANELSERVER_UNIT_ID = 255
DEVICE_ID_REGISTERS = list(range(505, 1001, 5))

# Wireless Device Register Map
REGISTERS = {
    "DeviceName": (31001, 10, "ascii"),
    "RFID": (31027, 4, "uint64"),
    "SerialNumber": (31089, 10, "ascii"),
    "ProductModel": (31107, 8, "ascii"),
}

def decode_ascii(registers):
    return ''.join(chr((r >> 8) & 0xFF) + chr(r & 0xFF) for r in registers).strip('\x00')

def decode_uint64(registers):
    result = 0
    for r in registers:
        result = (result << 16) | r
    return str(result)

def get_device_ids(client, log):
    ids = []
    for reg in DEVICE_ID_REGISTERS:
        res = client.read_holding_registers(reg, 1, unit=PANELSERVER_UNIT_ID)
        if not res.isError():
            device_id = res.registers[0]
            if device_id not in (0, 0xFFFF):
                ids.append(device_id)
                log(f"✓ DeviceID gefunden: {device_id} (Register {reg})")
            else:
                log(f"- Kein gültiger DeviceID-Wert in Register {reg}")
        else:
            log(f"⚠ Fehler beim Lesen von Register {reg}")
    return list(set(ids))

def read_device_data(client, device_id, log):
    data = {"DeviceID": device_id}
    for key, (reg, count, dtype) in REGISTERS.items():
        try:
            res = client.read_holding_registers(reg, count, unit=device_id)
            if res.isError():
                data[key] = "ERROR"
                log(f"⚠ Fehler beim Lesen {key} (DeviceID {device_id})")
            else:
                if dtype == "ascii":
                    data[key] = decode_ascii(res.registers)
                elif dtype == "uint64":
                    data[key] = decode_uint64(res.registers)
                log(f"✓ {key} gelesen für DeviceID {device_id}")
        except Exception as e:
            data[key] = "ERROR"
            log(f"⚠ Ausnahme bei {key}: {e}")
    return data

def run_query(ip, log):
    log(f"Starte Verbindung zu {ip}...")
    client = ModbusTcpClient(ip)
    if not client.connect():
        messagebox.showerror("Fehler", "Verbindung zum PanelServer fehlgeschlagen.")
        log("✗ Verbindung fehlgeschlagen.")
        return

    log("✓ Verbindung erfolgreich hergestellt.")
    log("→ Suche DeviceIDs in alternativen Registern (505, 510, 515, ...)")
    device_ids = get_device_ids(client, log)

    if not device_ids:
        log("⚠ Keine gültigen DeviceIDs gefunden.")
    else:
        log(f"✓ Gefundene DeviceIDs: {device_ids}")

    data = []
    for i, device_id in enumerate(device_ids):
        log(f"→ Lese Daten für DeviceID {device_id} ({i+1}/{len(device_ids)})")
        data.append(read_device_data(client, device_id, log))

    client.close()
    if not data:
        log("⚠ Keine Daten zum Exportieren vorhanden.")
        return

    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        with open(filename, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["DeviceID"] + list(REGISTERS.keys()))
            writer.writeheader()
            writer.writerows(data)
        log(f"✓ Daten gespeichert in {filename}")
        messagebox.showinfo("Fertig", f"Daten gespeichert in {filename}")

class App:
    def __init__(self, root):
        self.root = root
        root.title("Modbus Wireless Exporter")
        root.geometry("600x400")

        tk.Label(root, text="PanelServer IP-Adresse:").pack(pady=5)
        self.ip_entry = tk.Entry(root, width=30)
        self.ip_entry.pack()

        tk.Button(root, text="Abfragen & Exportieren", command=self.start).pack(pady=5)

        tk.Label(root, text="Debug-Log:").pack()
        self.log_area = scrolledtext.ScrolledText(root, height=15, width=80, state="disabled")
        self.log_area.pack(padx=10, pady=5)

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state="disabled")

    def start(self):
        ip = self.ip_entry.get()
        if not ip:
            messagebox.showwarning("Eingabe fehlt", "Bitte IP-Adresse eingeben.")
            return
        self.log_area.configure(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state="disabled")
        run_query(ip, self.log)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
