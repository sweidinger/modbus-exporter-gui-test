import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
try:
    # Try the newer import path first
    from pymodbus.client.tcp import ModbusTcpClient
except ImportError:
    try:
        # Fall back to the older import path
        from pymodbus.client import ModbusTcpClient
    except ImportError:
        # Last resort - very old import path
        from pymodbus.client.sync import ModbusTcpClient
import csv

PANELSERVER_UNIT_ID = 255
DEVICE_ID_REGISTERS = list(range(509, 560, 5))

REGISTERS_COMMON = {
    "DeviceName": (31000, 10, "ascii"),
    "Raw_31030": (31030, 1, "uint"),
    "Raw_31031": (31031, 1, "uint"),
    "Raw_31038": (31038, 1, "uint"),
    "Raw_31039": (31039, 1, "uint"),
}

def decode_ascii(registers):
    return ''.join(chr((r >> 8) & 0xFF) + chr(r & 0xFF) for r in registers).strip('\x00')

def decode_uint(registers):
    return str(registers[0]) if registers else ""

def get_device_ids(client, log):
    ids = []
    log("→ Suche DeviceIDs in alternativen Registern (509, 514, 519, ...)")
    for reg in DEVICE_ID_REGISTERS:
        try:
            # Try both parameter styles for compatibility
            try:
                res = client.read_holding_registers(reg, count=1, slave=PANELSERVER_UNIT_ID)
            except TypeError:
                res = client.read_holding_registers(reg, 1, unit=PANELSERVER_UNIT_ID)
            
            if not res.isError():
                val = res.registers[0]
                if val not in (0, 0xFFFF):
                    ids.append(val)
                    log(f"✓ Reg {reg}: DeviceID {val}")
        except Exception as e:
            log(f"⚠ Fehler bei Reg {reg}: {e}")
    if not ids:
        log("⚠ Keine DeviceIDs gefunden in alternativen Registern.")
    return ids

def read_device_data(client, device_id, log):
    data = {"DeviceID": device_id}
    log(f"→ Lese Daten von Device {device_id}")
    for key, (reg, count, dtype) in REGISTERS_COMMON.items():
        try:
            # Try both parameter styles for compatibility
            try:
                res = client.read_holding_registers(reg, count=count, slave=device_id)
            except TypeError:
                res = client.read_holding_registers(reg, count, unit=device_id)
            
            if res.isError():
                data[key] = "ERROR"
                log(f"  ⚠ {key}: Fehler beim Lesen")
            else:
                if dtype == "ascii":
                    data[key] = decode_ascii(res.registers)
                elif dtype == "uint":
                    data[key] = decode_uint(res.registers)
                log(f"  ✓ {key}: {data[key]}")
        except Exception as e:
            data[key] = "ERROR"
            log(f"  ⚠ {key}: Ausnahmefehler: {e}")
    return data

class App:
    def __init__(self, root):
        self.root = root
        root.title("Modbus Wireless Exporter")
        root.geometry("750x500")

        tk.Label(root, text="PanelServer IP-Adresse:").pack(pady=5)
        self.ip_entry = tk.Entry(root, width=30)
        self.ip_entry.pack()

        tk.Button(root, text="Abfragen & Exportieren", command=self.start).pack(pady=10)

        tk.Label(root, text="Debug Log:").pack()
        self.log_box = scrolledtext.ScrolledText(root, height=20, width=100)
        self.log_box.pack(padx=10, pady=5)

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        print(message)

    def start(self):
        ip = self.ip_entry.get()
        self.log_box.delete("1.0", tk.END)

        if not ip:
            messagebox.showwarning("Eingabe fehlt", "Bitte IP-Adresse eingeben.")
            return

        self.log(f"Starte Verbindung zu {ip}...")
        client = ModbusTcpClient(ip)
        if not client.connect():
            self.log("❌ Verbindung fehlgeschlagen!")
            messagebox.showerror("Fehler", "Keine Verbindung zum PanelServer möglich.")
            return
        self.log("✓ Verbindung erfolgreich hergestellt.")

        device_ids = get_device_ids(client, self.log)
        if not device_ids:
            self.log("⚠ Keine gültigen DeviceIDs gefunden.")
            client.close()
            return

        data = []
        for i, device_id in enumerate(device_ids):
            self.log(f"[{i+1}/{len(device_ids)}] Verarbeite Device ID {device_id}")
            data.append(read_device_data(client, device_id, self.log))

        client.close()
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if filename:
            with open(filename, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["DeviceID"] + list(REGISTERS_COMMON.keys()))
                writer.writeheader()
                writer.writerows(data)
            self.log(f"✓ CSV-Datei gespeichert: {filename}")
            messagebox.showinfo("Fertig", f"Daten gespeichert in {filename}")
        else:
            self.log("⚠ CSV-Speichern abgebrochen.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
