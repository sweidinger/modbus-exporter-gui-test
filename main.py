import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from pymodbus.client import ModbusTcpClient
from pymodbus.exceptions import ModbusIOException
import csv

# Relevante Register für Name, Typ, Seriennummer, RF-ID
REGISTERS = {
    "DeviceName": (31000, 10),
    "RFID": (31027, 4),
    "SerialNumber": (31089, 10),
    "ProductModel": (31107, 8),
}

ALTERNATIVE_ID_REGISTERS = [509, 514, 519, 524, 529, 534, 539]

def read_ascii(client, unit_id, address, length):
    try:
        rr = client.read_holding_registers(address, length, unit=unit_id)
        if isinstance(rr, ModbusIOException) or not rr or not rr.registers:
            return None
        data = bytearray()
        for reg in rr.registers:
            data.extend(reg.to_bytes(2, byteorder="big"))
        return data.decode("ascii", errors="ignore").strip("\x00").strip()
    except Exception:
        return None

def read_rfid(client, unit_id, address, length):
    try:
        rr = client.read_holding_registers(address, length, unit=unit_id)
        if isinstance(rr, ModbusIOException) or not rr or not rr.registers:
            return None
        value = ""
        for reg in rr.registers:
            value += format(reg, "04X")
        return value
    except Exception:
        return None

class ModbusExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus Export Tool")

        self.ip_label = ttk.Label(root, text="PanelServer IP-Adresse:")
        self.ip_label.pack(pady=2)

        self.ip_entry = ttk.Entry(root, width=30)
        self.ip_entry.pack(pady=2)
        self.ip_entry.insert(0, "10.0.1.110")

        self.export_button = ttk.Button(root, text="Export starten", command=self.export_data)
        self.export_button.pack(pady=10)

        self.debug_output = scrolledtext.ScrolledText(root, width=80, height=20)
        self.debug_output.pack(padx=10, pady=5)

    def log(self, message):
        self.debug_output.insert(tk.END, message + "\n")
        self.debug_output.see(tk.END)
        self.root.update()

    def export_data(self):
        ip = self.ip_entry.get()
        self.debug_output.delete(1.0, tk.END)

        self.log(f"Starte Verbindung zu {ip}...")
        client = ModbusTcpClient(ip)
        if not client.connect():
            self.log("❌ Verbindung fehlgeschlagen.")
            return
        self.log("✓ Verbindung erfolgreich hergestellt.")

        device_ids = []
        self.log("→ Suche DeviceIDs in alternativen Registern (509, 514, 519, ...)")
        for reg in ALTERNATIVE_ID_REGISTERS:
            try:
                rr = client.read_holding_registers(reg, 1, unit=255)
                if rr and rr.registers:
                    device_id = rr.registers[0]
                    device_ids.append(device_id)
                    self.log(f"✓ Reg {reg}: DeviceID {device_id}")
            except:
                pass

        if not device_ids:
            self.log("⚠ Keine gültigen DeviceIDs gefunden.")
            return

        results = []
        for idx, device_id in enumerate(device_ids, 1):
            self.log(f"[{idx}/{len(device_ids)}] Verarbeite Device ID {device_id}")
            row = {"DeviceID": device_id}
            for key, (address, length) in REGISTERS.items():
                self.log(f"→ Lese {key} von Device {device_id}")
                if key == "RFID":
                    val = read_rfid(client, device_id, address, length)
                else:
                    val = read_ascii(client, device_id, address, length)
                row[key] = val if val else "Fehler"
                self.log(f"  ✓ {key}: {row[key]}")
            results.append(row)

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV-Dateien", "*.csv")])
        if not save_path:
            return
        with open(save_path, "w", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=["DeviceID", "DeviceName", "ProductModel", "SerialNumber", "RFID"])
            writer.writeheader()
            writer.writerows(results)

        self.log(f"✓ CSV-Datei gespeichert: {save_path}")
        client.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModbusExporterApp(root)
    root.mainloop()
