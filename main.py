import sys
import csv
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout,
    QWidget, QFileDialog, QTextEdit, QLineEdit
)
from pymodbus.client.sync import ModbusTcpClient
from pymodbus.exceptions import ModbusIOException


DEVICE_ID_REGISTERS = list(range(501, 1000, 5))  # 501, 506, 511, ..., 996
DEVICE_ID_OFFSET = 4  # next register is always virtual address

DEVICE_TYPES = {
    "CL110": {
        "model_reg": 31107,
        "name_reg": 31000,
        "sn_reg": 31089,
        "rfid_reg": 31027
    },
    "TH110": {
        "model_reg": 31107,
        "name_reg": 31000,
        "sn_reg": 31089,
        "rfid_reg": 31027
    },
    "HeatTag": {
        "model_reg": 31107,
        "name_reg": 31000,
        "sn_reg": 31089,
        "rfid_reg": 31027
    }
}


def read_string(client, unit, address, length):
    try:
        rr = client.read_holding_registers(address, length, unit=unit)
        if not rr or not hasattr(rr, "registers"):
            return ""
        chars = [chr(reg) for reg in rr.registers if 32 <= reg <= 126]
        return ''.join(chars).strip()
    except Exception:
        return ""


class ModbusExporter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modbus Export Tool")

        self.ip_input = QLineEdit()
        self.ip_input.setPlaceholderText("PanelServer IP-Adresse (z.B. 192.168.1.100)")

        self.debug_output = QTextEdit()
        self.debug_output.setReadOnly(True)

        self.export_button = QPushButton("CSV exportieren")
        self.export_button.clicked.connect(self.export_data)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("PanelServer IP-Adresse:"))
        layout.addWidget(self.ip_input)
        layout.addWidget(self.export_button)
        layout.addWidget(QLabel("Debug-Ausgabe:"))
        layout.addWidget(self.debug_output)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def log(self, message):
        self.debug_output.append(message)
        QApplication.processEvents()

    def export_data(self):
        ip = self.ip_input.text().strip()
        if not ip:
            self.log("⚠ Bitte IP-Adresse eingeben.")
            return

        self.log(f"Starte Verbindung zu {ip}...")

        try:
            client = ModbusTcpClient(ip)
            if not client.connect():
                self.log("❌ Verbindung fehlgeschlagen.")
                return
            self.log("✓ Verbindung erfolgreich hergestellt.")
        except Exception as e:
            self.log(f"Fehler beim Verbinden: {str(e)}")
            return

        devices = []

        for rfid_reg in DEVICE_ID_REGISTERS:
            try:
                rfid_rr = client.read_holding_registers(rfid_reg, 4, unit=255)
                addr_rr = client.read_holding_registers(rfid_reg + DEVICE_ID_OFFSET, 1, unit=255)

                if not rfid_rr or not addr_rr or not hasattr(rfid_rr, "registers") or not hasattr(addr_rr, "registers"):
                    continue

                if all(reg == 0xFFFF for reg in rfid_rr.registers):
                    continue

                virtual_addr = addr_rr.registers[0]
                if virtual_addr == 0xFFFF:
                    continue

                devices.append(virtual_addr)
                self.log(f"✓ Gerät gefunden – Virtuelle Adresse: {virtual_addr}")

            except ModbusIOException:
                self.log(f"❌ Lesefehler bei Register {rfid_reg}")
            except Exception as e:
                self.log(f"⚠ Ausnahme bei Register {rfid_reg}: {str(e)}")

        if not devices:
            self.log("⚠ Keine gültigen DeviceIDs gefunden.")
            return

        output_data = [["Virtuelle Adresse", "Gerätename", "Produkttyp", "Seriennummer", "RF-ID"]]

        for dev_id in devices:
            self.log(f"→ Lese Gerät {dev_id}...")

            model = read_string(client, dev_id, DEVICE_TYPES["CL110"]["model_reg"], 8)
            name = read_string(client, dev_id, DEVICE_TYPES["CL110"]["name_reg"], 10)
            sn = read_string(client, dev_id, DEVICE_TYPES["CL110"]["sn_reg"], 10)
            rfid = read_string(client, dev_id, DEVICE_TYPES["CL110"]["rfid_reg"], 4)

            dev_type = "Unbekannt"
            for dtype, regs in DEVICE_TYPES.items():
                if model.startswith(dtype):
                    dev_type = dtype
                    break

            output_data.append([dev_id, name, dev_type, sn, rfid])

        client.close()

        path, _ = QFileDialog.getSaveFileName(self, "CSV speichern", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, "w", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerows(output_data)
                self.log(f"✓ CSV exportiert: {path}")
            except Exception as e:
                self.log(f"Fehler beim Speichern: {str(e)}")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = ModbusExporter()
        window.show()
        sys.exit(app.exec_())
    except Exception:
        print("Fehler beim Start der Anwendung:")
        traceback.print_exc()
