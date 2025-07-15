from pymodbus.client import ModbusTcpClient
from pymodbus.payload import BinaryPayloadDecoder
from pymodbus.constants import Endian
import csv
import sys

# Konfiguration
PANELSERVER_IP = '10.0.1.110'
PANELSERVER_ID = 255
OUTPUT_FILE = 'modbus_devices.csv'

# Feste Register-Adressen für die Device-IDs
RFID_START = 501
VIRTUAL_OFFSET = 4
DEVICE_COUNT = 100

# Geräte-spezifische Register
REGISTER_SETS = {
    "device_name": (31000, 10),
    "product_model": (31107, 8),
    "serial_number": (31089, 10),
    "rfid": (31027, 4)
}


def read_ascii_registers(client, unit_id, address, length):
    try:
        result = client.read_holding_registers(address, length, unit=unit_id)
        if result is None or result.isError():
            print(f"⚠ Fehler beim Lesen von Adresse {address} (Länge {length}) bei DeviceID {unit_id}")
            return ""
        decoder = BinaryPayloadDecoder.fromRegisters(result.registers, Endian.Big)
        return decoder.decode_string(length * 2).decode(errors="ignore").strip('\x00 ')
    except Exception as e:
        print(f"❌ Ausnahme beim Lesen von Adresse {address}: {e}")
        return ""


def get_device_ids(client):
    device_ids = []
    print("→ Suche DeviceIDs in festen Registern...")
    for i in range(DEVICE_COUNT):
        rfid_reg = RFID_START + i * 5
        virtual_reg = rfid_reg + VIRTUAL_OFFSET
        try:
            result = client.read_holding_registers(virtual_reg, 1, unit=PANELSERVER_ID)
            if result is None or result.isError():
                continue
            val = result.registers[0]
            if val != 0xFFFF:
                print(f"✓ Gefunden: DeviceID {val} bei Register {virtual_reg}")
                device_ids.append(val)
        except Exception as e:
            print(f"⚠ Fehler beim Lesen von Register {virtual_reg}: {e}")
            continue
    return device_ids


def poll_device_info(client, device_id):
    info = {"DeviceID": device_id}
    for key, (reg, length) in REGISTER_SETS.items():
        info[key] = read_ascii_registers(client, device_id, reg, length)
    return info


def main():
    print(f"Starte Verbindung zu {PANELSERVER_IP}...")
    client = ModbusTcpClient(PANELSERVER_IP)
    try:
        if not client.connect():
            print("✖ Verbindung fehlgeschlagen.")
            return

        print("✓ Verbindung erfolgreich hergestellt.")
        devices = get_device_ids(client)

        if not devices:
            print("⚠ Keine gültigen DeviceIDs gefunden.")
            return

        data = []
        for device_id in devices:
            print(f"→ Lese Daten von DeviceID {device_id}...")
            info = poll_device_info(client, device_id)
            data.append(info)

        # CSV schreiben
        fieldnames = ["DeviceID", "device_name", "product_model", "serial_number", "rfid"]
        with open(OUTPUT_FILE, mode='w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)

        print(f"✓ Export abgeschlossen: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Allgemeiner Fehler: {e}")
    finally:
        client.close()


if __name__ == "__main__":
    main()
