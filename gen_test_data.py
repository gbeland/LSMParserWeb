import openpyxl
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Log"

# Headers
ws.append(["Date", "Time", "Source", "IP", "Message"])

# Valid SBox Entry (ID=1)
sbox_ip = "192.168.1.10"
sbox_resp = "AA FF 01 06 41 8A 54 42 2D 48 4D 53 00"
ws.append(["2025-01-01", "12:00:00", "Dev", sbox_ip, sbox_resp])

# SBox Serial "SN12345"
sbox_sn = "AA FF 01 07 41 0B 53 4E 31 32 33 34 35 00"
ws.append(["", "", "", sbox_ip, sbox_sn])

# SBox FW "T-HMS 1000"
sbox_fw = "AA FF 01 0A 41 0E 54 2D 48 4D 53 20 31 30 30 30 00"
ws.append(["", "", "", sbox_ip, sbox_fw])

# SBox Group IPs
sbox_gips = "AA FF 01 13 41 1B 84 C0 A8 0A 01 C0 A8 14 01 00 00 00 00 00 00 00 00 00"
ws.append(["", "", "", sbox_ip, sbox_gips])

# Cabinet Entry
cab1_ip = "192.168.10.1"
cab_resp_hdr = "AA FF 05"
cab_model = "AA FF 05 03 41 8A 49 46 48 00"
ws.append(["", "", "", cab1_ip, cab_model])

cab_sn = "AA FF 05 06 41 0B 43 41 42 30 30 31 00"
ws.append(["", "", "", cab1_ip, cab_sn])

cab_layout = "AA FF 05 0B 41 8C A0 00 00 00 00 01 E0 01 0E 00"
ws.append(["", "", "", cab1_ip, cab_layout])

wb.save("test_log.xlsx")
print("Created test_log.xlsx")
