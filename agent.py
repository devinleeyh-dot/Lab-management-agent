"""
Lab Management Agent - Phase 1 + Phase 2 + Phase 3
Phase 1: request text -> parse items (qty at back) -> check Excel stock -> create tickets -> reply.
Phase 2: short stock -> draft PO email -> ask lab owner on WhatsApp -> APPROVE/REJECT
         -> email vendor (Gmail) -> update ticket -> tell requester.
Phase 3: Low-Stock Dashboard, Individual Delivery Records (DO Log), Strict MOQ,
         Completed/Pending tracking columns, Zero Inventory reservation while Pending > 0,
         Capped Completed Quantity, Requester Ticket Closure, Automatic Surplus Stock Credit,
         Strict Below-MinQty Shortage Alerts (< MinQty), Intelligent PO MOQ/MinQty Top-Up Recommendation,
         Sender-Restricted Ticket History & Status Queries, Command Cheat Sheets ('command'),
         Dynamic Admin Promotion ('GRANT'), Admin Access Withdrawal ('WITHDRAW'),
         Multiple Labs / Item Categories Support, Categorized R&D BOM Stock Catalog,
         and Integrated 'Categorization Guide' Sheet.
"""
import os, re, json, smtplib, threading, subprocess
from datetime import datetime
from email.message import EmailMessage

import requests
from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook

EXCEL_PATH = os.environ.get("LAB_EXCEL", "lab_inventory.xlsx")
GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

# --- Phase 2 settings ---
BRIDGE_URL = os.environ.get("BRIDGE_URL", "http://127.0.0.1:5001/send")
OWNER_PHONE = os.environ.get("OWNER_PHONE", "6589033626")        # lab owner WhatsApp (digits only)
BUYER_NAME = os.environ.get("BUYER_NAME", "YH Lee (SG)")
GMAIL_USER = os.environ.get("GMAIL_USER", "devinleeyh@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "")

lock = threading.Lock()  # only one write to Excel at a time
app = Flask(__name__)

# ==========================================
# 1. AUTOMATICALLY LAUNCH NODE.JS BRIDGE
# ==========================================
def launch_node_bridge():
    """
    Spawns bridge.js as a background child process so both Python
    and Node.js run inside a single host container.
    """
    try:
        print("[Agent] Spawning Node.js WhatsApp bridge (bridge.js)...")
        subprocess.Popen(["node", "bridge.js"])
        print("[Agent] Node.js bridge launched successfully in background.")
    except FileNotFoundError:
        print("[Agent Error] 'node' executable not found in PATH. Ensure Node.js is installed in this container.")
    except Exception as e:
        print(f"[Agent Error] Failed to launch bridge.js: {e}")

# Call the launcher before initializing workbooks and server
launch_node_bridge()

INV_HEADERS = ["Item", "Qty", "Unit", "MinQty", "MOQ", "Location", "Category", "Picture"]
REQ_HEADERS = ["TicketID", "Time", "Requester", "Item", "QtyRequested", "Completed", "Pending", "Status", "Location", "Category", "Note"]
VENDOR_HEADERS = ["Vendor", "Email", "Keywords"]
PO_HEADERS = ["POID", "Time", "TicketID", "Item", "Qty", "Vendor", "VendorEmail", "Status", "Note"]
DELIVERY_HEADERS = ["POID", "Time", "BatchReceived", "Status", "Note"]
REQUESTER_HEADERS = ["Name", "Phone", "WhatsAppID", "Role"]
UNKNOWN_HEADERS = ["WhatsAppID", "LastSeen", "LastText"]
CAT_GUIDE_HEADERS = ["Category Name", "Sub-Category / Scope", "Description Guidelines", "Example Components"]


def init_workbook():
    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()
        inv = wb.active
        inv.title = "Inventory"
        inv.append(INV_HEADERS)

        # Exact On-Hand Stock from Hardware_Components_BOM_v10.xlsx + Category Dummy Items
        inventory_data = [
            # 23 Real BOM Components
            ["TowerPro SG92R Digital Micro Servo Motor", 1, "pcs", 1, 2, "Singapore Lab", "Motors & Actuators", "MOT-SER-003"],
            ["TowerPro MG90S Digital Micro Servo Motor", 1, "pcs", 1, 2, "Singapore Lab", "Motors & Actuators", "MOT-SER-007"],
            ["Raspberry Pi Active Cooler for Pi 5", 1, "pcs", 1, 2, "Singapore Lab", "Cooling & Thermal", "5056561803357"],
            ["SanDisk Ultra microSDHC 32GB Card", 1, "pcs", 1, 2, "Singapore Lab", "Memory & Storage", "SDSUNR-032G"],
            ["Raspberry Pi M.2 HAT+ for Pi 5", 1, "pcs", 1, 2, "Singapore Lab", "Expansion Boards & HATs", "M.2 HAT+ M Key"],
            ["Kingston NV3 PCIe 4.0 NVMe M.2 2230 SSD 500GB", 1, "pcs", 1, 2, "Singapore Lab", "Memory & Storage", "9907705-008.A01G"],
            ["Raspberry Pi 5 Single Board Computer 8GB", 1, "pcs", 1, 2, "Singapore Lab", "Single Board Computers", "DFR1119 / RPI5-8GB"],
            ["Raspberry Pi Official Micro USB Power Supply 5.1V 2.5A", 1, "pcs", 1, 2, "Singapore Lab", "Power Supplies & Adapters", "0617588405754"],
            ["1.30 inch OLED Display Module 128x64 I2C", 1, "pcs", 1, 2, "Singapore Lab", "Displays & Optoelectronics", "MOD-DIS-020"],
            ["Raspberry Pi 27W USB-C Power Supply UK Plug", 1, "pcs", 1, 2, "Singapore Lab", "Power Supplies & Adapters", "27W USB-C UK"],
            ["Raspberry Pi Zero 2 W Single Board Computer", 1, "pcs", 1, 2, "Singapore Lab", "Single Board Computers", "5056561800004"],
            ["Lithium Battery Charger BMS Board 2S 2A", 1, "pcs", 1, 2, "Singapore Lab", "Power Management", "MOD-CHA-009"],
            ["9DOF 9-Axis Accel Gyro Mag Sensor BNO085", 1, "pcs", 1, 2, "Singapore Lab", "Sensors & Modules", "SEN-ACC-005"],
            ["DC-DC Step Down Buck Converter 7-32V to 5V 5A", 1, "pcs", 1, 2, "Singapore Lab", "Power Management", "MOD-DC-013"],
            ["Seeed Studio XIAO ESP32-C6 Microcontroller Board", 1, "pcs", 1, 2, "Singapore Lab", "Microcontrollers", "113991157"],
            ["Feetech / Waveshare STS3215 Serial Bus Smart Servo Motor", 14, "pcs", 2, 5, "Singapore Lab", "Motors & Actuators", "ST-3215-C001"],
            ["Seeed Studio ReSpeaker 2-Mics Pi HAT V2.0", 2, "pcs", 1, 2, "Singapore Lab", "Expansion Boards & HATs", "107100001"],
            ["DFRobot FireBeetle Board ESP32-S3 Microcontroller", 1, "pcs", 1, 2, "Singapore Lab", "Microcontrollers", "60068260624"],
            ["Waveshare Bus Servo Adapter (A) V1.1", 1, "pcs", 1, 2, "Singapore Lab", "Expansion Boards & HATs", "21263135"],
            ["1.28 inch Round IPS TFT LCD Display Module GC9A01", 1, "pcs", 1, 2, "Singapore Lab", "Displays & Optoelectronics", "MOD-DIS-128"],
            ["Heltec Automation WiFi kit 32 ESP32 Dev Board", 1, "pcs", 1, 2, "Singapore Lab", "Microcontrollers", "HTIT-WB32"],
            ["2.42 inch OLED Display Module 128x64 SPI/I2C", 1, "pcs", 1, 2, "Singapore Lab", "Displays & Optoelectronics", "MOD-DIS-029"],
            ["Syntiant TinyML Development Board NDP101 AI Processor", 4, "pcs", 1, 2, "Singapore Lab", "Microcontrollers & AI", "3944-SYNTIANTTINYML"],

            # Dummy Components for Categorization Coverage
            ["10k Ohm Carbon Film Resistor 1/4W", 200, "pcs", 50, 100, "Singapore Lab", "Passive Components", "RES-10K-025W"],
            ["100nF 50V Ceramic Capacitor", 150, "pcs", 30, 100, "Singapore Lab", "Passive Components", "CAP-100NF-50V"],
            ["LM358 Dual Operational Amplifier IC", 20, "pcs", 5, 10, "Singapore Lab", "Active Components", "IC-LM358-DIP8"],
            ["40-Pin Male to Female Jumper Wires 20cm", 30, "packs", 10, 20, "Singapore Lab", "Connectors & Cables", "CAB-JUMP-MF20"],
            ["Digital Multimeter 6000 Counts Auto-Ranging", 2, "pcs", 1, 1, "Singapore Lab", "Tools & Equipment", "TL-DMM-6000"]
        ]

        for item in inventory_data:
            inv.append(item)

        req = wb.create_sheet("Requests")
        req.append(REQ_HEADERS)
        wb.save(EXCEL_PATH)

    wb = load_workbook(EXCEL_PATH)
    changed = False

    # Ensure 'Categorization Guide' sheet exists
    if "Categorization Guide" not in wb.sheetnames:
        cat_ws = wb.create_sheet("Categorization Guide")
        cat_ws.append(CAT_GUIDE_HEADERS)
        cat_guide_data = [
            ["Motors & Actuators", "Servos, Stepper Motors, DC Motors, Linear Actuators", "Motion control devices, torque, rotation range, gear specs", "TowerPro SG92R, MG90S, Feetech STS3215 Serial Bus Servo"],
            ["Single Board Computers", "Full Linux/OS Computers, ARM SBCs", "Standalone processing units with RAM, OS storage support", "Raspberry Pi 5 8GB, Raspberry Pi Zero 2 W"],
            ["Microcontrollers & AI", "MCU Dev Boards, AI Processors, Wireless MCUs", "Low-power microcontrollers, WiFi/BLE modules, Neural NPUs", "ESP32-C6, ESP32-S3, Syntiant TinyML NDP101, Heltec WiFi Kit 32"],
            ["Expansion Boards & HATs", "Raspberry Pi HATs, Shield Boards, Servo Controllers", "Add-on PCBs that plug into SBCs or MCU pin headers", "Raspberry Pi M.2 HAT+, ReSpeaker 2-Mics HAT, Bus Servo Adapter"],
            ["Displays & Optoelectronics", "OLED, TFT LCD, E-Paper, LED Modules", "Visual feedback screens, panel drivers, screen resolution", "1.30\" OLED 128x64, 1.28\" Round IPS TFT, 2.42\" OLED"],
            ["Sensors & Modules", "IMU, Environmental, Acoustic, Optical Sensors", "Transducers for physical measurements (accel, gyro, temp)", "9DOF BNO085 IMU, DHT22 Temp/Humidity, Ultrasonic HC-SR04"],
            ["Power Management", "BMS Boards, Buck/Boost Converters, Regulators", "Voltage conversion, battery charging, power distribution", "Lithium 2S BMS 2A, DC-DC Buck Converter 7-32V to 5V 5A"],
            ["Power Supplies & Adapters", "AC-DC Adapters, Wall Chargers, Power Cables", "External power adapters supplying constant voltage/current", "Raspberry Pi 27W USB-C PSU, 5.1V 2.5A Micro USB PSU"],
            ["Memory & Storage", "MicroSD Cards, NVMe SSDs, Flash Storage", "Non-volatile storage media for OS and data retention", "SanDisk Ultra 32GB MicroSD, Kingston NV3 500GB NVMe SSD"],
            ["Cooling & Thermal", "Active Coolers, Heatsinks, Cooling Fans", "Thermal dissipation solutions for ICs and high-power processors", "Raspberry Pi 5 Active Cooler"],
            ["Passive Components", "Resistors, Capacitors, Inductors, Potentiometers", "Passive circuit elements with fixed electrical characteristics", "10k Ohm Resistor, 100nF Ceramic Capacitor"],
            ["Active Components", "Op-Amps, Diodes, Transistors, Logic ICs", "Active semiconductors and integrated circuit chips", "LM358 Op-Amp, NE555 Timer IC, 1N4007 Diode"],
            ["Connectors & Cables", "Jumper Wires, Terminal Blocks, Headers, Adapters", "Interconnect hardware, breadboards, wiring harnesses", "40-Pin Jumper Wires M-F, Breadboard 830 Points"],
            ["Tools & Equipment", "Multimeters, Soldering Irons, Oscilloscopes, Hand Tools", "Measurement, assembly, testing, and debugging instruments", "Digital Multimeter 6000 Counts, Soldering Station"]
        ]
        for row in cat_guide_data:
            cat_ws.append(row)
        changed = True

    if "Vendors" not in wb.sheetnames:
        v = wb.create_sheet("Vendors")
        v.append(VENDOR_HEADERS)
        v.append(["Devin", "devinleeyh@hotmail.com", ""])
        changed = True
    if "PO" not in wb.sheetnames:
        wb.create_sheet("PO").append(PO_HEADERS)
        changed = True
    if "Deliveries" not in wb.sheetnames:
        wb.create_sheet("Deliveries").append(DELIVERY_HEADERS)
        changed = True
    if "Requesters" not in wb.sheetnames:
        r = wb.create_sheet("Requesters")
        r.append(REQUESTER_HEADERS)
        r.append(["ABC", "60127110960", "269401924317247", "owner"])
        changed = True
    if "Unknown" not in wb.sheetnames:
        wb.create_sheet("Unknown").append(UNKNOWN_HEADERS)
        changed = True

    inv_ws = wb["Inventory"]
    if inv_ws.cell(1, 8).value != "Picture":
        inv_ws.cell(1, 6).value = "Location"
        inv_ws.cell(1, 7).value = "Category"
        inv_ws.cell(1, 8).value = "Picture"
        for row in inv_ws.iter_rows(min_row=2):
            if row[5].value is None:
                row[5].value = "Singapore Lab"
            if row[6].value is None:
                row[6].value = "General"
            if row[7].value is None:
                row[7].value = ""
        changed = True

    req_ws = wb["Requests"]
    if req_ws.cell(1, 9).value != "Location":
        req_ws.cell(1, 9).value = "Location"
        req_ws.cell(1, 10).value = "Category"
        req_ws.cell(1, 11).value = "Note"
        changed = True

    req_sheet = wb["Requesters"]
    if req_sheet.cell(1, 4).value != "Role":
        req_sheet.cell(1, 4).value = "Role"
        for r in range(2, req_sheet.max_row + 1):
            if req_sheet.cell(r, 1).value:
                req_sheet.cell(r, 4).value = "requester"
        changed = True

    if changed:
        wb.save(EXCEL_PATH)


# ---------- parsing ----------
def parse_simple(text):
    """Parser supporting formats like '10k resistor 10 Singapore Lab' or '10k resistor 10'."""
    items = []
    for line in re.split(r"[\n;]+", text):
        line = line.strip()
        if not line:
            continue
        m_loc = re.match(r"^(.+?)\s+(\d+)\s*(?:pcs|pc)?\s+(Lab\s+[A-Z]|Building\s+\d+|[A-Z0-9_-]+)$", line, re.I)
        m_back = re.match(r"^(.+?)\s+(?:x|pcs|pc)?\s*(\d+)\s*(?:pcs|pc)?$", line, re.I)
        m_front = re.match(r"^(\d+)\s*(?:x|pcs|pc)?\s+(.+)$", line, re.I)

        if m_loc:
            items.append({"item": m_loc.group(1).strip(), "qty": int(m_loc.group(2)), "location": m_loc.group(3).strip()})
        elif m_back and not m_back.group(1).isdigit():
            items.append({"item": m_back.group(1).strip(), "qty": int(m_back.group(2)), "location": None})
        elif m_front:
            items.append({"item": m_front.group(2).strip(), "qty": int(m_front.group(1)), "location": None})
        else:
            return []
    return items


def parse_gemini(text):
    """Use Gemini only when the simple parser cannot understand the message."""
    if not GEMINI_KEY:
        return []
    prompt = (
        "Extract lab item requests from the message (e.g. '10k resistor 10 Singapore Lab'). Reply ONLY with JSON: "
        '[{"item": "name", "qty": number, "location": "location name or null"}]. If none, reply [].\n\nMessage:\n' + text
    )
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_KEY}"
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"},
    }
    try:
        r = requests.post(url, json=body, timeout=30)
        r.raise_for_status()
        raw = r.json()["candidates"][0]["content"]["parts"][0]["text"]
        data = json.loads(raw)
        return [
            {"item": str(d["item"]), "qty": int(d["qty"]), "location": d.get("location")}
            for d in data
            if d.get("item") and int(d.get("qty", 0)) > 0
        ]
    except Exception as e:
        print("Gemini error:", e)
        return []


# ---------- WhatsApp / email helpers ----------
def wa_send(phone, text):
    """Send a WhatsApp message through the bridge (bridge.js must be running)."""
    try:
        r = requests.post(BRIDGE_URL, json={"phone": phone, "text": text}, timeout=20)
        if r.status_code != 200:
            print("WhatsApp send failed:", r.status_code, r.text)
            return False
        return True
    except Exception as e:
        print("WhatsApp send error:", e)
        return False


def draft_po_email(poid, vendor_name, item, qty):
    subject = f"Purchase Order {poid} - {item} x{qty}"
    body = (
        f"Dear {vendor_name},\n\n"
        f"Please supply the following item:\n\n"
        f"  PO number : {poid}\n"
        f"  Item      : {item}\n"
        f"  Quantity  : {qty}\n\n"
        f"Kindly reply with the price, availability and estimated delivery date (ETA).\n\n"
        f"Thank you.\n\n"
        f"Best regards,\n{BUYER_NAME}\n"
    )
    return subject, body


def send_email(to_addr, subject, body):
    """Send via Gmail SMTP. Returns (ok, error_text)."""
    if not GMAIL_APP_PASSWORD:
        return False, "GMAIL_APP_PASSWORD is not set"
    try:
        msg = EmailMessage()
        msg["From"] = f"{BUYER_NAME} <{GMAIL_USER}>"
        msg["To"] = to_addr
        msg["Subject"] = subject
        msg.set_content(body)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as s:
            s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            s.send_message(msg)
        return True, ""
    except Exception as e:
        return False, str(e)


# ---------- Excel logic ----------
def find_stock_full(ws, name, preferred_location=None):
    """Return (row_number, qty, unit, min_qty, moq, location, category, picture) of best match, or None."""
    name = name.lower()
    matches = []
    for row in ws.iter_rows(min_row=2):
        cell = str(row[0].value or "").lower()
        if cell and (name in cell or cell in name):
            qty = int(row[1].value or 0)
            min_qty = int(row[3].value or 0) if row[3].value is not None else 0
            moq = int(row[4].value or 1) if row[4].value is not None else 1
            loc = str(row[5].value or "Singapore Lab")
            cat = str(row[6].value or "General")
            pic = str(row[7].value or "") if len(row) > 7 else ""
            matches.append((row[0].row, qty, row[2].value or "pcs", min_qty, moq, loc, cat, pic))

    if not matches:
        return None

    if preferred_location:
        for m in matches:
            if preferred_location.lower() in m[5].lower():
                return m

    return matches[0]


def next_ticket_id(ws):
    n = ws.max_row
    return f"T{datetime.now():%y%m}-{n:04d}"


def find_ticket(ws, tid):
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value or "").upper() == tid.upper():
            return row[0].row
    return None


def find_po(ws, poid):
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value or "").upper() == poid.upper():
            return row[0].row
    return None


def load_vendors(wb):
    out = []
    for r in wb["Vendors"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[1]:
            out.append((str(r[0]).strip(), str(r[1]).strip(), str(r[2] or "").lower()))
    return out


def pick_vendor(vendors, item, preferred=None):
    """Owner's choice wins. Otherwise first vendor whose keywords match, else first vendor."""
    if preferred:
        for v in vendors:
            if v[0].lower() == preferred.strip().lower():
                return v
        return None
    for v in vendors:
        kws = [k.strip() for k in v[2].split(",") if k.strip()]
        if any(k in item.lower() for k in kws):
            return v
    return vendors[0] if vendors else None


def phone_from_requester(requester):
    m = re.search(r"\((\d{8,15})\)", str(requester))
    return m.group(1) if m else None


def is_matching_sender(req_sender, current_sender):
    """Verify if the ticket's requester matches the current sender digits/name."""
    if not current_sender:
        return True
    s_digits = re.sub(r"\D", "", str(current_sender))
    r_digits = re.sub(r"\D", "", str(req_sender))
    if s_digits and r_digits and (s_digits in r_digits or r_digits in s_digits):
        return True
    return str(current_sender).strip().lower() in str(req_sender).strip().lower()


def get_sender_role(sender):
    """Check whether sender is owner or requester based on Requesters sheet."""
    if not sender:
        return "requester"
    s_digits = re.sub(r"\D", "", str(sender))
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        ws = wb["Requesters"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            r_name, r_phone, r_wid, r_role = str(row[0] or "").strip(), str(row[1] or ""), str(row[2] or ""), str(row[3] or "requester").strip().lower()
            r_phone_digits = re.sub(r"\D", "", r_phone)
            if (s_digits and (s_digits == r_phone_digits or s_digits == r_wid)) or (sender.lower() in r_name.lower()):
                wb.close()
                return r_role if r_role in ("owner", "admin") else "requester"
        wb.close()
    return "requester"


def grant_admin_access(target_name_or_phone):
    """Promote a requester to owner/admin role."""
    with lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Requesters"]
        target = str(target_name_or_phone).strip().lower()
        target_digits = re.sub(r"\D", "", target)
        found = False
        for row in ws.iter_rows(min_row=2):
            r_name = str(row[0].value or "").strip().lower()
            r_phone = str(row[1].value or "").strip()
            if target == r_name or (target_digits and target_digits in r_phone):
                row[3].value = "owner"
                found = True
                name_val = row[0].value
                wb.save(EXCEL_PATH)
                return f"Success: Requester '{name_val}' has been granted Admin/Owner access! They can now execute owner commands."
        wb.save(EXCEL_PATH)
        if not found:
            return f"Requester '{target_name_or_phone}' not found in Requesters sheet."


def withdraw_admin_access(target_name_or_phone):
    """Revoke admin/owner role and set back to standard requester."""
    with lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Requesters"]
        target = str(target_name_or_phone).strip().lower()
        target_digits = re.sub(r"\D", "", target)
        found = False
        for row in ws.iter_rows(min_row=2):
            r_name = str(row[0].value or "").strip().lower()
            r_phone = str(row[1].value or "").strip()
            if target == r_name or (target_digits and target_digits in r_phone):
                row[3].value = "requester"
                found = True
                name_val = row[0].value
                wb.save(EXCEL_PATH)
                return f"Success: Admin/Owner access has been withdrawn from requester '{name_val}'. They are now set back to standard requester role."
        wb.save(EXCEL_PATH)
        if not found:
            return f"Requester '{target_name_or_phone}' not found in Requesters sheet."


def get_ticket_status_report(tid, current_sender=None, is_owner=False):
    """Check single ticket status, verifying ownership for requesters."""
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        req_ws = wb["Requests"]
        row_idx = find_ticket(req_ws, tid)
        if row_idx is None:
            wb.close()
            return f"Ticket {tid} not found."

        row_vals = [req_ws.cell(row_idx, col).value for col in range(1, 12)]
        wb.close()

    if not is_owner and not is_matching_sender(row_vals[2], current_sender):
        return f"Access Denied: Ticket {tid} does not belong to your account."

    return (
        f"📋 Status for Ticket {row_vals[0]}:\n"
        f"- Time: {row_vals[1]}\n"
        f"- Requester: {row_vals[2]}\n"
        f"- Item: {row_vals[3]}\n"
        f"- Location: {row_vals[8]} | Category: {row_vals[9]}\n"
        f"- Qty Requested: {row_vals[4]}\n"
        f"- Completed: {row_vals[5]}\n"
        f"- Pending: {row_vals[6]}\n"
        f"- Status: {row_vals[7]}\n"
        f"- Note: {row_vals[10]}"
    )


def get_requester_history_report(sender_info):
    """List all tickets associated ONLY with the specific requester who sent the query."""
    user_tickets = []
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        req_ws = wb["Requests"]
        for row in req_ws.iter_rows(min_row=2, values_only=True):
            tid, time_str, req_user, item, qty_req, completed, pending, status, loc, cat, note = row[:11]
            if tid and is_matching_sender(req_user, sender_info):
                user_tickets.append(f"- {tid}: {item} [{loc}] (Req: {qty_req}, Done: {completed}, Pending: {pending}) -> Status: [{status}]")
        wb.close()

    if not user_tickets:
        return f"No ticket history found for your account ({sender_info})."

    return f"📋 Ticket History for {sender_info}:\n" + "\n".join(user_tickets)


def get_all_tickets_report_for_owner():
    """List all tickets across the entire system for the lab owner."""
    tickets = []
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        req_ws = wb["Requests"]
        for row in req_ws.iter_rows(min_row=2, values_only=True):
            tid, time_str, req_user, item, qty_req, completed, pending, status, loc, cat, note = row[:11]
            if tid:
                tickets.append(f"- {tid} ({time_str}) | {req_user}: {item} [{loc}] (Req:{qty_req}, Done:{completed}, Pend:{pending}) -> [{status}]")
        wb.close()

    if not tickets:
        return "No tickets found in the system."

    return "📋 All System Tickets (Owner View):\n" + "\n".join(tickets)


def requester_command_help():
    return (
        "🤖 **Lab Requester Commands Guide**:\n"
        "  • **<Item> <Qty> [Location]** (e.g. `10k resistor 10 Singapore Lab`) -> Request items from specific or general location.\n"
        "  • **stock [Location/Category]** -> View current inventory levels (optional location/category filter).\n"
        "  • **lowstock [Location/Category]** -> Check items below minimum stock.\n"
        "  • **history** or **tickets** -> View your past and active item requests.\n"
        "  • **status <TicketID>** -> Check specific ticket details.\n"
        "  • **command** or **help** -> Show this command guide."
    )


def owner_command_help():
    return (
        "👑 **Lab Owner / Admin Commands Guide**:\n"
        "  • **stock [Location/Category]** -> View all lab stock levels.\n"
        "  • **lowstock [Location/Category]** -> View Low-Stock Dashboard.\n"
        "  • **status** or **tickets** -> View all system tickets across all users.\n"
        "  • **status <TicketID>** -> Inspect any ticket in detail.\n"
        "  • **APPROVE <TicketID>** -> Approve request using baseline MOQ PO.\n"
        "  • **APPROVE <TicketID> <VendorName>** -> Approve and pick a specific vendor.\n"
        "  • **TOPUP <TicketID> <Qty>** -> Approve request with custom top-up PO quantity.\n"
        "  • **REJECT <TicketID>** -> Reject a request ticket.\n"
        "  • **RECEIVE <POID> [Qty>** -> Log incoming delivery batch.\n"
        "  • **GRANT <Name or Phone>** -> Promote a requester to Admin/Owner access.\n"
        "  • **WITHDRAW <Name or Phone>** -> Withdraw Admin/Owner access from a requester.\n"
        "  • **command** or **help** -> Show this command guide."
    )


def process_request(sender, items):
    lines, po_jobs = [], []
    with lock:
        wb = load_workbook(EXCEL_PATH)
        inv, req = wb["Inventory"], wb["Requests"]
        vendors = load_vendors(wb)
        for it in items:
            tid = next_ticket_id(req)
            hit = find_stock_full(inv, it["item"], preferred_location=it.get("location"))

            requested_qty = it["qty"]
            min_qty = hit[3] if hit else 0
            moq = hit[4] if hit else 1
            loc = hit[5] if hit else (it.get("location") or "Singapore Lab")
            cat = hit[6] if hit else "General"

            if hit is None:
                completed = 0
                pending = requested_qty
                status, note = "Need PO", "Item not in inventory list"
                po_qty = moq
                shortage_qty = requested_qty
                suggested_topup = min_qty + pending
                lines.append(f"{tid}: {it['item']} {requested_qty}pcs [{loc}] - not in inventory, PO created following MOQ: {po_qty}")
            elif hit[1] >= requested_qty:
                new_qty = hit[1] - requested_qty
                inv.cell(hit[0], 2).value = new_qty
                completed = requested_qty
                pending = 0
                status, note = "Completed", f"Collected {requested_qty}, remaining stock at {loc}: {new_qty}"
                po_qty = requested_qty
                shortage_qty = 0
                suggested_topup = moq
                lines.append(f"{tid}: {it['item']} {requested_qty}pcs [{loc}] - collected from stock ({new_qty} left). Ticket closed!")
            else:
                available_qty = hit[1]
                shortage_qty = requested_qty - available_qty
                if available_qty > 0:
                    inv.cell(hit[0], 2).value = 0
                    completed = available_qty
                    pending = requested_qty - available_qty
                    status, note = "Partial Received", f"Collected partial {completed}, Pending {pending}, MOQ PO: {moq}"
                    lines.append(f"{tid}: {it['item']} {requested_qty}pcs [{loc}] - collected available {available_qty}pcs first! Shortage PO created following MOQ: {moq}")
                else:
                    completed = 0
                    pending = requested_qty
                    status, note = "Need PO", f"Out of stock at {loc}, MOQ PO: {moq}"
                    lines.append(f"{tid}: {it['item']} {requested_qty}pcs [{loc}] - out of stock, PO created following MOQ: {moq}")
                po_qty = moq
                suggested_topup = min_qty + pending

            req.append([tid, f"{datetime.now():%Y-%m-%d %H:%M}", sender, it["item"], requested_qty, completed, pending, status, loc, cat, note])
            if status in ("Need PO", "Partial Received") and pending > 0:
                po_jobs.append({
                    "tid": tid, "item": it["item"], "qty": po_qty, "shortage": shortage_qty,
                    "moq": moq, "min_qty": min_qty, "pending": pending, "suggested_topup": max(po_qty, suggested_topup),
                    "note": note, "requester": sender, "location": loc, "category": cat, "vendor": pick_vendor(vendors, it["item"]),
                    "vendor_names": [v[0] for v in vendors]
                })
        wb.save(EXCEL_PATH)

    for job in po_jobs:
        threading.Thread(target=notify_owner, args=(job,), daemon=True).start()
    return "Request received:\n" + "\n".join(lines) + "\nLab owner will review."


def notify_owner(job):
    v = job["vendor"]
    if v is None:
        wa_send(OWNER_PHONE, f"PO needed for {job['tid']} ({job['item']} {job['qty']}pcs [{job['location']}]) but the Vendors sheet is empty. Please add a vendor.")
        return
    poid = "PO-" + job["tid"]

    moq_note = ""
    if job["shortage"] < job["moq"]:
        topup_rec = max(job["moq"], job["suggested_topup"])
        moq_note = (
            f"\n\n⚠️ MOQ Advisory: Shortage quantity ({job['shortage']}pcs) is less than vendor MOQ ({job['moq']}pcs).\n"
            f"To satisfy MinQty ({job['min_qty']}pcs) after fulfilling requester's pending order ({job['pending']}pcs), "
            f"we advise topping up the PO quantity to **{topup_rec}pcs**.\n"
            f"Please confirm top-up decision prior to showing email for PO."
        )

    text = (
        f"PO needed - ticket {job['tid']}\n"
        f"Requester: {job['requester']}\n"
        f"Item: {job['item']} {job['qty']}pcs [{job['location']}] ({job['note']}){moq_note}\n"
        f"Vendor: {v[0]} <{v[1]}>\n\n"
        f"Reply to proceed:\n"
        f"APPROVE {job['tid']} (uses baseline MOQ qty: {job['qty']})\n"
        f"APPROVE {job['tid']} <vendor name>\n"
        f"TOPUP {job['tid']} <new qty> (e.g. TOPUP {job['tid']} {max(job['moq'], job['suggested_topup'])})\n"
        f"REJECT {job['tid']}"
    )
    wa_send(OWNER_PHONE, text)


OWNER_CMD = re.compile(r"^\s*(APPROVE|REJECT|TOPUP)\s+(T\d+-\d+)(?:\s+(\S.*?))?\s*$", re.I)
RECEIVE_CMD = re.compile(r"^\s*(RECEIVE|ARRIVE)\s+(PO-T\d+-\d+)(?:\s+(\d+))?\s*$", re.I)
GRANT_CMD = re.compile(r"^\s*GRANT\s+(.+?)\s*$", re.I)
WITHDRAW_CMD = re.compile(r"^\s*WITHDRAW\s+(.+?)\s*$", re.I)
STATUS_CMD = re.compile(r"^\s*(?:status|check|ticket)?\s*(T\d+-\d+)\s*$", re.I)


def handle_owner_command(text):
    text_clean = text.strip().lower()
    if text_clean.startswith("stock") or text_clean.startswith("inventory") or text_clean.startswith("/stock"):
        parts = text.split(maxsplit=1)
        filter_arg = parts[1] if len(parts) > 1 else None
        return stock_report(filter_arg)
    if text_clean.startswith("lowstock") or text_clean.startswith("alerts") or text_clean.startswith("alert"):
        parts = text.split(maxsplit=1)
        filter_arg = parts[1] if len(parts) > 1 else None
        return low_stock_report(filter_arg)
    if text_clean in ("status", "history", "tickets", "/status"):
        return get_all_tickets_report_for_owner()
    if text_clean in ("command", "commands", "help", "/help", "/command"):
        return owner_command_help()

    m_grant = GRANT_CMD.match(text)
    if m_grant:
        return grant_admin_access(m_grant.group(1))

    m_withdraw = WITHDRAW_CMD.match(text)
    if m_withdraw:
        return withdraw_admin_access(m_withdraw.group(1))

    m_status = STATUS_CMD.match(text)
    if m_status:
        return get_ticket_status_report(m_status.group(1), is_owner=True)

    m_recv = RECEIVE_CMD.match(text)
    if m_recv:
        poid = m_recv.group(2).upper()
        custom_qty = m_recv.group(3)
        with lock:
            wb = load_workbook(EXCEL_PATH)
            po_ws = wb["PO"]
            row = find_po(po_ws, poid)
            if row is None:
                return f"Purchase Order {poid} not found in PO sheet."

            item_name = po_ws.cell(row, 4).value
            ordered_qty = int(po_ws.cell(row, 5).value or 0)
            ticket_id = po_ws.cell(row, 3).value

            deliveries_ws = wb["Deliveries"]
            prev_received = 0
            for d_row in deliveries_ws.iter_rows(min_row=2, values_only=True):
                if str(d_row[0] or "").upper() == poid:
                    try:
                        note_txt = str(d_row[4] or "")
                        m_batch = re.search(r"\+([0-9]+)", note_txt)
                        if m_batch:
                            prev_received += int(m_batch.group(1))
                    except ValueError:
                        pass

            batch_qty = int(custom_qty) if custom_qty else (ordered_qty - prev_received)
            total_received = prev_received + batch_qty

            if total_received < ordered_qty:
                po_status = "Incomplete"
                batch_note = f"Batch received: +{batch_qty} (Total: {total_received}/{ordered_qty})"
            else:
                po_status = "Complete"
                batch_note = f"Batch received: +{batch_qty} (Fully Fulfilled: {total_received}/{ordered_qty})"

            deliveries_ws.append([
                poid,
                f"{datetime.now():%Y-%m-%d %H:%M}",
                batch_qty,
                po_status,
                batch_note
            ])

            po_ws.cell(row, 8).value = po_status
            po_ws.cell(row, 9).value = f"Received cumulative {total_received}/{ordered_qty}"

            req_ws = wb["Requests"]
            req_row = find_ticket(req_ws, ticket_id)
            requester_name = None
            ticket_closed = False
            ticket_already_closed = False
            surplus_stock = 0

            if req_row:
                qty_requested = int(req_ws.cell(req_row, 5).value or 0)
                current_completed = int(req_ws.cell(req_row, 6).value or 0)
                needed_qty = qty_requested - current_completed
                requester_name = req_ws.cell(req_row, 3).value

                if needed_qty <= 0:
                    new_completed = qty_requested
                    new_pending = 0
                    surplus_stock = batch_qty
                    ticket_already_closed = True
                elif batch_qty >= needed_qty:
                    new_completed = qty_requested
                    new_pending = 0
                    surplus_stock = batch_qty - needed_qty
                    ticket_closed = True
                else:
                    new_completed = current_completed + batch_qty
                    new_pending = qty_requested - new_completed
                    surplus_stock = 0

                req_ws.cell(req_row, 6).value = new_completed
                req_ws.cell(req_row, 7).value = new_pending

                if new_pending == 0 and not ticket_already_closed:
                    req_ws.cell(req_row, 8).value = "Completed"
                    req_ws.cell(req_row, 11).value = f"Collected {new_completed}/{qty_requested}. Ticket Closed!"
                elif ticket_already_closed:
                    req_ws.cell(req_row, 8).value = "Completed"
                    req_ws.cell(req_row, 11).value = f"Fully received ({qty_requested}/{qty_requested}) & surplus stored in stock"
                else:
                    req_ws.cell(req_row, 8).value = "Partial Received"
                    req_ws.cell(req_row, 11).value = f"Collected {new_completed}/{qty_requested}, Pending {new_pending}"

            inv_ws = wb["Inventory"]
            hit = find_stock_full(inv_ws, item_name)
            new_inv_qty = 0
            min_qty_val = 0
            if hit:
                inv_row, current_stock, unit, min_qty_val, moq, loc_val, cat_val, pic_val = hit
                if new_pending > 0 and not ticket_already_closed:
                    new_inv_qty = 0
                else:
                    new_inv_qty = current_stock + surplus_stock
                inv_ws.cell(inv_row, 2).value = new_inv_qty
            else:
                new_inv_qty = 0 if (new_pending > 0 and not ticket_already_closed) else surplus_stock
                inv_ws.append([item_name, new_inv_qty, "pcs", 5, 1, "Singapore Lab", "General", ""])

            wb.save(EXCEL_PATH)

        if requester_name:
            phone = phone_from_requester(requester_name)
            if phone:
                if ticket_closed:
                    wa_send(phone, f"Good news! Your remaining batch for '{item_name}' (Ticket {ticket_id}) has arrived ({new_completed}/{qty_requested}). You can now collect the balance quantity in the lab. Your task is now fully complete. Thank you, and have a great day! 👍")
                elif not ticket_already_closed and new_pending > 0 and batch_qty > 0:
                    wa_send(phone, f"📦 Notice for Ticket {ticket_id}: A delivery batch of {batch_qty} for '{item_name}' has arrived (Total received: {new_completed}/{qty_requested}). You can now collect this available batch from the lab while we await the rest!")

        owner_log = f"PO {poid} batch logged as '{po_status}'. "
        if ticket_closed:
            owner_log += f"Requester ticket {ticket_id} closed (Pending=0). Conversation ended for requester. Surplus stock (+{surplus_stock}) updated to inventory. "
        elif ticket_already_closed:
            owner_log += f"Surplus stock (+{batch_qty}) added to inventory. "
        else:
            owner_log += f"Updated requester ticket {ticket_id} (Completed: {new_completed}, Pending: {new_pending}). "

        if new_inv_qty < min_qty_val:
            owner_log += f"⚠️ Stock is below MinQty ({new_inv_qty} < {min_qty_val}), low-stock alert active."
        else:
            owner_log += f"Stock Qty is {new_inv_qty} (MinQty: {min_qty_val}), meeting minimum requirements."

        return owner_log

    m = OWNER_CMD.match(text)
    if not m:
        return None

    action, tid, vendor_arg = m.group(1).upper(), m.group(2).upper(), m.group(3)

    with lock:
        wb = load_workbook(EXCEL_PATH)
        req = wb["Requests"]
        row = find_ticket(req, tid)
        if row is None:
            return f"Ticket {tid} not found."

        status = req.cell(row, 8).value
        if status not in ("Need PO", "Partial Received", "Completed"):
            return f"Ticket {tid} is '{status}', so it cannot be {action.lower()}d now."

        requester, item, qty = req.cell(row, 3).value, req.cell(row, 4).value, req.cell(row, 5).value
        vendors = load_vendors(wb)
        inv_ws = wb["Inventory"]
        hit = find_stock_full(inv_ws, item)
        moq = hit[4] if hit else 1

        if action == "REJECT":
            req.cell(row, 8).value = "Rejected"
            req.cell(row, 11).value = "Rejected by lab owner"
            wb.save(EXCEL_PATH)
            result = f"Ticket {tid} rejected."
            requester_msg = f"Update {tid}: your request for {item} {qty}pcs was not approved by the lab owner."
        elif action == "TOPUP":
            try:
                topup_qty = int(vendor_arg.split()[-1])
            except Exception:
                return f"Please specify top up quantity, e.g. TOPUP {tid} 51"

            vendor_name = " ".join(vendor_arg.split()[:-1]) if len(vendor_arg.split()) > 1 else None
            v = pick_vendor(vendors, item, vendor_name)
            if v is None:
                v = vendors[0] if vendors else None
            if v is None:
                return "No vendors available."

            poid = "PO-" + tid
            subject, body = draft_po_email(poid, v[0], item, topup_qty)
            ok, err = send_email(v[1], subject, body)
            if not ok:
                return f"Email NOT sent: {err}."

            req.cell(row, 8).value = "PO Sent"
            req.cell(row, 11).value = f"{poid} emailed to {v[0]} (Top-up quantity: {topup_qty})"
            wb["PO"].append([poid, f"{datetime.now():%Y-%m-%d %H:%M}", tid, item, topup_qty, v[0], v[1], "Sent", "Owner top-up enforced"])
            wb.save(EXCEL_PATH)
            result = f"Approved with Top-up. {poid} (Top-up Qty: {topup_qty}) emailed to {v[0]} ({v[1]})."
            requester_msg = f"Update {tid}: {item} {qty}pcs is being purchased (PO top-up sent to vendor). We will notify you as batches arrive."
        else: # APPROVE
            v = pick_vendor(vendors, item, vendor_arg)
            if v is None:
                names = ", ".join(x[0] for x in vendors) or "(none)"
                return f"Vendor '{vendor_arg}' not found. Vendors available: {names}"

            po_qty = moq
            poid = "PO-" + tid
            subject, body = draft_po_email(poid, v[0], item, po_qty)
            ok, err = send_email(v[1], subject, body)
            if not ok:
                return f"Email NOT sent: {err}. Ticket {tid} stays pending."

            req.cell(row, 8).value = "PO Sent"
            req.cell(row, 11).value = f"{poid} emailed to {v[0]} (Strict MOQ qty: {po_qty})"
            wb["PO"].append([poid, f"{datetime.now():%Y-%m-%d %H:%M}", tid, item, po_qty, v[0], v[1], "Sent", "Strict MOQ enforced"])
            wb.save(EXCEL_PATH)
            result = f"Approved. {poid} (Strict MOQ Qty: {po_qty}) emailed to {v[0]} ({v[1]})."
            requester_msg = f"Update {tid}: {item} {qty}pcs is being purchased (PO sent to vendor). We will notify you as batches arrive."

    phone = phone_from_requester(requester)
    if phone:
        wa_send(phone, requester_msg)
    return result


def stock_report(filter_arg=None):
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        rows = list(wb["Inventory"].iter_rows(min_row=2, values_only=True))
    if not rows:
        return "Inventory is empty."

    filtered_rows = []
    for r in rows:
        if not r[0]:
            continue
        loc = str(r[5] or "Singapore Lab")
        cat = str(r[6] or "General")
        if filter_arg:
            f_clean = filter_arg.lower().strip()
            if f_clean in loc.lower() or f_clean in cat.lower():
                filtered_rows.append(r)
        else:
            filtered_rows.append(r)

    if not filtered_rows:
        return f"No inventory items found matching location/category filter '{filter_arg}'."

    filter_title = f" (Filtered: {filter_arg})" if filter_arg else ""
    return f"Inventory{filter_title}:\n" + "\n".join(
        f"- {r[0]}: {r[1]} {r[2]} [Loc: {r[5] or 'Singapore Lab'} | Cat: {r[6] or 'General'}] (MinQty: {r[3]}, MOQ: {r[4] if len(r) > 4 else 1})"
        for r in filtered_rows
    )


def low_stock_report(filter_arg=None):
    """Return list of items where Qty < MinQty (strictly below MinQty)."""
    alerts = []
    with lock:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        ws = wb["Inventory"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            item, qty, unit, min_qty, moq = row[0], row[1], row[2], row[3], (row[4] if len(row) > 4 else 1)
            loc = str(row[5] if len(row) > 5 else "Singapore Lab")
            cat = str(row[6] if len(row) > 6 else "General")

            if filter_arg:
                f_clean = filter_arg.lower().strip()
                if f_clean not in loc.lower() and f_clean not in cat.lower():
                    continue

            if item is not None:
                try:
                    q = int(qty or 0)
                    mq = int(min_qty or 0)
                    if q < mq:  # Strictly below MinQty
                        alerts.append(f"- {item} [{loc} | {cat}]: {q} {unit or 'pcs'} (Min required: {mq}, MOQ: {moq})")
                except ValueError:
                    continue
    if not alerts:
        filter_str = f" for '{filter_arg}'" if filter_arg else ""
        return f"Low-Stock Dashboard:\n  All inventory items{filter_str} are sufficiently stocked! 👍"

    filter_title = f" ({filter_arg})" if filter_arg else ""
    return f"⚠️ Low-Stock Dashboard{filter_title} (Needs Attention):\n" + "\n".join(alerts)


# ---------- web endpoint ----------
@app.route("/incoming", methods=["POST"])
def incoming():
    data = request.get_json(force=True)
    sender = data.get("sender", "unknown")
    text = (data.get("text") or "").strip()
    role = data.get("role", "requester")

    sender_role = get_sender_role(sender)
    if role == "owner" or sender_role in ("owner", "admin"):
        reply = handle_owner_command(text)
        if reply is not None:
            return jsonify({"reply": reply})
        return jsonify({})

    try:
        text_clean = text.lower()
        if text_clean.startswith("stock") or text_clean.startswith("inventory") or text_clean.startswith("/stock"):
            parts = text.split(maxsplit=1)
            filter_arg = parts[1] if len(parts) > 1 else None
            return jsonify({"reply": stock_report(filter_arg)})

        if text_clean.startswith("lowstock") or text_clean.startswith("alerts") or text_clean.startswith("alert"):
            parts = text.split(maxsplit=1)
            filter_arg = parts[1] if len(parts) > 1 else None
            return jsonify({"reply": low_stock_report(filter_arg)})

        if text_clean in ("history", "tickets", "my tickets", "status"):
            return jsonify({"reply": get_requester_history_report(sender)})

        if text_clean in ("command", "commands", "help", "/help", "/command"):
            return jsonify({"reply": requester_command_help()})

        m_status = STATUS_CMD.match(text)
        if m_status:
            return jsonify({"reply": get_ticket_status_report(m_status.group(1), current_sender=sender, is_owner=False)})

        items = parse_simple(text) or parse_gemini(text)
        if not items:
            return jsonify({"reply": "Sorry, I could not understand. Send 'command' to view available commands, 'stock' for inventory, or 'history' for your tickets."})
        return jsonify({"reply": process_request(sender, items)})
    except PermissionError:
        return jsonify({"reply": "The Excel file is open. Please close lab_inventory.xlsx and send again."})


# ---------- endpoints used by bridge.js ----------
def digits(v):
    if isinstance(v, float):
        v = int(v)
    return re.sub(r"\D", "", str(v or ""))


def load_requesters(wb):
    out = []
    for r in wb["Requesters"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[1]:
            out.append({"name": str(r[0]).strip(), "phone": digits(r[1]), "id": digits(r[2])})
    return out


_users_cache = []


@app.route("/users")
def users():
    """The bridge asks who is allowed to send requests."""
    global _users_cache
    try:
        with lock:
            wb = load_workbook(EXCEL_PATH, read_only=True)
            reqs = load_requesters(wb)
            wb.close()
        _users_cache = reqs
    except Exception as e:
        print("Could not read Requesters sheet, using last known list:", e)
        reqs = _users_cache
    return jsonify({"requesters": reqs})


@app.route("/unknown", methods=["POST"])
def unknown():
    """The bridge reports a sender who is not allowed yet, so it can be logged in Excel."""
    d = request.get_json(force=True)
    wid, text = digits(d.get("id")), str(d.get("text") or "")[:100]
    if not wid:
        return jsonify({"ok": False}), 400
    try:
        with lock:
            wb = load_workbook(EXCEL_PATH)
            ws = wb["Unknown"]
            now = f"{datetime.now():%Y-%m-%d %H:%M}"
            for row in ws.iter_rows(min_row=2):
                if digits(row[0].value) == wid:
                    row[1].value, row[2].value = now, text
                    break
            else:
                ws.append([wid, now, text])
            wb.save(EXCEL_PATH)
    except PermissionError:
        return jsonify({"ok": False, "error": "excel file is open"}), 503
    return jsonify({"ok": True})


# ---------- startup vendor menu ----------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def ask_yes_no(question):
    while True:
        ans = input(f"{question} (y/n): ").strip().lower()
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("Please type y or n.")


def add_vendor_to_excel(name, email):
    """Add a vendor row, or update the email if the vendor ID already exists."""
    while True:
        try:
            with lock:
                wb = load_workbook(EXCEL_PATH)
                ws = wb["Vendors"]
                for row in ws.iter_rows(min_row=2):
                    if str(row[0].value or "").strip().lower() == name.lower():
                        row[1].value = email
                        wb.save(EXCEL_PATH)
                        return "updated"
                ws.append([name, email, ""])
                wb.save(EXCEL_PATH)
                return "added"
        except PermissionError:
            input("lab_inventory.xlsx is open in Excel. Close it, then press Enter to retry...")


def get_vendors():
    wb = load_workbook(EXCEL_PATH, read_only=True)
    vendors = load_vendors(wb)
    wb.close()
    return vendors


def remove_vendor_from_excel(name):
    """Delete the vendor row. Returns True if a row was removed."""
    while True:
        try:
            with lock:
                wb = load_workbook(EXCEL_PATH)
                ws = wb["Vendors"]
                for row in ws.iter_rows(min_row=2):
                    if str(row[0].value or "").strip().lower() == name.lower():
                        ws.delete_rows(row[0].row)
                        wb.save(EXCEL_PATH)
                        return True
                return False
        except PermissionError:
            input("lab_inventory.xlsx is open in Excel. Close it, then press Enter to retry...")


def add_vendors_flow():
    while True:
        name = input("   Vendor ID (name): ").strip()
        if not name:
            print("   Vendor ID cannot be empty.")
            continue
        email = input("   Vendor email: ").strip()
        while not EMAIL_RE.match(email):
            print("   That email does not look valid. Please try again.")
            email = input("   Vendor email: ").strip()
        result = add_vendor_to_excel(name, email)
        print(f"   Vendor {name} <{email}> {result}.")
        if not ask_yes_no("   Add another vendor?"):
            return


def remove_vendors_flow():
    while True:
        vendors = get_vendors()
        if not vendors:
            print("   No vendors to remove.")
            return
        for i, v in enumerate(vendors, 1):
            print(f"   {i}. {v[0]} <{v[1]}>")
        choice = input("   Type the number or Vendor ID to remove (Enter to cancel): ").strip()
        if not choice:
            return
        name = None
        if choice.isdigit() and 1 <= int(choice) <= len(vendors):
            name = vendors[int(choice) - 1][0]
        else:
            for v in vendors:
                if v[0].lower() == choice.lower():
                    name = v[0]
        if name is None:
            print("   Vendor not found. Please try again.")
            continue
        warning = " This is your LAST vendor, so no PO can be drafted until you add one." if len(vendors) == 1 else ""
        if not ask_yes_no(f"   Remove {name}?{warning}"):
            continue
        if remove_vendor_from_excel(name):
            print(f"   Vendor {name} removed.")
        if not ask_yes_no("   Remove another vendor?"):
            return


def vendor_menu():
    while True:
        vendors = get_vendors()
        print("\nCurrent vendors:")
        for v in vendors:
            print(f"  - {v[0]} <{v[1]}>")
        if not vendors:
            print("  (none)")
        print("\nVendor menu:\n  1. Add vendor\n  2. Remove vendor\n  3. Back")
        choice = input("Choose 1, 2 or 3 (Enter = 3): ").strip()
        if choice in ("", "3"):
            return
        if choice == "1":
            add_vendors_flow()
        elif choice == "2":
            remove_vendors_flow()
        else:
            print("Please type 1, 2 or 3.")


# ---------- requester menu ----------
def edit_workbook(fn):
    """Run fn(wb) under the lock, save, and retry if Excel has the file open."""
    while True:
        try:
            with lock:
                wb = load_workbook(EXCEL_PATH)
                result = fn(wb)
                wb.save(EXCEL_PATH)
                return result
        except PermissionError:
            input("lab_inventory.xlsx is open in Excel. Close it, then press Enter to retry...")


def get_requesters():
    wb = load_workbook(EXCEL_PATH, read_only=True)
    out = []
    for r in wb["Requesters"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[1]:
            out.append({"name": str(r[0]).strip(), "phone": digits(r[1]), "id": digits(r[2]), "role": str(r[3] or "requester").strip().lower()})
    wb.close()
    return out


def get_unknown():
    wb = load_workbook(EXCEL_PATH, read_only=True)
    out = [r for r in wb["Unknown"].iter_rows(min_row=2, values_only=True) if r[0]]
    wb.close()
    return out


def show_requesters(reqs):
    for i, r in enumerate(reqs, 1):
        wid = r["id"] or "not linked yet"
        print(f"   {i}. {r['name']} (+{r['phone']}) | Role: {r['role']} | WhatsApp ID: {wid}")


def add_requester_flow():
    while True:
        name = input("   Requester name/ID: ").strip()
        if not name:
            print("   Name cannot be empty.")
            continue
        phone = digits(input("   Requester phone (country code, e.g. 60123456789): "))
        while not 8 <= len(phone) <= 15:
            print("   Phone must be 8-15 digits with country code. Please try again.")
            phone = digits(input("   Requester phone: "))
        wid = digits(input("   WhatsApp long ID (press Enter to skip, link it later with option 3): "))
        role_input = input("   Role (requester / owner) [default: requester]: ").strip().lower()
        if role_input not in ("owner", "admin"):
            role_input = "requester"

        def fn(wb):
            ws = wb["Requesters"]
            for row in ws.iter_rows(min_row=2):
                if digits(row[1].value) == phone:
                    row[0].value = name
                    if wid:
                        row[2].value = wid
                    row[3].value = role_input
                    return "updated"
            ws.append([name, phone, wid or None, role_input])
            return "added"

        print(f"   Requester {name} (+{phone}) {edit_workbook(fn)}.")
        if not ask_yes_no("   Add another requester?"):
            return


def remove_requester_flow():
    while True:
        reqs = get_requesters()
        if not reqs:
            print("   No requesters to remove.")
            return
        show_requesters(reqs)
        choice = input("   Type the number, name or phone to remove (Enter to cancel): ").strip()
        if not choice:
            return
        target = None
        if choice.isdigit() and 1 <= int(choice) <= len(reqs):
            target = reqs[int(choice) - 1]
        else:
            for r in reqs:
                if r["name"].lower() == choice.lower() or r["phone"] == digits(choice):
                    target = r
        if target is None:
            print("   Requester not found. Please try again.")
            continue
        if not ask_yes_no(f"   Remove {target['name']} (+{target['phone']})?"):
            continue

        def fn(wb):
            ws = wb["Requesters"]
            for row in wb["Requesters"].iter_rows(min_row=2):
                if digits(row[1].value) == target["phone"]:
                    ws.delete_rows(row[0].row)
                    return True
            return False

        if edit_workbook(fn):
            print(f"   Requester {target['name']} removed.")
        if not ask_yes_no("   Remove another requester?"):
            return


def link_requester_flow():
    """Attach a WhatsApp long ID (seen when an unknown person messaged) to a requester."""
    while True:
        unknown = get_unknown()
        reqs = get_requesters()
        if not unknown:
            print("   No unknown WhatsApp IDs yet. With agent and bridge running, ask the person to send any")
            print("   message to the linked number, then restart the agent and come back here.")
            return
        if not reqs:
            print("   Add the requester first (option 1), then link the ID.")
            return
        print("   Unknown senders:")
        for i, u in enumerate(unknown, 1):
            print(f"   {i}. ID {u[0]} | last seen {u[1]} | message: {u[2]}")
        pick = input("   Type the number of the sender to link (Enter to cancel): ").strip()
        if not pick:
            return
        if not (pick.isdigit() and 1 <= int(pick) <= len(unknown)):
            print("   Please type a number from the list.")
            continue
        wid = digits(unknown[int(pick) - 1][0])
        show_requesters(reqs)
        who = input("   Link to which requester? Type the number: ").strip()
        if not (who.isdigit() and 1 <= int(who) <= len(reqs)):
            print("   Please type a number from the list.")
            continue
        target = reqs[int(who) - 1]

        def fn(wb):
            for row in wb["Requesters"].iter_rows(min_row=2):
                if digits(row[1].value) == target["phone"]:
                    row[2].value = wid
            ws = wb["Unknown"]
            for row in wb["Unknown"].iter_rows(min_row=2):
                if digits(row[0].value) == wid:
                    ws.delete_rows(row[0].row)
                    break

        edit_workbook(fn)
        print(f"   Linked ID {wid} to {target['name']}. They can send requests now.")
        if not ask_yes_no("   Link another?"):
            return


def requester_menu():
    while True:
        print("\nCurrent requesters:")
        reqs = get_requesters()
        show_requesters(reqs)
        if not reqs:
            print("  (none)")
        print("\nRequester menu:\n  1. Add requester\n  2. Remove requester\n  3. Link WhatsApp ID to a requester\n  4. Back")
        choice = input("Choose 1, 2, 3 or 4 (Enter = 4): ").strip()
        if choice in ("", "4"):
            return
        if choice == "1":
            add_requester_flow()
        elif choice == "2":
            remove_requester_flow()
        elif choice == "3":
            link_requester_flow()
        else:
            print("Please type 1, 2, 3 or 4.")


def startup_menu():
    """Start-up menu: manage vendors and requesters, or start the agent."""
    print("\n" + "=" * 40)
    print(" Phase 3: Low-Stock Dashboard Audit")
    print("=" * 40)
    print(low_stock_report())
    print("=" * 40 + "\n")

    try:
        while True:
            print("Main menu:\n  1. Vendors (add / remove)\n  2. Requesters (add / remove / link WhatsApp ID)\n  3. Start agent")
            choice = input("Choose 1, 2 or 3 (Enter = 3): ").strip()
            if choice in ("", "3"):
                return
            if choice == "1":
                vendor_menu()
            elif choice == "2":
                requester_menu()
            else:
                print("Please type 1, 2 or 3.")
    except (EOFError, KeyboardInterrupt):
        print("\nSkipping setup menu.")


if __name__ == "__main__":
    init_workbook()
    startup_menu()
    app.run(host="127.0.0.1", port=5000)
