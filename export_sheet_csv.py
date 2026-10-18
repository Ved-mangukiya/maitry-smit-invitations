import json
import csv
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "invitation_contacts_18_oct.json")
CSV_OUT = os.path.join(BASE_DIR, "sync_to_google_sheet.csv")

with open(JSON_FILE, "r", encoding="utf-8") as f:
    contacts = json.load(f)

headers = [
    "Sr. No.",
    "Name",
    "Phone 1",
    "Phone 2",
    "Category / Note",
    "Invitation Status",
    "Sent From (Sender)",
    "PDF Sent",
    "Sent Timestamp",
    "Recipient Number",
    "વોટ્સએપ OK",
    "ફોન પર વાતચીત",
    "Notes"
]

with open(CSV_OUT, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=headers)
    writer.writeheader()
    for c in contacts:
        cat = c.get("Category / Note", "Sajode")
        row = {
            "Sr. No.": c.get("Sr. No.", ""),
            "Name": c.get("Name", ""),
            "Phone 1": f"'{c.get('Phone 1', '')}" if str(c.get('Phone 1', '')).startswith("+") else c.get("Phone 1", ""),
            "Phone 2": c.get("Phone 2", ""),
            "Category / Note": cat,
            "Invitation Status": c.get("Invitation Status", "Pending"),
            "Sent From (Sender)": c.get("Sent From (Sender)", ""),
            "PDF Sent": c.get("PDF Sent", ""),
            "Sent Timestamp": c.get("Sent Timestamp", ""),
            "Recipient Number": c.get("Recipient Number", ""),
            "વોટ્સએપ OK": c.get("વોટ્સએપ OK", ""),
            "ફોન પર વાતચીત": c.get("ફોન પર વાતચીત", ""),
            "Notes": c.get("Notes", "")
        }
        writer.writerow(row)

print(f"[OK] Generated {CSV_OUT} with {len(contacts)} contacts and {len(headers)} columns.")
