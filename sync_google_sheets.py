import os
import sys
import json
import csv
import urllib.request
import argparse
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "invitation_contacts_18_oct.json")
EXCEL_FILE = os.path.join(BASE_DIR, "Maitry_and_Smit_Invitation_List_18_Oct.xlsx")
CSV_EXPORT_FILE = os.path.join(BASE_DIR, "sync_to_google_sheet.csv")

SHEET_ID = "1Swo30hHUcKIUFQycJjUnn6ocMlWz5P2HiyQvM2UBnp8"
SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

def pull_from_google_sheet():
    print(f"Fetching latest data from Google Sheet...")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")

    req = urllib.request.Request(
        SHEET_CSV_URL,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read().decode("utf-8")
    except Exception as e:
        print(f"[Error] Failed to fetch Google Sheet: {e}")
        return False

    lines = content.splitlines()
    if not lines:
        print("[Error] Google Sheet returned empty response.")
        return False

    reader = csv.DictReader(lines)
    sheet_rows = list(reader)
    print(f"[OK] Downloaded {len(sheet_rows)} contact rows from Google Sheet.")

    # Load local JSON
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            local_contacts = json.load(f)
    else:
        local_contacts = []

    local_map = {str(c.get("Sr. No.")): c for c in local_contacts}

    updated_contacts = []
    added_count = 0
    modified_count = 0

    for r in sheet_rows:
        sr_raw = str(r.get("Sr. No.", "")).strip()
        if not sr_raw:
            continue

        try:
            sr_num = int(sr_raw)
        except ValueError:
            sr_num = sr_raw

        name = (r.get("Name") or "").strip()
        p1 = (r.get("Phone 1") or "").strip().lstrip("'")
        p2 = (r.get("Phone 2") or "").strip().lstrip("'")
        cat = (r.get("Category / Note") or "Sajode").strip().capitalize()
        if cat not in ["Sajode", "Sarvo"]:
            cat = "Sajode"

        # Check existing contact
        existing = local_map.get(str(sr_num))

        if existing:
            # Detect changes
            changed = False
            if existing.get("Name") != name:
                existing["Name"] = name
                changed = True
            if existing.get("Phone 1") != p1:
                existing["Phone 1"] = p1
                changed = True
            if existing.get("Phone 2") != p2:
                existing["Phone 2"] = p2
                changed = True
            if existing.get("Category / Note") != cat:
                existing["Category / Note"] = cat
                changed = True

            # Tracking fields from sheet if present
            if "Invitation Status" in r and r.get("Invitation Status"):
                existing["Invitation Status"] = r.get("Invitation Status").strip()
            if "Sent From (Sender)" in r and r.get("Sent From (Sender)"):
                existing["Sent From (Sender)"] = r.get("Sent From (Sender)").strip()
            if "PDF Sent" in r and r.get("PDF Sent"):
                existing["PDF Sent"] = r.get("PDF Sent").strip()
            if "Sent Timestamp" in r and r.get("Sent Timestamp"):
                existing["Sent Timestamp"] = r.get("Sent Timestamp").strip()
            if "Recipient Number" in r and r.get("Recipient Number"):
                existing["Recipient Number"] = r.get("Recipient Number").strip()
            if "વોટ્સએપ OK" in r:
                existing["વોટ્સએપ OK"] = (r.get("વોટ્સએપ OK") or "").strip()
            if "ફોન પર વાતચીત" in r:
                existing["ફોન પર વાતચીત"] = (r.get("ફોન પર વાતચીત") or "").strip()
            if "Notes" in r and r.get("Notes"):
                existing["Notes"] = (r.get("Notes") or "").strip()

            if changed:
                modified_count += 1
            updated_contacts.append(existing)
        else:
            # New contact added from Google Sheet
            new_contact = {
                "Sr. No.": sr_num,
                "Name": name,
                "Phone 1": p1,
                "Phone 2": p2,
                "Category / Note": cat,
                "WhatsApp Phone 1": f"+91{p1.replace(' ', '')}" if p1 and not p1.startswith('+') else p1,
                "WhatsApp Phone 2": "",
                "Invitation Status": (r.get("Invitation Status") or "Pending").strip(),
                "Delivery Method": "WhatsApp",
                "Notes": (r.get("Notes") or "").strip(),
                "Sent From (Sender)": (r.get("Sent From (Sender)") or "").strip(),
                "PDF Sent": (r.get("PDF Sent") or "").strip(),
                "Sent Timestamp": (r.get("Sent Timestamp") or "").strip(),
                "Recipient Number": (r.get("Recipient Number") or "").strip(),
                "વોટ્સએપ OK": (r.get("વોટ્સએપ OK") or "").strip(),
                "ફોન પર વાતચીત": (r.get("ફોન પર વાતચીત") or "").strip()
            }
            updated_contacts.append(new_contact)
            added_count += 1

    # Save to local JSON
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(updated_contacts, f, indent=2, ensure_ascii=False)
    print(f"[OK] Saved {len(updated_contacts)} contacts to {os.path.basename(JSON_FILE)}")
    if added_count > 0 or modified_count > 0:
        print(f"     Changes: {added_count} new contacts added, {modified_count} contacts updated.")

    # Re-save Excel, index.html, and compile A4 Portrait PDF
    try:
        from send_invitations import save_state
        save_state(updated_contacts)
    except Exception as e:
        print(f"[Warning] Failed to update Excel/site/PDF: {e}")

    print("[SUCCESS] Full two-way sync from Google Sheet completed successfully!")
    return True

def push_to_csv(contacts=None):
    if contacts is None:
        if not os.path.exists(JSON_FILE):
            print(f"Error: {JSON_FILE} not found!")
            return False
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

    with open(CSV_EXPORT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for c in contacts:
            row = {
                "Sr. No.": c.get("Sr. No.", ""),
                "Name": c.get("Name", ""),
                "Phone 1": f"'{c.get('Phone 1', '')}" if str(c.get('Phone 1', '')).startswith("+") else c.get("Phone 1", ""),
                "Phone 2": c.get("Phone 2", ""),
                "Category / Note": c.get("Category / Note", "Sajode"),
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

    print(f"[OK] Exported {len(contacts)} contacts to {os.path.basename(CSV_EXPORT_FILE)}")
    return True

def main():
    parser = argparse.ArgumentParser(description="Google Sheets Live Sync Engine for Maitry & Smit Invitations")
    parser.add_argument("--pull", action="store_true", help="Pull latest contacts from Google Sheet into local JSON and rebuild site/PDF")
    parser.add_argument("--export", action="store_true", help="Export local contacts to CSV for Google Sheet")
    args = parser.parse_args()

    if args.pull:
        pull_from_google_sheet()
    elif args.export:
        push_to_csv()
    else:
        # Default action: pull latest
        pull_from_google_sheet()

if __name__ == "__main__":
    main()
