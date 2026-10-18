import os
import sys
import json
import time
import random
import re
import subprocess
import argparse
from datetime import datetime

# Ensure UTF-8 output on Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "invitation_contacts_18_oct.json")
EXCEL_FILE = os.path.join(BASE_DIR, "Maitry_and_Smit_Invitation_List_18_Oct.xlsx")
SESSIONS_DIR = os.path.join(BASE_DIR, "sessions")

PDF_MAP = {
    "Sajode": os.path.join(BASE_DIR, "માંગુકિયા પરિવાર સજોડે.pdf"),
    "Sarvo": os.path.join(BASE_DIR, "માંગુકિયા પરિવાર સર્વો.pdf")
}

DEFAULT_ACCOUNTS = {
    "account1": {
        "name": "account1",
        "label": "9879881172 (Maheshbhai Mangukiya)",
        "session_dir": os.path.join(SESSIONS_DIR, "account1")
    },
    "account2": {
        "name": "account2",
        "label": "6351838451 (Varshaben Mangukiya)",
        "session_dir": os.path.join(SESSIONS_DIR, "account2")
    },
    "account3": {
        "name": "account3",
        "label": "Account 3",
        "session_dir": os.path.join(SESSIONS_DIR, "account3")
    }
}

def clean_recipient_number(phone_str):
    if not phone_str:
        return ""
    # strip quotes and non-digit except '+'
    s = str(phone_str).strip().strip('"').strip("'")
    digits = re.sub(r'\D', '', s)
    if not digits:
        return ""
    # UAE: +971...
    if s.startswith("+97") or (digits.startswith("971") and len(digits) >= 11):
        if not digits.startswith("971") and digits.startswith("97"):
            # e.g. 97 1585526281 -> 971585526281
            digits = "971" + digits[2:]
        return digits
    # USA: +1...
    if s.startswith("+1") or (len(digits) == 11 and digits.startswith("1")):
        return digits
    if len(digits) == 10:
        return f"91{digits}"
    if len(digits) == 12 and digits.startswith("91"):
        return digits
    return digits

NPX_CMD = "npx.cmd" if sys.platform == "win32" else "npx"

def check_account_status(session_dir):
    """Checks if mudslide session is logged in by running `mudslide -c <dir> me`."""
    if not os.path.exists(session_dir):
        return False, "Not initialized"
    cmd = [NPX_CMD, "mudslide", "-c", session_dir, "me"]
    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=20
        )
        output = proc.stdout.strip()
        if "Current user:" in output:
            for line in output.splitlines():
                if "Current user:" in line:
                    user_jid = line.split("Current user:")[1].strip()
                    return True, user_jid
            return True, "Logged in"
        if proc.returncode == 0:
            return True, output if output else "Logged in"
        return False, proc.stderr.strip() or "Not logged in"
    except Exception as e:
        return False, str(e)

def get_active_accounts():
    active = []
    for acc_id, acc_info in DEFAULT_ACCOUNTS.items():
        is_ok, details = check_account_status(acc_info["session_dir"])
        if is_ok:
            # check if details contain a phone number or name to augment label
            label = acc_info["label"]
            if acc_id != "account1" and details and details != "Logged in":
                label = f"{acc_id} ({details})"
            active.append({
                "id": acc_id,
                "label": label,
                "session_dir": acc_info["session_dir"]
            })
    return active

def save_state(contacts):
    # 1. Save JSON
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(contacts, f, indent=2, ensure_ascii=False)

    # 2. Update index.html embedded data
    html_file = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(html_file):
        try:
            with open(html_file, "r", encoding="utf-8") as f:
                content = f.read()
            pattern = r"let contactsData = \[.*?\];"
            replacement = "let contactsData = " + json.dumps(contacts, ensure_ascii=False) + ";"
            new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(new_content)
        except Exception:
            pass
    
    # 3. Save Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invitation Contacts"
    
    total = len(contacts)
    sent_count = sum(1 for c in contacts if c.get("Invitation Status") == "Sent")
    pending_count = total - sent_count
    
    # Title Block
    ws.merge_cells("A1:K1")
    ws["A1"] = "MAITRY & SMIT WEDDING INVITATION AUTOMATION TRACKER"
    ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:K2")
    ws["A2"] = f"Total Contacts: {total}  |  Sent: {sent_count}  |  Pending: {pending_count}  |  Updated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="DCE7F5")
    ws["A2"].fill = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    headers = [
        "Sr. No.",
        "Name",
        "Category / Note",
        "Phone 1",
        "Phone 2",
        "Recipient Number",
        "Invitation Status",
        "Sent From (Sender)",
        "PDF Sent",
        "Sent Timestamp",
        "Notes"
    ]
    ws.row_dimensions[3].height = 28
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E0'),
        right=Side(style='thin', color='CBD5E0'),
        top=Side(style='thin', color='CBD5E0'),
        bottom=Side(style='thin', color='CBD5E0')
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    sent_fill = PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid") # subtle mint green
    fail_fill = PatternFill(start_color="FFF5F5", end_color="FFF5F5", fill_type="solid") # subtle red
    alt_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    text_font = Font(name="Calibri", size=10)
    bold_green = Font(name="Calibri", size=10, bold=True, color="234E52")
    bold_red = Font(name="Calibri", size=10, bold=True, color="9B2C2C")

    for row_idx, r in enumerate(contacts, 4):
        status = r.get("Invitation Status", "Pending")
        if status == "Sent":
            fill = sent_fill
        elif status == "Failed":
            fill = fail_fill
        else:
            fill = alt_fill if row_idx % 2 == 0 else white_fill

        row_data = [
            r.get("Sr. No."),
            r.get("Name"),
            r.get("Category / Note"),
            r.get("Phone 1"),
            r.get("Phone 2"),
            r.get("Recipient Number", clean_recipient_number(r.get("Phone 1"))),
            status,
            r.get("Sent From (Sender)", ""),
            r.get("PDF Sent", ""),
            r.get("Sent Timestamp", ""),
            r.get("Notes", "")
        ]

        ws.row_dimensions[row_idx].height = 20
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = text_font
            cell.fill = fill
            cell.border = thin_border
            
            if col_idx in [1, 3, 7, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [4, 5, 6, 8]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '@'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if col_idx == 7: # Status column
                if status == "Sent":
                    cell.font = bold_green
                elif status == "Failed":
                    cell.font = bold_red

    # Set column widths
    widths = [10, 32, 16, 16, 16, 20, 18, 30, 26, 22, 25]
    for idx, w in enumerate(widths, 1):
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = w

    ws.views.sheetView[0].showGridLines = True
    wb.save(EXCEL_FILE)

def login_account(acc_id):
    if acc_id not in DEFAULT_ACCOUNTS:
        print(f"Error: Unknown account '{acc_id}'. Available: {list(DEFAULT_ACCOUNTS.keys())}")
        return
    session_dir = DEFAULT_ACCOUNTS[acc_id]["session_dir"]
    os.makedirs(session_dir, exist_ok=True)
    label = DEFAULT_ACCOUNTS[acc_id]["label"]
    print(f"\n=======================================================")
    print(f" Logging in for: {label}")
    print(f" Session Directory: {session_dir}")
    print(f" Scan the QR code below on your WhatsApp mobile app:")
    print(f"=======================================================\n")
    cmd = [NPX_CMD, "mudslide", "-c", session_dir, "login"]
    subprocess.run(cmd)

def send_invitation_pdf(session_dir, recipient_number, pdf_path):
    """Sends directly the PDF via Mudslide with NO caption text."""
    # Command: npx mudslide -c <dir> send-file <recipient> <file>
    cmd = [
        NPX_CMD, "mudslide",
        "-c", session_dir,
        "send-file",
        recipient_number,
        pdf_path
    ]
    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=60
        )
        if proc.returncode == 0:
            return True, proc.stdout.strip()
        else:
            err = proc.stderr.strip() or proc.stdout.strip()
            return False, err
    except Exception as e:
        return False, str(e)

def main():
    parser = argparse.ArgumentParser(description="Maitry & Smit WhatsApp Invitation Automator")
    parser.add_argument("--login", type=str, help="Login an account by scanning QR code (e.g. account1, account2, account3)")
    parser.add_argument("--check", action="store_true", help="Check login status of all accounts")
    parser.add_argument("--dry-run", action="store_true", help="Simulate distribution without sending messages")
    parser.add_argument("--send", action="store_true", help="Start sending PDF invitations")
    parser.add_argument("--account", type=str, help="Only send using a specific account (e.g. account1)")
    parser.add_argument("--delay-min", type=int, default=6, help="Minimum delay between sends in seconds (default: 6)")
    parser.add_argument("--delay-max", type=int, default=10, help="Maximum delay between sends in seconds (default: 10)")
    parser.add_argument("--batch-size", type=int, default=0, help="Limit number of contacts to send in this run (0 = all)")
    parser.add_argument("--start-sr", type=int, default=0, help="Start Sr. No. to send (e.g. 1)")
    parser.add_argument("--end-sr", type=int, default=0, help="End Sr. No. to send (e.g. 30)")
    args = parser.parse_args()

    # 1. Handle Login
    if args.login:
        login_account(args.login)
        return

    # 2. Handle Check
    if args.check:
        print("\n=== WhatsApp Accounts Status ===")
        for acc_id, acc in DEFAULT_ACCOUNTS.items():
            is_ok, details = check_account_status(acc["session_dir"])
            status_text = "LOGGED IN" if is_ok else "NOT LOGGED IN"
            print(f"[{status_text}] {acc_id}: {acc['label']}")
            if is_ok:
                print(f"       Details: {details}")
        print("================================\n")
        return

    # Verify PDFs exist
    for cat, p_path in PDF_MAP.items():
        if not os.path.exists(p_path):
            print(f"Error: PDF for category '{cat}' not found at: {p_path}")
            return

    # Load contacts
    if not os.path.exists(JSON_FILE):
        print(f"Error: {JSON_FILE} not found!")
        return
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        contacts = json.load(f)

    # Get active accounts
    all_active = get_active_accounts()
    if args.account:
        active_accounts = [a for a in all_active if a["id"] == args.account]
        if not active_accounts:
            print(f"Error: Specified account '{args.account}' is not logged in! Run with --login {args.account} first.")
            return
    else:
        active_accounts = all_active

    print(f"\nTotal Contacts in List: {len(contacts)}")
    pending_contacts = [c for c in contacts if c.get("Invitation Status") != "Sent"]
    print(f"Pending Contacts to Send: {len(pending_contacts)}")
    print(f"Active Sender Accounts: {len(active_accounts)}")
    for a in active_accounts:
        print(f"  • {a['id']} -> {a['label']}")

    if not active_accounts and not args.dry_run:
        print("\nNo WhatsApp accounts are currently logged in!")
        print("To log in Maheshbhai's account (9879881182), run:")
        print("   python send_invitations.py --login account1\n")
        return

    if args.dry_run:
        print("\n================ DRY RUN SIMULATION ================")
        if not active_accounts:
            # use account1 as dummy for dry run preview
            active_accounts = [DEFAULT_ACCOUNTS["account1"]]
            print("(Simulating with Account 1 as active)")

        if args.start_sr > 0 and args.end_sr > 0:
            to_process = [c for c in contacts if args.start_sr <= c.get("Sr. No.", 0) <= args.end_sr and c.get("Invitation Status") != "Sent"]
            print(f"Targeting range: Sr. {args.start_sr} to Sr. {args.end_sr} (Total: {len(to_process)})")
        elif args.start_sr > 0:
            to_process = [c for c in contacts if c.get("Sr. No.", 0) >= args.start_sr and c.get("Invitation Status") != "Sent"]
            print(f"Targeting from Sr. {args.start_sr} onwards (Total: {len(to_process)})")
        elif args.end_sr > 0:
            to_process = [c for c in contacts if c.get("Sr. No.", 0) <= args.end_sr and c.get("Invitation Status") != "Sent"]
            print(f"Targeting up to Sr. {args.end_sr} (Total: {len(to_process)})")
        else:
            limit = args.batch_size if args.batch_size > 0 else len(pending_contacts)
            to_process = pending_contacts[:limit]
        
        for idx, contact in enumerate(to_process):
            sender = active_accounts[idx % len(active_accounts)]
            cat = contact.get("Category / Note", "Sajode")
            pdf_file = os.path.basename(PDF_MAP.get(cat, PDF_MAP["Sajode"]))
            recipient = clean_recipient_number(contact.get("Phone 1"))
            print(f"[{idx+1}/{len(to_process)}] Contact #{contact['Sr. No.']}: {contact['Name']} ({cat})")
            print(f"   Recipient: {recipient}")
            print(f"   Sender:    {sender['label']}")
            print(f"   PDF File:  {pdf_file} (Direct PDF, No Text)")
            print("-" * 50)
        print(f"Dry run complete for {len(to_process)} contacts.")
        print("To actually send, make sure accounts are logged in and run with --send.")
        return

    if not args.send:
        print("\nPlease specify --send to begin sending, or --dry-run to preview.")
        print("Example: python send_invitations.py --send")
        return

    # Actual Sending Process
    if args.start_sr > 0 and args.end_sr > 0:
        to_process = [c for c in contacts if args.start_sr <= c.get("Sr. No.", 0) <= args.end_sr and c.get("Invitation Status") != "Sent"]
        print(f"Targeting range: Sr. {args.start_sr} to Sr. {args.end_sr} (Total: {len(to_process)})")
    elif args.start_sr > 0:
        to_process = [c for c in contacts if c.get("Sr. No.", 0) >= args.start_sr and c.get("Invitation Status") != "Sent"]
        print(f"Targeting from Sr. {args.start_sr} onwards (Total: {len(to_process)})")
    elif args.end_sr > 0:
        to_process = [c for c in contacts if c.get("Sr. No.", 0) <= args.end_sr and c.get("Invitation Status") != "Sent"]
        print(f"Targeting up to Sr. {args.end_sr} (Total: {len(to_process)})")
    else:
        limit = args.batch_size if args.batch_size > 0 else len(pending_contacts)
        to_process = pending_contacts[:limit]

    print(f"\nStarting live dispatch for {len(to_process)} contacts...")
    print("Direct PDF only (no text captions) as requested.")
    print("Press Ctrl+C at any time to pause safely.\n")

    sent_this_session = 0
    failed_this_session = 0

    try:
        for idx, contact in enumerate(to_process):
            sender = active_accounts[idx % len(active_accounts)]
            cat = contact.get("Category / Note", "Sajode")
            pdf_path = PDF_MAP.get(cat, PDF_MAP["Sajode"])
            pdf_name = os.path.basename(pdf_path)
            recipient = clean_recipient_number(contact.get("Phone 1"))

            if not recipient:
                print(f"[{idx+1}/{len(to_process)}] SKIPPING #{contact['Sr. No.']} {contact['Name']} - Missing Phone Number!")
                contact["Invitation Status"] = "Failed"
                contact["Notes"] = "Missing phone number"
                save_state(contacts)
                continue

            print(f"[{idx+1}/{len(to_process)}] Sending to #{contact['Sr. No.']} {contact['Name']} ({recipient})...")
            print(f"   Sender: {sender['label']} | Category: {cat} | PDF: {pdf_name}")

            success, output = send_invitation_pdf(sender["session_dir"], recipient, pdf_path)
            timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if success:
                print(f"   [SUCCESS] Delivered from {sender['label']}")
                contact["Invitation Status"] = "Sent"
                contact["Sent From (Sender)"] = sender["label"]
                contact["PDF Sent"] = pdf_name
                contact["Sent Timestamp"] = timestamp_str
                contact["Recipient Number"] = recipient
                contact["Notes"] = ""
                sent_this_session += 1
            else:
                print(f"   [FAILED] Error: {output}")
                contact["Invitation Status"] = "Failed"
                contact["Sent From (Sender)"] = sender["label"]
                contact["Notes"] = f"Error: {output[:100]}"
                failed_this_session += 1

            # Save state immediately to JSON and Excel after each contact
            save_state(contacts)

            # Anti-ban sleep before next message
            if idx < len(to_process) - 1:
                sleep_sec = random.uniform(args.delay_min, args.delay_max)
                print(f"   Pacing delay: waiting {sleep_sec:.1f}s before next contact...\n")
                time.sleep(sleep_sec)

    except KeyboardInterrupt:
        print("\n[PAUSED] Process interrupted by user. All progress has been saved!")
    
    print(f"\n================ SUMMARY ================")
    print(f"Sent in this run:   {sent_this_session}")
    print(f"Failed in this run: {failed_this_session}")
    total_sent = sum(1 for c in contacts if c.get("Invitation Status") == "Sent")
    print(f"Total Sent so far:  {total_sent} / {len(contacts)}")
    print(f"Files updated: {os.path.basename(JSON_FILE)} & {os.path.basename(EXCEL_FILE)}")
    print("=========================================\n")

if __name__ == "__main__":
    main()
