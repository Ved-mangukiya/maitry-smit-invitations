import os
import sys
import json
import subprocess
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "invitation_contacts_18_oct.json")
HTML_FILE = os.path.join(BASE_DIR, "index.html")
PDF_FILE = os.path.join(BASE_DIR, "Maitry_and_Smit_Invitations_A4_Portrait.pdf")
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def build_site_and_pdf(contacts=None):
    if contacts is None:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            contacts = json.load(f)

    total = len(contacts)
    sent_list = [c for c in contacts if c.get("Invitation Status") == "Sent"]
    pending_list = [c for c in contacts if c.get("Invitation Status") != "Sent"]
    sent_count = len(sent_list)
    pending_count = len(pending_list)
    sajode_count = sum(1 for c in contacts if c.get("Category / Note") == "Sajode")
    sarvo_count = sum(1 for c in contacts if c.get("Category / Note") == "Sarvo")
    mahesh_count = sum(1 for c in contacts if "9879881172" in (c.get("Sent From (Sender)") or ""))
    varsha_count = sum(1 for c in contacts if "6351838451" in (c.get("Sent From (Sender)") or ""))
    pct = round((sent_count / total) * 100) if total > 0 else 0

    # Build Pre-rendered Table Rows
    table_rows_html = []
    for c in contacts:
        sr = c.get("Sr. No.", "")
        name = c.get("Name", "")
        cat = c.get("Category / Note", "Sajode")
        phone = c.get("Phone 1") or c.get("Recipient Number") or "—"
        is_sent = c.get("Invitation Status") == "Sent"
        sender_raw = c.get("Sent From (Sender)") or ""
        pdf_sent = (c.get("PDF Sent") or "—").replace(".pdf", "")

        cat_badge = '<span class="badge b-sarvo">સર્વો</span>' if cat == 'Sarvo' else '<span class="badge b-sajode">સજોડે</span>'

        if is_sent:
            status_badge = '<span class="badge b-sent">✅ મોકલેલ</span>'
            if "9879881172" in sender_raw:
                sender_badge = '<span class="badge b-mahesh">મહેશભાઈ</span>'
            elif "6351838451" in sender_raw:
                sender_badge = '<span class="badge b-varsha">વર્ષાબેન</span>'
            else:
                sender_badge = '<span class="badge b-sent">મોકલેલ</span>'
        else:
            status_badge = '<span class="badge b-pending">⏳ બાકી</span>'
            sender_badge = '—'

        wa_ok = c.get("વોટ્સએપ OK", "")
        wa_mark = "✔" if wa_ok else ""
        call_ok = c.get("ફોન પર વાતચીત", "")
        call_mark = "✔" if call_ok else ""

        row = f"""            <tr>
              <td class="td-sr"><b>#{sr}</b></td>
              <td class="td-name">{name}</td>
              <td class="td-cat col-center">{cat_badge}</td>
              <td class="td-phone col-center">{phone}</td>
              <td class="td-status col-center">{status_badge}</td>
              <td class="td-sender col-center">{sender_badge}</td>
              <td class="td-pdf" style="font-size:0.75rem; color:#475569;">{pdf_sent}</td>
              <td class="check-cell"><span class="print-box">{wa_mark}</span></td>
              <td class="check-cell"><span class="print-box">{call_mark}</span></td>
            </tr>"""
        table_rows_html.append(row)

    table_rows_str = "\n".join(table_rows_html)

    # Build Pre-rendered Mobile Cards
    mobile_cards_html = []
    for c in contacts:
        sr = c.get("Sr. No.", "")
        name = c.get("Name", "")
        cat = c.get("Category / Note", "Sajode")
        phone = c.get("Phone 1") or c.get("Recipient Number") or "—"
        is_sent = c.get("Invitation Status") == "Sent"
        sender_raw = c.get("Sent From (Sender)") or ""

        cat_badge = '<span class="badge b-sarvo">👨‍👩‍👧‍👦 સર્વો</span>' if cat == 'Sarvo' else '<span class="badge b-sajode">👥 સજોડે</span>'
        status_badge = '<span class="badge b-sent">✅ મોકલાઈ ગયું</span>' if is_sent else '<span class="badge b-pending">⏳ બાકી</span>'

        sender_badge = ""
        if "9879881172" in sender_raw:
            sender_badge = '<span class="badge b-mahesh">મહેશભાઈ</span>'
        elif "6351838451" in sender_raw:
            sender_badge = '<span class="badge b-varsha">વર્ષાબેન</span>'

        wa_mark = "✅" if c.get("વોટ્સએપ OK") else ""
        call_mark = "✅" if c.get("ફોન પર વાતચીત") else ""

        card = f"""        <div class="guest-card">
          <div class="gc-header">
            <div class="gc-name">{name}</div>
            <div class="gc-sr">#{sr}</div>
          </div>
          <div class="gc-meta">
            {cat_badge}
            {status_badge}
            {sender_badge}
            <a href="tel:{phone}" class="gc-phone">📞 {phone}</a>
          </div>
          <div class="gc-checks">
            <div class="check-label">
              <span class="box-square">{wa_mark}</span>
              વોટ્સએપ OK
            </div>
            <div class="check-label">
              <span class="box-square">{call_mark}</span>
              ફોન પર આમંત્રણ
            </div>
          </div>
        </div>"""
        mobile_cards_html.append(card)

    mobile_cards_str = "\n".join(mobile_cards_html)
    contacts_json_str = json.dumps(contacts, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="gu">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>મૈત્રી અને સ્મિત — નિમંત્રણ પત્રિકા ટ્રેકર (Live Tracker)</title>
  
  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Noto+Sans+Gujarati:wght@500;600;700;800&display=swap" rel="stylesheet">
  
  <style>
    :root {{
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --card-border: #e2e8f0;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --primary: #1e3a8a;
      --primary-light: #3b82f6;
      --gold: #b45309;
      --gold-light: #fef3c7;
      --success: #059669;
      --success-bg: #ecfdf5;
      --pending: #d97706;
      --pending-bg: #fffbeb;
      --sajode: #7c3aed;
      --sajode-bg: #f5f3ff;
      --sarvo: #0891b2;
      --sarvo-bg: #ecfeff;
      --mahesh: #2563eb;
      --mahesh-bg: #eff6ff;
      --varsha: #db2777;
      --varsha-bg: #fdf2f8;
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Plus Jakarta Sans', 'Noto Sans Gujarati', 'Nirmala UI', 'Shruti', sans-serif;
      -webkit-tap-highlight-color: transparent;
    }}

    body {{
      background-color: var(--bg);
      color: var(--text-main);
      padding: 16px 12px;
      line-height: 1.5;
    }}

    .container {{
      max-width: 1320px;
      margin: 0 auto;
    }}

    /* Web Header */
    header {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 16px;
      padding: 20px 16px;
      margin-bottom: 16px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
      text-align: center;
    }}

    .badge-top {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 12px;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 700;
      background: #eff6ff;
      color: var(--primary);
      margin-bottom: 8px;
    }}

    .pulse-dot {{
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: var(--success);
      box-shadow: 0 0 8px var(--success);
    }}

    h1 {{
      font-size: 1.5rem;
      font-weight: 800;
      color: var(--primary);
      margin-bottom: 4px;
    }}

    .sub-title {{
      font-size: 0.85rem;
      color: var(--text-muted);
    }}

    /* Action Toolbar (Mobile-first) */
    .action-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 16px;
    }}

    .btn-action {{
      flex: 1;
      min-width: 150px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      padding: 13px 16px;
      border-radius: 12px;
      font-size: 0.92rem;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
      border: none;
      text-decoration: none;
      text-align: center;
    }}

    .btn-pdf {{
      background: #1e3a8a;
      color: #ffffff;
      box-shadow: 0 4px 12px rgba(30, 58, 138, 0.2);
    }}

    .btn-pdf:hover {{
      background: #1e40af;
      transform: translateY(-1px);
    }}

    .btn-print {{
      background: #0f172a;
      color: #ffffff;
      box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2);
    }}

    .btn-print:hover {{
      background: #334155;
    }}

    .btn-refresh {{
      background: #ffffff;
      color: var(--text-main);
      border: 1px solid var(--card-border);
    }}

    .btn-refresh:hover {{
      background: #f1f5f9;
    }}

    /* Quick Summary Cards */
    .stats-grid {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      margin-bottom: 16px;
    }}

    @media (min-width: 768px) {{
      .stats-grid {{
        grid-template-columns: repeat(4, 1fr);
        gap: 16px;
      }}
      h1 {{ font-size: 2rem; }}
    }}

    .stat-card {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      padding: 14px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
      border-top: 4px solid var(--primary);
    }}

    .stat-card.c-green {{ border-top-color: var(--success); }}
    .stat-card.c-gold {{ border-top-color: var(--gold); }}
    .stat-card.c-blue {{ border-top-color: var(--mahesh); }}
    .stat-card.c-pink {{ border-top-color: var(--varsha); }}

    .stat-title {{
      font-size: 0.75rem;
      font-weight: 600;
      color: var(--text-muted);
      margin-bottom: 4px;
    }}

    .stat-value {{
      font-size: 1.5rem;
      font-weight: 800;
      color: var(--text-main);
    }}

    .stat-sub {{
      font-size: 0.7rem;
      color: var(--text-muted);
    }}

    /* Progress bar */
    .progress-box {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      padding: 14px 16px;
      margin-bottom: 16px;
    }}

    .progress-labels {{
      display: flex;
      justify-content: space-between;
      font-size: 0.8rem;
      font-weight: 700;
      margin-bottom: 6px;
    }}

    .progress-track {{
      background: #e2e8f0;
      height: 10px;
      border-radius: 9999px;
      overflow: hidden;
    }}

    .progress-fill {{
      height: 100%;
      background: linear-gradient(90deg, #10b981, #06b6d4, #3b82f6);
      border-radius: 9999px;
      transition: width 0.8s ease;
    }}

    /* Search & Filter Controls */
    .controls {{
      display: flex;
      flex-direction: column;
      gap: 10px;
      margin-bottom: 16px;
    }}

    .search-input {{
      width: 100%;
      padding: 12px 14px;
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      font-size: 0.95rem;
      color: var(--text-main);
      outline: none;
    }}

    .search-input:focus {{
      border-color: var(--primary-light);
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);
    }}

    .filter-scroll {{
      display: flex;
      gap: 6px;
      overflow-x: auto;
      padding-bottom: 4px;
      -webkit-overflow-scrolling: touch;
    }}

    .filter-btn {{
      white-space: nowrap;
      padding: 8px 12px;
      border-radius: 10px;
      font-size: 0.8rem;
      font-weight: 600;
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      color: var(--text-muted);
      cursor: pointer;
    }}

    .filter-btn.active {{
      background: var(--primary);
      border-color: var(--primary);
      color: #ffffff;
    }}

    /* Mobile Cards View */
    .mobile-cards {{
      display: flex;
      flex-direction: column;
      gap: 10px;
    }}

    .guest-card {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      padding: 14px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
    }}

    .gc-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 8px;
    }}

    .gc-name {{
      font-size: 1.05rem;
      font-weight: 700;
      color: var(--text-main);
    }}

    .gc-sr {{
      font-size: 0.75rem;
      font-weight: 700;
      color: var(--text-muted);
      background: #f1f5f9;
      padding: 2px 8px;
      border-radius: 6px;
    }}

    .gc-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      align-items: center;
      margin-bottom: 10px;
    }}

    .gc-phone {{
      font-family: monospace;
      font-size: 0.9rem;
      color: #1e3a8a;
      font-weight: 700;
    }}

    .gc-checks {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      background: #f8fafc;
      padding: 10px;
      border-radius: 10px;
      border: 1px dashed #cbd5e1;
      margin-top: 8px;
    }}

    .check-label {{
      font-size: 0.8rem;
      font-weight: 700;
      color: #1e293b;
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .box-square {{
      width: 20px;
      height: 20px;
      border: 2px solid #64748b;
      border-radius: 4px;
      background: #ffffff;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 0.8rem;
    }}

    /* Badges */
    .badge {{
      display: inline-flex;
      align-items: center;
      gap: 3px;
      padding: 3px 8px;
      border-radius: 6px;
      font-size: 0.72rem;
      font-weight: 700;
    }}

    .b-sajode {{ background: var(--sajode-bg); color: var(--sajode); border: 1px solid #ddd6fe; }}
    .b-sarvo {{ background: var(--sarvo-bg); color: var(--sarvo); border: 1px solid #cffafe; }}
    .b-sent {{ background: var(--success-bg); color: var(--success); border: 1px solid #a7f3d0; }}
    .b-pending {{ background: var(--pending-bg); color: var(--pending); border: 1px solid #fde68a; }}
    .b-mahesh {{ background: var(--mahesh-bg); color: var(--mahesh); border: 1px solid #bfdbfe; }}
    .b-varsha {{ background: var(--varsha-bg); color: var(--varsha); border: 1px solid #fbcfe8; }}

    /* Desktop Table (Hidden on Mobile) */
    .table-box {{
      display: none;
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 16px;
      overflow-x: auto;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
    }}

    @media (min-width: 992px) {{
      .table-box {{ display: block; }}
      .mobile-cards {{ display: none; }}
    }}

    .custom-table {{
      width: 100%;
      border-collapse: collapse;
      text-align: left;
    }}

    .custom-table th {{
      background: #f1f5f9;
      padding: 12px 14px;
      font-size: 0.82rem;
      font-weight: 800;
      color: #1e293b;
      border-bottom: 1px solid var(--card-border);
    }}

    .custom-table td {{
      padding: 12px 14px;
      border-bottom: 1px solid #f1f5f9;
      font-size: 0.9rem;
      vertical-align: middle;
    }}

    .custom-table tr:hover {{
      background: #f8fafc;
    }}

    .col-center {{ text-align: center; }}

    .check-cell {{
      text-align: center;
      width: 110px;
    }}

    .print-box {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 22px;
      height: 22px;
      border: 2px solid #64748b;
      border-radius: 4px;
      background: #ffffff;
      font-size: 0.85rem;
      font-weight: bold;
    }}

    .print-header {{
      display: none;
    }}

    /* ========================================================= */
    /* A4 PORTRAIT PRINT & BIG FONTS OPTIMIZATION                */
    /* ========================================================= */
    @media print {{
      @page {{
        size: A4 portrait !important;
        margin: 8mm 6mm 8mm 6mm !important;
      }}

      body {{
        background: #ffffff !important;
        color: #000000 !important;
        padding: 0 !important;
        font-family: 'Noto Sans Gujarati', 'Nirmala UI', 'Shruti', 'Plus Jakarta Sans', Arial, sans-serif !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }}

      .container {{
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
      }}

      /* Hide interactive web elements */
      header,
      .action-bar,
      .controls,
      .badge-top,
      .filter-scroll,
      .btn-action,
      .search-box,
      .stats-grid,
      .progress-box,
      .mobile-cards {{
        display: none !important;
      }}

      /* Printable Compact Header */
      .print-header {{
        display: block !important;
        text-align: center;
        border-bottom: 2px solid #000000;
        padding-bottom: 5px;
        margin-bottom: 8px;
      }}

      .print-title {{
        font-size: 15pt !important;
        font-weight: 800 !important;
        color: #000000 !important;
        margin-bottom: 2px !important;
      }}

      .print-meta {{
        font-size: 9.5pt !important;
        font-weight: 700 !important;
        color: #000000 !important;
      }}

      /* Big Font Table Layout in Portrait */
      .table-box {{
        display: block !important;
        border: none !important;
        box-shadow: none !important;
        overflow: visible !important;
      }}

      .custom-table {{
        width: 100% !important;
        border-collapse: collapse !important;
      }}

      .custom-table thead {{
        display: table-header-group !important;
      }}

      .custom-table tr {{
        page-break-inside: avoid !important;
        page-break-after: auto !important;
      }}

      .custom-table th {{
        background: #e2e8f0 !important;
        color: #000000 !important;
        border: 1.5px solid #000000 !important;
        padding: 7px 4px !important;
        font-size: 11pt !important; /* BIG FONT HEADER */
        font-weight: 800 !important;
        text-align: center !important;
      }}

      .custom-table td {{
        border: 1px solid #000000 !important;
        padding: 6px 5px !important;
        font-size: 11.5pt !important; /* BIG FONT FOR BODY */
        color: #000000 !important;
        vertical-align: middle !important;
      }}

      /* Hide PDF filename column in portrait print to give maximum width for BIG fonts */
      .th-pdf, .td-pdf {{
        display: none !important;
      }}

      .td-sr {{
        font-size: 12pt !important;
        font-weight: 800 !important;
        text-align: center !important;
        width: 38px !important;
      }}

      .td-name {{
        font-size: 12pt !important; /* BIG FONT NAME */
        font-weight: 700 !important;
        color: #000000 !important;
      }}

      .td-cat {{
        font-size: 11pt !important;
        font-weight: 700 !important;
        text-align: center !important;
        width: 60px !important;
      }}

      .td-phone {{
        font-size: 11.5pt !important; /* BIG FONT PHONE */
        font-weight: 700 !important;
        font-family: monospace, sans-serif !important;
        text-align: center !important;
        white-space: nowrap !important;
        width: 130px !important;
      }}

      .td-status {{
        font-size: 10pt !important;
        font-weight: 700 !important;
        text-align: center !important;
        width: 100px !important;
      }}

      .td-sender {{
        font-size: 10pt !important;
        font-weight: 700 !important;
        text-align: center !important;
        width: 90px !important;
      }}

      .badge {{
        border: 1.5px solid #000000 !important;
        color: #000000 !important;
        background: transparent !important;
        font-size: 9pt !important;
        padding: 2px 4px !important;
        font-weight: 800 !important;
      }}

      .check-cell {{
        width: 65px !important;
        text-align: center !important;
      }}

      .print-box {{
        width: 22px !important;
        height: 22px !important;
        border: 2px solid #000000 !important;
        border-radius: 4px !important;
        display: inline-block !important;
        background: #ffffff !important;
      }}
    }}
  </style>
</head>
<body>

  <div class="container">
    
    <!-- Printable Header (Only appears on A4 Print / PDF) -->
    <div class="print-header">
      <div class="print-title">મૈત્રી & સ્મિત — લગ્ન / સગાઈ નિમંત્રણ પત્રિકા લિસ્ટ</div>
      <div class="print-meta">
        <span>કુલ મહેમાન: {total}</span> | 
        <span>મોકલેલ: {sent_count}</span> | 
        <span>બાકી: {pending_count}</span> | 
        <span>સજોડે: {sajode_count}</span> | 
        <span>સર્વો: {sarvo_count}</span> | 
        <span>તારીખ: 18-Oct-2026</span>
      </div>
    </div>

    <!-- Web Header -->
    <header>
      <div class="badge-top">
        <span class="pulse-dot"></span>
        લાઈવ નિમંત્રણ ટ્રેકર (Live Tracker)
      </div>
      <h1>મૈત્રી & સ્મિત — લગ્ન / સગાઈ નિમંત્રણ</h1>
      <p class="sub-title">WhatsApp નિમંત્રણ પત્રિકા વિતરણ અને કોલિંગ ટ્રેકર (સજોડે / સર્વો)</p>
    </header>

    <!-- Top Action Buttons (Mobile First) -->
    <div class="action-bar">
      <!-- 1-Click Direct Download of Pre-rendered A4 Portrait PDF -->
      <a href="Maitry_and_Smit_Invitations_A4_Portrait.pdf" download="Maitry_and_Smit_Invitations_A4_Portrait.pdf" class="btn-action btn-pdf">
        📥 A4 Portrait PDF ડાઉનલોડ
      </a>

      <!-- Instant Browser Print Dialogue configured in A4 Portrait -->
      <button class="btn-action btn-print" onclick="window.print()">
        🖨️ પ્રિન્ટ / સેવ PDF
      </button>

      <button class="btn-action btn-refresh" onclick="loadData()">
        🔄 રીફ્રેશ
      </button>
    </div>

    <!-- Quick Stats (2x2 Mobile Grid) -->
    <div class="stats-grid">
      <div class="stat-card c-green">
        <div class="stat-title">મોકલાઈ ગયું (Sent)</div>
        <div class="stat-value" id="stat-sent">{sent_count}</div>
        <div class="stat-sub" id="stat-sent-pct">{pct}% પૂર્ણ ({total} માંથી)</div>
      </div>

      <div class="stat-card c-gold">
        <div class="stat-title">બાકી છે (Pending)</div>
        <div class="stat-value" id="stat-pending">{pending_count}</div>
        <div class="stat-sub">મોકલવાના બાકી</div>
      </div>

      <div class="stat-card c-blue">
        <div class="stat-title">મહેશભાઈ Mangukiya</div>
        <div class="stat-value" id="stat-mahesh">{mahesh_count}</div>
        <div class="stat-sub">9879881172 વડે</div>
      </div>

      <div class="stat-card c-pink">
        <div class="stat-title">વર્ષાબેન Mangukiya</div>
        <div class="stat-value" id="stat-varsha">{varsha_count}</div>
        <div class="stat-sub">6351838451 વડે</div>
      </div>
    </div>

    <!-- Progress Bar -->
    <div class="progress-box">
      <div class="progress-labels">
        <span>કુલ કામગીરી પ્રગતિ (Total Progress)</span>
        <span id="progress-text">{sent_count} / {total} ({pct}%)</span>
      </div>
      <div class="progress-track">
        <div class="progress-fill" id="progress-bar" style="width: {pct}%;"></div>
      </div>
    </div>

    <!-- Search & Filter Controls -->
    <div class="controls">
      <input type="text" class="search-input" id="search-input" placeholder="🔍 નામ, ગામ કે ફોન નંબર શોધો...">

      <div class="filter-scroll">
        <button class="filter-btn active" data-filter="all">બધા (<span id="c-all">{total}</span>)</button>
        <button class="filter-btn" data-filter="sent">મોકલેલ (<span id="c-sent">{sent_count}</span>)</button>
        <button class="filter-btn" data-filter="pending">બાકી (<span id="c-pending">{pending_count}</span>)</button>
        <button class="filter-btn" data-filter="sajode">સજોડે (👥 {sajode_count})</button>
        <button class="filter-btn" data-filter="sarvo">સર્વો (👨‍👩‍👧‍👦 {sarvo_count})</button>
        <button class="filter-btn" data-filter="mahesh">મહેશભાઈ ({mahesh_count})</button>
        <button class="filter-btn" data-filter="varsha">વર્ષાબેન ({varsha_count})</button>
      </div>
    </div>

    <!-- Mobile Cards View -->
    <div class="mobile-cards" id="mobile-cards-container">
{mobile_cards_str}
    </div>

    <!-- Desktop & A4 Portrait Print Table View -->
    <div class="table-box">
      <table class="custom-table">
        <thead>
          <tr>
            <th class="td-sr">ક્રમ</th>
            <th class="td-name">મહેમાનનું નામ</th>
            <th class="td-cat">શ્રેણી</th>
            <th class="td-phone">ફોન નંબર</th>
            <th class="td-status">સ્થિતિ</th>
            <th class="td-sender">મોકલનાર</th>
            <th class="th-pdf" style="width: 110px;">PDF પત્રિકા</th>
            <th class="check-cell" style="background:#fef3c7; color:#92400e;">વોટ્સએપ OK?</th>
            <th class="check-cell" style="background:#dbeafe; color:#1e40af;">ફોન પર આમંત્રણ?</th>
          </tr>
        </thead>
        <tbody id="table-body">
{table_rows_str}
        </tbody>
      </table>
    </div>

  </div>

  <script>
    let contactsData = {contacts_json_str};
    let activeFilter = 'all';
    let searchQuery = '';

    async function loadData() {{
      try {{
        const res = await fetch('invitation_contacts_18_oct.json?t=' + Date.now());
        if (res.ok) {{
          contactsData = await res.json();
          render();
        }}
      }} catch (e) {{
        console.log("Local fetch fallback:", e);
      }}
    }}

    function render() {{
      const total = contactsData.length;
      const sentList = contactsData.filter(c => c["Invitation Status"] === "Sent");
      const pendingList = contactsData.filter(c => c["Invitation Status"] !== "Sent");
      
      const maheshSent = contactsData.filter(c => (c["Sent From (Sender)"] || "").includes("9879881172")).length;
      const varshaSent = contactsData.filter(c => (c["Sent From (Sender)"] || "").includes("6351838451")).length;
      
      const sentCount = sentList.length;
      const pendingCount = pendingList.length;
      const pct = total > 0 ? Math.round((sentCount / total) * 100) : 0;

      // Update Stats
      document.getElementById("stat-sent").innerText = sentCount;
      document.getElementById("stat-sent-pct").innerText = `${{pct}}% પૂર્ણ (${{total}} માંથી)`;
      document.getElementById("stat-pending").innerText = pendingCount;
      document.getElementById("stat-mahesh").innerText = maheshSent;
      document.getElementById("stat-varsha").innerText = varshaSent;

      // Update Progress
      document.getElementById("progress-bar").style.width = `${{pct}}%`;
      document.getElementById("progress-text").innerText = `${{sentCount}} / ${{total}} (${{pct}}%)`;

      // Update counts on filter buttons
      document.getElementById("c-all").innerText = total;
      document.getElementById("c-sent").innerText = sentCount;
      document.getElementById("c-pending").innerText = pendingCount;

      // Filter Logic
      let filtered = contactsData.filter(c => {{
        if (activeFilter === 'sent' && c["Invitation Status"] !== 'Sent') return false;
        if (activeFilter === 'pending' && c["Invitation Status"] === 'Sent') return false;
        if (activeFilter === 'sajode' && c["Category / Note"] !== 'Sajode') return false;
        if (activeFilter === 'sarvo' && c["Category / Note"] !== 'Sarvo') return false;
        if (activeFilter === 'mahesh' && !(c["Sent From (Sender)"] || "").includes("9879881172")) return false;
        if (activeFilter === 'varsha' && !(c["Sent From (Sender)"] || "").includes("6351838451")) return false;

        if (searchQuery) {{
          const q = searchQuery.toLowerCase();
          const name = (c["Name"] || "").toLowerCase();
          const phone = (c["Phone 1"] || "") + " " + (c["Recipient Number"] || "");
          const cat = (c["Category / Note"] || "").toLowerCase();
          if (!name.includes(q) && !phone.includes(q) && !cat.includes(q)) {{
            return false;
          }}
        }}
        return true;
      }});

      // 1. Render Mobile Cards View
      const mobileContainer = document.getElementById("mobile-cards-container");
      if (filtered.length === 0) {{
        mobileContainer.innerHTML = `<div style="text-align:center; padding:30px; color:#64748b;">કોઈ મહેમાન મળ્યા નથી.</div>`;
      }} else {{
        mobileContainer.innerHTML = filtered.map(c => {{
          const isSent = c["Invitation Status"] === "Sent";
          const cat = c["Category / Note"] || "Sajode";
          const catBadge = cat === 'Sarvo' 
            ? `<span class="badge b-sarvo">👨‍👩‍👧‍👦 સર્વો</span>` 
            : `<span class="badge b-sajode">👥 સજોડે</span>`;
          const statusBadge = isSent 
            ? `<span class="badge b-sent">✅ મોકલાઈ ગયું</span>` 
            : `<span class="badge b-pending">⏳ બાકી</span>`;
          
          let senderBadge = "";
          if (c["Sent From (Sender)"]) {{
            if (c["Sent From (Sender)"].includes("9879881172")) {{
              senderBadge = `<span class="badge b-mahesh">મહેશભાઈ</span>`;
            }} else if (c["Sent From (Sender)"].includes("6351838451")) {{
              senderBadge = `<span class="badge b-varsha">વર્ષાબેન</span>`;
            }}
          }}

          const phone = c["Phone 1"] || c["Recipient Number"] || "—";
          const waCheck = c["વોટ્સએપ OK"] ? "✅" : "";
          const callCheck = c["ફોન પર વાતચીત"] ? "✅" : "";

          return `
            <div class="guest-card">
              <div class="gc-header">
                <div class="gc-name">${{c["Name"]}}</div>
                <div class="gc-sr">#${{c["Sr. No."]}}</div>
              </div>
              <div class="gc-meta">
                ${{catBadge}}
                ${{statusBadge}}
                ${{senderBadge}}
                <a href="tel:${{phone}}" class="gc-phone">📞 ${{phone}}</a>
              </div>
              <div class="gc-checks">
                <div class="check-label">
                  <span class="box-square">${{waCheck}}</span>
                  વોટ્સએપ OK
                </div>
                <div class="check-label">
                  <span class="box-square">${{callCheck}}</span>
                  ફોન પર આમંત્રણ
                </div>
              </div>
            </div>
          `;
        }}).join('');
      }}

      // 2. Render Desktop / A4 Portrait Print Table View
      const tbody = document.getElementById("table-body");
      if (filtered.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:30px; color:#64748b;">કોઈ મહેમાન મળ્યા નથી.</td></tr>`;
      }} else {{
        tbody.innerHTML = filtered.map(c => {{
          const isSent = c["Invitation Status"] === "Sent";
          const cat = c["Category / Note"] || "Sajode";
          const catBadge = cat === 'Sarvo' 
            ? `<span class="badge b-sarvo">સર્વો</span>` 
            : `<span class="badge b-sajode">સજોડે</span>`;
          const statusBadge = isSent 
            ? `<span class="badge b-sent">✅ મોકલેલ</span>` 
            : `<span class="badge b-pending">⏳ બાકી</span>`;
          
          let senderBadge = "—";
          if (c["Sent From (Sender)"]) {{
            if (c["Sent From (Sender)"].includes("9879881172")) {{
              senderBadge = `<span class="badge b-mahesh">મહેશભાઈ</span>`;
            }} else if (c["Sent From (Sender)"].includes("6351838451")) {{
              senderBadge = `<span class="badge b-varsha">વર્ષાબેન</span>`;
            }}
          }}

          const pdfName = c["PDF Sent"] ? c["PDF Sent"].replace(".pdf", "") : "—";
          const phone = c["Phone 1"] || c["Recipient Number"] || "—";
          const waCheck = c["વોટ્સએપ OK"] ? "✔" : "";
          const callCheck = c["ફોન પર વાતચીત"] ? "✔" : "";

          return `
            <tr>
              <td class="td-sr"><b>#${{c["Sr. No."]}}</b></td>
              <td class="td-name">${{c["Name"]}}</td>
              <td class="td-cat col-center">${{catBadge}}</td>
              <td class="td-phone col-center">${{phone}}</td>
              <td class="td-status col-center">${{statusBadge}}</td>
              <td class="td-sender col-center">${{senderBadge}}</td>
              <td class="td-pdf" style="font-size:0.75rem; color:#475569;">${{pdfName}}</td>
              <td class="check-cell"><span class="print-box">${{waCheck}}</span></td>
              <td class="check-cell"><span class="print-box">${{callCheck}}</span></td>
            </tr>
          `;
        }}).join('');
      }}
    }}

    // Filter Buttons
    document.querySelectorAll(".filter-btn").forEach(btn => {{
      btn.addEventListener("click", () => {{
        document.querySelectorAll(".filter-btn").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        activeFilter = btn.getAttribute("data-filter");
        render();
      }});
    }});

    // Search Input
    document.getElementById("search-input").addEventListener("input", (e) => {{
      searchQuery = e.target.value.trim();
      render();
    }});

    // Initial Load & Refresh
    loadData();
    setInterval(loadData, 15000);
  </script>
</body>
</html>
"""

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] Saved index.html ({len(html_content)} bytes)")

    # Generate A4 Portrait PDF using Edge Headless
    html_url = "file:///" + HTML_FILE.replace("\\", "/")
    cmd = [
        EDGE_PATH,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={PDF_FILE}",
        html_url
    ]
    print(f"Running Edge headless to compile A4 Portrait PDF...")
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if res.returncode == 0 and os.path.exists(PDF_FILE):
        pdf_size = os.path.getsize(PDF_FILE)
        print(f"[SUCCESS] Generated {os.path.basename(PDF_FILE)} ({pdf_size} bytes)")
    else:
        print(f"[ERROR] PDF generation failed: {res.stderr}")

if __name__ == "__main__":
    build_site_and_pdf()
