import json
import os
import re
import shutil
import threading
from datetime import datetime
from typing import Optional
import xlsxwriter

from database import DATA_ROOT, IS_VERCEL, get_connection

EXCEL_DIR = os.path.join(DATA_ROOT, "Excel_Sheets")
SHOPKEEPERS_EXCEL_DIR = os.path.join(EXCEL_DIR, "Shopkeepers")
STOCK_EXCEL_PATH = os.path.join(EXCEL_DIR, "FANS_STOCK.xlsx")
MASTER_EXCEL_PATH = os.path.join(EXCEL_DIR, "Rajput_Traders_Master.xlsx")
MOBILE_VIEWER_PATH = os.path.join(EXCEL_DIR, "Rajpoot_Traders_Mobile_Viewer.html")
CLOUD_CONFIG_PATH = os.path.join(DATA_ROOT, "cloud_sync_config.json")

SYNC_LOCK = threading.Lock()


def safe_filename(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name or "")).strip(" .")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:100] if cleaned else "Shopkeeper"


def safe_sheet_title(name: str, existing_titles: set) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\\x00-\x1f]", "", str(name or "")).strip(" '")
    base = cleaned[:28] if len(cleaned) > 28 else (cleaned or "Sheet")
    candidate = base
    idx = 2
    while candidate.lower() in {t.lower() for t in existing_titles}:
        candidate = f"{base[:24]}_{idx}"
        idx += 1
    existing_titles.add(candidate)
    return candidate


def create_workbook_formats(wb: xlsxwriter.Workbook) -> dict:
    """Creates 100% Microsoft Excel compliant cell formats matching the client's sheets."""
    return {
        "brand_title": wb.add_format({
            "font_name": "Calibri", "font_size": 14, "bold": True,
            "font_color": "#FFFFFF", "bg_color": "#0F172A",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "brand_sub": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#E2E8F0", "bg_color": "#1E293B",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "mint_header": wb.add_format({
            "font_name": "Calibri", "font_size": 13, "bold": True,
            "font_color": "#000000", "bg_color": "#1FF0A6",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "gold_contact": wb.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True, "italic": True,
            "font_color": "#000000", "bg_color": "#EBCB8B",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "blue_cols": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True, "italic": True,
            "font_color": "#000000", "bg_color": "#1DA1F2",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "cyan_cat": wb.add_format({
            "font_name": "Calibri", "font_size": 12, "bold": True,
            "font_color": "#000000", "bg_color": "#2AF5D6",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "lime_cat": wb.add_format({
            "font_name": "Calibri", "font_size": 12, "bold": True,
            "font_color": "#000000", "bg_color": "#52F230",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "coral_cols": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True, "italic": True,
            "font_color": "#000000", "bg_color": "#D96B6B",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "opening_green": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": "#22C55E",
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "data_left": wb.add_format({
            "font_name": "Calibri", "font_size": 10,
            "align": "left", "valign": "vcenter", "border": 1
        }),
        "data_center": wb.add_format({
            "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter", "border": 1
        }),
        "data_num": wb.add_format({
            "font_name": "Calibri", "font_size": 10,
            "align": "right", "valign": "vcenter", "border": 1, "num_format": "#,##0"
        }),
        "data_bold_num": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "align": "right", "valign": "vcenter", "border": 1, "num_format": "#,##0"
        }),
        "total_label": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": "#F1F5F9", "align": "left", "valign": "vcenter",
            "top": 2, "bottom": 2, "left": 1, "right": 1
        }),
        "total_center": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": "#F1F5F9", "align": "center", "valign": "vcenter",
            "top": 2, "bottom": 2, "left": 1, "right": 1, "num_format": "#,##0"
        }),
        "total_num": wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": "#F1F5F9", "align": "right", "valign": "vcenter",
            "top": 2, "bottom": 2, "left": 1, "right": 1, "num_format": "#,##0"
        }),
        "stamp": wb.add_format({
            "font_name": "Calibri", "font_size": 9, "bold": True, "italic": True,
            "font_color": "#475569", "align": "center", "valign": "vcenter"
        }),
    }


def write_stock_worksheet(ws, fmts: dict, company: dict, products: list) -> None:
    official_cells = f"{company['ceo_name']} ({company['ceo_phone']}), {company['manager_name']} ({company['manager_phone']})"
    ws.set_header(f"&C&B{company['company_name']}")
    ws.set_footer(f"&CCell no. : {official_cells}")

    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 36)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 14)
    ws.set_column(4, 4, 22)

    ws.set_row(0, 26)
    ws.merge_range(0, 0, 0, 4, f"Name: {company['company_name']}", fmts["brand_title"])

    ws.set_row(1, 20)
    ws.merge_range(1, 0, 1, 4, f"Cell no. : {official_cells}   |   Address: {company['address']}", fmts["brand_sub"])

    categories = {}
    for p in products:
        categories.setdefault(p["category"], []).append(p)

    row_idx = 3
    cat_index = 0
    for cat_name, items in categories.items():
        is_even = cat_index % 2 == 0
        cat_fmt = fmts["cyan_cat"] if is_even else fmts["lime_cat"]
        col_fmt = fmts["blue_cols"] if is_even else fmts["coral_cols"]

        ws.set_row(row_idx, 22)
        ws.merge_range(row_idx, 0, row_idx, 4, cat_name, cat_fmt)
        row_idx += 1

        ws.set_row(row_idx, 20)
        headers = ["Date", "Product Name", "Stock In", "Stock Out", "Remaining Products"]
        for c_i, h_text in enumerate(headers):
            ws.write_string(row_idx, c_i, h_text, col_fmt)
        row_idx += 1

        start_excel_row = row_idx + 1
        total_in = 0
        total_out = 0
        total_rem = 0

        for item in items:
            s_in = int(item["stock_in"] or 0)
            s_out = int(item["stock_out"] or 0)
            rem = s_in - s_out
            total_in += s_in
            total_out += s_out
            total_rem += rem
            excel_r = row_idx + 1

            ws.set_row(row_idx, 19)
            ws.write_string(row_idx, 0, str(item["date_added"] or ""), fmts["data_center"])
            ws.write_string(row_idx, 1, str(item["name"] or ""), fmts["data_left"])
            ws.write_number(row_idx, 2, s_in, fmts["data_num"])
            ws.write_number(row_idx, 3, s_out, fmts["data_num"])
            ws.write_formula(row_idx, 4, f"=C{excel_r}-D{excel_r}", fmts["data_bold_num"], rem)
            row_idx += 1

        end_excel_row = row_idx

        ws.set_row(row_idx, 18)
        for c_i in range(5):
            ws.write_blank(row_idx, c_i, None, fmts["data_left"])
        row_idx += 1

        ws.set_row(row_idx, 21)
        ws.write_blank(row_idx, 0, None, fmts["total_label"])
        ws.write_string(row_idx, 1, "Total Products", fmts["total_label"])
        if end_excel_row >= start_excel_row:
            ws.write_formula(row_idx, 2, f"=SUM(C{start_excel_row}:C{end_excel_row})", fmts["total_num"], total_in)
            ws.write_formula(row_idx, 3, f"=SUM(D{start_excel_row}:D{end_excel_row})", fmts["total_num"], total_out)
            ws.write_formula(row_idx, 4, f"=SUM(E{start_excel_row}:E{end_excel_row})", fmts["total_num"], total_rem)
        else:
            ws.write_number(row_idx, 2, 0, fmts["total_num"])
            ws.write_number(row_idx, 3, 0, fmts["total_num"])
            ws.write_number(row_idx, 4, 0, fmts["total_num"])

        row_idx += 2
        cat_index += 1


def write_shopkeeper_worksheet(ws, fmts: dict, company: dict, shopkeeper: dict, entries: list) -> None:
    official_cells = f"{company['ceo_name']} ({company['ceo_phone']}), {company['manager_name']} ({company['manager_phone']})"
    ws.set_header(f"&C&B{company['company_name']}")
    ws.set_footer(f"&CCell no. : {official_cells}")

    col_widths = [13, 34, 11, 12, 14, 20, 18, 22, 14, 16]
    for idx, width in enumerate(col_widths):
        ws.set_column(idx, idx, width)

    ws.set_row(0, 24)
    ws.merge_range(0, 0, 0, 9, str(shopkeeper["shop_name"]), fmts["mint_header"])

    ws.set_row(1, 22)
    ws.write_string(1, 0, "Contact No:", fmts["gold_contact"])
    ws.merge_range(1, 1, 1, 9, f"{shopkeeper['contact_person']}: {shopkeeper['phone']}", fmts["gold_contact"])

    ws.set_row(2, 21)
    headers = [
        "Date",
        "Product Name",
        "Items No",
        "Rate",
        "Amount",
        "Delivery / Dispatch",
        "Payment Mode",
        "Note",
        "Received",
        "Total Balance",
    ]
    for c_i, h_text in enumerate(headers):
        ws.write_string(2, c_i, h_text, fmts["blue_cols"])

    ws.set_row(3, 19)
    for c_i in range(7):
        ws.write_blank(3, c_i, None, fmts["data_left"])
    ws.write_string(3, 7, "Opening B", fmts["opening_green"])
    ws.write_blank(3, 8, None, fmts["data_left"])
    ob_val = float(shopkeeper["opening_balance"] or 0)
    ws.write_number(3, 9, ob_val, fmts["data_bold_num"])

    row_idx = 4
    start_excel_row = 5
    running_balance = ob_val
    total_qty = 0
    total_amt = 0.0
    total_rec = 0.0

    for entry in entries:
        qty = int(entry["qty"] or 0)
        rate = float(entry["rate"] or 0)
        amt = float(entry["amount"] or 0)
        rec = float(entry["received"] or 0)
        running_balance = running_balance + amt - rec
        total_qty += qty
        total_amt += amt
        total_rec += rec
        excel_r = row_idx + 1

        ws.set_row(row_idx, 19)
        ws.write_string(row_idx, 0, str(entry["date"] or ""), fmts["data_center"])
        ws.write_string(row_idx, 1, str(entry["product_name"] or ""), fmts["data_left"])
        if qty > 0:
            ws.write_number(row_idx, 2, qty, fmts["data_center"])
        else:
            ws.write_blank(row_idx, 2, None, fmts["data_center"])

        if rate > 0:
            ws.write_number(row_idx, 3, rate, fmts["data_num"])
        else:
            ws.write_blank(row_idx, 3, None, fmts["data_num"])

        ws.write_number(row_idx, 4, amt, fmts["data_num"])
        ws.write_string(row_idx, 5, str(entry.get("delivery_info") or ""), fmts["data_left"])
        ws.write_string(row_idx, 6, str(entry.get("payment_mode") or ""), fmts["data_left"])
        ws.write_string(row_idx, 7, str(entry.get("note") or ""), fmts["data_left"])

        if rec > 0:
            ws.write_number(row_idx, 8, rec, fmts["data_num"])
        else:
            ws.write_blank(row_idx, 8, None, fmts["data_num"])

        ws.write_formula(row_idx, 9, f"=J{excel_r-1}+E{excel_r}-I{excel_r}", fmts["data_bold_num"], running_balance)
        row_idx += 1

    end_excel_row = row_idx

    for _ in range(2):
        ws.set_row(row_idx, 18)
        for c_i in range(10):
            ws.write_blank(row_idx, c_i, None, fmts["data_left"])
        row_idx += 1

    ws.set_row(row_idx, 22)
    excel_tot_r = row_idx + 1
    ws.write_blank(row_idx, 0, None, fmts["total_label"])
    ws.write_string(row_idx, 1, "Total Amount", fmts["total_label"])
    ws.write_blank(row_idx, 3, None, fmts["total_label"])
    ws.write_blank(row_idx, 5, None, fmts["total_label"])
    ws.write_blank(row_idx, 6, None, fmts["total_label"])
    ws.write_blank(row_idx, 7, None, fmts["total_label"])

    if end_excel_row >= start_excel_row:
        ws.write_formula(row_idx, 2, f"=SUM(C{start_excel_row}:C{end_excel_row})", fmts["total_center"], total_qty)
        ws.write_formula(row_idx, 4, f"=SUM(E{start_excel_row}:E{end_excel_row})", fmts["total_num"], total_amt)
        ws.write_formula(row_idx, 8, f"=SUM(I{start_excel_row}:I{end_excel_row})", fmts["total_num"], total_rec)
        ws.write_formula(row_idx, 9, f"=J4+E{excel_tot_r}-I{excel_tot_r}", fmts["total_num"], running_balance)
    else:
        ws.write_number(row_idx, 2, 0, fmts["total_center"])
        ws.write_number(row_idx, 4, 0, fmts["total_num"])
        ws.write_number(row_idx, 8, 0, fmts["total_num"])
        ws.write_number(row_idx, 9, ob_val, fmts["total_num"])

    row_idx += 2
    ws.set_row(row_idx, 18)
    ws.merge_range(
        row_idx, 0, row_idx, 9,
        f"Name: {company['company_name']}  |  Cell no. : {official_cells}  |  Address: {company['address']}",
        fmts["stamp"]
    )


def save_workbook_atomically(build_fn, target_path: str) -> tuple[bool, Optional[str]]:
    """
    Writes workbook to a temporary file first and atomically replaces target_path.
    Never raises an exception so one bad or locked Excel file cannot break saving.
    """
    tmp_path = f"{target_path}.tmp_{os.getpid()}_{threading.get_ident()}.xlsx"
    wb = None
    try:
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        wb = xlsxwriter.Workbook(tmp_path)
        build_fn(wb)
        wb.close()
        wb = None
        os.replace(tmp_path, target_path)
        return True, None
    except Exception as exc:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return False, str(exc)


def generate_mobile_offline_viewer(company: dict, products: list, shopkeepers: list, entries_by_shop: dict) -> str:
    payload = {
        "synced_at": datetime.now().strftime("%d-%b-%Y %I:%M %p"),
        "company": company,
        "products": [
            {
                **p,
                "remaining": int(p["stock_in"]) - int(p["stock_out"]),
            }
            for p in products
        ],
        "shopkeepers": [],
    }
    for s in shopkeepers:
        s_entries = entries_by_shop.get(s["id"], [])
        running = float(s["opening_balance"])
        enriched = []
        tot_amt = 0.0
        tot_rec = 0.0
        tot_qty = 0
        for e in s_entries:
            amt = float(e["amount"] or 0)
            rec = float(e["received"] or 0)
            running = running + amt - rec
            tot_amt += amt
            tot_rec += rec
            tot_qty += int(e["qty"] or 0)
            enriched.append({**e, "running_balance": running})
        payload["shopkeepers"].append(
            {
                **s,
                "total_qty": tot_qty,
                "total_amount": tot_amt,
                "total_received": tot_rec,
                "current_balance": running,
                "entries": enriched,
            }
        )

    data_json = json.dumps(payload, ensure_ascii=False)
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Rajpoot Traders — Mobile Khata & Stock Viewer</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: system-ui, -apple-system, sans-serif; }}
  body {{ background: #0f172a; color: #f8fafc; padding-bottom: 40px; }}
  .top {{ background: linear-gradient(135deg, #1e293b, #0f172a); padding: 16px; border-bottom: 2px solid #10b981; }}
  .top h1 {{ font-size: 18px; color: #38bdf8; }}
  .top p {{ font-size: 12px; color: #cbd5e1; margin-top: 4px; }}
  .wrap {{ padding: 14px; max-width: 720px; margin: 0 auto; }}
  .tabs {{ display: flex; gap: 8px; margin-bottom: 14px; }}
  .tab {{ flex: 1; padding: 10px; border-radius: 8px; border: 1px solid #334155; background: #1e293b; color: #fff; font-weight: 700; cursor: pointer; }}
  .tab.active {{ background: #10b981; color: #052e16; border-color: #10b981; }}
  .search {{ width: 100%; padding: 11px 14px; border-radius: 8px; border: 1px solid #334155; background: #1e293b; color: #fff; margin-bottom: 12px; font-size: 14px; }}
  .card {{ background: #1e293b; border: 1px solid #334155; border-radius: 10px; padding: 14px; margin-bottom: 10px; }}
  .row {{ display: flex; justify-content: space-between; align-items: center; margin-top: 6px; font-size: 13px; }}
  .btn-wa {{ display: inline-block; background: #22c55e; color: #052e16; font-weight: 800; padding: 8px 12px; border-radius: 6px; text-decoration: none; font-size: 12.5px; margin-top: 10px; border: none; cursor: pointer; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }}
  th, td {{ border: 1px solid #334155; padding: 6px; text-align: left; }}
  th {{ background: #0284c7; color: #fff; }}
</style>
</head>
<body>
  <div class="top">
    <h1>Name: Rajpoot Traders (Whole Sale Distributor)</h1>
    <p>Cell no. : Rana Alauddin (03077178887), Rana Burhan ud din (03045156760)</p>
    <p>Address: Ali Town near Zong Tower Budhla road, Multan.</p>
    <p id="sync-time" style="color:#10b981; font-weight:700; margin-top:6px;"></p>
  </div>
  <div class="wrap">
    <div class="tabs">
      <button class="tab active" id="t-shops" onclick="showTab('shops')">Shopkeepers Khata</button>
      <button class="tab" id="t-stock" onclick="showTab('stock')">Master Stock</button>
    </div>
    <input class="search" id="q" placeholder="Search shopkeeper or product..." oninput="render()" />
    <div id="content"></div>
  </div>
<script>
const DATA = {data_json};
let currentTab = 'shops';
document.getElementById('sync-time').textContent = 'Last Synced: ' + DATA.synced_at;
function fmt(n) {{ return Number(n || 0).toLocaleString('en-PK'); }}
function showTab(t) {{
  currentTab = t;
  document.getElementById('t-shops').classList.toggle('active', t === 'shops');
  document.getElementById('t-stock').classList.toggle('active', t === 'stock');
  render();
}}
function sendWA(idx) {{
  const s = DATA.shopkeepers[idx];
  let lines = [
    '*Name: Rajpoot Traders (Whole Sale Distributor)*',
    'Cell no. : Rana Alauddin (03077178887), Rana Burhan ud din (03045156760)',
    'Address: Ali Town near Zong Tower Budhla road, Multan.',
    '--------------------------------',
    '*Party:* ' + s.shop_name,
    '*Opening Balance:* Rs. ' + fmt(s.opening_balance)
  ];
  for (const e of s.entries) {{
    let p = ['📅 ' + e.date];
    if (e.product_name) p.push(e.product_name + ' (' + e.qty + 'x' + fmt(e.rate) + '=' + fmt(e.amount) + ')');
    if (e.delivery_info) p.push('🚚 ' + e.delivery_info);
    if (e.received) p.push('💵 Rec: ' + fmt(e.received) + (e.payment_mode ? ' (' + e.payment_mode + ')' : ''));
    p.push('*Bal: Rs. ' + fmt(e.running_balance) + '*');
    lines.push(p.join(' | '));
  }}
  lines.push('--------------------------------');
  lines.push('*NET BALANCE: Rs. ' + fmt(s.current_balance) + '*');
  let ph = (s.phone || '').replace(/[^0-9]/g, '');
  if (ph.startsWith('03')) ph = '92' + ph.slice(1);
  window.open('https://wa.me/' + ph + '?text=' + encodeURIComponent(lines.join('\\n')), '_blank');
}}
function render() {{
  const q = (document.getElementById('q').value || '').toLowerCase();
  const box = document.getElementById('content');
  if (currentTab === 'shops') {{
    box.innerHTML = DATA.shopkeepers.map((s, idx) => {{
      if (q && !s.shop_name.toLowerCase().includes(q) && !s.contact_person.toLowerCase().includes(q)) return '';
      const rows = s.entries.map(e => `<tr><td>${{e.date}}</td><td>${{e.product_name || e.note || '-'}}</td><td>${{e.qty || ''}}</td><td>${{fmt(e.amount)}}</td><td>${{fmt(e.received)}}</td><td><b>${{fmt(e.running_balance)}}</b></td></tr>`).join('');
      return `<div class="card">
        <div style="font-weight:800; font-size:15px; color:#38bdf8;">${{s.shop_name}}</div>
        <div class="row"><span>${{s.contact_person}}: ${{s.phone}}</span><span>Opening: Rs. ${{fmt(s.opening_balance)}}</span></div>
        <div class="row"><span>Total Sales: Rs. ${{fmt(s.total_amount)}}</span><span style="color:#4ade80;">Rec: Rs. ${{fmt(s.total_received)}}</span></div>
        <div class="row" style="font-size:15px; font-weight:800; color:#fbbf24;"><span>Current Balance:</span><span>Rs. ${{fmt(s.current_balance)}}</span></div>
        <table><thead><tr><th>Date</th><th>Item/Note</th><th>Qty</th><th>Amt</th><th>Rec</th><th>Bal</th></tr></thead><tbody>${{rows}}</tbody></table>
        <button class="btn-wa" onclick="sendWA(${{idx}})">📲 Forward Full Khata on WhatsApp</button>
      </div>`;
    }}).join('');
  }} else {{
    box.innerHTML = DATA.products.filter(p => !q || p.name.toLowerCase().includes(q) || p.category.toLowerCase().includes(q)).map(p => `
      <div class="card">
        <div style="font-size:11px; color:#94a3b8;">${{p.category}}</div>
        <div style="font-weight:800; font-size:15px;">${{p.name}}</div>
        <div class="row"><span>Stock In: ${{p.stock_in}}</span><span>Out: ${{p.stock_out}}</span><span style="color:#4ade80; font-weight:800;">Remaining: ${{p.remaining}}</span></div>
      </div>
    `).join('');
  }}
}}
render();
</script>
</body>
</html>"""
    with open(MOBILE_VIEWER_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    return MOBILE_VIEWER_PATH


def get_cloud_mirror_dir() -> str:
    if IS_VERCEL:
        return ""
    if os.path.exists(CLOUD_CONFIG_PATH):
        try:
            with open(CLOUD_CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                custom_path = cfg.get("cloud_mirror_dir", "").strip()
                if custom_path:
                    return custom_path
        except Exception:
            pass
    user_home = os.path.expanduser("~")
    candidates = [
        r"G:\My Drive\Rajpoot_Traders_Cloud_Sync",
        os.path.join(user_home, "Google Drive", "Rajpoot_Traders_Cloud_Sync"),
        os.path.join(user_home, "OneDrive", "Rajpoot_Traders_Cloud_Sync"),
    ]
    for cand in candidates:
        parent = os.path.dirname(cand)
        if os.path.exists(parent):
            return cand
    return ""


def set_cloud_mirror_dir(folder_path: str) -> str:
    folder_path = folder_path.strip()
    if folder_path and not IS_VERCEL:
        os.makedirs(folder_path, exist_ok=True)
    with open(CLOUD_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump({"cloud_mirror_dir": folder_path}, f, indent=2)
    return folder_path


def sync_excel_files(
    shopkeeper_id: Optional[int] = None,
    rebuild_stock: bool = False,
    rebuild_master: bool = False,
    rebuild_all_shops: bool = False,
    old_shop_name: Optional[str] = None,
) -> dict:
    """
    Incremental, fault-isolated Excel synchronizer:
      - When shopkeeper_id is provided, rebuilds ONLY that shopkeeper's Excel file
        (plus FANS_STOCK.xlsx if rebuild_stock=True).
      - Isolates every file write in save_workbook_atomically so a single bad/locked
        Excel file never breaks saving or blocks other files.
    """
    with SYNC_LOCK:
        os.makedirs(EXCEL_DIR, exist_ok=True)
        os.makedirs(SHOPKEEPERS_EXCEL_DIR, exist_ok=True)

        conn = get_connection()
        cur = conn.cursor()

        cur.execute("SELECT * FROM company_info WHERE id = 1;")
        company = dict(cur.fetchone())

        products = []
        if rebuild_stock or rebuild_master or rebuild_all_shops:
            cur.execute("SELECT * FROM products ORDER BY category, id;")
            products = [dict(r) for r in cur.fetchall()]

        if rebuild_all_shops or rebuild_master:
            cur.execute("SELECT * FROM shopkeepers ORDER BY id;")
            shopkeepers = [dict(r) for r in cur.fetchall()]
        elif shopkeeper_id is not None:
            cur.execute("SELECT * FROM shopkeepers WHERE id = ?;", (int(shopkeeper_id),))
            row = cur.fetchone()
            shopkeepers = [dict(row)] if row else []
        else:
            shopkeepers = []

        entries_by_shop = {}
        for s in shopkeepers:
            cur.execute(
                "SELECT * FROM ledger_entries WHERE shopkeeper_id = ? ORDER BY id;",
                (s["id"],),
            )
            entries_by_shop[s["id"]] = [dict(r) for r in cur.fetchall()]

        conn.close()

        locked_files = []
        failed_files = {}
        updated_files = []

        # Remove old shopkeeper Excel file if shopkeeper was renamed
        if old_shop_name and shopkeepers:
            old_fname = f"{safe_filename(old_shop_name)}.xlsx"
            new_fname = f"{safe_filename(shopkeepers[0]['shop_name'])}.xlsx"
            if old_fname.lower() != new_fname.lower():
                old_path = os.path.join(SHOPKEEPERS_EXCEL_DIR, old_fname)
                if os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                    except Exception:
                        pass

        # 1. Rebuild FANS_STOCK.xlsx only when requested
        if rebuild_stock or rebuild_all_shops:
            def _build_stock(wb):
                fmts = create_workbook_formats(wb)
                ws = wb.add_worksheet("FANS STOCK")
                write_stock_worksheet(ws, fmts, company, products)

            ok, err = save_workbook_atomically(_build_stock, STOCK_EXCEL_PATH)
            if ok:
                updated_files.append(STOCK_EXCEL_PATH)
            else:
                locked_files.append("FANS_STOCK.xlsx")
                failed_files["FANS_STOCK.xlsx"] = err

        # 2. Rebuild ONLY the affected shopkeeper(s)
        target_shops = shopkeepers if (rebuild_all_shops or shopkeeper_id is not None) else []
        for s in target_shops:
            shop_filename = f"{safe_filename(s['shop_name'])}.xlsx"
            shop_path = os.path.join(SHOPKEEPERS_EXCEL_DIR, shop_filename)

            def _build_shop(wb, shop_obj=s):
                fmts = create_workbook_formats(wb)
                ws = wb.add_worksheet("Ledger")
                write_shopkeeper_worksheet(ws, fmts, company, shop_obj, entries_by_shop[shop_obj["id"]])

            ok, err = save_workbook_atomically(_build_shop, shop_path)
            if ok:
                updated_files.append(shop_path)
            else:
                locked_files.append(shop_filename)
                failed_files[shop_filename] = err

        # 3. Rebuild Combined Master Workbook only when requested
        if rebuild_master or rebuild_all_shops:
            def _build_master(wb):
                fmts = create_workbook_formats(wb)
                ws_stock = wb.add_worksheet("FANS_STOCK")
                write_stock_worksheet(ws_stock, fmts, company, products)
                used_titles = {"FANS_STOCK"}
                for s in shopkeepers:
                    tab_title = safe_sheet_title(s["shop_name"], used_titles)
                    ws_tab = wb.add_worksheet(tab_title)
                    write_shopkeeper_worksheet(ws_tab, fmts, company, s, entries_by_shop[s["id"]])

            ok, err = save_workbook_atomically(_build_master, MASTER_EXCEL_PATH)
            if ok:
                updated_files.append(MASTER_EXCEL_PATH)
            else:
                locked_files.append("Rajput_Traders_Master.xlsx")
                failed_files["Rajput_Traders_Master.xlsx"] = err

        # 4. Rebuild Mobile Viewer HTML when doing a full sync
        if rebuild_all_shops or rebuild_master:
            try:
                viewer_path = generate_mobile_offline_viewer(company, products, shopkeepers, entries_by_shop)
                updated_files.append(viewer_path)
            except Exception as exc:
                failed_files["Rajpoot_Traders_Mobile_Viewer.html"] = str(exc)

        # 5. Mirror ONLY the updated files to Cloud Drive folder if available
        cloud_dir = get_cloud_mirror_dir()
        cloud_synced = False
        if cloud_dir and updated_files:
            try:
                os.makedirs(os.path.join(cloud_dir, "Shopkeepers"), exist_ok=True)
                for f_path in updated_files:
                    rel = os.path.relpath(f_path, EXCEL_DIR)
                    dest = os.path.join(cloud_dir, rel)
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    shutil.copy2(f_path, dest)
                cloud_synced = True
            except Exception:
                cloud_synced = False

        return {
            "synced_at": datetime.now().strftime("%d-%b-%Y %I:%M:%S %p"),
            "excel_dir": EXCEL_DIR,
            "mobile_viewer_path": MOBILE_VIEWER_PATH,
            "cloud_mirror_dir": cloud_dir,
            "cloud_synced": cloud_synced,
            "updated_count": len(updated_files),
            "locked_files": locked_files,
            "failed_files": failed_files,
        }


def sync_all_excel_files() -> dict:
    """Full sync of FANS_STOCK.xlsx, all Shopkeeper files, Master Workbook, and Mobile Viewer."""
    return sync_excel_files(
        shopkeeper_id=None,
        rebuild_stock=True,
        rebuild_master=True,
        rebuild_all_shops=True,
    )
