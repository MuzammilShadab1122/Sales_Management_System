import json
import mimetypes
import os
import re
import subprocess
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from database import BASE_DIR, IS_VERCEL, get_connection, init_db, validate_shop_name
from excel_sync import (
    EXCEL_DIR,
    MASTER_EXCEL_PATH,
    MOBILE_VIEWER_PATH,
    SHOPKEEPERS_EXCEL_DIR,
    STOCK_EXCEL_PATH,
    get_cloud_mirror_dir,
    safe_filename,
    set_cloud_mirror_dir,
    sync_all_excel_files,
    sync_excel_files,
)

STATIC_DIR = os.path.join(BASE_DIR, "static")

# Ensure DB is initialized (critical for Vercel serverless cold starts)
init_db()


def ensure_unique_shop_filename(cur, clean_name: str, exclude_id: int | None = None) -> None:
    """Prevents two different shop names from colliding onto the same .xlsx filename."""
    target_fname = safe_filename(clean_name).lower()
    if exclude_id is not None:
        cur.execute("SELECT id, shop_name FROM shopkeepers WHERE id != ?;", (exclude_id,))
    else:
        cur.execute("SELECT id, shop_name FROM shopkeepers;")
    for row in cur.fetchall():
        if safe_filename(row["shop_name"]).lower() == target_fname:
            raise ValueError(
                f"Shopkeeper name '{clean_name}' conflicts with existing shop '{row['shop_name']}'."
            )


def api_get_state() -> dict:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM company_info WHERE id = 1;")
    company = dict(cur.fetchone())

    cur.execute("SELECT * FROM products ORDER BY category, id;")
    products = []
    total_stock_in = 0
    total_stock_out = 0
    low_stock_count = 0
    for row in cur.fetchall():
        p = dict(row)
        p["remaining"] = int(p["stock_in"]) - int(p["stock_out"])
        total_stock_in += int(p["stock_in"])
        total_stock_out += int(p["stock_out"])
        if p["remaining"] <= 10:
            low_stock_count += 1
        products.append(p)

    cur.execute("SELECT * FROM shopkeepers ORDER BY id;")
    shopkeepers = []
    total_receivable = 0.0
    total_sales_amount = 0.0
    total_received_amount = 0.0

    for row in cur.fetchall():
        s = dict(row)
        cur.execute(
            """
            SELECT
                COUNT(*) AS entry_count,
                COALESCE(SUM(qty), 0) AS total_items,
                COALESCE(SUM(amount), 0) AS total_amount,
                COALESCE(SUM(received), 0) AS total_received
            FROM ledger_entries
            WHERE shopkeeper_id = ?;
            """,
            (s["id"],),
        )
        agg = dict(cur.fetchone())
        s["entry_count"] = agg["entry_count"]
        s["total_items"] = agg["total_items"]
        s["total_amount"] = float(agg["total_amount"])
        s["total_received"] = float(agg["total_received"])
        s["current_balance"] = float(s["opening_balance"]) + s["total_amount"] - s["total_received"]
        s["excel_file"] = f"{safe_filename(s['shop_name'])}.xlsx"

        total_receivable += s["current_balance"]
        total_sales_amount += s["total_amount"]
        total_received_amount += s["total_received"]
        shopkeepers.append(s)

    conn.close()

    return {
        "company": company,
        "products": products,
        "shopkeepers": shopkeepers,
        "metrics": {
            "shopkeeper_count": len(shopkeepers),
            "product_count": len(products),
            "total_stock_in": total_stock_in,
            "total_stock_out": total_stock_out,
            "total_stock_remaining": total_stock_in - total_stock_out,
            "low_stock_count": low_stock_count,
            "total_sales_amount": total_sales_amount,
            "total_received_amount": total_received_amount,
            "total_receivable": total_receivable,
        },
        "excel_dir": EXCEL_DIR,
        "mobile_viewer_path": MOBILE_VIEWER_PATH,
        "cloud_mirror_dir": get_cloud_mirror_dir(),
        "is_vercel": IS_VERCEL,
    }


def api_update_company(payload: dict) -> dict:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE company_info
        SET company_name = ?, tagline = ?, ceo_name = ?, ceo_phone = ?,
            manager_name = ?, manager_phone = ?, address = ?,
            watermark_text = ?, watermark_style = ?
        WHERE id = 1;
        """,
        (
            str(payload.get("company_name", "")).strip() or "Rajpoot Traders (Whole Sale Distributor)",
            str(payload.get("tagline", "")).strip(),
            str(payload.get("ceo_name", "")).strip(),
            str(payload.get("ceo_phone", "")).strip(),
            str(payload.get("manager_name", "")).strip(),
            str(payload.get("manager_phone", "")).strip(),
            str(payload.get("address", "")).strip(),
            str(payload.get("watermark_text", "")).strip() or "Rajpoot Traders",
            str(payload.get("watermark_style", "diagonal")),
        ),
    )
    conn.commit()
    conn.close()
    return {"status": "ok", "sync": sync_excel_files(rebuild_stock=True)}


def api_create_shopkeeper(payload: dict) -> dict:
    clean_shop_name = validate_shop_name(payload.get("shop_name"))
    contact_person = str(payload.get("contact_person", "")).strip()
    phone = str(payload.get("phone", "")).strip()
    city = str(payload.get("city", "")).strip()
    opening_balance = float(payload.get("opening_balance", 0.0) or 0.0)

    conn = get_connection()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        ensure_unique_shop_filename(cur, clean_shop_name)
        cur.execute(
            """
            INSERT INTO shopkeepers (shop_name, contact_person, phone, city, opening_balance, created_at)
            VALUES (?, ?, ?, ?, ?, ?);
            """,
            (clean_shop_name, contact_person, phone, city, opening_balance, now_str),
        )
        new_id = cur.lastrowid
        conn.commit()
    except ValueError:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise ValueError(f"Shopkeeper already exists or invalid data: {e}")
    conn.close()

    # Rebuild ONLY the newly created shopkeeper's Excel file
    sync_info = sync_excel_files(shopkeeper_id=new_id, rebuild_stock=False)
    return {"status": "ok", "shopkeeper_id": new_id, "sync": sync_info}


def api_update_shopkeeper(shopkeeper_id: int, payload: dict) -> dict:
    clean_shop_name = validate_shop_name(payload.get("shop_name"))
    contact_person = str(payload.get("contact_person", "")).strip()
    phone = str(payload.get("phone", "")).strip()
    city = str(payload.get("city", "")).strip()
    opening_balance = float(payload.get("opening_balance", 0.0) or 0.0)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT shop_name FROM shopkeepers WHERE id = ?;", (shopkeeper_id,))
    existing = cur.fetchone()
    if not existing:
        conn.close()
        raise ValueError("Shopkeeper not found")
    old_name = existing["shop_name"]

    try:
        ensure_unique_shop_filename(cur, clean_shop_name, exclude_id=shopkeeper_id)
        cur.execute(
            """
            UPDATE shopkeepers
            SET shop_name = ?, contact_person = ?, phone = ?, city = ?, opening_balance = ?
            WHERE id = ?;
            """,
            (clean_shop_name, contact_person, phone, city, opening_balance, shopkeeper_id),
        )
        conn.commit()
    except ValueError:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise ValueError(f"Could not update shopkeeper: {e}")
    conn.close()

    # Rebuild ONLY this shopkeeper's Excel file
    sync_info = sync_excel_files(shopkeeper_id=shopkeeper_id, rebuild_stock=False, old_shop_name=old_name)
    return {"status": "ok", "sync": sync_info}


def api_delete_shopkeeper(shopkeeper_id: int) -> dict:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT shop_name FROM shopkeepers WHERE id = ?;", (shopkeeper_id,))
    shop_row = cur.fetchone()
    old_name = shop_row["shop_name"] if shop_row else None

    cur.execute(
        "SELECT product_id, qty FROM ledger_entries WHERE shopkeeper_id = ? AND product_id IS NOT NULL AND qty > 0;",
        (shopkeeper_id,),
    )
    restored_rows = cur.fetchall()
    had_stock_change = len(restored_rows) > 0
    for row in restored_rows:
        cur.execute(
            "UPDATE products SET stock_out = MAX(0, stock_out - ?) WHERE id = ?;",
            (row["qty"], row["product_id"]),
        )
    cur.execute("DELETE FROM shopkeepers WHERE id = ?;", (shopkeeper_id,))
    conn.commit()
    conn.close()

    if old_name:
        old_path = os.path.join(SHOPKEEPERS_EXCEL_DIR, f"{safe_filename(old_name)}.xlsx")
        if os.path.exists(old_path):
            try:
                os.remove(old_path)
            except Exception:
                pass

    return {"status": "ok", "sync": sync_excel_files(shopkeeper_id=None, rebuild_stock=had_stock_change)}


def api_get_shopkeeper_ledger(shopkeeper_id: int) -> dict:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM shopkeepers WHERE id = ?;", (shopkeeper_id,))
    shop_row = cur.fetchone()
    if not shop_row:
        conn.close()
        raise ValueError("Shopkeeper not found")
    shop = dict(shop_row)

    cur.execute(
        "SELECT * FROM ledger_entries WHERE shopkeeper_id = ? ORDER BY id;",
        (shopkeeper_id,),
    )
    entries = []
    running_balance = float(shop["opening_balance"])
    total_qty = 0
    total_amount = 0.0
    total_received = 0.0

    for row in cur.fetchall():
        e = dict(row)
        amt = float(e["amount"] or 0)
        rec = float(e["received"] or 0)
        running_balance = running_balance + amt - rec
        e["running_balance"] = running_balance
        total_qty += int(e["qty"] or 0)
        total_amount += amt
        total_received += rec
        entries.append(e)

    conn.close()
    shop["total_qty"] = total_qty
    shop["total_amount"] = total_amount
    shop["total_received"] = total_received
    shop["current_balance"] = running_balance
    shop["excel_file"] = f"{safe_filename(shop['shop_name'])}.xlsx"

    return {"shopkeeper": shop, "entries": entries}


def api_add_ledger_entry(shopkeeper_id: int, payload: dict) -> dict:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT id FROM shopkeepers WHERE id = ?;", (shopkeeper_id,))
    if not cur.fetchone():
        conn.close()
        raise ValueError("Shopkeeper not found")

    product_id = payload.get("product_id")
    product_id = int(product_id) if product_id else None
    product_name = str(payload.get("custom_product_name", "")).strip()
    raw_qty = int(payload.get("qty", 0) or 0)
    if raw_qty < 0:
        conn.close()
        raise ValueError("Quantity cannot be negative.")
    qty = raw_qty
    rate = max(0.0, float(payload.get("rate", 0.0) or 0.0))

    stock_changed = False
    if product_id:
        if qty > 0:
            # Single-step atomic stock check and deduction so remaining stock can NEVER go negative
            cur.execute(
                """
                UPDATE products
                SET stock_out = stock_out + ?
                WHERE id = ? AND (stock_in - stock_out) >= ?;
                """,
                (qty, product_id, qty),
            )
            if cur.rowcount == 0:
                cur.execute("SELECT name, stock_in, stock_out FROM products WHERE id = ?;", (product_id,))
                prod = cur.fetchone()
                conn.rollback()
                conn.close()
                if not prod:
                    raise ValueError("Selected product not found in Stock")
                remaining = int(prod["stock_in"]) - int(prod["stock_out"])
                raise ValueError(
                    f"Insufficient stock for '{prod['name']}'! Available Remaining Stock: {remaining}, Requested: {qty}."
                )
            stock_changed = True

        cur.execute("SELECT name FROM products WHERE id = ?;", (product_id,))
        prod_row = cur.fetchone()
        if not prod_row:
            conn.rollback()
            conn.close()
            raise ValueError("Selected product not found in Stock")
        product_name = prod_row["name"]

    amount = round(qty * rate, 2) if qty > 0 else 0.0
    received = max(0.0, float(payload.get("received", 0.0) or 0.0))
    payment_mode = str(payload.get("payment_mode", "")).strip()
    delivery_info = str(payload.get("delivery_info", "")).strip()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry_date = str(payload.get("date", "")).strip() or datetime.now().strftime("%d-%m-%y")

    cur.execute(
        """
        INSERT INTO ledger_entries (
            shopkeeper_id, date, product_id, product_name, qty, rate, amount,
            payment_mode, delivery_info, note, received, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """,
        (
            shopkeeper_id,
            entry_date,
            product_id,
            product_name,
            qty,
            rate,
            amount,
            payment_mode,
            delivery_info,
            str(payload.get("note", "")).strip(),
            received,
            now_str,
        ),
    )
    conn.commit()
    conn.close()

    # Rebuild ONLY the affected shopkeeper's Excel file (and FANS_STOCK.xlsx if stock changed)
    sync_info = sync_excel_files(shopkeeper_id=shopkeeper_id, rebuild_stock=stock_changed)
    return {"status": "ok", "sync": sync_info}


def api_delete_ledger_entry(entry_id: int) -> dict:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM ledger_entries WHERE id = ?;", (entry_id,))
    entry = cur.fetchone()
    if not entry:
        conn.close()
        raise ValueError("Ledger entry not found")

    shopkeeper_id = int(entry["shopkeeper_id"])
    stock_changed = False
    if entry["product_id"] and int(entry["qty"] or 0) > 0:
        cur.execute(
            "UPDATE products SET stock_out = MAX(0, stock_out - ?) WHERE id = ?;",
            (int(entry["qty"]), entry["product_id"]),
        )
        stock_changed = True

    cur.execute("DELETE FROM ledger_entries WHERE id = ?;", (entry_id,))
    conn.commit()
    conn.close()

    # Rebuild ONLY the affected shopkeeper's Excel file (and FANS_STOCK.xlsx if stock was restored)
    sync_info = sync_excel_files(shopkeeper_id=shopkeeper_id, rebuild_stock=stock_changed)
    return {"status": "ok", "sync": sync_info}


def api_create_product(payload: dict) -> dict:
    name = re.sub(r"\s+", " ", str(payload.get("name", ""))).strip()
    if len(name) < 2:
        raise ValueError("Product name must be at least 2 characters long.")
    category = re.sub(r"\s+", " ", str(payload.get("category", ""))).strip() or "General Stock"
    stock_in = int(payload.get("stock_in", 0) or 0)
    if stock_in < 0:
        raise ValueError("Initial Stock In cannot be negative.")
    default_rate = max(0.0, float(payload.get("default_rate", 0.0) or 0.0))
    date_str = str(payload.get("date_added", "")).strip() or datetime.now().strftime("%d-%m-%y")

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute(
            """
            INSERT INTO products (category, name, default_rate, stock_in, stock_out, date_added)
            VALUES (?, ?, ?, ?, 0, ?);
            """,
            (category, name, default_rate, stock_in, date_str),
        )
        conn.commit()
    except Exception as e:
        conn.close()
        raise ValueError(f"Product already exists or invalid: {e}")
    conn.close()
    return {"status": "ok", "sync": sync_excel_files(shopkeeper_id=None, rebuild_stock=True)}


def api_adjust_product_stock(product_id: int, payload: dict) -> dict:
    add_stock_in = int(payload.get("add_stock_in", 0) or 0)
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM products WHERE id = ?;", (product_id,))
    prod = cur.fetchone()
    if not prod:
        conn.close()
        raise ValueError("Product not found")

    new_rate = (
        max(0.0, float(payload["new_default_rate"]))
        if payload.get("new_default_rate") is not None
        else float(prod["default_rate"])
    )
    new_date = str(payload.get("date_added", "")).strip() or prod["date_added"]

    # Atomic update ensuring stock_in + add_stock_in >= stock_out (so remaining stock never goes negative)
    cur.execute(
        """
        UPDATE products
        SET stock_in = stock_in + ?, default_rate = ?, date_added = ?
        WHERE id = ? AND (stock_in + ?) >= stock_out AND (stock_in + ?) >= 0;
        """,
        (add_stock_in, new_rate, new_date, product_id, add_stock_in, add_stock_in),
    )
    if cur.rowcount == 0:
        conn.rollback()
        conn.close()
        raise ValueError("Stock adjustment would make remaining stock negative.")

    conn.commit()
    conn.close()
    return {"status": "ok", "sync": sync_excel_files(shopkeeper_id=None, rebuild_stock=True)}


def launch_windows_path(target_path: str, is_folder: bool = False) -> None:
    if IS_VERCEL or os.name != "nt":
        return
    excel_exe = r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
    try:
        if is_folder:
            subprocess.Popen(["explorer.exe", target_path])
        elif target_path.lower().endswith(".xlsx") and os.path.exists(excel_exe):
            subprocess.Popen([excel_exe, target_path])
        else:
            subprocess.Popen(["cmd.exe", "/c", "start", "", target_path], shell=False)
    except Exception:
        try:
            os.startfile(target_path)
        except Exception:
            pass


def api_list_excel_files() -> dict:
    sync_info = sync_all_excel_files()
    files_list = []
    if os.path.exists(EXCEL_DIR):
        for root, _, files in os.walk(EXCEL_DIR):
            for f in sorted(files):
                if f.startswith("~$") or ".tmp_" in f:
                    continue
                full = os.path.join(root, f)
                rel = os.path.relpath(full, EXCEL_DIR)
                stat = os.stat(full)
                files_list.append(
                    {
                        "name": f,
                        "rel_path": rel.replace("\\", "/"),
                        "full_path": full,
                        "size_kb": round(stat.st_size / 1024, 1),
                        "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d-%b-%Y %I:%M:%S %p"),
                    }
                )
    return {"sync": sync_info, "excel_dir": EXCEL_DIR, "files": files_list}


def api_open_excel(payload: dict) -> dict:
    target = payload.get("target", "folder")
    shopkeeper_id = payload.get("shopkeeper_id")
    target_path = EXCEL_DIR
    is_folder = False

    if target == "stock":
        sync_excel_files(rebuild_stock=True)
        target_path = STOCK_EXCEL_PATH
    elif target == "master":
        sync_excel_files(rebuild_stock=True, rebuild_master=True)
        target_path = MASTER_EXCEL_PATH
    elif target == "mobile_viewer":
        sync_all_excel_files()
        target_path = MOBILE_VIEWER_PATH
    elif target == "cloud_folder":
        sync_all_excel_files()
        cdir = get_cloud_mirror_dir()
        target_path = cdir if cdir and os.path.exists(cdir) else EXCEL_DIR
        is_folder = True
    elif target == "shopkeeper" and shopkeeper_id:
        sync_excel_files(shopkeeper_id=int(shopkeeper_id), rebuild_stock=False)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT shop_name FROM shopkeepers WHERE id = ?;", (int(shopkeeper_id),))
        row = cur.fetchone()
        conn.close()
        if row:
            target_path = os.path.join(SHOPKEEPERS_EXCEL_DIR, f"{safe_filename(row['shop_name'])}.xlsx")
    else:
        sync_all_excel_files()
        is_folder = True

    if os.path.exists(target_path):
        launch_windows_path(target_path, is_folder=is_folder)
        return {"status": "opened", "path": target_path, "target": target, "is_vercel": IS_VERCEL}
    raise ValueError(f"File not found: {target_path}")


class RajputTradersRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        return json.loads(raw) if raw else {}

    def _parse_request_path(self) -> tuple[str, dict]:
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query, keep_blank_values=True)
        if "__path" in qs and qs["__path"]:
            raw_route = qs["__path"][0]
            sub_parsed = urlparse(raw_route)
            path = sub_parsed.path or "/"
            if not path.startswith("/"):
                path = "/" + path
            if sub_parsed.query:
                sub_qs = parse_qs(sub_parsed.query, keep_blank_values=True)
                for k, v in sub_qs.items():
                    if k not in qs:
                        qs[k] = v
        else:
            path = parsed.path or "/"
        if len(path) > 1 and path.endswith("/"):
            path = path.rstrip("/")
        return path, qs

    def do_GET(self):
        path, qs = self._parse_request_path()

        try:
            if path == "/api/state":
                return self._send_json(api_get_state())

            if path == "/api/excel-files":
                return self._send_json(api_list_excel_files())

            if path == "/mobile-viewer":
                sync_all_excel_files()
                if os.path.isfile(MOBILE_VIEWER_PATH):
                    with open(MOBILE_VIEWER_PATH, "rb") as f:
                        content = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(content)))
                    self.end_headers()
                    self.wfile.write(content)
                    return

            if path == "/api/download-excel":
                target = (qs.get("target") or ["master"])[0]
                shop_id = (qs.get("shopkeeper_id") or [None])[0]
                dl_path = MASTER_EXCEL_PATH
                if target == "stock":
                    sync_excel_files(rebuild_stock=True)
                    dl_path = STOCK_EXCEL_PATH
                elif target == "master":
                    sync_excel_files(rebuild_stock=True, rebuild_master=True)
                    dl_path = MASTER_EXCEL_PATH
                elif target == "mobile_viewer":
                    sync_all_excel_files()
                    dl_path = MOBILE_VIEWER_PATH
                elif target == "shopkeeper" and shop_id:
                    sync_excel_files(shopkeeper_id=int(shop_id), rebuild_stock=False)
                    conn = get_connection()
                    cur = conn.cursor()
                    cur.execute("SELECT shop_name FROM shopkeepers WHERE id = ?;", (int(shop_id),))
                    row = cur.fetchone()
                    conn.close()
                    if row:
                        dl_path = os.path.join(SHOPKEEPERS_EXCEL_DIR, f"{safe_filename(row['shop_name'])}.xlsx")
                elif target == "rel" and qs.get("file"):
                    rel_file = qs["file"][0].replace("..", "").strip("/\\")
                    candidate = os.path.join(EXCEL_DIR, rel_file)
                    if os.path.isfile(candidate):
                        dl_path = candidate

                if os.path.isfile(dl_path):
                    with open(dl_path, "rb") as f:
                        data_bytes = f.read()
                    fname = os.path.basename(dl_path)
                    mime = (
                        "text/html"
                        if fname.endswith(".html")
                        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    self.send_response(200)
                    self.send_header("Content-Type", mime)
                    self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
                    self.send_header("Content-Length", str(len(data_bytes)))
                    self.end_headers()
                    self.wfile.write(data_bytes)
                    return

            m_ledger = re.match(r"^/api/shopkeepers/(\d+)/ledger$", path)
            if m_ledger:
                return self._send_json(api_get_shopkeeper_ledger(int(m_ledger.group(1))))

            if path == "/" or path == "/index.html":
                file_path = os.path.join(STATIC_DIR, "index.html")
            elif path.startswith("/static/"):
                rel = path[len("/static/"):]
                file_path = os.path.join(STATIC_DIR, rel)
            else:
                file_path = os.path.join(STATIC_DIR, "index.html")

            if os.path.isfile(file_path):
                mime, _ = mimetypes.guess_type(file_path)
                with open(file_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", f"{mime or 'text/html'}; charset=utf-8")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self._send_json({"detail": "Not found"}, status=404)
        except Exception as e:
            self._send_json({"detail": str(e)}, status=400)

    def do_POST(self):
        path, _ = self._parse_request_path()
        try:
            payload = self._read_json()
            if path == "/api/shopkeepers":
                return self._send_json(api_create_shopkeeper(payload))

            m_ledger = re.match(r"^/api/shopkeepers/(\d+)/ledger$", path)
            if m_ledger:
                return self._send_json(api_add_ledger_entry(int(m_ledger.group(1)), payload))

            if path == "/api/products":
                return self._send_json(api_create_product(payload))

            m_stock = re.match(r"^/api/products/(\d+)/stock$", path)
            if m_stock:
                return self._send_json(api_adjust_product_stock(int(m_stock.group(1)), payload))

            if path == "/api/open-excel":
                return self._send_json(api_open_excel(payload))

            if path == "/api/sync-excel":
                return self._send_json(sync_all_excel_files())

            if path == "/api/cloud-config":
                folder = str(payload.get("cloud_mirror_dir", "")).strip()
                set_cloud_mirror_dir(folder)
                return self._send_json(sync_all_excel_files())

            self._send_json({"detail": "Endpoint not found"}, status=404)
        except Exception as e:
            self._send_json({"detail": str(e)}, status=400)

    def do_PUT(self):
        path, _ = self._parse_request_path()
        try:
            payload = self._read_json()
            if path == "/api/company":
                return self._send_json(api_update_company(payload))

            m_shop = re.match(r"^/api/shopkeepers/(\d+)$", path)
            if m_shop:
                return self._send_json(api_update_shopkeeper(int(m_shop.group(1)), payload))

            self._send_json({"detail": "Endpoint not found"}, status=404)
        except Exception as e:
            self._send_json({"detail": str(e)}, status=400)

    def do_DELETE(self):
        path, _ = self._parse_request_path()
        try:
            m_entry = re.match(r"^/api/ledger/(\d+)$", path)
            if m_entry:
                return self._send_json(api_delete_ledger_entry(int(m_entry.group(1))))

            m_shop = re.match(r"^/api/shopkeepers/(\d+)$", path)
            if m_shop:
                return self._send_json(api_delete_shopkeeper(int(m_shop.group(1))))

            self._send_json({"detail": "Endpoint not found"}, status=404)
        except Exception as e:
            self._send_json({"detail": str(e)}, status=400)
