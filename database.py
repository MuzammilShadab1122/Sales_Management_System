import os
import re
import sqlite3
import tempfile
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IS_VERCEL = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV")) or not os.access(BASE_DIR, os.W_OK)

if IS_VERCEL:
    DATA_ROOT = os.path.join(tempfile.gettempdir(), "rajpoot_traders_data")
    os.makedirs(DATA_ROOT, exist_ok=True)
else:
    DATA_ROOT = BASE_DIR

DB_PATH = os.path.join(DATA_ROOT, "rajput_traders.db")


def validate_shop_name(raw_name: str) -> str:
    """Validates and normalizes shopkeeper name on the server."""
    if raw_name is None:
        raise ValueError("Shopkeeper name is required.")
    name = re.sub(r"\s+", " ", str(raw_name)).strip()
    if len(name) < 2:
        raise ValueError("Shopkeeper name must be at least 2 characters long.")
    if len(name) > 120:
        raise ValueError("Shopkeeper name cannot exceed 120 characters.")
    if not re.search(r"[A-Za-z0-9]", name):
        raise ValueError("Shopkeeper name must contain letters or numbers.")
    if name in {".", ".."} or name.upper() in {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }:
        raise ValueError("Shopkeeper name is reserved by the operating system.")
    return name


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS company_info (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            company_name TEXT NOT NULL,
            tagline TEXT NOT NULL,
            ceo_name TEXT NOT NULL,
            ceo_phone TEXT NOT NULL,
            manager_name TEXT NOT NULL,
            manager_phone TEXT NOT NULL,
            address TEXT NOT NULL,
            watermark_text TEXT NOT NULL,
            watermark_style TEXT NOT NULL DEFAULT 'diagonal'
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            name TEXT NOT NULL UNIQUE,
            default_rate REAL NOT NULL DEFAULT 0,
            stock_in INTEGER NOT NULL DEFAULT 0,
            stock_out INTEGER NOT NULL DEFAULT 0,
            date_added TEXT NOT NULL,
            CHECK (stock_in >= 0),
            CHECK (stock_out >= 0),
            CHECK (stock_in >= stock_out)
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS shopkeepers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shop_name TEXT NOT NULL UNIQUE,
            contact_person TEXT NOT NULL,
            phone TEXT NOT NULL,
            city TEXT NOT NULL DEFAULT '',
            opening_balance REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL
        );
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS ledger_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shopkeeper_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            product_id INTEGER,
            product_name TEXT NOT NULL DEFAULT '',
            qty INTEGER NOT NULL DEFAULT 0,
            rate REAL NOT NULL DEFAULT 0,
            amount REAL NOT NULL DEFAULT 0,
            payment_mode TEXT NOT NULL DEFAULT '',
            delivery_info TEXT NOT NULL DEFAULT '',
            note TEXT NOT NULL DEFAULT '',
            received REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY (shopkeeper_id) REFERENCES shopkeepers(id) ON DELETE CASCADE,
            FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE SET NULL
        );
        """
    )

    # Migrate existing database if payment_mode or delivery_info columns are missing
    cur.execute("PRAGMA table_info(ledger_entries);")
    existing_cols = {row["name"] for row in cur.fetchall()}
    if "payment_mode" not in existing_cols:
        cur.execute("ALTER TABLE ledger_entries ADD COLUMN payment_mode TEXT NOT NULL DEFAULT '';")
    if "delivery_info" not in existing_cols:
        cur.execute("ALTER TABLE ledger_entries ADD COLUMN delivery_info TEXT NOT NULL DEFAULT '';")

    # Enforce fixed official company details for Rajpoot Traders
    cur.execute("SELECT COUNT(*) AS cnt FROM company_info;")
    if cur.fetchone()["cnt"] == 0:
        cur.execute(
            """
            INSERT INTO company_info (
                id, company_name, tagline, ceo_name, ceo_phone,
                manager_name, manager_phone, address, watermark_text, watermark_style
            ) VALUES (
                1,
                'Rajpoot Traders (Whole Sale Distributor)',
                'Whole Sale Distributor',
                'Rana Alauddin',
                '03077178887',
                'Rana Burhan ud din',
                '03045156760',
                'Ali Town near Zong Tower Budhla road, Multan.',
                'Rajpoot Traders',
                'diagonal'
            );
            """
        )
    else:
        cur.execute(
            """
            UPDATE company_info
            SET company_name = 'Rajpoot Traders (Whole Sale Distributor)',
                tagline = 'Whole Sale Distributor',
                ceo_name = 'Rana Alauddin',
                ceo_phone = '03077178887',
                manager_name = 'Rana Burhan ud din',
                manager_phone = '03045156760',
                address = 'Ali Town near Zong Tower Budhla road, Multan.',
                watermark_text = 'Rajpoot Traders'
            WHERE id = 1;
            """
        )

    # Seed products (matching client's video FANS_STOCK & shopkeeper sales)
    cur.execute("SELECT COUNT(*) AS cnt FROM products;")
    if cur.fetchone()["cnt"] == 0:
        initial_products = [
            ("SFC Invertor 30 watt Stock", "Sparkel Black 30 watt", 8500, 30, 0, "20-10-26"),
            ("SFC Invertor 30 watt Stock", "Atlas Black 30 watt", 8500, 60, 0, "20-10-26"),
            ("SFC Invertor 30 watt Stock", "Atlas Black 30 watt (Series II)", 8500, 15, 0, "20-10-26"),
            ("SFC Invertor 30 watt Stock", "Atlas Cream 30 watt", 8500, 10, 0, "20-10-26"),
            ("SFC Invertor 30 watt Stock", "SFC 24 Model 30 watt", 8800, 212, 0, "20-10-26"),
            ("SFC Invertor 30 watt Stock", "SFC Designing Series", 9200, 19, 0, "20-10-26"),
            ("Cel SV C.C Stock", "SV C.C Cream", 6200, 148, 2, "20-10-26"),
            ("Cel SV C.C Stock", "SV C.C Black", 6200, 25, 0, "20-10-26"),
            ("Cel SV C.C Stock", "A.N.C C.C Black", 6500, 35, 10, "20-10-26"),
            ("Cel SV C.C Stock", "A.N.C C.C Cream", 6500, 74, 10, "20-10-26"),
            ("Cel SV C.C Stock", "A.N.C (4.B) C.C Cream", 6800, 46, 0, "20-10-26"),
            ("Pedestal & Ceiling Fans Stock", "Ped Sv c.c 4.B", 10700, 40, 1, "17-10-26"),
            ("Pedestal & Ceiling Fans Stock", "Ped Sv c.c 6.B", 10900, 40, 1, "17-10-26"),
            ("Pedestal & Ceiling Fans Stock", "Cel Sv c.c", 6800, 60, 2, "17-10-26"),
            ("Washing Machines (Raj National)", "W.M Metal 603 Silver Raj National", 16500, 25, 1, "17-10-26"),
            ("Washing Machines (Raj National)", "W.M Plastic 700 Silver Raj National", 16000, 25, 1, "17-10-26"),
            ("Washing Machines (Raj National)", "W.M Steel 803 Copper Raj National", 23000, 20, 1, "17-10-26"),
        ]
        cur.executemany(
            """
            INSERT INTO products (category, name, default_rate, stock_in, stock_out, date_added)
            VALUES (?, ?, ?, ?, ?, ?);
            """,
            initial_products,
        )

    # Seed shopkeepers and ledger entries from client video
    cur.execute("SELECT COUNT(*) AS cnt FROM shopkeepers;")
    if cur.fetchone()["cnt"] == 0:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        shopkeepers_data = [
            ("Hassan Husain Electronics Setal Mari Multan", "Rana Zeshan", "03006334742", "Multan", 30000, now_str),
            ("Nawaz Electronics Salar wan", "Asif Nawaz Sb", "03009841757", "Salar Wan", 0, now_str),
            ("W.H Electronics Muneer Abad", "Hamza Sb", "03000469643", "Muneer Abad", 20000, now_str),
        ]
        cur.executemany(
            """
            INSERT INTO shopkeepers (shop_name, contact_person, phone, city, opening_balance, created_at)
            VALUES (?, ?, ?, ?, ?, ?);
            """,
            shopkeepers_data,
        )

        cur.execute("SELECT id, name FROM products;")
        prod_map = {row["name"]: row["id"] for row in cur.fetchall()}

        cur.execute("SELECT id, shop_name FROM shopkeepers;")
        shop_map = {row["shop_name"]: row["id"] for row in cur.fetchall()}

        s1 = shop_map["Hassan Husain Electronics Setal Mari Multan"]
        s2 = shop_map["Nawaz Electronics Salar wan"]
        s3 = shop_map["W.H Electronics Muneer Abad"]

        initial_entries = [
            (s1, "17-10-26", prod_map["Ped Sv c.c 4.B"], "Ped Sv c.c 4.B", 1, 10700, 10700, "", "Self Pickup (Khud Aaye)", "", 0, now_str),
            (s1, "17-10-26", prod_map["Ped Sv c.c 6.B"], "Ped Sv c.c 6.B", 1, 10900, 10900, "", "Self Pickup (Khud Aaye)", "", 0, now_str),
            (s1, "17-10-26", prod_map["Cel Sv c.c"], "Cel Sv c.c", 2, 6800, 13600, "", "Self Pickup (Khud Aaye)", "", 0, now_str),
            (s1, "17-10-26", prod_map["W.M Metal 603 Silver Raj National"], "W.M Metal 603 Silver Raj National", 1, 16500, 16500, "", "Cargo / Gari", "", 0, now_str),
            (s1, "17-10-26", prod_map["W.M Plastic 700 Silver Raj National"], "W.M Plastic 700 Silver Raj National", 1, 16000, 16000, "", "Cargo / Gari", "", 0, now_str),
            (s1, "17-10-26", prod_map["W.M Steel 803 Copper Raj National"], "W.M Steel 803 Copper Raj National", 1, 23000, 23000, "", "Cargo / Gari", "", 0, now_str),
            (s1, "22-10-26", None, "", 0, 0, 0, "Cash", "", "Khod S (Cash Received)", 10000, now_str),
            (s2, "24-10-26", prod_map["SV C.C Cream"], "SV C.C Cream", 2, 6200, 12400, "", "Rickshaw Wala", "", 0, now_str),
            (s3, "25-10-26", None, "", 0, 0, 0, "JazzCash / EasyPaisa", "", "J.C (JazzCash Payment)", 5000, now_str),
        ]
        cur.executemany(
            """
            INSERT INTO ledger_entries (
                shopkeeper_id, date, product_id, product_name, qty, rate, amount,
                payment_mode, delivery_info, note, received, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """,
            initial_entries,
        )

    conn.commit()
    conn.close()
