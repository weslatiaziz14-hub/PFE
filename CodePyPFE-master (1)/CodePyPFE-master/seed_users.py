# -*- coding: utf-8 -*-
# =============================================================================
# USER SEEDING SCRIPT
# Extracts unique operator names from Excel files and populates MySQL users table
# =============================================================================

import sys
import os

# Force UTF-8 output on Windows (avoids cp1252 UnicodeEncodeError)
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass

import pandas as pd
import mysql.connector
from mysql.connector import Error as MySQLError
from werkzeug.security import generate_password_hash

# =============================================================================
# CONFIGURATION
# =============================================================================

DATA_FOLDER = "data"

CNC_FILE            = "Suivi_defauts_CNC.xlsx"
FINAL_CONTROL_FILE  = "Saisie_controle_Finale.xlsx"
MMC_PRODUCTION_FILE = "01-Rapport_quantite_MMC.xlsx"

CNC_SHEET            = "Feuil2"
FINAL_CONTROL_SHEET  = "Feuil3"
MMC_PRODUCTION_SHEET = "\u041b\u0438\u0441\u04421"   # Лист1

# Column names
CNC_COL_OPERATOR    = "Controler par"
CNC_COL_DETECTED_BY = "detecter par"
FC_COL_TECHNICIAN   = "Technicien QM"
MMC_COL_OPERATOR    = "Op\u00e9rateur"   # Opérateur

MYSQL_CONFIG = {
    'host'             : 'localhost',
    'user'             : 'root',
    'password'         : 'root',
    'database'         : 'kpi_dashboard',
    'port'             : 3306,
    'raise_on_warnings': False,
}

DEFAULT_PASSWORD = "password123"

# =============================================================================
# EXCEL HELPERS
# =============================================================================

def load_excel(filename, sheet_name, skip_rows=0):
    filepath = os.path.join(DATA_FOLDER, filename)
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=skip_rows)
        print(f"[OK] Loaded '{filename}' — {len(df)} rows")
        return df
    except Exception as e:
        print(f"[WARNING] Could not read '{filename}': {e}")
        return None


def extract_unique_operators():
    """Extract all unique operator/technician names from Excel files."""
    operators = set()

    def split_and_add(name_string):
        if not name_string:
            return
        s = str(name_string).strip()
        if not s:
            return
        # Split on '/' or '+' delimiters (combined shifts like 'AICHA/SAMI')
        for sep in ('/', '+'):
            if sep in s:
                for part in s.split(sep):
                    p = part.strip()
                    if p:
                        operators.add(p)
                return
        operators.add(s)

    # CNC file
    df_cnc = load_excel(CNC_FILE, CNC_SHEET)
    if df_cnc is not None:
        for col in (CNC_COL_OPERATOR, CNC_COL_DETECTED_BY):
            if col in df_cnc.columns:
                for val in df_cnc[col].dropna().unique():
                    split_and_add(val)

    # Final Control file
    df_fc = load_excel(FINAL_CONTROL_FILE, FINAL_CONTROL_SHEET)
    if df_fc is not None:
        if FC_COL_TECHNICIAN in df_fc.columns:
            for val in df_fc[FC_COL_TECHNICIAN].dropna().unique():
                split_and_add(val)

    # MMC Production file
    df_mmc = load_excel(MMC_PRODUCTION_FILE, MMC_PRODUCTION_SHEET, skip_rows=1)
    if df_mmc is not None:
        if MMC_COL_OPERATOR in df_mmc.columns:
            for val in df_mmc[MMC_COL_OPERATOR].dropna().unique():
                split_and_add(val)

    result = sorted(op for op in operators if op)

    print(f"\n[OK] Extracted {len(result)} unique individual operators:")
    for op in result:
        print(f"     - {op}")

    return result


# =============================================================================
# MYSQL HELPER
# =============================================================================

def get_mysql_connection():
    try:
        conn = mysql.connector.connect(**MYSQL_CONFIG)
        return conn
    except MySQLError as err:
        print(f"[ERROR] MySQL connection failed: {err}")
        return None


def fetch_one(conn, query, params=None):
    """
    Execute a SELECT and return the first row.
    Uses its own cursor that is always closed cleanly.
    """
    cur = conn.cursor()
    try:
        cur.execute(query, params or ())
        row = cur.fetchone()
        cur.fetchall()   # drain any remaining rows so the cursor is clean
        return row
    finally:
        cur.close()


def fetch_all(conn, query, params=None):
    """
    Execute a SELECT and return all rows.
    Uses its own cursor that is always closed cleanly.
    """
    cur = conn.cursor()
    try:
        cur.execute(query, params or ())
        rows = cur.fetchall()
        return rows
    finally:
        cur.close()


def execute_dml(conn, query, params=None):
    """
    Execute an INSERT / UPDATE / DELETE.
    Uses its own cursor; commits nothing (caller decides).
    Returns lastrowid.
    """
    cur = conn.cursor()
    try:
        cur.execute(query, params or ())
        lastrowid = cur.lastrowid
        cur.fetchall()   # some connectors leave state even after DML
        return lastrowid
    finally:
        cur.close()


# =============================================================================
# SEEDING
# =============================================================================

def seed_users(operators):
    """Insert unregistered employee accounts and ensure a manager exists."""
    conn = get_mysql_connection()
    if not conn:
        return False

    try:
        password_hash = generate_password_hash(DEFAULT_PASSWORD, method='pbkdf2:sha256')

        # ── 1. Fetch existing unregistered usernames ───────────────────────
        rows = fetch_all(conn, "SELECT username FROM users WHERE is_registered = 0")
        existing_unregistered = {row[0] for row in rows}

        # ── 2. Insert new employees ────────────────────────────────────────
        inserted = 0
        for operator in operators:
            if operator in existing_unregistered:
                continue
            try:
                execute_dml(
                    conn,
                    "INSERT INTO users (username, password_hash, role, is_registered, created_at) "
                    "VALUES (%s, %s, %s, %s, NOW())",
                    (operator, password_hash, 'employee', 0)
                )
                inserted += 1
            except MySQLError as e:
                if "Duplicate entry" in str(e):
                    print(f"[INFO] User '{operator}' already exists — skipped")
                else:
                    print(f"[WARNING] Failed to insert '{operator}': {e}")

        conn.commit()
        print(f"\n[OK] Seeded {inserted} new employee accounts")

        # ── 3. Ensure a manager account exists ─────────────────────────────
        manager_row = fetch_one(conn, "SELECT id FROM users WHERE role = 'manager' LIMIT 1")

        if not manager_row:
            print("\n[INFO] No manager found — creating default manager account...")
            manager_hash = generate_password_hash("manager123", method='pbkdf2:sha256')
            try:
                execute_dml(
                    conn,
                    "INSERT INTO users (username, password_hash, role, is_registered, created_at) "
                    "VALUES (%s, %s, %s, %s, NOW())",
                    ("manager", manager_hash, 'manager', 1)
                )
                conn.commit()
                print("[OK]  Manager account created:")
                print("       username : manager")
                print("       password : manager123")
                print("[WARNING] Change the manager password after first login!")
            except MySQLError as e:
                print(f"[ERROR] Failed to create manager account: {e}")
        else:
            print("[OK] Manager account already exists")

        return True

    except Exception as e:
        print(f"[ERROR] Seeding failed: {e}")
        import traceback; traceback.print_exc()
        return False

    finally:
        conn.close()


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 60)
    print("  USER SEEDING SCRIPT")
    print("=" * 60)

    print("\n[*] Extracting unique operators from Excel files...")
    operators = extract_unique_operators()

    if not operators:
        print("[WARNING] No operators found in Excel files. Check that the files")
        print("          exist in the 'data/' folder and have the expected columns.")
        return

    print("\n[*] Seeding users table in MySQL...")
    success = seed_users(operators)

    if success:
        print("\n" + "=" * 60)
        print("  SEEDING COMPLETE")
        print("=" * 60)
        print()
        print("Next steps:")
        print("  1. python kpi_automation_system.py   — populate KPI tables")
        print("  2. cd backend && python app.py        — start the Flask server")
        print()
        print("Employee registration flow:")
        print("  - Employees go to /register")
        print("  - Enter their username (must match name seeded above)")
        print("  - Set a password (min 6 chars)")
        print("  - Login at /login")


if __name__ == "__main__":
    main()