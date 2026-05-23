# -*- coding: utf-8 -*-
# =============================================================================
# KPI AUTOMATION SYSTEM — MMC
# Project : Digitalisation de KPI - MMC
# Date    : 2026
# =============================================================================

import sys
import os

# Force UTF-8 output on Windows (avoids cp1252 UnicodeEncodeError for Cyrillic chars)
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass  # Python < 3.7 fallback — errors will be replaced silently

# =============================================================================
# DESCRIPTION:
#   Reads operational Excel files, calculates KPIs and exports to MySQL
#   (or Excel) ready to be connected to Power BI.
#
#   SOURCE FILES (place in data/ folder):
#     - "01-KPI'S 2027.xlsx"           -> Suivi Journalier des fautes  (daily defect log)
#                                      -> Depuis 2018 LEFT table only  (monthly Qte livrer + defauts)
#     - "01-Rapport_quantite_MMC.xlsx" -> Лист1                        (daily Finish Good / Qte)
#
# FORMULAS (all values in %):
#
#   DAILY  (denominator = Qte from Лист1; defects from Suivi Journalier):
#     RFT CNC              = (Qte - Nbre_defauts_CNC)              / Qte              * 100
#     Taux Erreur CNC      = Nbre_defauts_CNC                      / Qte              * 100
#     Efficacite CNC       = (Nbre_pieces_ctrl - Nbre_defauts_CNC) / Nbre_pieces_ctrl * 100
#     RFT Assembly         = (Qte - Quantite_Assembly)             / Qte              * 100
#     Taux Erreur Assembly = Quantite_Assembly                     / Qte              * 100
#     RFT MMC Global       = (Qte - (Nbre_defauts_CNC + Quantite_Assembly)) / Qte    * 100
#     Finish Good          = SUM(Qte) per date from Лист1 (prototype = NaN only)
#
#   MONTHLY / YEARLY  (denominator = Quantite livrer from Depuis 2018):
#     RFT CNC              = (Qte_livrer - Nbre_defauts_interne)   / Qte_livrer * 100
#     Taux Erreur CNC      = Nbre_defauts_interne                  / Qte_livrer * 100
#     Efficacite CNC       = (Nbre_pieces_ctrl - Nbre_defauts_CNC) / Nbre_pieces_ctrl * 100
#                            (only in CNC module; from Suivi Journalier aggregated monthly)
#     RFT Assembly         = (Qte_livrer - Nbre_defauts_externe)   / Qte_livrer * 100
#     Taux Erreur Assembly = Nbre_defauts_externe                  / Qte_livrer * 100
#     RFT MMC Global       = (Qte_livrer - (interne + externe))    / Qte_livrer * 100
#     Finish Good          = Qte_livrer from Depuis 2018
#
#   COLUMN MAPPING:
#     "Nbre_defauts_CNC"     (daily)   = SUM(Nombre des defauts) WHERE detecter par = 'CNC'
#     "Quantite_Assembly"    (daily)   = SUM(Nombre des defauts) WHERE detecter par = 'Assemblage mecanique'
#     "Nbre_pieces_ctrl"     (daily)   = SUM(Nombre des pieces controlees) WHERE detecter par = 'CNC'
#     "Nbre_defauts_interne" (monthly) = row "Nombre des defauts interne" in Depuis 2018 left table
#     "Nbre_defauts_externe" (monthly) = row "Nombre des defauts externe" in Depuis 2018 left table
#
# OUTPUT SHEETS:
#   -- CNC QUALITY --
#   1.  Suivi_Journalier          Daily log
#   2.  CNC_Mensuel               Monthly CNC KPIs + Efficacite CNC
#   3.  CNC_Journalier            Daily CNC KPIs + Efficacite CNC
#   4.  CNC_RFT_Hebdomadaire      RFT-CNC weekly
#   5.  CNC_Par_Operateur         Defects by operator
#   6.  CNC_Par_Statut            Defects by status
#   7.  CNC_Pareto_Defauts        Pareto of defect types
#   8.  CNC_Pareto_Actions        Pareto of corrective actions
#   9.  CNC_Par_Origine           Defects by origin
#   -- FINAL CONTROL (ASSEMBLY) --
#   10. CF_Mensuel                Monthly Assembly KPIs
#   11. CF_Journalier             Daily Assembly KPIs
#   12. CF_Scrap_Rework           Scrap/Rework tracking
#   13. CF_Pareto_Actions         Pareto of corrective actions
#   14. CF_Par_Origine            Non-conformities by origin
#   15. CF_Par_Technicien         Non-conformities by technician
#   -- MMC PRODUCTION --
#   16. MMC_Journalier            Daily Finish Good
#   17. MMC_FinishGood_Mensuel    Monthly Finish Good (= Qte_livrer) + targets
#   18. MMC_Par_Client            Production by client
#   19. MMC_Par_Operateur         Production by operator
#   20. MMC_Tests                 Test results summary
#   -- GLOBAL (main chart = RFT_MMC_Global) --
#   21. RFT_Global                Monthly: RFT MMC Global + RFT CNC + RFT Assembly
#   22. RFT_Global_Journalier     Daily: RFT MMC Global + RFT CNC + RFT Assembly
#   23. RFT_YTD                   Yearly / YTD KPIs
#   24. KPI_Mensuel               Combined monthly KPIs
#   25. KPI_Global                Annual summary
#   -- ADDITIONAL FILES --
#   26+. [Dynamic Tabs]
# =============================================================================

import pandas as pd
import os
import sys
import re
import unicodedata
from datetime import datetime
import mysql.connector
from mysql.connector import Error as MySQLError

try:
    import xlrd  # noqa: F401
    _XLRD_AVAILABLE = True
except ImportError:
    _XLRD_AVAILABLE = False
    print("[INFO] xlrd not installed - legacy .xls files will be skipped.")

# =============================================================================
# CONFIGURATION
# =============================================================================

DATA_FOLDER   = "data"
OUTPUT_FOLDER = "output"

KPI_FILE            = "01-KPI'S 2027.xlsx"
MMC_PRODUCTION_FILE = "01-Rapport_quantite_MMC.xlsx"
CNC_DEFAUTS_FILE    = "Suivi_defauts_CNC.xlsx"       # NEW — CNC defects (Feuil2)
ASSEMBLY_FILE       = "Saisie_controle_Finale.xlsx"  # NEW — Assembly defects (Feuil3)
OUTPUT_FILE         = "KPI_Output.xlsx"

# Path to dashboard.html — update_dashboard_raw() rewrites its RAW=[...] block
# Relative to the project root (same folder as this script)
DASHBOARD_HTML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "frontend", "templates", "dashboard.html"
)

SUIVI_SHEET  = "Suivi Journalier des fautes"
DEPUIS_SHEET = "Depuis 2018"
MMC_SHEET    = "\u041b\u0438\u0441\u04421"   # Лист1

MYSQL_CONFIG = {
    'host'    : 'localhost',
    'user'    : 'root',
    'password': 'root',
    'database': 'kpi_dashboard',
    'port'    : 3306,
    'raise_on_warnings': False,
}

EXPORT_MODE = 'mysql'   # 'mysql' or 'excel'

# =============================================================================
# TARGETS
# =============================================================================

def load_targets_from_db():
    defaults = {
        'TARGET_FINISH_GOOD'        : 40,
        'TARGET_STRETCH_FINISH_GOOD': 45,
        'TARGET_INTERVENTION_LIMIT' : 35,
        'TARGET_ERROR_RATE_CNC'     : 6.0,
        'TARGET_ERROR_RATE_ASSEMBLY': 10.0,
        'TARGET_RFT_CNC'            : 94.0,
        'TARGET_RFT_ASSEMBLY'       : 90.0,
        'TARGET_RFT_MMC'            : 80.0,
        'TARGET_EFFICACITE_CNC'     : 94.0,
    }
    try:
        conn = mysql.connector.connect(**MYSQL_CONFIG, connection_timeout=5)
        cur  = conn.cursor(dictionary=True)
        cur.execute("SELECT target_name, target_value FROM kpi_targets")
        for row in cur.fetchall():
            k, v = row['target_name'], float(row['target_value'])
            if k in defaults:
                if k.startswith(('TARGET_RFT_', 'TARGET_ERROR_RATE_', 'TARGET_EFFICACITE_')) and v < 1.5:
                    v *= 100
                defaults[k] = v
                print(f"[DB] {k} = {v}")
        cur.close(); conn.close()
    except Exception as e:
        print(f"[WARNING] Targets from DB failed ({e}). Using defaults.")
    return defaults


_T = load_targets_from_db()

TARGET_FINISH_GOOD          = _T['TARGET_FINISH_GOOD']
TARGET_STRETCH_FINISH_GOOD  = _T['TARGET_STRETCH_FINISH_GOOD']
TARGET_INTERVENTION_LIMIT   = _T['TARGET_INTERVENTION_LIMIT']
TARGET_ERROR_RATE_CNC       = _T['TARGET_ERROR_RATE_CNC']
TARGET_ERROR_RATE_ASSEMBLY  = _T['TARGET_ERROR_RATE_ASSEMBLY']
TARGET_RFT_CNC              = _T['TARGET_RFT_CNC']
TARGET_RFT_ASSEMBLY         = _T['TARGET_RFT_ASSEMBLY']
TARGET_RFT_MMC              = _T['TARGET_RFT_MMC']
TARGET_EFFICACITE_CNC       = _T['TARGET_EFFICACITE_CNC']

# =============================================================================
# CONSTANTS
# =============================================================================

MOIS_NOMS = {
    1:"Janvier", 2:"F\u00e9vrier", 3:"Mars",    4:"Avril",
    5:"Mai",     6:"Juin",         7:"Juillet", 8:"Ao\u00fbt",
    9:"Septembre",10:"Octobre",   11:"Novembre",12:"D\u00e9cembre",
}

_MOIS_NORM_MAP = {
    "janvier":1,"fevrier":2,"f\u00e9vrier":2,"mars":3,"avril":4,
    "mai":5,"juin":6,"juillet":7,"aout":8,"ao\u00fbt":8,
    "septembre":9,"octobre":10,"novembre":11,"decembre":12,"d\u00e9cembre":12,
}

# Exact column names in "Suivi Journalier des fautes"
SJ_DATE    = "Date"
SJ_DETECT  = "detecter par"
SJ_CTRL    = "Controler par2"
SJ_ORDER   = "N\u00b0Order"
SJ_STATU   = "Statu"
SJ_PIECES  = "Nombre des pi\u00e9ces control\u00e9s"
SJ_DEFAUTS = "Nombre des d\u00e9fauts"
SJ_NIO     = "Pi\u00e9ce NIO"
SJ_TYPE    = "type technique du defaut"
SJ_DESC    = "Description du defaut"
SJ_QRQC   = "N\u00b0QRQC"
SJ_ACTION  = "Action de correction"
SJ_DETAIL  = "Detail de l'action"
SJ_ORIGINE = "Origine de d\u00e9faut"

SJ_CNC_LABEL = "CNC"
SJ_ASM_LABEL = "Assemblage m\u00e9canique"

# Exact column names in MMC Лист1 (header at row index 1)
MMC_DATE      = "DATE"
MMC_QTE       = "Qte"
MMC_CLIENT    = "Client"
MMC_OPERATOR  = "Op\u00e9rateur"
MMC_PROTOTYPE = "prototype"
MMC_SEC_LOCK  = "Secondary locking"
MMC_VERZT     = "VerZt-Prufung"
MMC_AUSTRAST  = "Austrast"
MMC_SEC_INLK  = "Secondary inlocking"
MMC_OFFSET    = "Offset test:"
MMC_PUSHBACK  = "Push-back test"
MMC_UNLOCK    = "Unlocking test:"
MMC_HOUSING   = "Hosuing attachments:"
MMC_ETANCH    = "etanchiete"
MMC_COULEUR   = "couleur"

# =============================================================================
# UTILITIES
# =============================================================================

def _norm(text):
    t = str(text).strip().lower()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9 ]+", " ", t).strip()


def _month_num(label):
    """Return month number (1-12), 0 for YTD, None if unknown."""
    n = _norm(label).replace(" ", "")
    if n == "ytd":
        return 0
    return _MOIS_NORM_MAP.get(n, None)


def create_folders():
    os.makedirs(DATA_FOLDER,   exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    print("[OK] Folders verified.")


def safe_pct(numerator, denominator):
    """
    (numerator / denominator) * 100.
    Returns 0 wherever denominator is 0 or NaN.
    Works on scalars and pandas Series.
    """
    if isinstance(denominator, pd.Series) or isinstance(numerator, pd.Series):
        num = pd.to_numeric(numerator,   errors="coerce")
        den = pd.to_numeric(denominator, errors="coerce")
        return (num / den.replace(0, pd.NA) * 100).fillna(0).round(2)
    try:
        if not denominator or pd.isna(denominator):
            return 0.0
        return round((numerator / denominator) * 100, 2)
    except Exception:
        return 0.0


def build_pareto(df, group_col, value_col, top_n=20):
    pareto = df.groupby(group_col)[value_col].sum().reset_index()
    pareto = pareto.sort_values(value_col, ascending=False).head(top_n)
    total  = pareto[value_col].sum()
    pareto["Cumul_%"] = (pareto[value_col].cumsum() / total * 100).round(1) if total > 0 else 0
    return pareto


# =============================================================================
# FILE READERS
# Each source file has a completely different layout, so each gets its own
# dedicated reader function rather than a generic one.
# =============================================================================

# ─────────────────────────────────────────────────────────────────────────────
# READER A — "Suivi Journalier des fautes"
#
# Layout : flat tabular table, headers in row 0, no skipping needed.
# Filter : column "detecter par"
#   -> 'CNC'                  = CNC defects
#   -> 'Assemblage mecanique' = Assembly defects
# ─────────────────────────────────────────────────────────────────────────────

def read_suivi_journalier(kpi_filepath):
    """
    Reads 'Suivi Journalier des fautes' from the KPI file.
    Returns cleaned DataFrame or empty DataFrame on failure.
    """
    print(f"\n[READ-A] Suivi Journalier -> {kpi_filepath}")
    try:
        df = pd.read_excel(kpi_filepath, sheet_name=SUIVI_SHEET)
    except Exception as e:
        print(f"  [ERROR] Cannot open sheet: {e}")
        return pd.DataFrame()

    required = [SJ_DATE, SJ_DETECT, SJ_PIECES, SJ_DEFAUTS]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        print(f"  [ERROR] Missing columns: {missing}")
        print(f"  Found: {list(df.columns)}")
        return pd.DataFrame()

    df[SJ_DATE]    = pd.to_datetime(df[SJ_DATE],   errors="coerce")
    df[SJ_PIECES]  = pd.to_numeric(df[SJ_PIECES],  errors="coerce").fillna(0)
    df[SJ_DEFAUTS] = pd.to_numeric(df[SJ_DEFAUTS], errors="coerce").fillna(0)
    df = df.dropna(subset=[SJ_DATE])
    df = df[df[SJ_DATE].dt.year >= 2022].copy()

    df["Annee"]  = df[SJ_DATE].dt.year
    df["Mois_N"] = df[SJ_DATE].dt.month
    df["Mois"]   = df["Mois_N"].map(MOIS_NOMS)
    df["KW"]     = df[SJ_DATE].dt.isocalendar().week.astype(int)

    print(f"  {len(df)} rows | detecter par: {df[SJ_DETECT].unique().tolist()}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# READER B — "Depuis 2018"  (LEFT TABLE ONLY)
#
# Layout: wide/pivoted.
#   Row 0  : headers — col0='Mois', col1='Annee', col2-9=years 2018-2025,
#             col10='Valeur moyenne...' (SKIP), col11=2026
#   Each month block = 5 rows (col0 has month name only on first row):
#     +0 : 'Quantite livrer'
#     +1 : 'Nombre des defauts externe'   -> Assembly defects
#     +2 : "Taux d'erreur externe"        (skip, we recalculate)
#     +3 : 'Nombre des defauts interne'   -> CNC defects
#     +4 : "Taux d'erreur interne"        (skip, we recalculate)
#   YTD block: col0='YTD', same row labels
#
# We parse only left table (cols 0-11, skip col10 'Valeur moyenne').
# ─────────────────────────────────────────────────────────────────────────────

def read_depuis_2018(kpi_filepath):
    """
    Reads ONLY the left table of 'Depuis 2018'.
    Returns DataFrame with one row per (scope, year, month) containing:
      Scope | Annee | Mois_N | Mois | Qte_livrer |
      Nbre_defauts_interne | Nbre_defauts_externe
    """
    print(f"\n[READ-B] Depuis 2018 -> {kpi_filepath}")
    try:
        raw = pd.read_excel(kpi_filepath, sheet_name=DEPUIS_SHEET, header=None)
    except Exception as e:
        print(f"  [ERROR] Cannot open sheet: {e}")
        return pd.DataFrame()

    # Step 1: find header row (col0='Mois', col1='Annee')
    hdr_row = None
    for i, row in raw.iterrows():
        if _norm(str(row.iloc[0])) == "mois" and _norm(str(row.iloc[1])) in ("annee", "annee"):
            hdr_row = i
            break
    # fallback: also accept 'année'
    if hdr_row is None:
        for i, row in raw.iterrows():
            c0 = _norm(str(row.iloc[0]))
            c1 = _norm(str(row.iloc[1]))
            if "mois" in c0 and ("ann" in c1 or "year" in c1):
                hdr_row = i
                break
    if hdr_row is None:
        print("  [ERROR] Header row (Mois/Annee) not found.")
        return pd.DataFrame()

    # Step 2: map col index -> year (cols 2-14, skip 'Valeur moyenne' text col)
    year_col = {}   # {year_int: col_index}
    for j in range(2, min(15, raw.shape[1])):
        cell = raw.iloc[hdr_row, j]
        try:
            y = int(float(str(cell)))
            if 2018 <= y <= 2030:
                year_col[y] = j
        except Exception:
            pass    # skip 'Valeur moyenne...', NaN, text

    if not year_col:
        print("  [ERROR] No year columns found.")
        return pd.DataFrame()
    print(f"  Years detected: {sorted(year_col)}")

    # Step 3: parse data rows into month blocks
    records       = []
    current_label = None
    block         = {}    # {'qte': {col_j: float}, 'interne': ..., 'externe': ...}

    def _cell_float(row_series, col_j):
        try:
            v = row_series.iloc[col_j]
            return float(v) if pd.notna(v) else 0.0
        except Exception:
            return 0.0

    def flush_block(label, blk):
        """Emit one record per year for the completed month/YTD block."""
        mois_n = _month_num(label)
        is_ytd = (mois_n == 0)
        if mois_n is None:
            print(f"  [WARN] Unknown label '{label}' — skipped.")
            return
        for year, col_j in sorted(year_col.items()):
            if year < 2022:
                continue
            qte     = blk.get("qte",     {}).get(col_j, 0.0) or 0.0
            interne = blk.get("interne", {}).get(col_j, 0.0) or 0.0
            externe = blk.get("externe", {}).get(col_j, 0.0) or 0.0
            if qte <= 0 and interne == 0 and externe == 0:
                continue
            records.append({
                "Scope"               : "YTD" if is_ytd else "Monthly",
                "Annee"               : year,
                "Mois_N"              : 0 if is_ytd else mois_n,
                "Mois"                : "YTD" if is_ytd else MOIS_NOMS.get(mois_n, ""),
                "Qte_livrer"          : qte,
                "Nbre_defauts_interne": interne,
                "Nbre_defauts_externe": externe,
            })

    for i in range(hdr_row + 1, len(raw)):
        row  = raw.iloc[i]
        col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        col1n = _norm(col1)

        # New month/YTD block starts when col0 is non-empty
        if col0 and col0.lower() not in ("nan", ""):
            if current_label is not None and block:
                flush_block(current_label, block)
            current_label = col0
            block = {}

        # Accumulate metric rows
        if "quantite livrer" in col1n or "quantit" in col1n and "livrer" in col1n:
            block["qte"]     = {j: _cell_float(row, j) for j in year_col.values()}
        elif "nombre des defauts interne" in col1n:
            block["interne"] = {j: _cell_float(row, j) for j in year_col.values()}
        elif "nombre des defauts externe" in col1n:
            block["externe"] = {j: _cell_float(row, j) for j in year_col.values()}

    # Flush last block
    if current_label is not None and block:
        flush_block(current_label, block)

    if not records:
        print("  [ERROR] No data parsed from Depuis 2018.")
        return pd.DataFrame()

    df = pd.DataFrame(records).sort_values(["Annee", "Mois_N"]).reset_index(drop=True)
    print(f"  {len(df)} records (Monthly + YTD) from {sorted(year_col)} columns.")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# READER C — "Лист1" in 01-Rapport_quantite_MMC.xlsx
#
# Layout: flat tabular table.
#   Row 0  : blank / title row -> skip with header=1
#   Row 1  : column headers (DATE, CAD, Qte, ..., prototype)
# Finish Good = rows where 'prototype' column is NaN / empty string
# ─────────────────────────────────────────────────────────────────────────────

def read_mmc_production(mmc_filepath):
    """
    Reads Лист1 from 01-Rapport_quantite_MMC.xlsx.
    Uses header=1 because row 0 is a blank/title row.
    Finish Good = rows where prototype IS NaN or blank.
    """
    print(f"\n[READ-C] MMC Production -> {mmc_filepath}")
    try:
        df = pd.read_excel(mmc_filepath, sheet_name=MMC_SHEET, header=1)
    except Exception as e:
        print(f"  [ERROR] Cannot open sheet: {e}")
        return pd.DataFrame()

    if MMC_DATE not in df.columns or MMC_QTE not in df.columns:
        print(f"  [ERROR] Missing DATE or Qte. Columns: {list(df.columns)}")
        return pd.DataFrame()

    df[MMC_DATE] = pd.to_datetime(df[MMC_DATE], errors="coerce")
    df[MMC_QTE]  = pd.to_numeric(df[MMC_QTE],   errors="coerce").fillna(0)
    df = df.dropna(subset=[MMC_DATE])
    df = df[df[MMC_DATE].dt.year >= 2022].copy()

    df["Annee"]  = df[MMC_DATE].dt.year
    df["Mois_N"] = df[MMC_DATE].dt.month
    df["Mois"]   = df["Mois_N"].map(MOIS_NOMS)

    if MMC_PROTOTYPE in df.columns:
        # Use isna() because prototype column uses StringDtype where NaN is pd.NA
        # (not float NaN), so astype(str).isin(['nan']) does not match correctly.
        # Finish Good = prototype IS empty/null (not a prototype run)
        proto_filled = df[MMC_PROTOTYPE].fillna("").astype(str).str.strip().str.lower()
        df["Is_Finish_Good"] = df[MMC_PROTOTYPE].isna() | proto_filled.isin(["", "nan", "none"])
    else:
        df["Is_Finish_Good"] = True

    fg = int(df["Is_Finish_Good"].sum())
    print(f"  {len(df)} rows | Finish Good: {fg} | Prototype: {len(df)-fg}")
    return df


# =============================================================================
# ADDITIONAL DATA FILES (scans data/ for any extra .xlsx files)
# Unchanged logic from original — generic reader for unknown structures
# =============================================================================

def _detect_excel_engine(filepath, filename=''):
    XLS_MAGIC  = b'\xd0\xcf\x11\xe0'
    XLSX_MAGIC = b'PK'
    try:
        with open(filepath, 'rb') as f:
            header = f.read(8)
        if header[:4] == XLS_MAGIC:
            if not _XLRD_AVAILABLE:
                raise ImportError(f"'{os.path.basename(filepath)}' is .xls. pip install xlrd>=2.0.1")
            return 'xlrd'
        if header[:2] == XLSX_MAGIC:
            return 'openpyxl'
    except ImportError:
        raise
    except Exception:
        pass
    ext = os.path.splitext(filename or filepath)[1].lower()
    if ext == '.xls':
        if not _XLRD_AVAILABLE:
            raise ImportError(f"'{os.path.basename(filepath)}' needs xlrd. pip install xlrd>=2.0.1")
        return 'xlrd'
    return 'openpyxl'


def identify_target_rows(df):
    df = df.copy()
    df['_ROW_TYPE'] = 'data'
    if len(df) > 0 and len(df.columns) > 0:
        first_col   = df.columns[0]
        target_mask = df[first_col].astype(str).str.lower().str.contains('target', na=False)
        df.loc[target_mask, '_ROW_TYPE'] = 'target'
    target_cols = [c for c in df.columns if 'target' in str(c).lower()]
    return df, target_cols


def get_additional_excel_files():
    skip = {KPI_FILE.lower(), MMC_PRODUCTION_FILE.lower(), CNC_DEFAUTS_FILE.lower(), ASSEMBLY_FILE.lower()}
    result = []
    if not os.path.exists(DATA_FOLDER):
        return result
    try:
        for fn in os.listdir(DATA_FOLDER):
            if fn.startswith('~$'):
                print(f"    [SKIP] Lock file: {fn}")
                continue
            if fn.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')) and fn.lower() not in skip:
                result.append((fn, os.path.join(DATA_FOLDER, fn)))
        print(f"[OK] Found {len(result)} additional data files.")
        return sorted(result)
    except Exception as e:
        print(f"[WARNING] Error scanning data folder: {e}")
        return result


def process_additional_excel_files():
    files = get_additional_excel_files()
    if not files:
        print("[OK] No additional data files.")
        return {}
    all_sheets = {}
    for filename, full_path in files:
        try:
            engine = _detect_excel_engine(full_path, filename)
            xls    = pd.ExcelFile(full_path, engine=engine)
            for sheet_name in xls.sheet_names:
                if sheet_name.startswith('_') or sheet_name.lower() in ['print_area', 'print_titles']:
                    continue
                try:
                    df = pd.read_excel(full_path, sheet_name=sheet_name, engine=engine)
                    if df is None or df.empty:
                        continue
                    df = df.dropna(how="all").dropna(axis=1, how="all")
                    if df.empty:
                        continue
                    new_cols = []
                    for col in df.columns:
                        cs = str(col)
                        if cs.startswith('Unnamed:') or cs in ('nan', 'None'):
                            new_cols.append(f"Column_{len(new_cols)+1}")
                        else:
                            new_cols.append(cs.replace('\n', ' ').replace('\r', ' ').strip()[:50])
                    df.columns = new_cols
                    df, _ = identify_target_rows(df)
                    base = os.path.splitext(filename)[0]
                    key  = re.sub(r'[^a-z0-9_]', '_',
                                  f"{base}_{sheet_name}".replace(' ','_').replace('-','_').lower())[:64]
                    all_sheets[key] = df
                    print(f"    [+] {filename} > {sheet_name} ({len(df)} rows)")
                except Exception as e:
                    print(f"    [WARNING] Sheet '{sheet_name}' from '{filename}': {e}")
        except Exception as e:
            print(f"    [WARNING] File '{filename}': {e}")
    print(f"[OK] Processed {len(all_sheets)} additional sheets.")
    return all_sheets


# =============================================================================
# MODULE 1 — SUIVI JOURNALIER (export table)
# =============================================================================

def build_suivi_journalier(df_sj):
    print("\n--- Building Suivi Journalier ---")
    rename = {
        SJ_DATE   : "Date",
        SJ_DETECT : "Detecter_par",
        SJ_CTRL   : "Controler_par",
        SJ_ORDER  : "N_Order",
        SJ_STATU  : "Statu",
        SJ_PIECES : "Nombre_pieces_controlees",
        SJ_DEFAUTS: "Nombre_defauts",
        SJ_NIO    : "Piece_NIO",
        SJ_TYPE   : "Type_defaut",
        SJ_DESC   : "Description_defaut",
        SJ_QRQC  : "QRQC",
        SJ_ACTION : "Action_correction",
        SJ_ORIGINE: "Origine_defaut",
    }
    cols = ["KW","Mois"] + [c for c in rename if c in df_sj.columns]
    suivi = df_sj[cols].rename(columns=rename)
    print(f"[OK] Suivi Journalier — {len(suivi)} rows.")
    return suivi


# =============================================================================
# MODULE 2 — CNC KPIs
#
# DAILY:
#   Source   : SJ rows WHERE detecter par = 'CNC'
#   Denom    : Qte (daily Finish Good from MMC)
#   RFT CNC              = (Qte - Nbre_defauts_CNC) / Qte * 100
#   Taux Err CNC         = Nbre_defauts_CNC / Qte * 100
#   Efficacite CNC       = (Nbre_pieces_ctrl - Nbre_defauts_CNC) / Nbre_pieces_ctrl * 100
#
# MONTHLY:
#   RFT CNC / Taux Err   : from Depuis 2018 (Qte_livrer as denominator)
#   Efficacite CNC       : from Suivi Journalier CNC rows aggregated monthly
#                          (Nbre_pieces_ctrl - Nbre_defauts_CNC) / Nbre_pieces_ctrl * 100
# =============================================================================

def build_cnc_kpis(df_sj, df_depuis, daily_fg):
    print("\n--- Building CNC KPIs ---")

    cnc_rows = df_sj[df_sj[SJ_DETECT] == SJ_CNC_LABEL].copy()

    # ── DAILY CNC ─────────────────────────────────────────────────────────────
    d_cnc = cnc_rows.groupby(SJ_DATE).agg(
        Nbre_Pieces_Controlees=(SJ_PIECES,  "sum"),
        Nbre_Defauts_CNC      =(SJ_DEFAUTS, "sum"),
        Nb_Ordres             =(SJ_ORDER,   "count"),
    ).reset_index().rename(columns={SJ_DATE: "Date"})

    d_cnc["Annee"]  = d_cnc["Date"].dt.year
    d_cnc["Mois_N"] = d_cnc["Date"].dt.month
    d_cnc["Mois"]   = d_cnc["Mois_N"].map(MOIS_NOMS)

    # Merge daily Finish Good (Qte) as denominator
    if not daily_fg.empty:
        d_cnc = d_cnc.merge(daily_fg[["Date","Qte_FG"]], on="Date", how="left")
        d_cnc["Qte_FG"] = d_cnc["Qte_FG"].fillna(0)
    else:
        d_cnc["Qte_FG"] = d_cnc["Nbre_Pieces_Controlees"]

    # RFT CNC daily = (Qte - Nbre_defauts_CNC) / Qte * 100
    d_cnc["RFT_CNC_%"] = safe_pct(
        d_cnc["Qte_FG"] - d_cnc["Nbre_Defauts_CNC"], d_cnc["Qte_FG"]
    )
    # Taux Erreur CNC daily = Nbre_defauts_CNC / Qte * 100
    d_cnc["Taux_Erreur_CNC_%"] = safe_pct(
        d_cnc["Nbre_Defauts_CNC"], d_cnc["Qte_FG"]
    )
    # Efficacite CNC = (Nbre_pieces_ctrl - Nbre_defauts_CNC) / Nbre_pieces_ctrl * 100
    d_cnc["Efficacite_CNC_%"] = safe_pct(
        d_cnc["Nbre_Pieces_Controlees"] - d_cnc["Nbre_Defauts_CNC"],
        d_cnc["Nbre_Pieces_Controlees"]
    )
    d_cnc["Target_RFT_CNC"]        = TARGET_RFT_CNC
    d_cnc["Target_Erreur_CNC"]     = TARGET_ERROR_RATE_CNC
    d_cnc["Target_Efficacite_CNC"] = TARGET_EFFICACITE_CNC

    # ── MONTHLY Efficacite CNC (from Suivi Journalier, used in monthly table) ─
    m_eff = cnc_rows.groupby(["Annee","Mois_N","Mois"]).agg(
        Nbre_Pieces_Controlees=(SJ_PIECES,  "sum"),
        Nbre_Defauts_CNC      =(SJ_DEFAUTS, "sum"),
        Nb_Ordres             =(SJ_ORDER,   "count"),
    ).reset_index()
    m_eff["Efficacite_CNC_%"] = safe_pct(
        m_eff["Nbre_Pieces_Controlees"] - m_eff["Nbre_Defauts_CNC"],
        m_eff["Nbre_Pieces_Controlees"]
    )

    # ── MONTHLY RFT CNC / Taux Erreur (from Depuis 2018) ─────────────────────
    monthly_src = df_depuis[df_depuis["Scope"] == "Monthly"].copy() if df_depuis is not None and not df_depuis.empty else pd.DataFrame()

    if not monthly_src.empty:
        monthly = monthly_src[["Annee","Mois_N","Mois","Qte_livrer","Nbre_defauts_interne"]].copy()

        # RFT CNC monthly = (Qte_livrer - Nbre_defauts_interne) / Qte_livrer * 100
        monthly["RFT_CNC_%"] = safe_pct(
            monthly["Qte_livrer"] - monthly["Nbre_defauts_interne"], monthly["Qte_livrer"]
        )
        # Taux Erreur CNC monthly = Nbre_defauts_interne / Qte_livrer * 100
        monthly["Taux_Erreur_CNC_%"] = safe_pct(
            monthly["Nbre_defauts_interne"], monthly["Qte_livrer"]
        )
        # Merge Efficacite CNC from Suivi Journalier
        monthly = monthly.merge(
            m_eff[["Annee","Mois_N","Mois","Efficacite_CNC_%",
                   "Nbre_Pieces_Controlees","Nbre_Defauts_CNC","Nb_Ordres"]],
            on=["Annee","Mois_N","Mois"], how="left"
        )
        monthly["Efficacite_CNC_%"]      = monthly["Efficacite_CNC_%"].fillna(0)
        monthly["Target_RFT_CNC"]        = TARGET_RFT_CNC
        monthly["Target_Erreur_CNC"]     = TARGET_ERROR_RATE_CNC
        monthly["Target_Efficacite_CNC"] = TARGET_EFFICACITE_CNC
        monthly["Finish_Good"]           = monthly["Qte_livrer"]
    else:
        monthly = m_eff.copy()
        monthly["RFT_CNC_%"]             = 0.0
        monthly["Taux_Erreur_CNC_%"]     = 0.0
        monthly["Target_RFT_CNC"]        = TARGET_RFT_CNC
        monthly["Target_Erreur_CNC"]     = TARGET_ERROR_RATE_CNC
        monthly["Target_Efficacite_CNC"] = TARGET_EFFICACITE_CNC

    # ── WEEKLY ────────────────────────────────────────────────────────────────
    weekly = cnc_rows.groupby(["Annee","KW","Mois"]).agg(
        Nbre_Pieces_Controlees=(SJ_PIECES,  "sum"),
        Nbre_Defauts          =(SJ_DEFAUTS, "sum"),
        Nb_Ordres             =(SJ_ORDER,   "count"),
    ).reset_index()
    weekly["Taux_Erreur_%"] = safe_pct(weekly["Nbre_Defauts"], weekly["Nbre_Pieces_Controlees"])
    weekly["Efficacite_%"]  = safe_pct(
        weekly["Nbre_Pieces_Controlees"] - weekly["Nbre_Defauts"],
        weekly["Nbre_Pieces_Controlees"]
    )
    weekly["Target_RFT_CNC"] = TARGET_RFT_CNC

    # ── BREAKDOWNS ────────────────────────────────────────────────────────────
    defects_only = cnc_rows[cnc_rows[SJ_DEFAUTS] > 0]

    pareto_desc = build_pareto(defects_only, SJ_DESC, SJ_DEFAUTS)
    pareto_desc.rename(columns={SJ_DESC:"Description_Defaut", SJ_DEFAUTS:"Nombre_Defauts"}, inplace=True)

    pareto_act = build_pareto(defects_only, SJ_ACTION, SJ_DEFAUTS)
    pareto_act.rename(columns={SJ_ACTION:"Action_Correction", SJ_DEFAUTS:"Nombre_Defauts"}, inplace=True)

    by_op = cnc_rows.groupby(SJ_CTRL).agg(
        Nbre_Pieces_Ctrl=(SJ_PIECES,  "sum"),
        Nbre_Defauts    =(SJ_DEFAUTS, "sum"),
    ).reset_index().rename(columns={SJ_CTRL:"Operateur"})
    by_op["Taux_Defauts_%"] = safe_pct(by_op["Nbre_Defauts"], by_op["Nbre_Pieces_Ctrl"])

    by_status = cnc_rows.groupby(SJ_STATU).agg(
        Nbre_Pieces_Ctrl=(SJ_PIECES,  "sum"),
        Nbre_Defauts    =(SJ_DEFAUTS, "sum"),
    ).reset_index().rename(columns={SJ_STATU:"Statu"})
    by_status["Taux_Defauts_%"] = safe_pct(by_status["Nbre_Defauts"], by_status["Nbre_Pieces_Ctrl"])

    by_origin = cnc_rows.groupby(SJ_ORIGINE).agg(
        Nbre_Defauts=(SJ_DEFAUTS,"sum"),
    ).reset_index().rename(columns={SJ_ORIGINE:"Origine"}).sort_values("Nbre_Defauts", ascending=False)

    print(f"[OK] CNC KPIs — {len(monthly)} monthly | {len(d_cnc)} daily records.")
    print(f"     Avg RFT CNC    : {monthly['RFT_CNC_%'].mean():.2f}%")
    print(f"     Avg Taux Err   : {monthly['Taux_Erreur_CNC_%'].mean():.2f}%")
    print(f"     Avg Efficacite : {monthly['Efficacite_CNC_%'].mean():.2f}%")

    return {
        "CNC_Mensuel"         : monthly,
        "CNC_Journalier"      : d_cnc,
        "CNC_RFT_Hebdomadaire": weekly,
        "CNC_Pareto_Defauts"  : pareto_desc,
        "CNC_Pareto_Actions"  : pareto_act,
        "CNC_Par_Operateur"   : by_op,
        "CNC_Par_Statut"      : by_status,
        "CNC_Par_Origine"     : by_origin,
    }


# =============================================================================
# MODULE 3 — ASSEMBLY (Final Control) KPIs
#
# DAILY:
#   Source   : SJ rows WHERE detecter par = 'Assemblage mecanique'
#   Denom    : Qte (daily Finish Good from MMC)
#   RFT Assembly         = (Qte - Quantite_Assembly) / Qte * 100
#   Taux Err Assembly    = Quantite_Assembly          / Qte * 100
#   NOTE: Efficacite CNC is NOT shown here (CNC module only)
#
# MONTHLY:
#   Denom    : Qte_livrer from Depuis 2018
#   RFT Assembly         = (Qte_livrer - Nbre_defauts_externe) / Qte_livrer * 100
#   Taux Err Assembly    = Nbre_defauts_externe                / Qte_livrer * 100
# =============================================================================

def build_assembly_kpis(df_sj, df_depuis, daily_fg):
    print("\n--- Building Assembly KPIs ---")

    asm_rows = df_sj[df_sj[SJ_DETECT] == SJ_ASM_LABEL].copy()

    # ── DAILY ASSEMBLY ────────────────────────────────────────────────────────
    d_asm = asm_rows.groupby(SJ_DATE).agg(
        Quantite_Assembly=(SJ_DEFAUTS, "sum"),
        Nb_Total         =(SJ_ORDER,   "count"),
    ).reset_index().rename(columns={SJ_DATE:"Date"})

    d_asm["Annee"]  = d_asm["Date"].dt.year
    d_asm["Mois_N"] = d_asm["Date"].dt.month
    d_asm["Mois"]   = d_asm["Mois_N"].map(MOIS_NOMS)

    if not daily_fg.empty:
        d_asm = d_asm.merge(daily_fg[["Date","Qte_FG"]], on="Date", how="left")
        d_asm["Qte_FG"] = d_asm["Qte_FG"].fillna(0)
    else:
        d_asm["Qte_FG"] = 0

    # RFT Assembly daily = (Qte - Quantite_Assembly) / Qte * 100
    d_asm["RFT_Assembly_%"] = safe_pct(
        d_asm["Qte_FG"] - d_asm["Quantite_Assembly"], d_asm["Qte_FG"]
    )
    # Taux Erreur Assembly daily = Quantite_Assembly / Qte * 100
    d_asm["Taux_Erreur_Assembly_%"] = safe_pct(
        d_asm["Quantite_Assembly"], d_asm["Qte_FG"]
    )
    d_asm["Target_RFT_Assembly"]    = TARGET_RFT_ASSEMBLY
    d_asm["Target_Erreur_Assembly"] = TARGET_ERROR_RATE_ASSEMBLY

    # ── MONTHLY ASSEMBLY ──────────────────────────────────────────────────────
    monthly_src = df_depuis[df_depuis["Scope"] == "Monthly"].copy() if df_depuis is not None and not df_depuis.empty else pd.DataFrame()

    if not monthly_src.empty:
        monthly = monthly_src[["Annee","Mois_N","Mois","Qte_livrer","Nbre_defauts_externe"]].copy()

        # RFT Assembly monthly = (Qte_livrer - Nbre_defauts_externe) / Qte_livrer * 100
        monthly["RFT_Assembly_%"] = safe_pct(
            monthly["Qte_livrer"] - monthly["Nbre_defauts_externe"], monthly["Qte_livrer"]
        )
        # Taux Erreur Assembly monthly = Nbre_defauts_externe / Qte_livrer * 100
        monthly["Taux_Erreur_Assembly_%"] = safe_pct(
            monthly["Nbre_defauts_externe"], monthly["Qte_livrer"]
        )
        monthly["Target_RFT_Assembly"]    = TARGET_RFT_ASSEMBLY
        monthly["Target_Erreur_Assembly"] = TARGET_ERROR_RATE_ASSEMBLY
        monthly["Finish_Good"]            = monthly["Qte_livrer"]
    else:
        monthly = pd.DataFrame()

    # ── SCRAP / REWORK ────────────────────────────────────────────────────────
    all_asm = df_sj[df_sj[SJ_DETECT] == SJ_ASM_LABEL].copy()
    all_asm["Action_Lower"] = all_asm[SJ_ACTION].fillna("").astype(str).str.lower()

    def classify_action(a):
        if not a or a == "nan": return "Autre"
        if "rebut"   in a: return "Mise en rebut"
        if "rempla"  in a: return "Demande de remplacement"
        if "refaire" in a or "repar" in a: return "Refaire la piece"
        if "annul"   in a: return "Annulation"
        return "Autre"

    all_asm["Type_Action"] = all_asm["Action_Lower"].apply(classify_action)
    scrap = all_asm.groupby(["Annee","Mois_N","Mois","Type_Action"]).size().reset_index(name="Total")

    pareto_act = build_pareto(all_asm[all_asm[SJ_DEFAUTS] > 0], SJ_ACTION, SJ_DEFAUTS)
    pareto_act.rename(columns={SJ_ACTION:"Action_Correction", SJ_DEFAUTS:"Quantite"}, inplace=True)

    by_origin = all_asm.groupby(SJ_ORIGINE).agg(
        Nbre_Defauts=(SJ_DEFAUTS,"sum")
    ).reset_index().rename(columns={SJ_ORIGINE:"Origine"}).sort_values("Nbre_Defauts", ascending=False)

    by_ctrl = all_asm.groupby(SJ_CTRL).agg(
        Nbre_Defauts=(SJ_DEFAUTS,"sum"),
        Nb_Total    =(SJ_ORDER,  "count"),
    ).reset_index().rename(columns={SJ_CTRL:"Technicien"})
    by_ctrl["Taux_%"] = safe_pct(by_ctrl["Nbre_Defauts"], by_ctrl["Nb_Total"])

    print(f"[OK] Assembly KPIs — {len(monthly) if not monthly.empty else 0} monthly | {len(d_asm)} daily records.")
    if not monthly.empty:
        print(f"     Avg RFT Assembly : {monthly['RFT_Assembly_%'].mean():.2f}%")

    return {
        "CF_Mensuel"       : monthly,
        "CF_Journalier"    : d_asm,
        "CF_Scrap_Rework"  : scrap,
        "CF_Pareto_Actions": pareto_act,
        "CF_Par_Origine"   : by_origin,
        "CF_Par_Technicien": by_ctrl,
    }


# =============================================================================
# MODULE 4 — MMC PRODUCTION (Finish Good)
#
# Daily Finish Good  = SUM(Qte) WHERE prototype IS NaN, per date, from Лист1
# Monthly Finish Good= Qte_livrer from Depuis 2018 left table
#
# Returns both the full mmc_results dict AND daily_fg (the shared denominator).
# =============================================================================

def build_mmc_kpis(df_mmc, df_depuis):
    print("\n--- Building MMC / Finish Good KPIs ---")

    daily_fg      = pd.DataFrame()   # [Date, Qte_FG] — shared denominator for all daily KPIs
    daily_fg_full = pd.DataFrame()   # enriched daily table for export
    by_client     = pd.DataFrame()
    by_op         = pd.DataFrame()
    test_results  = pd.DataFrame()

    # ── Daily Finish Good from Лист1 ──────────────────────────────────────────
    if df_mmc is not None and not df_mmc.empty:
        fg_rows  = df_mmc[df_mmc["Is_Finish_Good"]].copy()
        daily_fg = (
            fg_rows.groupby(MMC_DATE)[MMC_QTE]
            .sum()
            .reset_index()
            .rename(columns={MMC_DATE:"Date", MMC_QTE:"Qte_FG"})
        )
        daily_fg["Annee"]  = daily_fg["Date"].dt.year
        daily_fg["Mois_N"] = daily_fg["Date"].dt.month
        daily_fg["Mois"]   = daily_fg["Mois_N"].map(MOIS_NOMS)

        daily_fg_full = daily_fg.copy()
        daily_fg_full["Target_Finish_Good"]         = TARGET_FINISH_GOOD
        daily_fg_full["Target_Stretch_Finish_Good"] = TARGET_STRETCH_FINISH_GOOD
        daily_fg_full["Target_Intervention_Limit"]  = TARGET_INTERVENTION_LIMIT

        if MMC_CLIENT in df_mmc.columns:
            by_client = df_mmc.groupby(MMC_CLIENT)[MMC_QTE].sum().reset_index()\
                            .rename(columns={MMC_CLIENT:"Client", MMC_QTE:"Quantite_Produite"})\
                            .sort_values("Quantite_Produite", ascending=False)

        if MMC_OPERATOR in df_mmc.columns:
            by_op = df_mmc.groupby(MMC_OPERATOR)[MMC_QTE].sum().reset_index()\
                        .rename(columns={MMC_OPERATOR:"Operateur", MMC_QTE:"Quantite_Produite"})\
                        .sort_values("Quantite_Produite", ascending=False)

        test_cols = [MMC_SEC_LOCK, MMC_VERZT, MMC_AUSTRAST, MMC_SEC_INLK,
                     MMC_OFFSET, MMC_PUSHBACK, MMC_UNLOCK, MMC_HOUSING, MMC_ETANCH, MMC_COULEUR]
        test_rows = []
        for col in test_cols:
            if col in df_mmc.columns:
                total  = len(df_mmc)
                passed = int(df_mmc[col].notna().sum())
                test_rows.append({
                    "Test"            : col,
                    "Total_Pieces"    : total,
                    "Reussi"          : passed,
                    "Non_Teste"       : total - passed,
                    "Taux_Reussite_%": round(passed/total*100,2) if total > 0 else 0,
                    "Objectif_%"      : 95.0,
                })
        test_results = pd.DataFrame(test_rows)

    # ── Monthly Finish Good = Qte_livrer from Depuis 2018 ─────────────────────
    monthly_fg = pd.DataFrame()
    if df_depuis is not None and not df_depuis.empty:
        m_src = df_depuis[df_depuis["Scope"] == "Monthly"].copy()
        if not m_src.empty:
            monthly_fg = m_src[["Annee","Mois_N","Mois","Qte_livrer"]].copy()
            monthly_fg.rename(columns={"Qte_livrer":"Finish_Good"}, inplace=True)
            monthly_fg["Target_Finish_Good"]         = TARGET_FINISH_GOOD
            monthly_fg["Target_Stretch_Finish_Good"] = TARGET_STRETCH_FINISH_GOOD
            monthly_fg["Target_Intervention_Limit"]  = TARGET_INTERVENTION_LIMIT
            monthly_fg["Target_RFT_MMC"]             = TARGET_RFT_MMC

    print(f"[OK] MMC — {len(daily_fg)} daily FG dates | {len(monthly_fg)} monthly records.")
    if not daily_fg.empty:
        print(f"     Total daily Finish Good: {daily_fg['Qte_FG'].sum():.0f} pieces")

    return (
        {
            "MMC_Journalier"        : daily_fg_full,
            "MMC_FinishGood_Mensuel": monthly_fg,
            "MMC_Par_Client"        : by_client,
            "MMC_Par_Operateur"     : by_op,
            "MMC_Tests"             : test_results,
        },
        daily_fg    # returned separately — used as denominator in CNC & Assembly
    )


# =============================================================================
# MODULE 5 — GLOBAL KPIs
#
# Main chart = RFT MMC Global
#
# MONTHLY:
#   RFT MMC Global = (Qte_livrer - (Nbre_defauts_interne + Nbre_defauts_externe))
#                   / Qte_livrer * 100
#
# DAILY:
#   RFT MMC Global = (Qte - (Nbre_defauts_CNC + Quantite_Assembly)) / Qte * 100
#
# YEARLY (YTD):
#   Same formulas as monthly, applied to the YTD row from Depuis 2018
# =============================================================================

def build_global_kpis(cnc_results, asm_results, mmc_results, df_depuis, daily_fg):
    print("\n--- Building Global KPIs ---")

    # ── MONTHLY GLOBAL ────────────────────────────────────────────────────────
    rft_global = pd.DataFrame()
    monthly_src = (
        df_depuis[df_depuis["Scope"] == "Monthly"].copy()
        if df_depuis is not None and not df_depuis.empty
        else pd.DataFrame()
    )

    if not monthly_src.empty:
        rft = monthly_src[["Annee","Mois_N","Mois",
                            "Qte_livrer","Nbre_defauts_interne","Nbre_defauts_externe"]].copy()
        rft["Total_Defauts"] = rft["Nbre_defauts_interne"] + rft["Nbre_defauts_externe"]

        # RFT MMC Global monthly = (Qte_livrer - (interne + externe)) / Qte_livrer * 100
        rft["RFT_MMC_Global_%"] = safe_pct(
            rft["Qte_livrer"] - rft["Total_Defauts"], rft["Qte_livrer"]
        )
        # RFT CNC monthly = (Qte_livrer - interne) / Qte_livrer * 100
        rft["RFT_CNC_%"] = safe_pct(
            rft["Qte_livrer"] - rft["Nbre_defauts_interne"], rft["Qte_livrer"]
        )
        # RFT Assembly monthly = (Qte_livrer - externe) / Qte_livrer * 100
        rft["RFT_Assembly_%"] = safe_pct(
            rft["Qte_livrer"] - rft["Nbre_defauts_externe"], rft["Qte_livrer"]
        )
        rft["Target_RFT_MMC"]      = TARGET_RFT_MMC
        rft["Target_RFT_CNC"]      = TARGET_RFT_CNC
        rft["Target_RFT_Assembly"] = TARGET_RFT_ASSEMBLY
        rft["Finish_Good"]         = rft["Qte_livrer"]
        rft_global = rft.sort_values(["Annee","Mois_N"]).reset_index(drop=True)

    # ── YEARLY / YTD ──────────────────────────────────────────────────────────
    rft_ytd = pd.DataFrame()
    ytd_src = (
        df_depuis[df_depuis["Scope"] == "YTD"].copy()
        if df_depuis is not None and not df_depuis.empty
        else pd.DataFrame()
    )

    if not ytd_src.empty:
        ytd = ytd_src[["Annee","Qte_livrer","Nbre_defauts_interne","Nbre_defauts_externe"]].copy()
        ytd["Total_Defauts"] = ytd["Nbre_defauts_interne"] + ytd["Nbre_defauts_externe"]

        # Yearly formulas = same as monthly
        ytd["RFT_MMC_Global_%"] = safe_pct(
            ytd["Qte_livrer"] - ytd["Total_Defauts"], ytd["Qte_livrer"]
        )
        ytd["RFT_CNC_%"] = safe_pct(
            ytd["Qte_livrer"] - ytd["Nbre_defauts_interne"], ytd["Qte_livrer"]
        )
        ytd["RFT_Assembly_%"] = safe_pct(
            ytd["Qte_livrer"] - ytd["Nbre_defauts_externe"], ytd["Qte_livrer"]
        )
        ytd["Taux_Erreur_CNC_%"] = safe_pct(
            ytd["Nbre_defauts_interne"], ytd["Qte_livrer"]
        )
        ytd["Taux_Erreur_Assembly_%"] = safe_pct(
            ytd["Nbre_defauts_externe"], ytd["Qte_livrer"]
        )
        ytd["Target_RFT_MMC"]      = TARGET_RFT_MMC
        ytd["Target_RFT_CNC"]      = TARGET_RFT_CNC
        ytd["Target_RFT_Assembly"] = TARGET_RFT_ASSEMBLY
        ytd["Finish_Good"]         = ytd["Qte_livrer"]
        rft_ytd = ytd

    # ── DAILY GLOBAL ──────────────────────────────────────────────────────────
    rft_daily = pd.DataFrame()
    d_cnc = cnc_results.get("CNC_Journalier", pd.DataFrame())
    d_asm = asm_results.get("CF_Journalier",  pd.DataFrame())

    if not daily_fg.empty:
        rft_d = daily_fg.copy().rename(columns={"Qte_FG":"Qte"})

        rft_d = rft_d.merge(
            d_cnc[["Date","Nbre_Defauts_CNC"]].rename(columns={"Nbre_Defauts_CNC":"Nbre_Defauts_CNC"}),
            on="Date", how="left"
        ) if not d_cnc.empty else rft_d.assign(Nbre_Defauts_CNC=0)

        rft_d = rft_d.merge(
            d_asm[["Date","Quantite_Assembly"]],
            on="Date", how="left"
        ) if not d_asm.empty else rft_d.assign(Quantite_Assembly=0)

        rft_d["Nbre_Defauts_CNC"]  = rft_d["Nbre_Defauts_CNC"].fillna(0)
        rft_d["Quantite_Assembly"] = rft_d["Quantite_Assembly"].fillna(0)
        rft_d["Total_Defauts"]     = rft_d["Nbre_Defauts_CNC"] + rft_d["Quantite_Assembly"]

        # RFT MMC Global daily = (Qte - (CNC_defects + Assembly_defects)) / Qte * 100
        rft_d["RFT_MMC_Global_%"] = safe_pct(
            rft_d["Qte"] - rft_d["Total_Defauts"], rft_d["Qte"]
        )
        # RFT CNC daily = (Qte - Nbre_defauts_CNC) / Qte * 100
        rft_d["RFT_CNC_%"] = safe_pct(
            rft_d["Qte"] - rft_d["Nbre_Defauts_CNC"], rft_d["Qte"]
        )
        # RFT Assembly daily = (Qte - Quantite_Assembly) / Qte * 100
        rft_d["RFT_Assembly_%"] = safe_pct(
            rft_d["Qte"] - rft_d["Quantite_Assembly"], rft_d["Qte"]
        )
        rft_d["Target_RFT_MMC"]      = TARGET_RFT_MMC
        rft_d["Target_RFT_CNC"]      = TARGET_RFT_CNC
        rft_d["Target_RFT_Assembly"] = TARGET_RFT_ASSEMBLY
        rft_daily = rft_d

    # ── COMBINED MONTHLY KPI TABLE ────────────────────────────────────────────
    cnc_m = cnc_results.get("CNC_Mensuel", pd.DataFrame())
    asm_m = asm_results.get("CF_Mensuel",  pd.DataFrame())
    kpi_mensuel = pd.DataFrame()
    if not cnc_m.empty or not asm_m.empty:
        kpi_mensuel = pd.merge(
            cnc_m, asm_m, on=["Annee","Mois_N","Mois"], how="outer"
        ).fillna(0).sort_values(["Annee","Mois_N"]).reset_index(drop=True)

    print(f"[OK] Global — {len(rft_global)} monthly | {len(rft_daily)} daily | {len(rft_ytd)} YTD.")

    return {
        "RFT_Global"           : rft_global,
        "RFT_Global_Journalier": rft_daily,
        "RFT_YTD"              : rft_ytd,
        "KPI_Mensuel"          : kpi_mensuel,
        "KPI_Global"           : rft_ytd.copy() if not rft_ytd.empty else pd.DataFrame(),
    }


# =============================================================================
# EXPORT — MySQL
# =============================================================================

def get_mysql_connection():
    try:
        return mysql.connector.connect(**MYSQL_CONFIG, connection_timeout=10)
    except Exception as e:
        print(f"[ERROR] MySQL: {e}")
        return None


def insert_dataframe_to_table(conn, df, table_name, column_mapping=None):
    """Drop + recreate table, then insert all rows."""
    if df is None or df.empty:
        return 0
    cur = conn.cursor()
    try:
        df_clean = df.copy()
        df_clean.columns = [
            re.sub(r"[^a-z0-9_]", "_",
                   str(c).lower().replace("%","pct").replace(" ","_")
                          .replace("/","_").replace("-","_"))
            for c in df_clean.columns
        ]
        df_clean = df_clean.where(pd.notna(df_clean), None)

        if column_mapping:
            m = {re.sub(r"[^a-z0-9_]","_",k.lower()): v for k,v in column_mapping.items()}
            df_clean = df_clean.rename(columns=m)

        db_cols  = df_clean.columns.tolist()
        col_defs = ", ".join(f"`{c}` TEXT" for c in db_cols)

        cur.execute(f"DROP TABLE IF EXISTS `{table_name}`")
        cur.execute(f"CREATE TABLE `{table_name}` ({col_defs})")
        conn.commit()

        ph  = ", ".join(["%s"] * len(db_cols))
        sql = f"INSERT INTO `{table_name}` VALUES ({ph})"
        count = 0
        for tup in df_clean.itertuples(index=False, name=None):
            vals = []
            for v in tup:
                if isinstance(v, (datetime, pd.Timestamp)):
                    vals.append(str(v))
                elif hasattr(v, "item"):
                    vals.append(v.item())
                else:
                    vals.append(v)
            try:
                cur.execute(sql, vals)
                count += 1
            except Exception:
                pass
        conn.commit()
        return count
    except Exception as e:
        print(f"[ERROR] insert {table_name}: {e}")
        return 0
    finally:
        cur.close()


def export_to_mysql(suivi, cnc_results, asm_results, mmc_results,
                    global_results, additional_results=None):
    print("\n--- Exporting to MySQL ---")
    conn = get_mysql_connection()
    if conn is None:
        print("[SKIP] MySQL unavailable.")
        return False
    try:
        all_dfs = {"kpi_suivi_journalier": suivi}
        all_dfs.update({f"kpi_{k.lower()}": v for k,v in cnc_results.items()})
        all_dfs.update({f"kpi_{k.lower()}": v for k,v in asm_results.items()})
        all_dfs.update({f"kpi_{k.lower()}": v for k,v in mmc_results.items()})
        all_dfs.update({f"kpi_{k.lower()}": v for k,v in global_results.items()})
        if additional_results:
            for k, v in additional_results.items():
                tbl = re.sub(r"[^a-z0-9_]","_",f"kpi_additional_{k}".lower())
                all_dfs[tbl] = v

        for tbl, df in all_dfs.items():
            n = insert_dataframe_to_table(conn, df, tbl)
            print(f"    {tbl}: {n} rows")
        print("[OK] MySQL export complete.")
        return True
    finally:
        conn.close()


# =============================================================================
# EXPORT — Excel
# =============================================================================

def export_to_excel(suivi, cnc_results, asm_results, mmc_results,
                    global_results, additional_results=None):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILE)
    print(f"\n--- Exporting to Excel: {output_path} ---")

    all_sheets = {"Suivi_Journalier": suivi}
    all_sheets.update(cnc_results)
    all_sheets.update(asm_results)
    all_sheets.update(mmc_results)
    all_sheets.update(global_results)
    if additional_results:
        all_sheets.update(additional_results)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=name[:31], index=False)
                print(f"    [+] {name} ({len(df)} rows)")
    print(f"[OK] Saved: {output_path}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 65)
    print("  KPI AUTOMATION SYSTEM — MMC")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)
    print()
    print("  SOURCE FILES:")
    print(f"    KPI file  : {KPI_FILE}")
    print(f"      Sheet 1 : {SUIVI_SHEET}  (flat table, header=row0)")
    print(f"      Sheet 2 : {DEPUIS_SHEET} (wide/pivoted, left table only, header=row0)")
    print(f"    MMC file  : {MMC_PRODUCTION_FILE}")
    print(f"      Sheet   : List1 (Cyrillic: \\u041b\\u0438\\u0441\\u04421)  (flat table, header=row1  <-- skip 1 blank row)")
    print()
    print("  FORMULAS (all in %):")
    print("  RFT CNC daily        = (Qte - CNC_defects) / Qte * 100")
    print("  Taux Err CNC daily   = CNC_defects / Qte * 100")
    print("  Efficacite CNC       = (pieces_ctrl - CNC_defects) / pieces_ctrl * 100")
    print("                         [CNC module only]")
    print("  RFT CNC monthly      = (Qte_livrer - defauts_interne) / Qte_livrer * 100")
    print("  Taux Err CNC monthly = defauts_interne / Qte_livrer * 100")
    print("  RFT Assembly daily   = (Qte - Assembly_defects) / Qte * 100")
    print("  Taux Err Asm daily   = Assembly_defects / Qte * 100")
    print("  RFT Assembly monthly = (Qte_livrer - defauts_externe) / Qte_livrer * 100")
    print("  Taux Err Asm monthly = defauts_externe / Qte_livrer * 100")
    print("  RFT MMC Global daily = (Qte - (CNC + Assembly)) / Qte * 100")
    print("  RFT MMC Global month = (Qte_livrer - (interne + externe)) / Qte_livrer * 100")
    print("  Finish Good daily    = SUM(Qte) WHERE prototype IS NaN  (from MMC List1)")
    print("  Finish Good monthly  = Quantite livrer  (from Depuis 2018 left table)")
    print("=" * 65)

    create_folders()

    kpi_path = os.path.join(DATA_FOLDER, KPI_FILE)
    mmc_path = os.path.join(DATA_FOLDER, MMC_PRODUCTION_FILE)

    # Each file has a dedicated reader because their structures are fundamentally different:
    #   - Suivi Journalier : flat table, headers in row 0
    #   - Depuis 2018      : wide/pivoted, left table only, years in columns
    #   - MMC List1      : flat table, headers in row 1 (row 0 is blank)
    df_sj     = read_suivi_journalier(kpi_path) if os.path.exists(kpi_path)  else pd.DataFrame()
    df_depuis = read_depuis_2018(kpi_path)       if os.path.exists(kpi_path)  else pd.DataFrame()
    df_mmc    = read_mmc_production(mmc_path)    if os.path.exists(mmc_path)  else pd.DataFrame()

    if not os.path.exists(kpi_path):
        print(f"[ERROR] KPI file not found: {kpi_path}")
    if not os.path.exists(mmc_path):
        print(f"[ERROR] MMC file not found: {mmc_path}")

    # Build Suivi Journalier export table
    suivi = build_suivi_journalier(df_sj) if not df_sj.empty else pd.DataFrame()

    # Build MMC first — daily_fg is the shared daily denominator for CNC and Assembly
    mmc_results, daily_fg = build_mmc_kpis(df_mmc, df_depuis)

    cnc_results = build_cnc_kpis(df_sj, df_depuis, daily_fg) if not df_sj.empty else {}
    asm_results = build_assembly_kpis(df_sj, df_depuis, daily_fg) if not df_sj.empty else {}
    global_results = build_global_kpis(
        cnc_results, asm_results, mmc_results, df_depuis, daily_fg
    ) if (not df_sj.empty or not df_depuis.empty) else {}

    # Scan data/ for additional Excel files
    print("\n--- Processing Additional Data Files ---")
    additional_results = process_additional_excel_files()

    # Export
    if EXPORT_MODE == 'mysql':
        export_to_mysql(suivi, cnc_results, asm_results, mmc_results,
                        global_results, additional_results)
    else:
        export_to_excel(suivi, cnc_results, asm_results, mmc_results,
                        global_results, additional_results)

    # ── Update dashboard.html RAW array automatically ─────────────────────────
    cnc_path = os.path.join(DATA_FOLDER, CNC_DEFAUTS_FILE)
    asm_path = os.path.join(DATA_FOLDER, ASSEMBLY_FILE)
    update_dashboard_raw(cnc_path, asm_path, mmc_path, DASHBOARD_HTML_PATH)

    print("\n" + "=" * 65)
    print(f"  DONE — {'MySQL kpi_dashboard' if EXPORT_MODE=='mysql' else OUTPUT_FILE}")
    print("=" * 65)



# =============================================================================
# READER D — Suivi_defauts_CNC.xlsx  (Feuil2)
#
# Layout : flat table, headers in row 0.
# All rows from 2026 onwards are used regardless of 'detecter par' value
# because in this file all rows are CNC defects by definition.
# Key columns: Date | Nbre piéces controlés | Nbre des défauts
# =============================================================================

def read_cnc_defauts(cnc_filepath):
    """Read Suivi_defauts_CNC.xlsx -> Feuil2 for 2026+ data."""
    print(f"\n[READ-D] CNC Defauts -> {cnc_filepath}")
    try:
        df = pd.read_excel(cnc_filepath, sheet_name="Feuil2", header=0)
    except Exception as e:
        print(f"  [ERROR] Cannot open: {e}")
        return pd.DataFrame()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Nbre des défauts"]       = pd.to_numeric(df["Nbre des défauts"],       errors="coerce").fillna(0)
    df["Nbre piéces controlés"]  = pd.to_numeric(df["Nbre piéces controlés"],  errors="coerce").fillna(0)
    df = df.dropna(subset=["Date"])
    df = df[df["Date"].dt.year >= 2026].copy()
    print(f"  {len(df)} rows | months: {sorted(df['Date'].dt.month.unique().tolist())}")
    return df


# =============================================================================
# READER E — Saisie_controle_Finale.xlsx  (Feuil3)
#
# Layout : flat table, headers in row 0.
# Key columns: Date | Quantité  (= number of defective assembly pieces)
# All 2026+ rows are used.
# =============================================================================

def read_assembly_defauts(asm_filepath):
    """Read Saisie_controle_Finale.xlsx -> Feuil3 for 2026+ data."""
    print(f"\n[READ-E] Assembly Defauts -> {asm_filepath}")
    try:
        df = pd.read_excel(asm_filepath, sheet_name="Feuil3", header=0)
    except Exception as e:
        print(f"  [ERROR] Cannot open: {e}")
        return pd.DataFrame()

    df["Date"]     = pd.to_datetime(df["Date"], errors="coerce")
    df["Quantité"] = pd.to_numeric(df["Quantité"], errors="coerce").fillna(0)
    df = df.dropna(subset=["Date"])
    df = df[df["Date"].dt.year >= 2026].copy()
    print(f"  {len(df)} rows | months: {sorted(df['Date'].dt.month.unique().tolist())}")
    return df


# =============================================================================
# UPDATE DASHBOARD RAW ARRAY
#
# Reads the three new source files, computes daily KPIs using the formulas
# from requirements, then rewrites the RAW=[...] block inside dashboard.html.
#
# This is called at the end of main() so every 5-minute automation cycle
# automatically updates the dashboard with any new months added to the files.
#
# Column mapping to RAW fields:
#   qte      = SUM(Qte) per date from MMC Лист1 (ALL rows including prototype)
#   cnc_def  = SUM(Nbre des défauts) per date from Suivi_defauts_CNC Feuil2
#   cnc_pcs  = SUM(Nbre piéces controlés) per date from Suivi_defauts_CNC Feuil2
#   asm_def  = SUM(Quantité) per date from Saisie_controle_Finale Feuil3
#   fg       = SUM(Qte) WHERE prototype IS NULL per date from MMC Лист1
# =============================================================================

def update_dashboard_raw(cnc_filepath, asm_filepath, mmc_filepath, dashboard_path):
    """
    Read the three source Excel files, compute daily KPIs for 2026+,
    and rewrite the RAW=[...] block inside dashboard.html.
    """
    print("\n--- Updating dashboard.html RAW data ---")

    if not os.path.exists(dashboard_path):
        print(f"  [SKIP] dashboard.html not found at: {dashboard_path}")
        return False

    # ── Load CNC defects ──────────────────────────────────────────────────────
    if os.path.exists(cnc_filepath):
        df_cnc = read_cnc_defauts(cnc_filepath)
    else:
        print(f"  [WARN] CNC file not found: {cnc_filepath} — CNC columns will be 0")
        df_cnc = pd.DataFrame()

    # ── Load Assembly defects ─────────────────────────────────────────────────
    if os.path.exists(asm_filepath):
        df_asm = read_assembly_defauts(asm_filepath)
    else:
        print(f"  [WARN] Assembly file not found: {asm_filepath} — Assembly columns will be 0")
        df_asm = pd.DataFrame()

    # ── Load MMC production (Qte + Finish Good) ───────────────────────────────
    if not os.path.exists(mmc_filepath):
        print(f"  [ERROR] MMC file not found: {mmc_filepath}")
        return False

    try:
        df_mmc = pd.read_excel(mmc_filepath, sheet_name="\u041b\u0438\u0441\u04421", header=1)
    except Exception as e:
        print(f"  [ERROR] Cannot read MMC file: {e}")
        return False

    df_mmc["DATE"] = pd.to_datetime(df_mmc["DATE"], errors="coerce")
    df_mmc["Qte"]  = pd.to_numeric(df_mmc["Qte"],  errors="coerce").fillna(0)
    df_mmc = df_mmc.dropna(subset=["DATE"])
    df_mmc = df_mmc[df_mmc["DATE"].dt.year >= 2026].copy()

    if df_mmc.empty:
        print("  [ERROR] No 2026+ data in MMC file")
        return False

    # Total Qte per day (all rows)
    qte_daily = df_mmc.groupby("DATE")["Qte"].sum().reset_index()
    qte_daily.columns = ["Date", "qte"]

    # Finish Good = rows where prototype IS NULL / empty
    proto_col = "prototype"
    if proto_col in df_mmc.columns:
        proto_str = df_mmc[proto_col].fillna("").astype(str).str.strip().str.lower()
        fg_mask   = df_mmc[proto_col].isna() | proto_str.isin(["", "nan", "none"])
    else:
        fg_mask = pd.Series(True, index=df_mmc.index)
    fg_daily = df_mmc[fg_mask].groupby("DATE")["Qte"].sum().reset_index()
    fg_daily.columns = ["Date", "fg"]

    # ── Aggregate CNC per day ─────────────────────────────────────────────────
    if not df_cnc.empty:
        cnc_daily = df_cnc.groupby("Date").agg(
            cnc_def=("Nbre des défauts",      "sum"),
            cnc_pcs=("Nbre piéces controlés", "sum"),
        ).reset_index()
    else:
        cnc_daily = pd.DataFrame(columns=["Date", "cnc_def", "cnc_pcs"])

    # ── Aggregate Assembly per day ────────────────────────────────────────────
    if not df_asm.empty:
        asm_daily = df_asm.groupby("Date")["Quantité"].sum().reset_index()
        asm_daily.columns = ["Date", "asm_def"]
    else:
        asm_daily = pd.DataFrame(columns=["Date", "asm_def"])

    # ── Merge all into one daily table ────────────────────────────────────────
    daily = qte_daily.copy()
    daily = daily.merge(cnc_daily, on="Date", how="left")
    daily = daily.merge(asm_daily, on="Date", how="left")
    daily = daily.merge(fg_daily,  on="Date", how="left")
    daily = daily.fillna(0).sort_values("Date").reset_index(drop=True)

    # Cast to int for clean JS output
    for col in ["qte", "cnc_def", "cnc_pcs", "asm_def", "fg"]:
        daily[col] = daily[col].astype(int)

    if daily.empty:
        print("  [ERROR] No daily data to write")
        return False

    print(f"  {len(daily)} days | months: {sorted(daily['Date'].dt.month.unique().tolist())}")

    # ── Build JS RAW array string ─────────────────────────────────────────────
    MO_NAMES = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jun",
                7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
    MONTH_COMMENTS = {}
    prev_month = None
    for _, r in daily.iterrows():
        mn = int(r["Date"].month)
        if mn != prev_month:
            MONTH_COMMENTS[int(r["Date"].strftime("%j"))] = MO_NAMES.get(mn, str(mn)) + " " + str(int(r["Date"].year))
            prev_month = mn

    lines = ["const RAW=["]
    prev_m = None
    for _, r in daily.iterrows():
        mn = int(r["Date"].month)
        yr = int(r["Date"].year)
        if mn != prev_m:
            mo_name = MO_NAMES.get(mn, str(mn))
            lines.append(f"  // \u2500\u2500 {mo_name} {yr} " + "\u2500" * 40)
            prev_m = mn
        d_str = r["Date"].strftime("%d/%m")
        lines.append(
            f"  {{d:'{d_str}',m:{mn},y:{yr},"
            f"qte:{int(r['qte'])},cnc_def:{int(r['cnc_def'])},"
            f"cnc_pcs:{int(r['cnc_pcs'])},asm_def:{int(r['asm_def'])},"
            f"fg:{int(r['fg'])}}},")

    lines.append("  // \u2500\u2500 ADD NEW DAYS/MONTHS HERE \u2014 system auto-discovers them")
    lines.append("];")
    new_raw_block = "\n".join(lines)

    # ── Rewrite dashboard.html ─────────────────────────────────────────────────
    with open(dashboard_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Find the RAW=[...]; block using regex
    pattern = re.compile(r"const RAW=\[.*?^\];", re.DOTALL | re.MULTILINE)
    if not pattern.search(html):
        print("  [ERROR] RAW=[...]; block not found in dashboard.html")
        print("  Make sure dashboard.html contains 'const RAW=[' and '];'")
        return False

    new_html = pattern.sub(new_raw_block, html, count=1)

    with open(dashboard_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    months_found = sorted(daily["Date"].dt.month.unique().tolist())
    month_names  = [MO_NAMES.get(m, str(m)) for m in months_found]
    print(f"  [OK] dashboard.html updated — {len(daily)} rows, months: {month_names}")
    return True


if __name__ == "__main__":
    main()