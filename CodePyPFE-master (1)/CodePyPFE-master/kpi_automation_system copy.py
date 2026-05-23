# -*- coding: utf-8 -*-
# =============================================================================
# KPI AUTOMATION SYSTEM
# Project : Digitalisation de KPI - MMC
# Author  : [Your Name]
# Date    : 2026
# =============================================================================
# DESCRIPTION:
#   Reads three operational Excel files, calculates KPIs matching the
#   supervisor's Monthly PARETO Report style, and exports to Excel
#   ready to be connected to Power BI.
#
# OUTPUT SHEETS:
#   -- CNC QUALITY --
#   1.  Suivi_Journalier         Daily log (mirrors KPI_S_2027)
#   2.  CNC_Mensuel              Monthly CNC KPIs + targets
#   3.  CNC_RFT_Mensuel          RFT-CNC monthly (Right First Time)
#   4.  CNC_Par_Operateur        Defects by operator
#   5.  CNC_Par_Statut           Defects by status
#   6.  CNC_Par_Piece            Defects by piece
#   7.  CNC_Pareto_Defauts       Pareto of defect types
#   8.  CNC_Pareto_Actions       Pareto of corrective actions
#   9.  CNC_Par_Origine          Defects by origin (CNC/FAO/CAO...)
#   -- FINAL CONTROL --
#   10. CF_Mensuel               Monthly Final Control KPIs
#   11. CF_RFT_Mensuel           RFT-Assembly monthly
#   12. CF_Par_Origine           Non-conformities by origin
#   13. CF_Par_Technicien        Non-conformities by technician
#   14. CF_Pareto_Actions        Pareto of corrective actions
#   15. CF_Scrap_Rework          Scrap/Rework tracking
#   -- MMC PRODUCTION --
#   16. MMC_Journalier           Daily production
#   17. MMC_FinishGood_Mensuel   Finish Good rate monthly + targets
#   18. MMC_Par_Client           Production by client
#   19. MMC_Tests                Test results summary
#   -- GLOBAL --
#   20. KPI_Mensuel              Combined monthly KPIs
#   21. KPI_Global               Annual summary
#   22. RFT_Global               Global RFT rates monthly
#
# HOW TO USE:
#   1. Place Excel files in the 'data/' folder
#   2. Run: py kpi_automation_system.py
#   3. Open output/KPI_Output.xlsx and connect to Power BI
#
# TO ADD A NEW DATA SOURCE:
#   1. Add file/sheet/column names in CONFIGURATION section
#   2. Create a new process_xxx() function following existing patterns
#   3. Call it in main() and add to export_to_excel()
# =============================================================================

import pandas as pd
import os
from datetime import datetime

# =============================================================================
# CONFIGURATION
# =============================================================================

DATA_FOLDER   = "data"
OUTPUT_FOLDER = "output"

CNC_FILE            = "Suivi_defauts_CNC.xlsx"
FINAL_CONTROL_FILE  = "Saisie_controle_Finale.xlsx"
MMC_PRODUCTION_FILE = "01-Rapport_quantite_MMC.xlsx"
OUTPUT_FILE         = "KPI_Output.xlsx"

CNC_SHEET            = "Feuil2"
FINAL_CONTROL_SHEET  = "Feuil3"
MMC_PRODUCTION_SHEET = "\u041b\u0438\u0441\u04421"

# Targets (from supervisor's PARETO report)
TARGET_FINISH_GOOD          = 40    # pieces per day target
TARGET_STRETCH_FINISH_GOOD  = 45    # stretch target
TARGET_INTERVENTION_LIMIT   = 35    # intervention limit
TARGET_ERROR_RATE_CNC       = 0.06  # 6% max error rate CNC
TARGET_ERROR_RATE_ASSEMBLY  = 0.10  # 10% max error rate Assembly
TARGET_RFT_CNC              = 0.94  # 94% Right First Time CNC
TARGET_RFT_ASSEMBLY         = 0.90  # 90% Right First Time Assembly
TARGET_RFT_MMC              = 0.80  # 80% Right First Time Global

# French month names
MOIS_NOMS = {
    1: "Janvier",  2: "F\u00e9vrier", 3: "Mars",
    4: "Avril",    5: "Mai",           6: "Juin",
    7: "Juillet",  8: "Ao\u00fbt",    9: "Septembre",
    10: "Octobre", 11: "Novembre",    12: "D\u00e9cembre"
}

# Origin categories for grouping
CNC_ORIGINS    = ["CNC", "cnc", "CNC+FAO", "CNC + CAO", "CNC+FAO+CAO"]
FAO_ORIGINS    = ["FAO", "fao", "CAO+FAO"]
CAO_ORIGINS    = ["CAO", "cao", "CAO+FAO"]
MONTAGE_ORIGINS = ["MONTAGE", "montage", "magazin", "mag+cnc"]

# ---------------------------------------------------------------------------
# CNC FILE — Column names
# ---------------------------------------------------------------------------
CNC_COL_DATE        = "Date"
CNC_COL_OPERATOR    = "Controler par"
CNC_COL_ORDER       = "ordre"
CNC_COL_STATUS      = "Statu"
CNC_COL_PIECES      = "Nbre pi\u00e9ces control\u00e9s"
CNC_COL_PIECE_NAME  = "Noms des pieces"
CNC_COL_DEFECTS     = "Nbre des d\u00e9fauts"
CNC_COL_NIO         = "Pi\u00e9ce non conforme"
CNC_COL_DESC        = "Description du defaut"
CNC_COL_QRQC        = "N\u00b0QRQC"
CNC_COL_ACTION      = "action de correction"
CNC_COL_CAO         = "CAO"
CNC_COL_FAUTIF      = "Fautif"
CNC_COL_DETECTED_BY = "detecter par"
CNC_COL_ETAT        = "Etat"

# ---------------------------------------------------------------------------
# FINAL CONTROL FILE — Column names
# ---------------------------------------------------------------------------
FC_COL_DATE         = "Date"
FC_COL_DETECTED_IN  = "d\u00e9tecter dans"
FC_COL_TECHNICIAN   = "Technicien QM"
FC_COL_ORDER        = "Work Order"
FC_COL_STATUS       = "Statu"
FC_COL_QTY_CTRL     = "quantit\u00e9 control\u00e9"
FC_COL_QUANTITY     = "Quantit\u00e9"
FC_COL_NIO          = "Pi\u00e9ce NIO"
FC_COL_DESC         = "D\u00e9scription du probl\u00e9me "
FC_COL_QRQC         = "N\u00b0 QRQC"
FC_COL_ACTION       = "Action de correction"
FC_COL_ORIGIN       = "Origine du probl\u00e9me"
FC_COL_FEEDBACK     = "Retour feedback"
FC_COL_ETAT         = "Etat"
FC_COL_STUFF        = "Stuff NM"

# ---------------------------------------------------------------------------
# MMC PRODUCTION FILE — Column names
# ---------------------------------------------------------------------------
MMC_COL_DATE        = "DATE"
MMC_COL_QTY         = "Qte"
MMC_COL_PART        = "LEONI-part number:"
MMC_COL_TYPE        = "type"
MMC_COL_PRODUCER    = "Producer:"
MMC_COL_CLIENT      = "Client"
MMC_COL_OPERATOR    = "Op\u00e9rateur"
MMC_COL_CAD         = "CAD"
MMC_COL_SEC_LOCK    = "Secondary locking"
MMC_COL_VERZT       = "VerZt-Prufung"
MMC_COL_AUSTRAST    = "Austrast"
MMC_COL_SEC_INLOCK  = "Secondary inlocking"
MMC_COL_OFFSET      = "Offset test:"
MMC_COL_PUSHBACK    = "Push-back test"
MMC_COL_UNLOCK      = "Unlocking test:"
MMC_COL_HOUSING     = "Hosuing attachments:"
MMC_COL_ETANCH      = "etanchiete"
MMC_COL_COULEUR     = "couleur"
MMC_COL_ORDER       = "N work order"
MMC_COL_PROTOTYPE   = "prototype"

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def create_folders():
    os.makedirs(DATA_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    print("[OK] Folders verified.")


def load_excel(filename, sheet_name, skip_rows=0):
    filepath = os.path.join(DATA_FOLDER, filename)
    if not os.path.exists(filepath):
        print(f"[WARNING] File not found: {filepath}")
        return None
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=skip_rows)
        print(f"[OK] Loaded '{filename}' — {len(df)} rows found.")
        return df
    except Exception as e:
        print(f"[ERROR] Could not read '{filename}': {e}")
        return None


def clean_dataframe(df, date_column):
    df = df.dropna(how="all").copy()
    df[date_column] = pd.to_datetime(df[date_column], errors="coerce")
    df = df.dropna(subset=[date_column])
    df["Annee"]   = df[date_column].dt.year
    df["Mois_N"]  = df[date_column].dt.month
    df["Mois"]    = df["Mois_N"].map(MOIS_NOMS)
    df["Semaine"] = df[date_column].dt.isocalendar().week.astype(int)
    df["KW"]      = df["Semaine"]
    return df


def safe_rate(numerator, denominator):
    """Safe division returning 0 when denominator is 0."""
    return (numerator / denominator.replace(0, pd.NA) * 100).fillna(0).round(2)


def add_targets(df, col_map):
    """Add target columns to any dataframe."""
    for col_name, value in col_map.items():
        df[col_name] = value
    return df


def build_pareto(df, group_col, value_col, top_n=20):
    """Build a Pareto table with cumulative % column."""
    pareto = df.groupby(group_col)[value_col].sum().reset_index()
    pareto = pareto.sort_values(value_col, ascending=False).head(top_n)
    pareto["Cumul_%"] = (
        pareto[value_col].cumsum() / pareto[value_col].sum() * 100
    ).round(1)
    return pareto


# =============================================================================
# MODULE 1 — SUIVI JOURNALIER
# =============================================================================

def build_suivi_journalier(df_cnc):
    """Daily log — mirrors 'Suivi Journalier des fautes' in KPI_S_2027."""
    print("\n--- Building Suivi Journalier ---")

    df = df_cnc.copy()
    df[CNC_COL_OPERATOR] = df[CNC_COL_OPERATOR].ffill()
    df[CNC_COL_PIECES]   = pd.to_numeric(df[CNC_COL_PIECES],  errors="coerce").fillna(0)
    df[CNC_COL_DEFECTS]  = pd.to_numeric(df[CNC_COL_DEFECTS], errors="coerce").fillna(0)

    suivi = pd.DataFrame({
        "KW"                       : df["KW"],
        "Mois"                     : df["Mois"],
        "Date"                     : df[CNC_COL_DATE],
        "Detecter par"             : df[CNC_COL_DETECTED_BY],
        "Controler par"            : df[CNC_COL_OPERATOR],
        "N\u00b0Order"            : df[CNC_COL_ORDER],
        "Statu"                    : df[CNC_COL_STATUS],
        "Nombre pieces controlees" : df[CNC_COL_PIECES],
        "Nombre defauts"           : df[CNC_COL_DEFECTS],
        "Piece NIO"                : df[CNC_COL_NIO],
        "Type technique du defaut" : df[CNC_COL_CAO],
        "Description du defaut"    : df[CNC_COL_DESC],
        "N\u00b0QRQC"             : df[CNC_COL_QRQC],
        "Action de correction"     : df[CNC_COL_ACTION],
        "Origine de defaut"        : df[CNC_COL_CAO],
        "Etat"                     : df[CNC_COL_ETAT],
    })

    print(f"[OK] Suivi Journalier — {len(suivi)} rows.")
    return suivi


# =============================================================================
# MODULE 2 — CNC KPIs (Error Rate + RFT + Pareto)
# =============================================================================

def build_cnc_kpis(df_cnc):
    """
    Build all CNC KPI tables matching the supervisor's PARETO report:
    - Monthly error rate vs target 6%
    - RFT-CNC monthly vs target 94%
    - Pareto of defect types
    - Pareto of corrective actions
    - Defects by operator / status / piece / origin
    """
    print("\n--- Building CNC KPIs ---")

    df = df_cnc.copy()
    df[CNC_COL_OPERATOR] = df[CNC_COL_OPERATOR].ffill()
    df[CNC_COL_PIECES]   = pd.to_numeric(df[CNC_COL_PIECES],  errors="coerce").fillna(0)
    df[CNC_COL_DEFECTS]  = pd.to_numeric(df[CNC_COL_DEFECTS], errors="coerce").fillna(0)

    # RFT = orders with ZERO defects / total orders
    df["Est_RFT"] = (df[CNC_COL_DEFECTS] == 0).astype(int)

    # Normalize origin column
    df["Origine_Norm"] = df[CNC_COL_CAO].str.strip().str.upper()
    df["Origine_Norm"] = df["Origine_Norm"].replace({
        "CNC": "CNC", "CNc": "CNC", "cnc": "CNC",
        "FAO": "FAO", "fao": "FAO",
        "CAO": "CAO", "cao": "CAO",
        "MONTAGE": "MONTAGE",
    })

    # --- Monthly CNC KPIs ---
    monthly = df.groupby(["Annee", "Mois_N", "Mois"]).agg(
        Pieces_Controlees = (CNC_COL_PIECES,  "sum"),
        Nombre_Defauts    = (CNC_COL_DEFECTS, "sum"),
        Nb_Ordres         = (CNC_COL_ORDER,   "count"),
        Nb_RFT            = ("Est_RFT",        "sum"),
    ).reset_index()

    monthly["Taux_Erreur_CNC_%"]    = safe_rate(monthly["Nombre_Defauts"], monthly["Pieces_Controlees"])
    monthly["Objectif_%_Pieces"]    = TARGET_ERROR_RATE_CNC * 100
    monthly["RFT_CNC"]              = (monthly["Nb_RFT"] / monthly["Nb_Ordres"].replace(0, pd.NA)).fillna(0).round(4)
    monthly["Objectif_RFT_CNC"]     = TARGET_RFT_CNC
    monthly["Efficacite_CNC"]       = (1 - monthly["Taux_Erreur_CNC_%"] / 100).round(4)
    monthly["Objectif_Efficacite"]  = 1 - TARGET_ERROR_RATE_CNC

    # --- RFT Weekly ---
    weekly_rft = df.groupby(["Annee", "KW", "Mois"]).agg(
        Pieces_Controlees = (CNC_COL_PIECES,  "sum"),
        Nombre_Defauts    = (CNC_COL_DEFECTS, "sum"),
        Nb_Ordres         = (CNC_COL_ORDER,   "count"),
        Nb_RFT            = ("Est_RFT",        "sum"),
    ).reset_index()
    weekly_rft["Taux_Erreur_%"] = safe_rate(weekly_rft["Nombre_Defauts"], weekly_rft["Pieces_Controlees"])
    weekly_rft["RFT_CNC"]       = (weekly_rft["Nb_RFT"] / weekly_rft["Nb_Ordres"].replace(0, pd.NA)).fillna(0).round(4)
    weekly_rft["Objectif_RFT"]  = TARGET_RFT_CNC

    # --- Pareto Defect Types ---
    defects_only = df[df[CNC_COL_DEFECTS] > 0]
    pareto_defauts = build_pareto(defects_only, CNC_COL_DESC, CNC_COL_DEFECTS)
    pareto_defauts.rename(columns={
        CNC_COL_DESC: "Description_Defaut",
        CNC_COL_DEFECTS: "Nombre_Defauts"
    }, inplace=True)

    # --- Pareto Corrective Actions ---
    pareto_actions = build_pareto(defects_only, CNC_COL_ACTION, CNC_COL_DEFECTS)
    pareto_actions.rename(columns={
        CNC_COL_ACTION: "Action_Correction",
        CNC_COL_DEFECTS: "Nombre_Defauts"
    }, inplace=True)

    # --- By Operator ---
    by_op = df.groupby(CNC_COL_OPERATOR).agg(
        Pieces_Controlees = (CNC_COL_PIECES,  "sum"),
        Nombre_Defauts    = (CNC_COL_DEFECTS, "sum"),
    ).reset_index()
    by_op["Taux_Defauts_%"] = safe_rate(by_op["Nombre_Defauts"], by_op["Pieces_Controlees"])
    by_op.rename(columns={CNC_COL_OPERATOR: "Operateur"}, inplace=True)

    # --- By Status ---
    by_status = df.groupby(CNC_COL_STATUS).agg(
        Pieces_Controlees = (CNC_COL_PIECES,  "sum"),
        Nombre_Defauts    = (CNC_COL_DEFECTS, "sum"),
    ).reset_index()
    by_status["Taux_Defauts_%"] = safe_rate(by_status["Nombre_Defauts"], by_status["Pieces_Controlees"])
    by_status.rename(columns={CNC_COL_STATUS: "Statut"}, inplace=True)

    # --- By Piece ---
    by_piece = df.groupby(CNC_COL_PIECE_NAME).agg(
        Pieces_Controlees = (CNC_COL_PIECES,  "sum"),
        Nombre_Defauts    = (CNC_COL_DEFECTS, "sum"),
    ).reset_index().sort_values("Nombre_Defauts", ascending=False)
    by_piece["Taux_Defauts_%"] = safe_rate(by_piece["Nombre_Defauts"], by_piece["Pieces_Controlees"])
    by_piece.rename(columns={CNC_COL_PIECE_NAME: "Nom_Piece"}, inplace=True)

    # --- By Origin ---
    by_origin = df.groupby("Origine_Norm").agg(
        Nombre_Defauts = (CNC_COL_DEFECTS, "sum"),
    ).reset_index().sort_values("Nombre_Defauts", ascending=False)
    by_origin.rename(columns={"Origine_Norm": "Origine"}, inplace=True)

    print(f"[OK] CNC KPIs — {len(monthly)} monthly records.")
    print(f"     Avg RFT-CNC     : {monthly['RFT_CNC'].mean():.2%}")
    print(f"     Avg Error Rate  : {monthly['Taux_Erreur_CNC_%'].mean():.2f}%")

    return {
        "CNC_Mensuel"         : monthly,
        "CNC_RFT_Hebdomadaire": weekly_rft,
        "CNC_Pareto_Defauts"  : pareto_defauts,
        "CNC_Pareto_Actions"  : pareto_actions,
        "CNC_Par_Operateur"   : by_op,
        "CNC_Par_Statut"      : by_status,
        "CNC_Par_Piece"       : by_piece,
        "CNC_Par_Origine"     : by_origin,
    }


# =============================================================================
# MODULE 3 — FINAL CONTROL KPIs (Assembly Error Rate + RFT + Scrap/Rework)
# =============================================================================

def build_fc_kpis(df_fc):
    """
    Build Final Control KPI tables matching supervisor's report:
    - Assembly error rate vs target 10%
    - RFT-Assembly monthly vs target 90%
    - Scrap/Rework tracking
    - Pareto of corrective actions
    - By origin / technician
    """
    print("\n--- Building Final Control KPIs ---")

    df = df_fc.copy()
    df[FC_COL_QUANTITY] = pd.to_numeric(df[FC_COL_QUANTITY], errors="coerce").fillna(0)
    df[FC_COL_QTY_CTRL] = pd.to_numeric(df[FC_COL_QTY_CTRL], errors="coerce").fillna(0)

    # Non-conformity flag
    df["Est_NIO"] = df[FC_COL_DESC].notna() & (df[FC_COL_DESC].astype(str).str.strip() != "")

    # RFT = rows that are conforming (no problem)
    df["Est_RFT"] = (~df["Est_NIO"]).astype(int)

    # --- Monthly Final Control KPIs ---
    monthly = df.groupby(["Annee", "Mois_N", "Mois"]).agg(
        Quantite_Inspectee     = (FC_COL_QUANTITY, "sum"),
        Nb_Non_Conformites     = ("Est_NIO",        "sum"),
        Nb_Conforme            = ("Est_RFT",        "sum"),
        Nb_Total               = (FC_COL_ORDER,     "count"),
    ).reset_index()

    # Alias for backward compatibility with existing Power BI dashboards
    monthly["Quantite_NIO"]           = monthly["Quantite_Inspectee"]
    monthly["Nombre_Non_Conformites"] = monthly["Nb_Non_Conformites"]

    monthly["Taux_Erreur_Assembly_%"] = safe_rate(
        monthly["Nb_Non_Conformites"], monthly["Nb_Total"]
    )
    monthly["Objectif_%_Module"]      = TARGET_ERROR_RATE_ASSEMBLY * 100
    monthly["RFT_Assembly"]           = (
        monthly["Nb_Conforme"] / monthly["Nb_Total"].replace(0, pd.NA)
    ).fillna(0).round(4)
    monthly["Objectif_RFT_Assembly"]  = TARGET_RFT_ASSEMBLY

    # --- Scrap / Rework Tracking ---
    scrap_keywords = ["rebut", "remplacement", "refaire", "annul", "r\u00e9paration"]
    df["Action_Lower"] = df[FC_COL_ACTION].fillna("").astype(str).str.lower()

    def classify_action(action):
        if not action or action == "nan":
            return "Autre"
        if "rebut" in action:
            return "Mise en rebut"
        elif "remplacement" in action or "remplacer" in action:
            return "Demande de remplacement"
        elif "refaire" in action or "r\u00e9parer" in action or "r\u00e9paration" in action:
            return "Refaire la pi\u00e8ce"
        elif "annul" in action:
            return "Annulation"
        else:
            return "Autre"

    df["Type_Action"] = df["Action_Lower"].apply(classify_action)

    scrap_monthly = df.groupby(["Annee", "Mois_N", "Mois", "Type_Action"]).size().reset_index(name="Total")
    scrap_monthly = scrap_monthly.sort_values(["Annee", "Mois_N", "Type_Action"])

    # --- Pareto Corrective Actions ---
    pareto_actions = build_pareto(
        df[df["Est_NIO"]], FC_COL_ACTION, FC_COL_QUANTITY
    )
    pareto_actions.rename(columns={
        FC_COL_ACTION: "Action_Correction",
        FC_COL_QUANTITY: "Quantite"
    }, inplace=True)

    # --- By Origin ---
    by_origin = df.groupby(FC_COL_ORIGIN).agg(
        Nb_Non_Conformites = ("Est_NIO", "sum"),
    ).reset_index().sort_values("Nb_Non_Conformites", ascending=False)
    by_origin.rename(columns={FC_COL_ORIGIN: "Origine"}, inplace=True)

    # --- By Technician ---
    by_tech = df.groupby(FC_COL_TECHNICIAN).agg(
        Nb_Non_Conformites = ("Est_NIO", "sum"),
        Nb_Total           = (FC_COL_ORDER, "count"),
    ).reset_index()
    by_tech["Taux_NC_%"] = safe_rate(by_tech["Nb_Non_Conformites"], by_tech["Nb_Total"])
    by_tech.rename(columns={FC_COL_TECHNICIAN: "Technicien"}, inplace=True)

    print(f"[OK] Final Control KPIs — {len(monthly)} monthly records.")
    print(f"     Avg RFT-Assembly : {monthly['RFT_Assembly'].mean():.2%}")

    return {
        "CF_Mensuel"        : monthly,
        "CF_Scrap_Rework"   : scrap_monthly,
        "CF_Pareto_Actions" : pareto_actions,
        "CF_Par_Origine"    : by_origin,
        "CF_Par_Technicien" : by_tech,
    }


# =============================================================================
# MODULE 4 — MMC PRODUCTION (Finish Good Rate + RFT Global)
# =============================================================================

def build_mmc_kpis(df_mmc):
    """
    Build MMC Production KPI tables matching supervisor's report:
    - Finish Good rate monthly vs targets (35/40/45)
    - Prototyping rate monthly
    - RFT-MMC global
    - Production by client
    - Test results summary
    """
    print("\n--- Building MMC Production KPIs ---")

    df = df_mmc.copy()
    df[MMC_COL_QTY] = pd.to_numeric(df[MMC_COL_QTY], errors="coerce").fillna(0)

    # Finish Good = type 'N' (normal production)
    # Prototype = type 'P' or CAD column has value
    df["Is_Prototype"]    = df[MMC_COL_PROTOTYPE].notna()
    df["Is_Finish_Good"]  = ~df["Is_Prototype"]

    # --- Daily Production ---
    daily = df.groupby(MMC_COL_DATE).agg(
        Quantite_Produite  = (MMC_COL_QTY, "sum"),
        Nb_Finish_Good     = ("Is_Finish_Good", "sum"),
        Nb_Prototype       = ("Is_Prototype",   "sum"),
    ).reset_index()
    daily.rename(columns={MMC_COL_DATE: "Date"}, inplace=True)
    daily["Taux_Prototypage"]  = (daily["Nb_Prototype"] / daily["Quantite_Produite"].replace(0, pd.NA)).fillna(0).round(4)
    daily["Tested_Good"]       = daily["Quantite_Produite"] - daily["Nb_Prototype"]

    # --- Monthly Finish Good ---
    monthly_fg = df.groupby(["Annee", "Mois_N", "Mois"]).agg(
        Quantite_Produite = (MMC_COL_QTY,        "sum"),
        Nb_Prototype      = ("Is_Prototype",      "sum"),
        Nb_Finish_Good    = ("Is_Finish_Good",    "sum"),
    ).reset_index()
    monthly_fg["Tested_Good"]              = monthly_fg["Quantite_Produite"] - monthly_fg["Nb_Prototype"]
    monthly_fg["Taux_Prototypage"]         = (
        monthly_fg["Nb_Prototype"] / monthly_fg["Quantite_Produite"].replace(0, pd.NA)
    ).fillna(0).round(4)
    monthly_fg["Quantity_Target"]          = TARGET_FINISH_GOOD
    monthly_fg["Quantity_Stretch_Target"]  = TARGET_STRETCH_FINISH_GOOD
    monthly_fg["Quantity_Intervention"]    = TARGET_INTERVENTION_LIMIT
    monthly_fg["Objectif_RFT_MMC"]        = TARGET_RFT_MMC

    # --- By Client ---
    by_client = df.groupby(MMC_COL_CLIENT).agg(
        Quantite_Produite = (MMC_COL_QTY, "sum"),
    ).reset_index().sort_values("Quantite_Produite", ascending=False)
    by_client.rename(columns={MMC_COL_CLIENT: "Client"}, inplace=True)

    # --- By Operator ---
    by_op = df.groupby(MMC_COL_OPERATOR).agg(
        Quantite_Produite = (MMC_COL_QTY, "sum"),
    ).reset_index().sort_values("Quantite_Produite", ascending=False)
    by_op.rename(columns={MMC_COL_OPERATOR: "Operateur"}, inplace=True)

    # --- Test Results ---
    test_cols = [
        MMC_COL_SEC_LOCK, MMC_COL_VERZT,   MMC_COL_AUSTRAST,
        MMC_COL_SEC_INLOCK, MMC_COL_OFFSET, MMC_COL_PUSHBACK,
        MMC_COL_UNLOCK, MMC_COL_HOUSING,   MMC_COL_ETANCH,
        MMC_COL_COULEUR
    ]
    test_rows = []
    for col in test_cols:
        if col in df.columns:
            total  = len(df)
            passed = int(df[col].notna().sum())
            test_rows.append({
                "Test"            : col,
                "Total_Pieces"    : total,
                "Reussi"          : passed,
                "Non_Teste"       : total - passed,
                "Taux_Reussite_%" : round(passed / total * 100, 2) if total > 0 else 0,
                "Objectif_%"      : 95.0,
            })
    test_results = pd.DataFrame(test_rows)

    print(f"[OK] MMC Production KPIs — {len(daily)} daily records.")
    print(f"     Avg Finish Good : {monthly_fg['Tested_Good'].mean():.1f} pieces/month")

    return {
        "MMC_Journalier"        : daily,
        "MMC_FinishGood_Mensuel": monthly_fg,
        "MMC_Par_Client"        : by_client,
        "MMC_Par_Operateur"     : by_op,
        "MMC_Tests"             : test_results,
    }


# =============================================================================
# MODULE 5 — GLOBAL RFT + KPI SUMMARY
# =============================================================================

def build_global_kpis(cnc_results, fc_results, mmc_results):
    """
    Build global KPI summary combining all modules:
    - RFT Global (CNC + Assembly + MMC) monthly
    - KPI Mensuel combined
    - KPI Global annual
    """
    print("\n--- Building Global KPIs ---")

    cnc_monthly = cnc_results.get("CNC_Mensuel",          pd.DataFrame())
    fc_monthly  = fc_results.get("CF_Mensuel",            pd.DataFrame())
    mmc_monthly = mmc_results.get("MMC_FinishGood_Mensuel", pd.DataFrame())

    # --- Combined Monthly ---
    kpi = pd.merge(cnc_monthly, fc_monthly,  on=["Annee", "Mois_N", "Mois"], how="outer")
    kpi = pd.merge(kpi,         mmc_monthly, on=["Annee", "Mois_N", "Mois"], how="outer")
    kpi = kpi.fillna(0)
    kpi = kpi.sort_values(["Annee", "Mois_N"]).reset_index(drop=True)

    # Global defect rate
    total_defects = kpi.get("Nombre_Defauts", 0) + kpi.get("Nb_Non_Conformites", 0)
    total_pieces  = kpi.get("Pieces_Controlees", 0) + kpi.get("Quantite_Inspectee", 0)
    kpi["Taux_Defauts_Global_%"] = safe_rate(total_defects, total_pieces)

    # --- RFT Global Monthly ---
    rft_global = pd.DataFrame()
    if "RFT_CNC" in cnc_monthly.columns and "RFT_Assembly" in fc_monthly.columns:
        rft = pd.merge(
            cnc_monthly[["Annee", "Mois_N", "Mois", "RFT_CNC", "Objectif_RFT_CNC"]],
            fc_monthly[["Annee", "Mois_N", "Mois", "RFT_Assembly", "Objectif_RFT_Assembly"]],
            on=["Annee", "Mois_N", "Mois"], how="outer"
        )
        rft = rft.fillna(0).sort_values(["Annee", "Mois_N"]).reset_index(drop=True)
        rft["RFT_Global"]      = (rft["RFT_CNC"] * rft["RFT_Assembly"]).round(4)
        rft["Objectif_RFT_MMC"] = TARGET_RFT_MMC
        rft_global = rft

    # --- Annual Summary ---
    annual = kpi.groupby("Annee").agg(
        Quantite_Produite  = ("Quantite_Produite",  "sum"),
        Pieces_Controlees  = ("Pieces_Controlees",  "sum"),
        Nombre_Defauts     = ("Nombre_Defauts",     "sum"),
        Nb_Non_Conformites = ("Nb_Non_Conformites", "sum"),
    ).reset_index()
    annual["Taux_Defauts_%"]      = safe_rate(annual["Nombre_Defauts"],     annual["Pieces_Controlees"])
    annual["Taux_NC_%"]           = safe_rate(annual["Nb_Non_Conformites"], annual["Pieces_Controlees"])

    print(f"[OK] Global KPIs — {len(kpi)} monthly records, {len(annual)} years.")

    return {
        "KPI_Mensuel" : kpi,
        "RFT_Global"  : rft_global,
        "KPI_Global"  : annual,
    }


# =============================================================================
# EXPORT
# =============================================================================

def export_to_excel(suivi, cnc_results, fc_results, mmc_results, global_results):
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILE)
    print(f"\n--- Exporting to Excel: {output_path} ---")

    all_sheets = {}
    all_sheets["Suivi_Journalier"] = suivi
    all_sheets.update(cnc_results)
    all_sheets.update(fc_results)
    all_sheets.update(mmc_results)
    all_sheets.update(global_results)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in all_sheets.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"    [+] Sheet: {sheet_name} ({len(df)} rows)")

    print(f"\n[SUCCESS] KPI Output saved: {output_path}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 60)
    print("  KPI AUTOMATION SYSTEM — MMC")
    print(f"  Run date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    create_folders()

    # Load
    df_cnc = load_excel(CNC_FILE,            CNC_SHEET)
    df_fc  = load_excel(FINAL_CONTROL_FILE,  FINAL_CONTROL_SHEET)
    df_mmc = load_excel(MMC_PRODUCTION_FILE, MMC_PRODUCTION_SHEET, skip_rows=1)

    # Clean
    if df_cnc is not None: df_cnc = clean_dataframe(df_cnc, CNC_COL_DATE)
    if df_fc  is not None: df_fc  = clean_dataframe(df_fc,  FC_COL_DATE)
    if df_mmc is not None: df_mmc = clean_dataframe(df_mmc, MMC_COL_DATE)

    # Process
    suivi       = build_suivi_journalier(df_cnc) if df_cnc is not None else pd.DataFrame()
    cnc_results = build_cnc_kpis(df_cnc)         if df_cnc is not None else {}
    fc_results  = build_fc_kpis(df_fc)           if df_fc  is not None else {}
    mmc_results = build_mmc_kpis(df_mmc)         if df_mmc is not None else {}

    global_results = {}
    if df_cnc is not None and df_fc is not None and df_mmc is not None:
        global_results = build_global_kpis(cnc_results, fc_results, mmc_results)

    # Export
    export_to_excel(suivi, cnc_results, fc_results, mmc_results, global_results)

    print("\n" + "=" * 60)
    print("  DONE. Open KPI_Output.xlsx and connect to Power BI.")
    print("=" * 60)


if __name__ == "__main__":
    main()
