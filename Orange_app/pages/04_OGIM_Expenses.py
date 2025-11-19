# ============================================================
#                 IMPORTS OGIM + STREAMLIT
# ============================================================
import streamlit as st
from io import BytesIO
from pathlib import Path

# Librairies OGIM
import pandas as pd
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
import unicodedata


# ============================================================
#              VARIABLES + DOSSIER UPLOAD
# ============================================================
TZ = ZoneInfo("Europe/Paris")

OUTPUT_DIR = Path("uploads")
OUTPUT_DIR.mkdir(exist_ok=True)

EXPENSES_FILE = None
FX_FILE = None


# ============================================================
#                   FONCTIONS UTILITAIRES OGIM
# ============================================================
def strip_accents_punct_ws(text: str) -> str:
    if pd.isna(text):
        return ""
    s = str(text)
    s = ''.join(c for c in unicodedata.normalize('NFKD', s)
                if not unicodedata.combining(c))
    s = re.sub(r"[.,;:!?\"'()\[\]{}_•·–—-]+", " ", s)
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE).strip()
    return s.lower()


def normalize_month_fr(value: str) -> str:
    base = strip_accents_punct_ws(value)
    base = base.replace(".", "").replace("-", " ").strip()
    base = base.rstrip(". ")
    base_no_space = base.replace(" ", "")

    FRENCH_MONTHS = {
        "janv": "janvier", "jan": "janvier", "janvier": "janvier",
        "fevr": "fevrier", "fevrier": "fevrier", "fev": "fevrier",
        "mars": "mars",
        "avr": "avril", "avril": "avril",
        "mai": "mai", "juin": "juin",
        "juil": "juillet", "juillet": "juillet",
        "aout": "aout",
        "sept": "septembre", "septembre": "septembre",
        "oct": "octobre", "octobre": "octobre",
        "nov": "novembre", "novembre": "novembre",
        "dec": "decembre", "decembre": "decembre"
    }

    return FRENCH_MONTHS.get(base_no_space, base_no_space)


def is_airplus(val: str) -> bool:
    return strip_accents_punct_ws(val).replace(" ", "") == "airplus"


def is_hors_paye(val: str) -> bool:
    if val is None:
        return False
    return strip_accents_punct_ws(val).replace(" ", "") == "horspaye"


def extract_currency_from_sheet(sheet_name: str) -> str | None:
    tokens = re.split(r"[^A-Za-z0-9]+", sheet_name)
    candidates = [t for t in tokens if t and t.isalpha() and len(t) in (2, 3)]
    for t in candidates:
        if len(t) == 3:
            return t.upper()
    return candidates[-1].upper() if candidates else None


def to_float_force(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except:
        return np.nan


def get_target_months(dt: datetime):
    y = dt.year
    m = dt.month - 1
    if m == 0:
        m = 12
        y -= 1

    EN_MONTH_ABBR = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
        7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
    }
    FR_CANONICAL = {
        1: "janvier", 2: "fevrier", 3: "mars", 4: "avril", 5: "mai", 6: "juin",
        7: "juillet", 8: "aout", 9: "septembre", 10: "octobre", 11: "novembre", 12: "decembre"
    }

    return {
        "year": str(y),
        "month_fr": FR_CANONICAL[m],
        "month_en_abbr": EN_MONTH_ABBR[m],
    }


def complete_fx(df: pd.DataFrame, fx_table: pd.DataFrame, month_col: str) -> pd.DataFrame:
    out = df.copy()
    if "CURRENCY" not in out.columns:
        out["CURRENCY"] = np.nan
    out["CURRENCY"] = out["CURRENCY"].astype(str).str.upper().str.strip()

    fx_idx = fx_table.set_index("ISO-CODE")

    rates = []
    for cur in out["CURRENCY"]:
        if pd.isna(cur) or cur == "":
            rates.append(np.nan)
            continue
        if cur == "EUR":
            rates.append(1.0)
            continue
        try:
            rate = fx_idx.at[cur, month_col]
            rates.append(to_float_force(rate))
        except KeyError:
            rates.append(np.nan)

    out["CURRENCY RATE"] = rates
    return out


def prepare_amount_flags(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["TOTAL AMOUNT IN CURRENCY"] = out.get("TOTAL AMOUNT IN CURRENCY", np.nan).apply(to_float_force)
    out["CURRENCY RATE"] = out.get("CURRENCY RATE", np.nan).apply(to_float_force)

    out["_AMOUNT_OK_"] = (
        out["TOTAL AMOUNT IN CURRENCY"].notna() &
        out["CURRENCY RATE"].notna() &
        (out["CURRENCY RATE"] != 0)
    )
    return out


def clean_people_assignment(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    BAD_TOKENS = {"vide", "missing", "assignee", "entity"}

    def clean_name(v):
        s = strip_accents_punct_ws(v)
        return "DMI" if s == "" or s in BAD_TOKENS else str(v).upper()

    out["FIRST NAME"] = out.get("FIRST NAME", "DMI").apply(clean_name)
    out["LAST NAME"] = out.get("LAST NAME", "DMI").apply(clean_name)

    def clean_assign(v):
        return "N/A" if strip_accents_punct_ws(v) in ("", "cost estimate") else str(v).upper()

    out["TYPE OF ASSIGNMENT"] = out.get("TYPE OF ASSIGNMENT", "N/A").apply(clean_assign)

    return out


def write_all_sheets(main_dict, air_dict, hp_dict, fx_table, target) -> Path:
    out_path = OUTPUT_DIR / f"EXPENSES_OGIM_{target['month_fr']}_{target['year']}.xlsx"
    if out_path.exists():
        out_path.unlink()

    # Écriture sans formules
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in main_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
        for name, df in air_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
        for name, df in hp_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)

    return out_path


# ============================================================
#                   PIPELINE PRINCIPAL
# ============================================================
def run_pipeline(exp_path: Path, fx_path: Path) -> Path:
    global EXPENSES_FILE, FX_FILE
    EXPENSES_FILE = exp_path
    FX_FILE = fx_path

    # ------------- FX -------------
    raw_fx = pd.read_excel(FX_FILE, header=None)
    fx = raw_fx.iloc[7:, :].copy().iloc[:, 2:]
    fx.columns = fx.iloc[0].tolist()
    fx = fx.iloc[1:].reset_index(drop=True)

    if "ISO-CODE" not in fx.columns:
        raise ValueError("Colonne ISO-CODE introuvable")

    fx["ISO-CODE"] = fx["ISO-CODE"].astype(str).str.upper()

    for col in fx.columns:
        if col != "ISO-CODE":
            fx[col] = fx[col].apply(to_float_force)

    # ------------ EXPENSES ------------
    xls = pd.ExcelFile(EXPENSES_FILE)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(EXPENSES_FILE, sheet_name=name)
        if df.dropna(how="all").shape[0] > 1:
            sheets[name] = df.copy()

    # ------------ FILTRE MOIS ------------
    now = datetime.now(TZ)
    target = get_target_months(now)
    target_fr = target["month_fr"]
    target_en_abbr = target["month_en_abbr"]

    def match_month(col):
        return col.astype(str).apply(normalize_month_fr) == target_fr

    filtered_sheets = {}
    for name, df in sheets.items():
        if "MONTH" in df.columns:
            mask = match_month(df["MONTH"])
            if mask.any():
                filtered_sheets[name] = df[mask].copy()

    # ------------ AIRPLUS & HORS PAYE ------------
    airplus_frames = {}
    cleaned_sheets = {}
    for name, df in filtered_sheets.items():
        if "SUPPLIER NAME" not in df.columns:
            cleaned_sheets[name] = df.copy()
            continue
        mask_air = df["SUPPLIER NAME"].astype(str).apply(is_airplus)
        air = df[mask_air].copy()
        base = df[~mask_air].copy()
        dev = extract_currency_from_sheet(name) or "UNK"
        if not air.empty:
            key = f"AIRPLUS – {dev}"
            airplus_frames.setdefault(key, pd.DataFrame())
            airplus_frames[key] = pd.concat([airplus_frames[key], air], ignore_index=True)
        cleaned_sheets[name] = base.reset_index(drop=True)

    horspaye_frames = {}
    cleaned_sheets_hp = {}
    for name, df in cleaned_sheets.items():
        if "SALARIE OGIM" not in df.columns:
            cleaned_sheets_hp[name] = df.copy()
            continue
        mask_hp = df["SALARIE OGIM"].astype(str).apply(is_hors_paye)
        hp = df[mask_hp].copy()
        base = df[~mask_hp].copy()
        dev = extract_currency_from_sheet(name) or "UNK"
        if not hp.empty:
            key = f"HORS PAYE – {dev}"
            horspaye_frames.setdefault(key, pd.DataFrame())
            horspaye_frames[key] = pd.concat([horspaye_frames[key], hp], ignore_index=True)
        cleaned_sheets_hp[name] = base.reset_index(drop=True)

    # ------------ FX ------------
    fx_main = {n: complete_fx(df, fx, target_en_abbr) for n, df in cleaned_sheets_hp.items()}
    fx_air = {n: complete_fx(df, fx, target_en_abbr) for n, df in airplus_frames.items()}
    fx_hp = {n: complete_fx(df, fx, target_en_abbr) for n, df in horspaye_frames.items()}

    # ----------- PREP FLAGS -------
    prep_main = {n: prepare_amount_flags(df) for n, df in fx_main.items()}
    prep_air = {n: prepare_amount_flags(df) for n, df in fx_air.items()}
    prep_hp = {n: prepare_amount_flags(df) for n, df in fx_hp.items()}

    # ----------- CLEAN PEOPLE -------
    clean_main = {n: clean_people_assignment(df) for n, df in prep_main.items()}
    clean_air = {n: clean_people_assignment(df) for n, df in prep_air.items()}
    clean_hp = {n: clean_people_assignment(df) for n, df in prep_hp.items()}

    # ----------- OUTPUT -------
    out_file = write_all_sheets(clean_main, clean_air, clean_hp, fx, target)

    return out_file


# ============================================================
#                 INTERFACE STREAMLIT
# ============================================================
def main_streamlit():
    st.title("🧾OGIM – Traitement Automatisé des Expenses")
    st.write(
        "Cette application permet de **traiter automatiquement les fichiers de dépenses OGIM** "
        "et de générer un fichier final conforme au processus standard."
    )

    st.markdown("""
    **Fonctionnalités principales :**
    - Import du fichier des dépenses OGIM et du fichier des taux de change  
    - Filtrage automatique du mois précédent  
    - Extraction des lignes AIRPLUS et HORS PAYE  
    - Application des taux de change  
    - Nettoyage et normalisation des données salariés  
    - Préparation des montants à convertir  
    - Génération d’un fichier Excel final structuré  
    """)


    
    st.header("1️⃣ Importer vos fichiers")

    fx_file = st.file_uploader("Fichier de taux de devise (ex : Equant Corporate rates_October_2025.xlsx)", type=["xlsx"])
    exp_file = st.file_uploader("Fichier Expenses OGIM (ex : expenses5026.xlsx)", type=["xlsx"])

    if st.button("🚀 Lancer le traitement"):
        if fx_file is None or exp_file is None:
            st.error("Merci de charger les deux fichiers.")
            return

        tmp_fx = OUTPUT_DIR / "fx.xlsx"
        tmp_exp = OUTPUT_DIR / "exp.xlsx"

        with open(tmp_fx, "wb") as f:
            f.write(fx_file.getbuffer())
        with open(tmp_exp, "wb") as f:
            f.write(exp_file.getbuffer())

        with st.spinner("Traitement OGIM en cours…"):
            out_path = run_pipeline(tmp_exp, tmp_fx)

        st.success("Traitement terminé 🎉")

        with open(out_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier OGIM final",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == "__main__":
    main_streamlit()


