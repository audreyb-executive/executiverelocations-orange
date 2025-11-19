# --------------------------------------------------------------------
#              AJOUT STREAMLIT POUR RENDRE LE PIPELINE 100% UI
# --------------------------------------------------------------------
import streamlit as st
from io import BytesIO
from pathlib import Path  

def run_pipeline(exp_path: Path, fx_path: Path) -> Path:
    """
    Pipeline OGIM identique au notebook, mais exécutable après upload Streamlit.
    Seules modifications : EXPENSES_FILE et FX_FILE remplacés par exp_path et fx_path.
    """

    global EXPENSES_FILE, FX_FILE
    EXPENSES_FILE = exp_path
    FX_FILE = fx_path

    # ------------------------------
    # Ré-exécution EXACTE du pipeline notebook
    # ------------------------------

    # %% [load_fx] Chargement FX
    raw_fx = pd.read_excel(FX_FILE, header=None)
    fx = raw_fx.iloc[7:, :].copy()
    fx = fx.iloc[:, 2:].copy()

    fx.columns = fx.iloc[0].tolist()
    fx = fx.iloc[1:, :].reset_index(drop=True)

    if "ISO-CODE" not in fx.columns:
        candidates = [c for c in fx.columns if strip_accents_punct_ws(c).replace(" ", "") in ("iso-code","isocode","iso")]
        fx = fx.rename(columns={candidates[0]: "ISO-CODE"})

    fx["ISO-CODE"] = fx["ISO-CODE"].astype(str).str.upper().str.strip()

    month_cols = [c for c in fx.columns if c != "ISO-CODE"]
    for c in month_cols:
        fx[c] = fx[c].apply(to_float_force)

    # %% [load_expenses]
    xls = pd.ExcelFile(EXPENSES_FILE)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(EXPENSES_FILE, sheet_name=name)
        if df.dropna(how='all').shape[0] <= 1:
            continue
        sheets[name] = df.copy()

    # %% [step1] Filtre mois
    now = datetime.now(TZ)
    target = get_target_months(now)
    target_fr = target["month_fr"]
    target_en_abbr = target["month_en_abbr"]
    target_year = target["year"]

    def month_match_series(s: pd.Series) -> pd.Series:
        norm = s.astype(str).apply(normalize_month_fr)
        return norm == target_fr

    filtered_sheets = {}
    for name, df in sheets.items():
        if "MONTH" not in df.columns:
            continue
        mask = month_match_series(df["MONTH"])
        filtered = df[mask].copy()
        if not filtered.empty:
            filtered_sheets[name] = filtered.reset_index(drop=True)

    # %% [step2] AIRPLUS
    airplus_frames = {}
    cleaned_sheets = {}

    for name, df in filtered_sheets.items():
        if "SUPPLIER NAME" not in df.columns:
            cleaned_sheets[name] = df.copy()
            continue

        cur_df = df.copy()
        sup_norm = df["SUPPLIER NAME"].astype(str).apply(is_airplus)
        air = cur_df[sup_norm].copy()
        base = cur_df[~sup_norm].copy()

        dev = extract_currency_from_sheet(name) or "UNK"

        if not air.empty:
            key = f"AIRPLUS – {dev}"
            if key in airplus_frames:
                airplus_frames[key] = pd.concat([airplus_frames[key], air], ignore_index=True)
            else:
                airplus_frames[key] = air.reset_index(drop=True)

        cleaned_sheets[name] = base.reset_index(drop=True)

    # %% [step2bis] HORS PAYE
    horspaye_frames = {}
    cleaned_sheets_hp = {}

    for name, df in cleaned_sheets.items():
        if "SALARIE OGIM" not in df.columns:
            cleaned_sheets_hp[name] = df.copy()
            continue

        cur_df = df.copy()
        hp_mask = cur_df["SALARIE OGIM"].astype(str).apply(is_hors_paye)
        hors = cur_df[hp_mask].copy()
        base = cur_df[~hp_mask].copy()

        dev = extract_currency_from_sheet(name) or "UNK"

        if not hors.empty:
            key = f"HORS PAYE – {dev}"
            if key in horspaye_frames:
                horspaye_frames[key] = pd.concat([horspaye_frames[key], hors], ignore_index=True)
            else:
                horspaye_frames[key] = hors.reset_index(drop=True)

        cleaned_sheets_hp[name] = base.reset_index(drop=True)

    # %% [step3] FX
    month_col_fx = target_en_abbr

    fx_completed_sheets = {n: complete_fx(df, fx, month_col_fx) for n, df in cleaned_sheets_hp.items()}
    fx_completed_air    = {n: complete_fx(df, fx, month_col_fx) for n, df in airplus_frames.items()}
    fx_completed_hp     = {n: complete_fx(df, fx, month_col_fx) for n, df in horspaye_frames.items()}

    # %% [step4] Flags
    prepared_main = {n: prepare_amount_flags(df) for n, df in fx_completed_sheets.items()}
    prepared_air  = {n: prepare_amount_flags(df) for n, df in fx_completed_air.items()}
    prepared_hp   = {n: prepare_amount_flags(df) for n, df in fx_completed_hp.items()}

    # %% [step5] Nettoyage
    cleaned_main_final = {n: clean_people_assignment(df) for n, df in prepared_main.items()}
    cleaned_air_final  = {n: clean_people_assignment(df) for n, df in prepared_air.items()}
    cleaned_hp_final   = {n: clean_people_assignment(df) for n, df in prepared_hp.items()}

    # %% [step6] Output
    out_file = write_all_sheets(cleaned_main_final, cleaned_air_final, cleaned_hp_final, fx, target)

    return out_file


# --------------------------------------------------------------------
#                         INTERFACE STREAMLIT
# --------------------------------------------------------------------
def main_streamlit():
    st.title("📊 OGIM – Traitement Automatisé des Expenses")
    st.write("Cette application exécute **exactement** le même pipeline que ton notebook OGIM.")

    st.header("1️⃣ Charger les fichiers d'entrée")

    fx_file = st.file_uploader("Fichier de taux (FX)", type=["xlsx"])
    exp_file = st.file_uploader("Fichier des dépenses OGIM", type=["xlsx"])

    if st.button("🚀 Lancer le traitement"):
        if fx_file is None or exp_file is None:
            st.error("Merci de charger les deux fichiers.")
            return

        # Sauvegarde temporaire pour réutiliser la logique existante
        tmp_fx = OUTPUT_DIR / "fx_uploaded.xlsx"
        tmp_exp = OUTPUT_DIR / "exp_uploaded.xlsx"

        with open(tmp_fx, "wb") as f:
            f.write(fx_file.getbuffer())
        with open(tmp_exp, "wb") as f:
            f.write(exp_file.getbuffer())

        with st.spinner("Traitement OGIM en cours..."):
            out_path = run_pipeline(tmp_exp, tmp_fx)

        st.success("Traitement terminé 🎉")

        # Téléchargement
        with open(out_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier OGIM final",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# Lancer Streamlit en mode script
if __name__ == "__main__":
    main_streamlit()

