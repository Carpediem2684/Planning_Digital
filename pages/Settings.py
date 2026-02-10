
# ============================================
# Settings.py — Dashboard de Gestion des Plannings
# ============================================

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path

BASE_PATH = Path(r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main")
SUIVI_OF_FILE = BASE_PATH / "SUIVI_OF.xlsx"

LIGNE_MAPPING = {
    "L06 - 4M-LIGNE1": {"name": "L1", "file": "OFs_L1.xlsx", "sheet": "Feuil1"},
    "L08 - 4M-LIGNE2": {"name": "L2", "file": "OFs_L2.xlsx", "sheet": "Sheet1"},
    "L09 - 4M-IMPRIMERIE": {"name": "Imprimerie", "file": "OFs_Imprimerie.xlsx", "sheet": "Feuil1"},
    "L10 - 4M-VISITAGE": {"name": "Visitage", "file": "OFs_Visitage.xlsx", "sheet": "Feuil1"},
}

STATUT_NON_LANCE = [5, 10, 15]
STATUT_LANCE = [30]
STATUT_EN_COURS = [40, 50]
STATUT_TERMINE = [60, 61, 99]

MAX_CAMPAGNES = 20
DEFAULT_ML_MIN = 20
DEFAULT_HEURE = 60
DEFAULT_TRG = 0.8
OFFSET_TEMPS = 0.75


def load_suivi_of():
    return pd.read_excel(SUIVI_OF_FILE, engine="openpyxl")


def load_calcul_duree(ligne_name):
    file_map = {"L1": "OFs_L1.xlsx", "L2": "OFs_L2.xlsx", "Imprimerie": "OFs_Imprimerie.xlsx", "Visitage": "OFs_Visitage.xlsx"}
    try:
        df = pd.read_excel(BASE_PATH / file_map[ligne_name], sheet_name="calcul_durée", engine="openpyxl")
        df.columns = df.columns.str.strip()
        if "Famille " in df.columns:
            df = df.rename(columns={"Famille ": "Famille"})
        return df
    except:
        return pd.DataFrame()


def get_ml_min(calcul_df, search):
    if calcul_df.empty or not search:
        return DEFAULT_ML_MIN
    search = str(search).upper().strip()
    for _, row in calcul_df.iterrows():
        fam = str(row.get("Famille", "")).upper().strip()
        if fam == search or fam in search or search in fam:
            return float(row.get("ml/min", DEFAULT_ML_MIN))
    return DEFAULT_ML_MIN


def calc_temps(ml, ml_min):
    if not ml or ml <= 0 or not ml_min or ml_min <= 0:
        return 0
    return (ml / ml_min / DEFAULT_HEURE * DEFAULT_TRG) + OFFSET_TEMPS


def get_statut_label(statut):
    if statut in STATUT_NON_LANCE:
        return "Non lancé", "🔴"
    elif statut in STATUT_LANCE:
        return "Lancé", "🟢"
    elif statut in STATUT_EN_COURS:
        return "En cours", "🔵"
    elif statut in STATUT_TERMINE:
        return "Terminé", "⚫"
    return "Inconnu", "⚪"


# IMPORTANT : on considère désormais ACTIF = STATUT <= 50
# (exclut explicitement tout ce qui est > 50)

def filter_active(df):
    return df[df["STATUT"] <= 50]


def extract_laise(lib_format):
    if not lib_format or pd.isna(lib_format):
        return 4
    s = str(lib_format).upper()
    if "2M" in s:
        return 2
    elif "3M" in s:
        return 3
    return 4


def extract_type(campagne):
    if not campagne:
        return ""
    c = str(campagne).upper()
    for t in ["TEXLINE", "NERA", "PRIMETEX", "TARABUS", "BOOSTER", "SPORISOL", "TMAX", "START", "FUSION", "LOFTEX"]:
        if t in c:
            return t
    return ""


# ---------------------------------
# L1 — Applique STATUT <= 50 et ML = COMMANDE - FABRIQUE ; ML < 150 => skip
# ---------------------------------

def generate_L1(df, campagnes, urgents, journee, calcul_df):
    rows = []
    done = set()

    # 1) OFs du jour et urgents
    for of in journee + [u for u in urgents if u not in journee]:
        d = df[(df["LIB_LIGNE"] == "L06 - 4M-LIGNE1") & (df["NUM_OF"] == of)]
        if d.empty:
            continue
        r = d.iloc[0]

        # Règle STATUT
        if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
            continue

        done.add(of)

        lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
        produit = ""
        for code in ["CICDMD01", "CICD01", "CICD02", "CICD03", "CICD04", "CICD05", "CICD06", "CIMD02", "CIMD03"]:
            if code in lib.upper():
                largeur = "3M" if "3M" in lib.upper() else "4M"
                produit = f"{code} {largeur}"
                break

        # ML = COMMANDE - FABRIQUE
        commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
        fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
        ml = commande - fabrique
        if ml < 150:
            continue

        ml_min = get_ml_min(calcul_df, produit)

        rows.append({
            "Campagne": lib, "Colonne2": None, "Produit": produit, "Ml": ml,
            "Ofs": r["NUM_OF"], "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min)
        })

    # 2) Campagnes
    for camp in campagnes:
        if not camp:
            continue
        sub = df[(df["LIB_LIGNE"] == "L06 - 4M-LIGNE1") & (df["LIB_CAMPAGNE"] == camp)]
        sub = filter_active(sub)
        for _, r in sub.iterrows():
            if r["NUM_OF"] in done:
                continue

            # Règle STATUT
            if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
                continue

            lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
            produit = ""
            for code in ["CICDMD01", "CICD01", "CICD02", "CICD03", "CICD04", "CICD05", "CICD06", "CIMD02", "CIMD03"]:
                if code in lib.upper():
                    largeur = "3M" if "3M" in lib.upper() else "4M"
                    produit = f"{code} {largeur}"
                    break

            commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
            fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
            ml = commande - fabrique
            if ml < 150:
                continue

            ml_min = get_ml_min(calcul_df, produit)

            rows.append({
                "Campagne": lib, "Colonne2": None, "Produit": produit, "Ml": ml,
                "Ofs": r["NUM_OF"], "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min)
            })

    return pd.DataFrame(rows)


# ---------------------------------
# L2 (MODIFIÉ) — FAMILLE = DESCRIPTION (colonne K)
# Applique STATUT <= 50 et ML = COMMANDE - FABRIQUE ; ML < 150 => skip
# ---------------------------------

def generate_L2(df, campagnes, urgents, journee, calcul_df, col_k_name="DESCRIPTION"):
    rows = []
    done = set()

    # 1) OFs du jour et urgents
    for of in journee + [u for u in urgents if u not in journee]:
        d = df[(df["LIB_LIGNE"] == "L08 - 4M-LIGNE2") & (df["NUM_OF"] == of)]
        if d.empty:
            continue
        r = d.iloc[0]

        if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
            continue

        done.add(of)

        lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
        famille_calc = extract_type(lib)  # utilisé uniquement pour ml/min

        commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
        fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
        ml = commande - fabrique
        if ml < 150:
            continue

        ml_min = get_ml_min(calcul_df, famille_calc)

        desc = r[col_k_name] if (col_k_name in r.index and pd.notna(r[col_k_name])) else ""

        rows.append({
            "COLORIS": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
            "GRAIN": r["GRAIN"] if pd.notna(r["GRAIN"]) else "",
            "FAMILLE": desc,  # Remplacement par DESCRIPTION
            "ML": ml, "Ofs": r["NUM_OF"],
            "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min), "Campagne": lib
        })

    # 2) Campagnes
    for camp in campagnes:
        if not camp:
            continue
        sub = df[(df["LIB_LIGNE"] == "L08 - 4M-LIGNE2") & (df["LIB_CAMPAGNE"] == camp)]
        sub = filter_active(sub)
        for _, r in sub.iterrows():
            if r["NUM_OF"] in done:
                continue

            if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
                continue

            lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
            famille_calc = extract_type(lib)

            commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
            fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
            ml = commande - fabrique
            if ml < 150:
                continue

            ml_min = get_ml_min(calcul_df, famille_calc)
            desc = r[col_k_name] if (col_k_name in r.index and pd.notna(r[col_k_name])) else ""

            rows.append({
                "COLORIS": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
                "GRAIN": r["GRAIN"] if pd.notna(r["GRAIN"]) else "",
                "FAMILLE": desc,
                "ML": ml, "Ofs": r["NUM_OF"],
                "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min), "Campagne": lib
            })

    return pd.DataFrame(rows)


# ---------------------------------
# Imprimerie — STATUT <= 50 ; ML = COMMANDE - FABRIQUE ; ML < 150 => skip
# ---------------------------------

def generate_Imp(df, campagnes, urgents, journee, calcul_df):
    rows = []
    done = set()

    # 1) OFs du jour et urgents
    for of in journee + [u for u in urgents if u not in journee]:
        d = df[(df["LIB_LIGNE"] == "L09 - 4M-IMPRIMERIE") & (df["NUM_OF"] == of)]
        if d.empty:
            continue
        r = d.iloc[0]

        if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
            continue

        done.add(of)

        lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
        support = str(r.get("COMPOSANT", ""))[:6] if pd.notna(r.get("COMPOSANT")) else ""

        commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
        fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
        ml = commande - fabrique
        if ml < 150:
            continue

        type_camp = extract_type(lib)
        ml_min = get_ml_min(calcul_df, type_camp)

        rows.append({
            "Coloris": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
            "Support": support, "Campagne": lib, "Ml": ml, "Ofs": r["NUM_OF"],
            "Temp/min": ml_min, "Temps en h": calc_temps(ml, ml_min)
        })

    # 2) Campagnes
    for camp in campagnes:
        if not camp:
            continue
        sub = df[(df["LIB_LIGNE"] == "L09 - 4M-IMPRIMERIE") & (df["LIB_CAMPAGNE"] == camp)]
        sub = filter_active(sub)
        for _, r in sub.iterrows():
            if r["NUM_OF"] in done:
                continue

            if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
                continue

            lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
            support = str(r.get("COMPOSANT", ""))[:6] if pd.notna(r.get("COMPOSANT")) else ""

            commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
            fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
            ml = commande - fabrique
            if ml < 150:
                continue

            type_camp = extract_type(lib)
            ml_min = get_ml_min(calcul_df, type_camp)

            rows.append({
                "Coloris": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
                "Support": support, "Campagne": lib, "Ml": ml, "Ofs": r["NUM_OF"],
                "Temp/min": ml_min, "Temps en h": calc_temps(ml, ml_min)
            })

    return pd.DataFrame(rows)


# ---------------------------------
# Visitage — STATUT <= 50 ; ML = COMMANDE - FABRIQUE ; ML < 150 => skip
# ---------------------------------

def generate_Vis(df, campagnes, urgents, journee, calcul_df):
    rows = []
    done = set()

    # 1) OFs du jour et urgents
    for of in journee + [u for u in urgents if u not in journee]:
        d = df[(df["LIB_LIGNE"] == "L10 - 4M-VISITAGE") & (df["NUM_OF"] == of)]
        if d.empty:
            continue
        r = d.iloc[0]

        if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
            continue

        done.add(of)

        lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
        couleur = extract_type(lib)
        laise = extract_laise(r.get("LIB_FORMAT"))

        commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
        fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
        ml = commande - fabrique
        if ml < 150:
            continue

        ml_min = get_ml_min(calcul_df, couleur)

        rows.append({
            "Coloris": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
            "Laise": laise, "Campagne": lib, "Ml": ml, "Ofs": r["NUM_OF"],
            "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min),
            "trg": DEFAULT_TRG, "Heure": DEFAULT_HEURE, "COULEUR": couleur
        })

    # 2) Campagnes
    for camp in campagnes:
        if not camp:
            continue
        sub = df[(df["LIB_LIGNE"] == "L10 - 4M-VISITAGE") & (df["LIB_CAMPAGNE"] == camp)]
        sub = filter_active(sub)
        for _, r in sub.iterrows():
            if r["NUM_OF"] in done:
                continue

            if pd.notna(r.get("STATUT")) and r["STATUT"] > 50:
                continue

            lib = str(r["LIB_CAMPAGNE"]) if pd.notna(r["LIB_CAMPAGNE"]) else ""
            couleur = extract_type(lib)
            laise = extract_laise(r.get("LIB_FORMAT"))

            commande = float(r["COMMANDE"]) if pd.notna(r["COMMANDE"]) else 0
            fabrique = float(r["FABRIQUE"]) if "FABRIQUE" in r.index and pd.notna(r["FABRIQUE"]) else 0
            ml = commande - fabrique
            if ml < 150:
                continue

            ml_min = get_ml_min(calcul_df, couleur)

            rows.append({
                "Coloris": r["COLORIS"] if pd.notna(r["COLORIS"]) else "",
                "Laise": laise, "Campagne": lib, "Ml": ml, "Ofs": r["NUM_OF"],
                "ml / min": ml_min, "Temps en h": calc_temps(ml, ml_min),
                "trg": DEFAULT_TRG, "Heure": DEFAULT_HEURE, "COULEUR": couleur
            })

    return pd.DataFrame(rows)


# ---------------------------------
# Sauvegarde
# ---------------------------------

def save_with_calcul(df_export, path, ligne_name):
    try:
        calcul_df = load_calcul_duree(ligne_name)
        sheet = "Sheet1" if ligne_name == "L2" else "Feuil1"

        # Forcer FAMILLE en colonne C pour L2
        if ligne_name == "L2" and "FAMILLE" in df_export.columns:
            cols = list(df_export.columns)
            cols.remove("FAMILLE")
            cols.insert(2, "FAMILLE")  # index 2 => colonne C
            df_export = df_export[cols]

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df_export.to_excel(writer, sheet_name=sheet, index=False)
            if not calcul_df.empty:
                calcul_df.to_excel(writer, sheet_name="calcul_durée", index=False)
        return True
    except Exception as e:
        st.error(f"Erreur: {e}")
        return False


# ---------------------------------
# UI
# ---------------------------------

def show_settings():
    st.title("⚙️ Dashboard de Gestion des Plannings")

    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            st.rerun()

    try:
        df_suivi = load_suivi_of()
        st.success(f"✅ SUIVI_OF.xlsx : {len(df_suivi)} OFs")
        # Détecter la colonne K = DESCRIPTION (sécurisé)
        col_k_name = "DESCRIPTION"
        if col_k_name not in df_suivi.columns:
            try:
                col_k_name = df_suivi.columns[10]  # fallback colonne K
                st.info(f"ℹ️ Colonne K détectée : {col_k_name}")
            except Exception:
                st.warning("⚠️ Impossible d'identifier la colonne K — valeur vide utilisée pour L2")
                col_k_name = None
    except Exception as e:
        st.error(f"❌ Erreur chargement SUIVI_OF: {e}")
        return

    calcul_tables = {info["name"]: load_calcul_duree(info["name"]) for info in LIGNE_MAPPING.values()}

    # Vue globale (applique filter_active => STATUT <= 50)
    st.header("📊 Vue Globale")
    cols = st.columns(4)
    for i, (lib_ligne, info) in enumerate(LIGNE_MAPPING.items()):
        with cols[i]:
            sub = filter_active(df_suivi[df_suivi["LIB_LIGNE"] == lib_ligne])
            st.metric(info["name"], f"{len(sub)} OFs", f"{sub['COMMANDE'].sum():,.0f} ML")

    st.divider()

    # Gestion par ligne
    st.header("🔧 Gestion par Ligne")
    tabs = st.tabs(["L1", "L2", "Imprimerie", "Visitage"])

    for tab_idx, (lib_ligne, info) in enumerate(LIGNE_MAPPING.items()):
        with tabs[tab_idx]:
            ligne_name = info["name"]
            ligne_file = info["file"]

            # IMPORTANT : n'afficher que les OF ACTIFS (STATUT <= 50)
            df_ligne = filter_active(df_suivi[df_suivi["LIB_LIGNE"] == lib_ligne])
            campagnes = sorted(df_ligne["LIB_CAMPAGNE"].dropna().unique().tolist())

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 📋 Campagnes disponibles")
                st.info(f"{len(campagnes)} campagnes")
                with st.expander("Détail"):
                    for c in campagnes:
                        sub = df_ligne[df_ligne["LIB_CAMPAGNE"] == c]
                        st.write(f"**{c}**: {len(sub)} OFs, {sub['COMMANDE'].sum():,.0f} ML")

            with col2:
                st.markdown("#### 🚨 OFs Urgents / Journée")
                ofs_list = df_ligne["NUM_OF"].tolist()

                urgents = st.multiselect(f"🚨 Urgents {ligne_name}", ofs_list, key=f"urg_{ligne_name}")
                journee = st.multiselect(f"☀️ Journée {ligne_name}", ofs_list, key=f"jour_{ligne_name}")

            st.markdown("#### 🔢 Ordre des Campagnes")

            key_ordre = f"ordre_{ligne_name}"
            if key_ordre not in st.session_state:
                st.session_state[key_ordre] = [""] * MAX_CAMPAGNES

            options = ["(vide)"] + campagnes
            cols_camp = st.columns(5)
            for i in range(MAX_CAMPAGNES):
                with cols_camp[i % 5]:
                    cur = st.session_state[key_ordre][i]
                    idx = options.index(cur) if cur in options else 0
                    sel = st.selectbox(f"#{i+1}", options, index=idx, key=f"c_{ligne_name}_{i}")
                    st.session_state[key_ordre][i] = "" if sel == "(vide)" else sel

            ordre_final = [c for c in st.session_state[key_ordre] if c]

            st.markdown("#### 📝 Prévisualisation")

            if ordre_final or urgents or journee:
                calcul_df = calcul_tables.get(ligne_name, pd.DataFrame())

                if ligne_name == "L1":
                    preview = generate_L1(df_suivi, ordre_final, urgents, journee, calcul_df)
                elif ligne_name == "L2":
                    preview = generate_L2(df_suivi, ordre_final, urgents, journee, calcul_df, col_k_name)
                elif ligne_name == "Imprimerie":
                    preview = generate_Imp(df_suivi, ordre_final, urgents, journee, calcul_df)
                else:
                    preview = generate_Vis(df_suivi, ordre_final, urgents, journee, calcul_df)

                st.success(f"✅ {len(preview)} OFs prêts")
                with st.expander("Voir détail"):
                    st.dataframe(preview, use_container_width=True)

            if st.button(f"📥 Générer {ligne_file}", key=f"gen_{ligne_name}"):
                if not ordre_final and not urgents and not journee:
                    st.error("Aucune config!")
                else:
                    calcul_df = calcul_tables.get(ligne_name, pd.DataFrame())

                    if ligne_name == "L1":
                        export = generate_L1(df_suivi, ordre_final, urgents, journee, calcul_df)
                    elif ligne_name == "L2":
                        export = generate_L2(df_suivi, ordre_final, urgents, journee, calcul_df, col_k_name)
                    elif ligne_name == "Imprimerie":
                        export = generate_Imp(df_suivi, ordre_final, urgents, journee, calcul_df)
                    else:
                        export = generate_Vis(df_suivi, ordre_final, urgents, journee, calcul_df)

                    path = BASE_PATH / ligne_file

                    # Backup
                    if path.exists():
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        import shutil
                        shutil.copy(path, BASE_PATH / f"{ligne_file.replace('.xlsx', '')}_backup_{ts}.xlsx")

                    if save_with_calcul(export, path, ligne_name):
                        st.success(f"✅ {ligne_file} généré! ({len(export)} OFs)")
                        st.cache_data.clear()
                        st.info("💡 Le cache a été vidé. Les plannings utiliseront les nouvelles données.")

    st.divider()

    # Génération globale
    st.header("🚀 Génération Globale")

    if st.button("🚀 GÉNÉRER TOUS LES FICHIERS", type="primary"):
        results = []
        for lib_ligne, info in LIGNE_MAPPING.items():
            ligne_name = info["name"]
            ligne_file = info["file"]

            key_ordre = f"ordre_{ligne_name}"
            ordre = [c for c in st.session_state.get(key_ordre, []) if c]
            urgents = st.session_state.get(f"urg_{ligne_name}", [])
            journee = st.session_state.get(f"jour_{ligne_name}", [])

            if not ordre and not urgents and not journee:
                results.append(f"⏭️ {ligne_name}: ignoré")
                continue

            try:
                calcul_df = calcul_tables.get(ligne_name, pd.DataFrame())

                if ligne_name == "L1":
                    export = generate_L1(df_suivi, ordre, urgents, journee, calcul_df)
                elif ligne_name == "L2":
                    export = generate_L2(df_suivi, ordre, urgents, journee, calcul_df, col_k_name)
                elif ligne_name == "Imprimerie":
                    export = generate_Imp(df_suivi, ordre, urgents, journee, calcul_df)
                else:
                    export = generate_Vis(df_suivi, ordre, urgents, journee, calcul_df)

                path = BASE_PATH / ligne_file
                save_with_calcul(export, path, ligne_name)
                results.append(f"✅ {ligne_name}: {len(export)} OFs")
            except Exception as e:
                results.append(f"❌ {ligne_name}: {e}")

        for r in results:
            st.write(r)
        st.success("🎉 Terminé!")
        st.cache_data.clear()
        st.info("💡 Le cache a été vidé. Retournez sur les plannings pour voir les changements.")
