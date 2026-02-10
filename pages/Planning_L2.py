# ============================================
# Planning_L2.py — Planning L2 (Streamlit + Plotly)
# Version sécurisée : 2026-02-05 — Yannick (TY) + M365 Copilot
# ============================================

# 1) IMPORTS & CONFIG
# --------------------------------------------
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
from pages.utils import (
    load_suivi_of,
    get_ofs_exclus,
    get_statut_dict,
    get_fabrique_dict,
    is_statut_actif,
    STATUT_ACTIF,
)

# 👉 Adapter ce chemin si besoin sur ton poste
BASE_PATH = Path(
    r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main"
)
OFS_L2_FILE = BASE_PATH / "OFs_L2.xlsx"
CAL_FILE = BASE_PATH / "Calendrier 2026.xlsx"

LIGNE_NAME = "Ligne 2"  # nom affiché sur le Gantt

# ============================================
# 2) COULEURS PAR FAMILLE
# ============================================
FAMILLE_COLOR_MAP = {
    "INTRO": ("#FFFFFF", "#000000"),
    "TIMBERLINE/NEROK 3M": ("#00008B", "#FFFFFF"),
    "START 4M": ("#D3D3D3", "#000000"),
    "SPORISOL 4M": ("#C0C0C0", "#000000"),
    "TARASTEP PRO": ("#FFFF66", "#000000"),
    "PRIMETEX GRT 4M": ("#CCCC00", "#000000"),
    "TEXLINE GRIP'X 3M": ("#FFB6C1", "#000000"),
    "TEXLINE GRIP'X 4M": ("#8B008B", "#FFFFFF"),
    "NEROK TEX": ("#000000", "#FFFFFF"),
    "NEROK 50 TEX": ("#000000", "#FFFFFF"),
    "BAGNOSTAR 4M": ("#FF6666", "#FFFFFF"),
    "BAGNOSTAR METAL 4M": ("#FF6666", "#FFFFFF"),
    "BAGNOSTAR 3M": ("#FF6666", "#FFFFFF"),
    "SRA ACOUSTIC 3M": ("#808080", "#FFFFFF"),
    "FUSION 3M": ("#006400", "#FFFFFF"),
    "TARABUS HARMONIA 1/2": ("#50C878", "#000000"),
    "RECONCEPTION TARABUS H 1/2": ("#50C878", "#000000"),
    "TRADIFLOR 2S2 3M": ("#483C32", "#FFFFFF"),
    "TRADIFLOR 2S2 4M": ("#483C32", "#FFFFFF"),
    "TRANSIT-TEX MAX 33-43 2/2 4M": ("#654321", "#FFFFFF"),
    "TARABUS HARMONIA INTER": ("#ADD8E6", "#000000"),
    "RECONCEPTION TARABUS H 2/2": ("#ADD8E6", "#000000"),
    "GERBAD EVOLUTION 3M": ("#E0FFFF", "#000000"),
    "GERBAD EVOLUTION 4M": ("#E0FFFF", "#000000"),
    "TIMBERLINE 4M": ("#E0FFFF", "#000000"),
    "SRA ACOUSTIC 4M": ("#F5F5F5", "#000000"),
    "TRANSIT-TEX MAX 33-43 1/2 4M": ("#CD853F", "#FFFFFF"),
    "PRIMETEX GRT 3M": ("#C3B091", "#000000"),
    "BAGNOSTAR 2.5 3M": ("#E6E6FA", "#000000"),
    "BAGNOSTAR 2.5 4M": ("#E6E6FA", "#000000"),
    "BAGNOSTAR 2.5 METAL 4M": ("#E6E6FA", "#000000"),
    "TRANSIT TEX MAX 2S3 1/2 4M": ("#D2B48C", "#000000"),
    "BOOSTER 2.6 DIAM 4M": ("#FA8072", "#000000"),
    "MELODY 4M": ("#FA8072", "#000000"),
    "LOFTEX NATURE 4M": ("#FF8C00", "#000000"),
    "BAGNOSTAR MATT 4M": ("#FF6961", "#FFFFFF"),
    "SRA ACOUSTIC PU 4M": ("#FF6961", "#FFFFFF"),
    "TRANSIT-TEX 4M": ("#FF77FF", "#000000"),
    "TRANSIT-TEX 2/2 4M": ("#FF77FF", "#000000"),
    "TEXLINE NATURE 4M": ("#FFA500", "#000000"),
    "TEXLINE NATURE 3M": ("#FFA500", "#000000"),
    "LOFTEX GRT 3M": ("#FFA500", "#000000"),
    "LOFTEX GRT 4M": ("#FFA500", "#000000"),
    "LOFTEX NATURE 3M": ("#FFA500", "#000000"),
    "BOOSTER 2.6 3M": ("#FFA500", "#000000"),
    "PRIMETEX 4M": ("#FDFD96", "#000000"),
    "PRIMETEX 3M": ("#FDFD96", "#000000"),
    "PRIMETEX MATT 3M": ("#FDFD96", "#000000"),
    "PRIMETEX MATT 4M": ("#FDFD96", "#000000"),
    "START 3M": ("#2F4F4F", "#FFFFFF"),
    "GRAFIC PU 3M": ("#2F4F4F", "#FFFFFF"),
    "GRAFIC PU 4M": ("#2F4F4F", "#FFFFFF"),
    "TARABUS HARMONIA": ("#000080", "#FFFFFF"),
    "TIMBERLINE/NEROK 4M": ("#000080", "#FFFFFF"),
    "TRANSIT TEX MAX 2S3 2/2 4M": ("#483C32", "#FFFFFF"),
    "TEXLINE HQR 3M": ("#008000", "#FFFFFF"),
    "TEXLINE HQR 4M": ("#008000", "#FFFFFF"),
    "TEXLINE GRT 3M": ("#008000", "#FFFFFF"),
    "FUSION 4M": ("#CCFFCC", "#000000"),
    "TEXLINE GRT 4M": ("#006400", "#FFFFFF"),
    "NEROK TEX NATURE": ("#FFFFFF", "#000000"),
    "NERA FIRST 4M": ("#FFFFFF", "#000000"),
    "TRANSIT-TEX PLUS 2/2 4M": ("#FFFFFF", "#000000"),
    "BOOSTER 2.6 GRAIN 3M PUR BLANC": ("#FFFFFF", "#000000"),
    "TRANSIT-TEX PLUS 1/2 4M": ("#FFFFFF", "#000000"),
    "ESSAI UAP4M L2 1/2": ("#FFFFFF", "#000000"),
    "BOOSTER PUR BLANC 4M": ("#FFFFFF", "#000000"),
    "BOOSTER 2.6 GRAIN 3M": ("#FFC0CB", "#000000"),
    "ESSAI UAP4M L2 1/1": ("#FFFFFF", "#000000"),
    "BOOSTER 2.6 4M": ("#FF0000", "#FFFFFF"),
}

DEFAULT_BAR_COLOR = "#FFFFFF"
DEFAULT_TEXT_COLOR = "#000000"

INTRO_DUREE_H = 2.3
INTRO_COLOR = ("#FFFFFF", "#000000")
# ============================================
# 3) FONCTIONS UTILITAIRES & SÉCURITÉ
# ============================================

def parse_horaire(jour, horaire_str):
    """'jour' = Timestamp (date du jour)
    'horaire_str' = '12h30-20h30' -> (start_datetime, end_datetime)"""
    debut_str, fin_str = horaire_str.split("-")
    debut_str = debut_str.replace("h", ":")
    fin_str = fin_str.replace("h", ":")
    start = datetime.combine(jour.date(), datetime.strptime(debut_str, "%H:%M").time())
    end = datetime.combine(jour.date(), datetime.strptime(fin_str, "%H:%M").time())
    if end <= start:
        end += timedelta(days=1)
    return start, end


def build_open_slots_from_now(cal_df, horizon_days=14, now=None):
    """Retourne une liste ordonnée de créneaux OUVERT :
    [{'start': datetime, 'end': datetime}, ...] à partir de 'now' sur horizon_days."""
    if now is None:
        now = datetime.now()
    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])

    start_date = now.date()
    end_date = (now + timedelta(days=horizon_days)).date()

    mask = (cal["Jour"].dt.date >= start_date) & (cal["Jour"].dt.date <= end_date)
    sub = cal.loc[mask]

    horizon_end = now + timedelta(days=horizon_days)
    slots = []

    for _, row in sub.iterrows():
        jour = row["Jour"]
        for i in [1, 2, 3]:
            etat_col = f"Etat_{i}"
            hor_col = f"Horaire_{i}"
            if row.get(etat_col, "FERME") == "OUVERT":
                h = row[hor_col]
                if isinstance(h, str) and "-" in h:
                    start, end = parse_horaire(jour, h)
                    if end <= now:
                        continue
                    if start < now < end:
                        start = now
                    if start < horizon_end:
                        end = min(end, horizon_end)
                        slots.append({"start": start, "end": end})

    slots = sorted(slots, key=lambda x: x["start"])
    return slots


def schedule_ofs_from_slots(ofs_df, slots):
    """Planifie les OFs (dans l'ordre fourni) dans les créneaux 'slots'.
    Insère une INTRO automatiquement entre les campagnes.
    Retourne un DataFrame de segments."""
    from datetime import timedelta

    planning_rows = []
    if not slots:
        return pd.DataFrame()

    slot_idx = 0
    current_slot_start = slots[0]["start"]
    current_slot_end = slots[0]["end"]

    previous_campagne = None
    intro_counter = 0

    def consume_time(duration_h):
        """Consomme du temps dans les créneaux et retourne les segments créés."""
        nonlocal slot_idx, current_slot_start, current_slot_end
        segments = []
        remaining = timedelta(hours=duration_h)
        seg = 1

        while remaining > timedelta(0):
            if slot_idx >= len(slots):
                return segments, False

            available = current_slot_end - current_slot_start
            if available <= timedelta(0):
                slot_idx += 1
                if slot_idx >= len(slots):
                    return segments, False
                current_slot_start = slots[slot_idx]["start"]
                current_slot_end = slots[slot_idx]["end"]
                continue

            use = min(available, remaining)
            start = current_slot_start
            end = start + use

            segments.append(
                {
                    "start": start,
                    "end": end,
                    "duree_h": use.total_seconds() / 3600.0,
                    "segment": seg,
                }
            )

            seg += 1
            remaining -= use
            current_slot_start = end

        return segments, True

    for _, row in ofs_df.iterrows():
        current_campagne = row.get("Campagne", None)

        if previous_campagne is not None and current_campagne != previous_campagne:
            intro_counter += 1
            intro_segments, ok = consume_time(INTRO_DUREE_H)
            for seg_data in intro_segments:
                planning_rows.append(
                    {
                        "Ofs": f"INTRO_{intro_counter}",
                        "Segment": seg_data["segment"],
                        "COLORIS": "INTRO",
                        "GRAIN": "",
                        "FAMILLE": "INTRO",
                        "ML": "",
                        "Campagne": "",
                        "start": seg_data["start"],
                        "end": seg_data["end"],
                        "duree_h": seg_data["duree_h"],
                        "is_intro": True,
                    }
                )
            if not ok:
                return pd.DataFrame(planning_rows)

        previous_campagne = current_campagne

        of_id = row["ID_PLAN"]
        coloris = row["COLORIS"]
        grain = row["GRAIN"]
        famille = row["FAMILLE"]
        ml = row["ML"]
        campagne = row.get("Campagne", "")
        duree_h = float(row["Temps en h"]) if pd.notna(row["Temps en h"]) else 0.0

        of_segments, ok = consume_time(duree_h)
        for seg_data in of_segments:
            planning_rows.append(
                {
                    "Ofs": of_id,
                    "Segment": seg_data["segment"],
                    "COLORIS": coloris,
                    "GRAIN": grain,
                    "FAMILLE": famille,
                    "ML": ml,
                    "Campagne": campagne,
                    "start": seg_data["start"],
                    "end": seg_data["end"],
                    "duree_h": seg_data["duree_h"],
                    "is_intro": False,
                }
            )

        if not ok:
            return pd.DataFrame(planning_rows)

    return pd.DataFrame(planning_rows)


def split_coloris(value: str):
    """'322670-SHERWOOD BLOND' -> ('322670', 'SHERWOOD BLOND')"""
    if isinstance(value, str) and "-" in value:
        code, desc = value.split("-", 1)
        return code.strip(), desc.strip()
    return value, ""


def has_consecutive_duplicates(seq):
    return any(a == b for a, b in zip(seq, seq[1:]))


def has_any_duplicates(seq):
    return len(seq) != len(set(seq))


def normalize_of(x):
    try:
        return str(int(float(x)))
    except Exception:
        return str(x).strip()


# ============================================
# 4) CHARGEMENT DONNÉES
# ============================================

@st.cache_data(ttl=7200)
def load_data():
    ofs_l2 = pd.read_excel(OFS_L2_FILE, engine="openpyxl")
    cal = pd.read_excel(CAL_FILE, engine="openpyxl")
    return ofs_l2, cal
# ============================================
# 5) FONCTION PRINCIPALE D'AFFICHAGE
# ============================================

def show_planning_l2():

    # Configuration de la page
    st.set_page_config(
        page_title="Planning Ligne 2",
        layout="wide",
    )

    # Titre
    st.title("📅 Planning LIGNE 2 - Style Excel (ruban)")

    # --- Bouton retour menu ---
    st.markdown(""" """, unsafe_allow_html=True)
    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            return

    # Charger les données
    ofs_l2_df, cal_df = load_data()
    suivi_df = load_suivi_of()

    # Filtrer les OF exclus
    ofs_exclus = get_ofs_exclus(suivi_df, "LIGNE2|L08")
    ofs_l2_df = ofs_l2_df[~ofs_l2_df["Ofs"].isin(ofs_exclus)]

    # Dictionnaires statuts / impression
    statut_dict = get_statut_dict(suivi_df, "LIGNE2|L08")
    fabrique_dict = get_fabrique_dict(suivi_df)

    # Créer un ID unique pour chaque ligne
    ofs_l2_df["ID_PLAN"] = ofs_l2_df.index.astype(str) + "_" + ofs_l2_df["Ofs"].astype(str)

    # Label affiché dans l’UI
    def make_display_label(row):
        coloris_short = str(row["COLORIS"])[:25] if pd.notna(row["COLORIS"]) else ""
        return f"{row['Ofs']} - {coloris_short}"

    ofs_l2_df["DISPLAY_LABEL"] = ofs_l2_df.apply(make_display_label, axis=1)

    with st.expander("Données OFs L2"):
        st.dataframe(ofs_l2_df)

    # Sauvegarde de la liste originale
    if "ofs_list_original" not in st.session_state:
        st.session_state.ofs_list_original = list(ofs_l2_df["ID_PLAN"].values)

    id_to_label = dict(zip(ofs_l2_df["ID_PLAN"], ofs_l2_df["DISPLAY_LABEL"]))
    id_to_campagne = dict(zip(ofs_l2_df["ID_PLAN"], ofs_l2_df["Campagne"]))

    # ---- Réorganisation manuelle ----
    st.markdown("### 🔁 Réorganisation manuelle (OF ou Campagne)")

    if "ordre_ofs" not in st.session_state:
        st.session_state.ordre_ofs = list(ofs_l2_df["ID_PLAN"].values)
        st.session_state.ordre_ofs_origine = st.session_state.ordre_ofs.copy()

    ordre = st.session_state.ordre_ofs
    ordre_labels = [id_to_label.get(x, x) for x in ordre]
    label_to_id = {id_to_label.get(x, x): x for x in ordre}

    def get_campagnes_in_order(ordre_ids):
        seen = set()
        campagnes = []
        for id_plan in ordre_ids:
            camp = id_to_campagne.get(id_plan, "")
            if camp and camp not in seen:
                seen.add(camp)
                campagnes.append(camp)
        return campagnes

    mode_deplacement = st.radio(
        "Déplacer :",
        ["Un OF", "Une Campagne entière"],
        horizontal=True,
        key="mode_deplacement"
    )

    # ------------------------------
    # MODE : Déplacement d’un OF
    # ------------------------------
    if mode_deplacement == "Un OF":
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])

        with col1:
            of_to_move_label = st.selectbox(
                "OF à déplacer",
                ordre_labels,
                index=0,
                key="of_to_move_select"
            )
            of_to_move = label_to_id.get(of_to_move_label, of_to_move_label)

        with col2:
            cibles_labels = [id_to_label.get(x, x) for x in ordre if x != of_to_move]
            cible_label = st.selectbox(
                "Le placer par rapport à",
                cibles_labels,
                key="cible_select"
            )
            cible = label_to_id.get(cible_label, cible_label)

        with col3:
            position = st.radio("Position", ["Avant", "Après"], index=1, horizontal=True)

        with col4:
            appliquer = st.button("Appliquer", use_container_width=True)

        if appliquer:
            new_order = ordre.copy()
            try:
                new_order.remove(of_to_move)
            except:
                pass

            if cible in new_order:
                idx = new_order.index(cible)
                insert_pos = idx if position == "Avant" else idx + 1
                new_order.insert(insert_pos, of_to_move)
                st.session_state.ordre_ofs = new_order
                st.success("Déplacement effectué !")

    # ------------------------------
    # MODE : Déplacement d’une campagne
    # ------------------------------
    else:
        campagnes_ordre = get_campagnes_in_order(ordre)
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])

        with col1:
            campagne_to_move = st.selectbox(
                "Campagne à déplacer",
                campagnes_ordre,
                key="campagne_to_move_select"
            )

        with col2:
            cible_campagne = st.selectbox(
                "La placer par rapport à",
                [c for c in campagnes_ordre if c != campagne_to_move],
                key="cible_campagne_select"
            )

        with col3:
            position_camp = st.radio(
                "Position", ["Avant", "Après"], index=1, horizontal=True
            )

        with col4:
            appliquer_camp = st.button("Appliquer", use_container_width=True)

        if appliquer_camp:
            ids_campagne_to_move = [
                id_plan for id_plan in ordre if id_to_campagne.get(id_plan) == campagne_to_move
            ]

            new_order = [
                id_plan for id_plan in ordre if id_to_campagne.get(id_plan) != campagne_to_move
            ]

            ids_cible = [
                id_plan for id_plan in new_order if id_to_campagne.get(id_plan) == cible_campagne
            ]

            if ids_cible:
                insert_idx = (
                    new_order.index(ids_cible[0])
                    if position_camp == "Avant"
                    else new_order.index(ids_cible[-1]) + 1
                )

                for i, id_plan in enumerate(ids_campagne_to_move):
                    new_order.insert(insert_idx + i, id_plan)

                st.session_state.ordre_ofs = new_order
                st.success("Campagne déplacée !")

    # -------------------------
    # Boutons de sauvegarde
    # -------------------------
    col_reset, col_backup, col_save = st.columns([1, 2, 3])

    with col_reset:
        if st.button("Réinitialiser l'ordre"):
            st.session_state.ordre_ofs = st.session_state.ordre_ofs_origine.copy()
            st.info("Ordre réinitialisé.")

    with col_backup:
        backup_before_save = st.checkbox("Créer une sauvegarde", value=True)

    with col_save:
        if st.button("📥 Valider et écrire"):
            try:
                ordre_final = st.session_state.ordre_ofs

                df_excel = pd.read_excel(OFS_L2_FILE, engine="openpyxl")
                df_excel["ID_PLAN"] = df_excel.index.astype(str) + "_" + df_excel["Ofs"].astype(str)

                if backup_before_save:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup = BASE_PATH / f"OFs_L2_backup_{ts}.xlsx"
                    df_excel.drop(columns=["ID_PLAN"]).to_excel(backup, index=False)

                df_excel = df_excel.set_index("ID_PLAN")
                df_excel = df_excel.loc[ordre_final].reset_index(drop=True)
                df_excel.to_excel(OFS_L2_FILE, index=False)

                st.success("Ordre mis à jour avec succès !")

            except Exception as e:
                st.error(f"Erreur : {e}")

    # ---- Créneaux ouverts ----
    now = datetime.now()
    horizon_days = 14
    horizon_end = now + timedelta(days=horizon_days)

    st.markdown(f"**🕒 Maintenant : {now.strftime('%d/%m/%Y %H:%M')}**")

    slots = build_open_slots_from_now(cal_df, horizon_days, now)
    if not slots:
        st.error("Aucun créneau OUVERT trouvé.")
        st.stop()

    # -------------------------
    # Sécurisation des ID_PLAN
    # -------------------------
    ofs_l2_df["ID_PLAN"] = ofs_l2_df.index.astype(str) + "_" + ofs_l2_df["Ofs"].astype(str)

    current_ids = set(ofs_l2_df["ID_PLAN"])
    saved_order = st.session_state.ordre_ofs

    def rebuild_order_from_ofs(old_ids, df_ids, ofs_series):
        ids_by_of = {}
        for id_val, ofs in zip(df_ids, ofs_series.astype(str)):
            ids_by_of.setdefault(ofs, []).append(id_val)

        counters = {k: 0 for k in ids_by_of}

        new_order = []
        for old in old_ids:
            ofs = str(old).split("_")[-1]
            if ofs in ids_by_of and counters[ofs] < len(ids_by_of[ofs]):
                new_order.append(ids_by_of[ofs][counters[ofs]])
                counters[ofs] += 1

        remaining = [i for i in df_ids if i not in set(new_order)]
        return new_order + remaining

    ids_missing = [x for x in saved_order if x not in current_ids]

    if ids_missing:
        try:
            repaired = rebuild_order_from_ofs(
                saved_order,
                list(ofs_l2_df["ID_PLAN"]),
                ofs_l2_df["Ofs"],
            )
            if len(repaired) == len(ofs_l2_df):
                st.session_state.ordre_ofs = repaired
                st.success("Ordre réparé automatiquement !")
        except:
            st.session_state.ordre_ofs = list(ofs_l2_df["ID_PLAN"])
            st.warning("Ordre incompatible — réinitialisation.")

    ordre_final_l2 = [x for x in st.session_state.ordre_ofs if x in current_ids]
    st.session_state.ordre_ofs = ordre_final_l2

    # --------------------------------------------------------
    # PLANIFICATION DES OFS SUR LES CRÉNEAUX OUVERTS
    # --------------------------------------------------------

    ofs_l2_work = ofs_l2_df.set_index("ID_PLAN")
    ofs_l2_work = ofs_l2_work.loc[st.session_state.ordre_ofs].reset_index()

    planning_df = schedule_ofs_from_slots(ofs_l2_work, slots)

    if planning_df.empty:
        st.warning("Planning vide : pas assez de créneaux.")
        st.stop()

    planning_df[["Code_Produit", "Description"]] = planning_df["COLORIS"].apply(
        lambda x: pd.Series(split_coloris(x))
    )

    # --------------------------------------------------------
    # LABEL AFFICHÉ DANS LE GANTT
    # --------------------------------------------------------

    def build_label(row):
        if row.get("is_intro", False) or row.get("FAMILLE") == "INTRO":
            return f"<b>INTRO</b><br>{row['duree_h']:.1f} h"

        of_display = row["Ofs"]
        if isinstance(of_display, str) and "_" in of_display:
            of_display = of_display.split("_")[-1]

        coloris = row.get("COLORIS", "")
        imprime_label = ""
        if coloris in fabrique_dict:
            imprime_label = f"<br><b style='color:lime;'>✓ {fabrique_dict[coloris]:.0f} Imprimé</b>"

        return (
            f"{row['Code_Produit']}<br>"
            f"{row['Description']}<br>"
            f"{row['FAMILLE']} "
            f"{row['GRAIN'] if pd.notna(row['GRAIN']) else ''} - "
            f"{int(row['ML']) if pd.notna(row['ML']) else ''} ML / OF {of_display} - "
            f"{row['duree_h']:.1f} h"
            f"{imprime_label}"
        )

    # --------------------------------------------------------
    # GÉNÉRATION DU GANTT
    # --------------------------------------------------------

    plot_df = planning_df.copy()
    plot_df["Ligne"] = LIGNE_NAME
    plot_df["Label"] = plot_df.apply(build_label, axis=1)

    st.markdown("### 🔎 Fenêtre d'affichage")

    largeur_heures = st.selectbox(
        "Largeur de vue",
        [12, 24, 36, 48, 72],
        index=1,
        format_func=lambda x: f"{x}h ({x//24}j {x%24}h)" if x >= 24 else f"{x}h",
    )

    total_hours = (horizon_end - now).total_seconds() / 3600
    max_offset = max(0, total_hours - largeur_heures)

    if "offset_heures" not in st.session_state:
        st.session_state.offset_heures = 0.0

    offset_heures = st.session_state.offset_heures
    view_start = now + timedelta(hours=offset_heures)
    view_end = view_start + timedelta(hours=largeur_heures)

    if view_end > horizon_end:
        view_end = horizon_end

    fig = px.timeline(
        plot_df,
        x_start="start",
        x_end="end",
        y="Ligne",
        color="FAMILLE",
        text="Label",
    )

    fig.update_yaxes(autorange="reversed")

    for trace in fig.data:
        fam = trace.name
        bar_color, text_color = FAMILLE_COLOR_MAP.get(
            fam, (DEFAULT_BAR_COLOR, DEFAULT_TEXT_COLOR)
        )
        trace.marker.color = bar_color
        trace.marker.line.color = "black"
        trace.marker.line.width = 1
        trace.textfont.color = text_color
        trace.textfont.size = 9
        trace.textfont.family = "Arial"
        trace.textposition = "inside"
        trace.insidetextanchor = "middle"
        trace.hovertemplate = "<b>Détail :</b><br>%{text}<extra></extra>"
        trace.showlegend = False

    # --------------------------------------------------------
    # PASTILLES VERTES POUR STATUTS
    # --------------------------------------------------------

    for _, row in plot_df.iterrows():
        if row.get("is_intro", False):
            continue

        of_num = row["Ofs"]
        if isinstance(of_num, str) and "_" in of_num:
            try:
                of_num = int(of_num.split("_")[-1])
            except:
                continue

        if is_statut_actif(of_num, statut_dict):
            mid_time = row["start"] + (row["end"] - row["start"]) / 2
            fig.add_annotation(
                x=mid_time,
                y=LIGNE_NAME,
                yshift=-40,
                text="🟢",
                showarrow=False,
                font=dict(size=12),
            )

    # --------------------------------------------------------
    # ARRETS (FERMÉ) EN ROUGE
    # --------------------------------------------------------

    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])

    start_date_cal = now.date()
    end_date_cal = horizon_end.date()

    mask_cal = (cal["Jour"].dt.date >= start_date_cal) & (
        cal["Jour"].dt.date <= end_date_cal
    )
    sub_cal = cal.loc[mask_cal]

    for _, row_cal in sub_cal.iterrows():
        jour = row_cal["Jour"]
        for i in [1, 2, 3]:
            etat_col = f"Etat_{i}"
            hor_col = f"Horaire_{i}"

            if row_cal.get(etat_col, "FERME") == "FERME":
                h = row_cal[hor_col]

                if isinstance(h, str) and "-" in h:
                    start_arret, end_arret = parse_horaire(jour, h)

                    if end_arret <= now or start_arret >= horizon_end:
                        continue

                    if start_arret < now:
                        start_arret = now

                    if end_arret > horizon_end:
                        end_arret = horizon_end

                    fig.add_shape(
                        type="rect",
                        x0=start_arret,
                        x1=end_arret,
                        y0=0.05,
                        y1=0.95,
                        xref="x",
                        yref="paper",
                        fillcolor="#ff0000",
                        opacity=0.95,
                        line_width=0,
                    )

                    fig.add_annotation(
                        x=start_arret + (end_arret - start_arret) / 2,
                        y=0.5,
                        xref="x",
                        yref="paper",
                        text="ARRET 2x8",
                        showarrow=False,
                        font=dict(color="white", size=12, family="Arial Black"),
                    )

    # --------------------------------------------------------
    # AXE TEMPS
    # --------------------------------------------------------

    fig.update_xaxes(
        type="date",
        range=[view_start, view_end],
        showgrid=True,
        gridcolor="#777777",
        tickformat="%Hh",
        ticks="outside",
        ticklen=5,
        dtick=3600000,
        fixedrange=True,
        side="bottom",
    )

    fig.update_yaxes(title="", showticklabels=False, fixedrange=True)

    # --------------------------------------------------------
    # ANNOTATIONS DES JOURS
    # --------------------------------------------------------

    JOURS_FR = [
        "Lundi",
        "Mardi",
        "Mercredi",
        "Jeudi",
        "Vendredi",
        "Samedi",
        "Dimanche",
    ]

    current_day = view_start.date()
    end_day = view_end.date()
    day_annotations = []

    while current_day <= end_day:
        day_middle = datetime.combine(current_day, datetime.min.time()) + timedelta(
            hours=12
        )

        jour_semaine = JOURS_FR[current_day.weekday()]
        day_label = f"{jour_semaine} {current_day.strftime('%d/%m/%y')}"

        day_annotations.append(
            dict(
                x=day_middle,
                y=1.08,
                xref="x",
                yref="paper",
                text=f"<b>{day_label}</b>",
                showarrow=False,
                font=dict(color="white", size=11, family="Arial"),
                xanchor="center",
            )
        )

        current_day += timedelta(days=1)

    for annot in day_annotations:
        fig.add_annotation(annot)

    fig.add_vline(x=now, line_color="white", line_width=2, line_dash="dot")

    fig.update_layout(
        height=350,
        margin=dict(l=10, r=10, t=60, b=40),
        plot_bgcolor="#444444",
        paper_bgcolor="#444444",
        showlegend=False,
        dragmode=False,
        font=dict(color="white"),
    )

    st.plotly_chart(fig, use_container_width=True)

    new_offset = st.slider(
        "⬅️ Faire défiler le planning ➡️",
        min_value=0.0,
        max_value=max(1.0, max_offset),
        value=float(st.session_state.offset_heures),
        step=1.0,
        format="%.0fh",
    )

    if new_offset != st.session_state.offset_heures:
        st.session_state.offset_heures = new_offset
        st.rerun()

    view_start_display = now + timedelta(hours=new_offset)
    view_end_display = view_start_display + timedelta(hours=largeur_heures)

    st.caption(
        f"📍 Vue : {view_start_display.strftime('%d/%m %Hh')} → "
        f"{view_end_display.strftime('%d/%m %Hh')}"
    )

    # Debug final
    with st.expander("Segments planifiés L2 (debug)"):
        df_debug = planning_df.copy()
        df_debug["ML"] = pd.to_numeric(df_debug["ML"], errors="coerce")
        st.dataframe(df_debug)
