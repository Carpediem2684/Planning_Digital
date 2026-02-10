# ============================================
# Planning_L1.py — Planning L1 (Streamlit + Plotly)
# Version : 2026-02-05
# ============================================

# 1) IMPORTS & CONFIG
# --------------------------------------------
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
from pages.utils import load_suivi_of, get_ofs_exclus, get_statut_dict, is_statut_actif, STATUT_ACTIF

# st.set_page_config dans app.py

# 👉 Adapter ce chemin si besoin sur ton poste
BASE_PATH = Path(r"C:\\Users\\yannick.tetard\\OneDrive - GERFLOR\\Desktop\\Planning Streamlit\\xarpediem2684-repo-main")
OFS_L1_FILE = BASE_PATH / "OFs_L1.xlsx"
CAL_FILE = BASE_PATH / "Calendrier 2026 L1.xlsx"

LIGNE_NAME = "Ligne 1"  # nom affiché sur le Gantt

# ============================================
# 2) COULEURS PAR PRODUIT (mapping depuis VBA -> HEX)
# ============================================

def rgb_to_hex(r, g, b):
    return f"#{r:02X}{g:02X}{b:02X}"

PRODUIT_COLOR_MAP = {
    "INTRO": ("#FFFFFF", "#000000"),  # INTRO = fond blanc, texte noir
    "CICD03 3M": (rgb_to_hex(204, 204, 0), "#000000"),
    "CICD03 4M": (rgb_to_hex(255, 255, 0), "#000000"),
    "CICD05 4M": (rgb_to_hex(205, 133, 63), "#000000"),
    "CICD06 3M": (rgb_to_hex(255, 182, 193), "#000000"),
    "CICD06 4M": (rgb_to_hex(255, 119, 255), "#000000"),
    "CICD02 4M": (rgb_to_hex(211, 211, 211), "#000000"),
    "CICDMD01 4M": (rgb_to_hex(169, 169, 169), "#000000"),
    "CIMD02 3M": (rgb_to_hex(0, 128, 0), "#FFFFFF"),
    "CIMD02 4M": (rgb_to_hex(144, 238, 144), "#000000"),
    "CIMD03 3M": (rgb_to_hex(0, 255, 255), "#000000"),
    "CIMD03 4M": (rgb_to_hex(0, 100, 0), "#FFFFFF"),
    "INTRO POUR ENDUCTION L1-4140": (rgb_to_hex(0, 0, 0), "#FFFFFF"),
    "INTRO POUR ENDUCTION 4M": (rgb_to_hex(0, 0, 0), "#FFFFFF"),
    "INTRO POUR IMPRIMERIE 4M": (rgb_to_hex(0, 0, 0), "#FFFFFF"),
    "CICD01 4M": (rgb_to_hex(0, 139, 139), "#FFFFFF"),
    "CICD04 4M": (rgb_to_hex(173, 216, 230), "#000000"),
}

DEFAULT_BAR_COLOR = "#FFFFFF"
DEFAULT_TEXT_COLOR = "#000000"

# Durée de l'INTRO en heures (insérée automatiquement entre les campagnes)
INTRO_DUREE_H = 2.3

# ============================================
# 3) FONCTIONS UTILITAIRES
# ============================================

def parse_horaire(jour, horaire_str):
    """'jour' = Timestamp (date du jour)
    'horaire_str' = '12h30-20h30' -> (start_datetime, end_datetime)"""
    debut_str, fin_str = horaire_str.split('-')
    debut_str = debut_str.replace('h', ':')
    fin_str = fin_str.replace('h', ':')
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
    horizon_end = now + timedelta(days=horizon_days)
    end_date = horizon_end.date()

    mask = (cal["Jour"].dt.date >= start_date) & (cal["Jour"].dt.date <= end_date)
    sub = cal.loc[mask]

    slots = []
    for _, row in sub.iterrows():
        jour = row["Jour"]
        for i in [1, 2, 3]:
            etat_col = f"Etat_{i}"
            hor_col = f"Horaire_{i}"
            if etat_col in row and hor_col in row:
                etat = row[etat_col]
                h = row[hor_col]
                if etat == "OUVERT" and isinstance(h, str) and "-" in h:
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


def needs_intro(prev_produit, curr_produit):
    """Détermine si une INTRO est nécessaire entre deux produits.
    INTRO seulement quand :
    - Changement entre CICD et CIMD (dans les deux sens)
    - Changement entre CICDMD et CICD (dans les deux sens)
    - Passage de 3M à 4M (pas l'inverse)
    """
    if not prev_produit or not curr_produit:
        return False

    prev = str(prev_produit).upper()
    curr = str(curr_produit).upper()

    def get_support_type(p):
        if "CICDMD" in p:
            return "CICDMD"
        elif "CIMD" in p:
            return "CIMD"
        elif "CICD" in p:
            return "CICD"
        return "OTHER"

    def get_largeur(p):
        if "3M" in p:
            return "3M"
        elif "4M" in p:
            return "4M"
        return "UNKNOWN"

    prev_type = get_support_type(prev)
    curr_type = get_support_type(curr)
    prev_larg = get_largeur(prev)
    curr_larg = get_largeur(curr)

    if (prev_type == "CICD" and curr_type == "CIMD") or (prev_type == "CIMD" and curr_type == "CICD"):
        return True
    if (prev_type == "CICDMD" and curr_type == "CICD") or (prev_type == "CICD" and curr_type == "CICDMD"):
        return True
    if (prev_type == "CICDMD" and curr_type == "CIMD") or (prev_type == "CIMD" and curr_type == "CICDMD"):
        return True
    if prev_larg == "3M" and curr_larg == "4M":
        return True
    return False


def schedule_ofs_from_slots(ofs_df, slots):
    """Planifie les OFs (dans l'ordre fourni) dans les créneaux 'slots'.
    Insère automatiquement une INTRO selon les règles métier.
    Retourne un DataFrame de segments."""
    from datetime import timedelta

    planning_rows = []
    if not slots:
        return pd.DataFrame()

    slot_idx = 0
    current_slot_start = slots[0]["start"]
    current_slot_end = slots[0]["end"]

    previous_produit = None
    intro_counter = 0

    def consume_time(duration_h):
        """Consomme du temps dans les créneaux et retourne les segments créés."""
        nonlocal slot_idx, current_slot_start, current_slot_end
        segments = []
        remaining = timedelta(hours=duration_h)
        seg = 1

        while remaining > timedelta(0):
            if slot_idx >= len(slots):
                return segments, False  # Plus de créneaux

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

            segments.append({
                "start": start,
                "end": end,
                "duree_h": use.total_seconds() / 3600.0,
                "segment": seg,
            })

            seg += 1
            remaining -= use
            current_slot_start = end

        return segments, True

    for _, row in ofs_df.iterrows():
        current_produit = row.get("Produit", None)

        # INTRO ?
        if needs_intro(previous_produit, current_produit):
            intro_counter += 1
            intro_segments, ok = consume_time(INTRO_DUREE_H)
            for seg_data in intro_segments:
                planning_rows.append({
                    "Ofs": f"INTRO_{intro_counter}",
                    "Segment": seg_data["segment"],
                    "Produit": "INTRO",
                    "Ml": "",
                    "Campagne": "",
                    "start": seg_data["start"],
                    "end": seg_data["end"],
                    "duree_h": seg_data["duree_h"],
                    "is_intro": True,
                })
            if not ok:
                return pd.DataFrame(planning_rows)

        previous_produit = current_produit

        # Planifier l'OF
        of_id = row["ID_PLAN"]
        produit = row["Produit"]
        ml = row["Ml"]
        campagne = row.get("Campagne", "")
        duree_h = float(row["Temps en h"]) if pd.notna(row["Temps en h"]) else 0.0

        of_segments, ok = consume_time(duree_h)
        for seg_data in of_segments:
            planning_rows.append({
                "Ofs": of_id,
                "Segment": seg_data["segment"],
                "Produit": produit,
                "Ml": ml,
                "Campagne": campagne,
                "start": seg_data["start"],
                "end": seg_data["end"],
                "duree_h": seg_data["duree_h"],
                "is_intro": False,
            })
        if not ok:
            return pd.DataFrame(planning_rows)

    return pd.DataFrame(planning_rows)


# ============================================
# 4) CHARGEMENT DONNÉES
# ============================================
@st.cache_data(ttl=7200)
def load_data():
    ofs_l1 = pd.read_excel(OFS_L1_FILE, engine="openpyxl")
    cal = pd.read_excel(CAL_FILE, engine="openpyxl")
    return ofs_l1, cal


# ============================================
# 5) FONCTION PRINCIPALE D'AFFICHAGE L1
# ============================================

def show_planning_l1():
    # 5) UI STREAMLIT & RUBAN LIGNE 1
    st.title("📅 Planning LIGNE 1")

    # === Bouton retour menu ===
    st.markdown(""" """, unsafe_allow_html=True)
    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            st.rerun()

    ofs_l1_df, cal_df = load_data()

    # Charger SUIVI_OF pour statuts
    suivi_df = load_suivi_of()

    # Filtrer les OFs terminés (60, 61, 99) et < 200 ML
    ofs_exclus = get_ofs_exclus(suivi_df, "LIGNE1|L06")
    ofs_l1_df = ofs_l1_df[~ofs_l1_df["Ofs"].isin(ofs_exclus)]

    # Dict des statuts pour pastilles vertes
    statut_dict = get_statut_dict(suivi_df, "LIGNE1|L06")

    # Génère un identifiant unique pour chaque ligne
    ofs_l1_df["ID_PLAN"] = ofs_l1_df.index.astype(str) + "_" + ofs_l1_df["Ofs"].astype(str)

    # Créer un label lisible pour l'interface (OF + Produit)
    def make_display_label(row):
        produit = str(row["Produit"])[:20] if pd.notna(row["Produit"]) else ""
        return f"{row['Ofs']} - {produit}"

    ofs_l1_df["DISPLAY_LABEL"] = ofs_l1_df.apply(make_display_label, axis=1)

    with st.expander("Données OFs L1"):
        st.dataframe(ofs_l1_df)

    # Conserve la liste originale d'ID_PLAN
    if "ofs_list_original_L1" not in st.session_state:
        st.session_state.ofs_list_original_L1 = list(ofs_l1_df["ID_PLAN"].values)

    # Mapping ID_PLAN <-> label lisible
    id_to_label = dict(zip(ofs_l1_df["ID_PLAN"], ofs_l1_df["DISPLAY_LABEL"]))

    # Mapping ID_PLAN <-> Campagne
    id_to_campagne = dict(zip(ofs_l1_df["ID_PLAN"], ofs_l1_df["Campagne"]))

    # ---- 5.0 Réorganisation manuelle (OF ou Campagne)
    st.markdown("### 🔁 Réorganisation manuelle (OF ou Campagne)")

    # Initialisation de l'ordre dans la session
    if "ordre_ofs_L1" not in st.session_state:
        st.session_state.ordre_ofs_L1 = list(ofs_l1_df["ID_PLAN"].values)
        st.session_state.ordre_ofs_origine_L1 = st.session_state.ordre_ofs_L1.copy()

    ordre = st.session_state.ordre_ofs_L1

    # Créer les options lisibles pour les selectbox
    ordre_labels = [id_to_label.get(x, x) for x in ordre]
    label_to_id = {id_to_label.get(x, x): x for x in ordre}

    # Fonction pour obtenir les campagnes uniques dans l'ordre actuel
    def get_campagnes_in_order(ordre_ids):
        seen = set()
        campagnes = []
        for id_plan in ordre_ids:
            camp = id_to_campagne.get(id_plan, "")
            if camp and camp not in seen:
                seen.add(camp)
                campagnes.append(camp)
        return campagnes

    # Choix du mode : OF ou Campagne
    mode_deplacement = st.radio(
        "Déplacer :",
        ["Un OF", "Une Campagne entière"],
        horizontal=True,
        key="mode_deplacement_L1"
    )

    if mode_deplacement == "Un OF":
        # ---- Mode OF individuel ----
        col_move1, col_move2, col_move3, col_move4 = st.columns([2, 2, 1, 1])
        with col_move1:
            of_to_move_label = st.selectbox(
                "OF à déplacer",
                ordre_labels if ordre_labels else ["(aucun)"],
                index=0,
                key="of_to_move_select_L1"
            )
            of_to_move = label_to_id.get(of_to_move_label, of_to_move_label)
        with col_move2:
            cibles_labels = [id_to_label.get(x, x) for x in ordre if x != of_to_move]
            cible_label = st.selectbox(
                "Le placer par rapport à",
                cibles_labels if cibles_labels else ["(aucune cible)"],
                key="cible_select_L1"
            )
            cible = label_to_id.get(cible_label, cible_label)
        with col_move3:
            position = st.radio(
                "Position",
                ["Avant", "Après"],
                index=1,
                horizontal=True,
                key="position_radio_L1"
            )
        with col_move4:
            appliquer = st.button("Appliquer", use_container_width=True, key="btn_appliquer_of_L1")

        if appliquer and cibles_labels and of_to_move_label != "(aucun)" and cible_label != "(aucune cible)":
            new_order = ordre.copy()
            try:
                new_order.remove(of_to_move)
            except ValueError:
                pass
            if cible in new_order:
                idx = new_order.index(cible)
                insert_pos = idx if position == "Avant" else idx + 1
                new_order.insert(insert_pos, of_to_move)
                st.session_state.ordre_ofs_L1 = new_order
                st.success(f"OF déplacé {position.lower()} la cible sélectionnée.")
            else:
                st.warning("Cible introuvable dans l'ordre courant.")

    else:
        # ---- Mode Campagne entière ----
        campagnes_ordre = get_campagnes_in_order(ordre)
        col_camp1, col_camp2, col_camp3, col_camp4 = st.columns([2, 2, 1, 1])
        with col_camp1:
            campagne_to_move = st.selectbox(
                "Campagne à déplacer",
                campagnes_ordre if campagnes_ordre else ["(aucune)"],
                index=0,
                key="campagne_to_move_select_L1"
            )
        with col_camp2:
            cibles_campagnes = [c for c in campagnes_ordre if c != campagne_to_move]
            cible_campagne = st.selectbox(
                "La placer par rapport à",
                cibles_campagnes if cibles_campagnes else ["(aucune cible)"],
                key="cible_campagne_select_L1"
            )
        with col_camp3:
            position_camp = st.radio(
                "Position",
                ["Avant", "Après"],
                index=1,
                horizontal=True,
                key="position_campagne_radio_L1"
            )
        with col_camp4:
            appliquer_camp = st.button("Appliquer", use_container_width=True, key="btn_appliquer_camp_L1")

        if appliquer_camp and cibles_campagnes and campagne_to_move != "(aucune)" and cible_campagne != "(aucune cible)":
            ids_campagne_to_move = [id_plan for id_plan in ordre if id_to_campagne.get(id_plan) == campagne_to_move]
            new_order = [id_plan for id_plan in ordre if id_to_campagne.get(id_plan) != campagne_to_move]
            ids_cible = [id_plan for id_plan in new_order if id_to_campagne.get(id_plan) == cible_campagne]
            if ids_cible:
                if position_camp == "Avant":
                    insert_idx = new_order.index(ids_cible[0])
                else:
                    insert_idx = new_order.index(ids_cible[-1]) + 1
                for i, id_plan in enumerate(ids_campagne_to_move):
                    new_order.insert(insert_idx + i, id_plan)
                st.session_state.ordre_ofs_L1 = new_order
                st.success(f"Campagne '{campagne_to_move}' déplacée {position_camp.lower()} '{cible_campagne}'.")
            else:
                st.warning("Campagne cible introuvable.")

    # Boutons Réinitialiser + Valider
    col_reset, col_backup, col_save = st.columns([1, 2, 3])
    with col_reset:
        if st.button("Réinitialiser l'ordre (Excel)", use_container_width=True, key="reset_L1"):
            st.session_state.ordre_ofs_L1 = st.session_state.ordre_ofs_origine_L1.copy()
            st.info("Ordre réinitialisé à l'ordre Excel d'origine.")
    with col_backup:
        backup_before_save = st.checkbox("Créer une sauvegarde avant d'écrire", value=True, key="backup_L1")
    with col_save:
        if st.button("📥 Valider l'ordre et mettre à jour OFs_L1.xlsx", use_container_width=True, key="save_L1"):
            try:
                ordre_final = st.session_state.ordre_ofs_L1
                if sorted(ordre_final) != sorted(st.session_state.ofs_list_original_L1):
                    st.error("❌ L'ordre ne correspond pas exactement aux OFs d'origine.")
                    st.stop()
                df_excel = pd.read_excel(OFS_L1_FILE, engine="openpyxl")
                df_excel["ID_PLAN"] = df_excel.index.astype(str) + "_" + df_excel["Ofs"].astype(str)
                if backup_before_save:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_path = BASE_PATH / f"OFs_L1_backup_{timestamp}.xlsx"
                    df_excel.drop(columns=["ID_PLAN"]).to_excel(backup_path, index=False, engine="openpyxl")
                df_excel = df_excel.set_index("ID_PLAN")
                df_excel = df_excel.loc[ordre_final].reset_index(drop=True)
                df_excel.to_excel(OFS_L1_FILE, index=False, engine="openpyxl")
                st.success("✅ L'ordre a été appliqué et OFs_L1.xlsx a été mis à jour !")
                st.session_state.ofs_list_original_L1 = ordre_final.copy()
                st.session_state.ordre_ofs_origine_L1 = ordre_final.copy()
            except Exception as e:
                st.error(f"❌ Erreur lors de la mise à jour du fichier : {e}")

    # ---- 5.1 Créneaux ouverts ----
    now = datetime.now()
    horizon_days = 14
    horizon_end = now + timedelta(days=horizon_days)

    st.markdown(f"**🕒 Maintenant : {now.strftime('%d/%m/%Y %H:%M')}**")

    slots = build_open_slots_from_now(cal_df, horizon_days=horizon_days, now=now)
    if not slots:
        st.error("Aucun créneau OUVERT trouvé dans les 14 prochains jours.")
        st.stop()

    # ============================================
    # 🔒 4-BIS — SÉCURISATION DES ID POUR ÉVITER KeyError
    # ============================================
    # 1) Recréer ID_PLAN frais (au cas où Excel a bougé)
    ofs_l1_df["ID_PLAN"] = ofs_l1_df.index.astype(str) + "_" + ofs_l1_df["Ofs"].astype(str)

    # 2) Initialiser ordre si absent
    if "ordre_ofs_L1" not in st.session_state:
        st.session_state.ordre_ofs_L1 = list(ofs_l1_df["ID_PLAN"].values)

    current_ids = set(ofs_l1_df["ID_PLAN"])
    saved_order = st.session_state.ordre_ofs_L1

    # 3) Identifier les IDs manquants (cause du KeyError)
    ids_missing = [x for x in saved_order if x not in current_ids]

    def rebuild_order_from_ofs(old_ids, df_ids, ofs_series):
        """Reconstruit un ordre compatible à partir des OFs."""
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
        remaining = [x for x in df_ids if x not in set(new_order)]
        return new_order + remaining

    # 4) Correction automatique si incohérence
    if ids_missing:
        try:
            st.info("🔁 Ordre L1 invalide — tentative de réparation...")
            repaired = rebuild_order_from_ofs(
                saved_order,
                list(ofs_l1_df["ID_PLAN"]),
                ofs_l1_df["Ofs"]
            )
            if len(repaired) == len(ofs_l1_df) and len(set(repaired)) == len(repaired):
                st.session_state.ordre_ofs_L1 = repaired
                st.success("✅ Ordre réparé automatiquement.")
            else:
                raise Exception("Réparation partielle")
        except Exception:
            st.warning("⚠️ Impossible de réparer — ordre réinitialisé.")
            st.session_state.ordre_ofs_L1 = list(ofs_l1_df["ID_PLAN"])

    # 5) Dernier filet de sécurité
    ordre_final_l1 = st.session_state.ordre_ofs_L1
    ordre_final_l1 = [x for x in ordre_final_l1 if x in current_ids]
    st.session_state.ordre_ofs_L1 = ordre_final_l1

    # ---- 5.2 Planification des OFs ----
    ofs_l1_work = ofs_l1_df.set_index("ID_PLAN")
    ofs_l1_work = ofs_l1_work.loc[st.session_state.ordre_ofs_L1].reset_index()

    planning_df = schedule_ofs_from_slots(ofs_l1_work, slots)
    if planning_df.empty:
        st.warning("Planning L1 vide : pas assez de créneaux ouverts pour placer les OFs.")
        st.stop()

    # -------- Label utilisé dans les barres --------
    def build_label(row):
        # Si c'est une INTRO
        if row.get("is_intro", False) or row.get("Produit") == "INTRO":
            return f"<b>INTRO</b><br>{row['duree_h']:.1f} h"
        # Extraire le numéro d'OF depuis ID_PLAN
        of_display = row['Ofs']
        if isinstance(of_display, str) and "_" in of_display:
            of_display = of_display.split("_")[-1]
        return (
            f"<b>{row['Produit']}</b><br>"
            f"{int(row['Ml']) if pd.notna(row['Ml']) and row['Ml'] != '' else ''} ML<br>"
            f"OF {of_display} - {row['duree_h']:.1f} h"
        )

    # DataFrame pour le Gantt
    plot_df = planning_df.copy()
    plot_df["Ligne"] = LIGNE_NAME
    plot_df["Label"] = plot_df.apply(build_label, axis=1)

    # ---- 5.3 Paramètres d'affichage ----
    st.markdown("### 🔎 Fenêtre d'affichage")
    largeur_heures = st.selectbox(
        "Largeur de vue",
        [12, 24, 36, 48, 72],
        index=1,
        format_func=lambda x: f"{x}h ({x//24}j {x%24}h)" if x >= 24 else f"{x}h",
        key="largeur_vue_L1"
    )

    total_hours = (horizon_end - now).total_seconds() / 3600
    max_offset = max(0, total_hours - largeur_heures)

    if "offset_heures_L1" not in st.session_state:
        st.session_state.offset_heures_L1 = 0.0

    offset_heures = st.session_state.offset_heures_L1
    view_start = now + timedelta(hours=offset_heures)
    view_end = view_start + timedelta(hours=largeur_heures)
    if view_end > horizon_end:
        view_end = horizon_end

    # ---- 5.4 Gantt des OFs ----
    fig = px.timeline(
        plot_df,
        x_start="start",
        x_end="end",
        y="Ligne",
        color="Produit",
        text="Label",
    )

    fig.update_yaxes(autorange="reversed")

    # Appliquer les couleurs par Produit
    for trace in fig.data:
        prod = trace.name
        bar_color, text_color = PRODUIT_COLOR_MAP.get(
            prod, (DEFAULT_BAR_COLOR, DEFAULT_TEXT_COLOR)
        )
        trace.marker.color = bar_color
        trace.marker.line.color = "black"
        trace.marker.line.width = 1
        trace.textfont = dict(size=9, family="Arial", color=text_color)
        trace.textposition = "inside"
        trace.insidetextanchor = "middle"
        trace.hovertemplate = "<b>Détail :</b><br>%{text}<extra></extra>"
        trace.showlegend = False

    # ---- Pastilles vertes ----
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
                y="Ligne 1",
                yshift=-35,
                text="🟢",
                showarrow=False,
                font=dict(size=12),
                xref="x",
                yref="y",
            )

    # ---- ARRETS = créneaux FERMÉ (en rouge) ----
    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])
    start_date_cal = now.date()
    end_date_cal = horizon_end.date()
    mask_cal = (cal["Jour"].dt.date >= start_date_cal) & (cal["Jour"].dt.date <= end_date_cal)
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

    # Axe temps
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

    # Jours en haut
    JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    current_day = view_start.date()
    end_day = view_end.date()
    day_annotations = []
    while current_day <= end_day:
        day_middle = datetime.combine(current_day, datetime.min.time()) + timedelta(hours=12)
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

    # Slider de défilement horizontal
    new_offset = st.slider(
        "⬅️ Faire défiler le planning ➡️",
        min_value=0.0,
        max_value=max(1.0, max_offset),
        value=float(st.session_state.offset_heures_L1),
        step=1.0,
        format="%.0fh",
        key="slider_defilement_L1"
    )

    if new_offset != st.session_state.offset_heures_L1:
        st.session_state.offset_heures_L1 = new_offset
        st.rerun()

    view_start_display = now + timedelta(hours=new_offset)
    view_end_display = view_start_display + timedelta(hours=largeur_heures)
    st.caption(f"📍 Vue : {view_start_display.strftime('%d/%m %Hh')} → {view_end_display.strftime('%d/%m %Hh')}")

    # Tableau debug
    with st.expander("Segments planifiés L1 (debug)"):
        df_debug = planning_df.copy()
        df_debug["Ml"] = pd.to_numeric(df_debug["Ml"], errors="coerce")
        st.dataframe(df_debug)
