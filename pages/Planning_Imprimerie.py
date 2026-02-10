# ============================================
# Planning_Imprimerie.py — Planning Imprimerie (Streamlit + Plotly)
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

from pages.utils import (
    load_suivi_of,
    get_ofs_exclus,
    get_statut_dict,
    get_stock_supports,
    is_statut_actif,
    STATUT_ACTIF,
    SUPPORTS_L1,
)

# Adapter le chemin si besoin
BASE_PATH = Path(
    r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main"
)
OFS_IMP_FILE = BASE_PATH / "OFs_Imprimerie.xlsx"
CAL_FILE = BASE_PATH / "Calendrier 2026 imprimerie.xlsx"

LIGNE_NAME = "Imprimerie"

# ============================================
# 2) COULEURS PAR TYPE DE CAMPAGNE
# ============================================

def rgb_to_hex(r, g, b):
    return f"#{r:02X}{g:02X}{b:02X}"

CAMPAGNE_TYPE_COLOR = {
    "PRIMETEX": rgb_to_hex(204, 204, 0),
    "TEXLINE": rgb_to_hex(0, 128, 0),
    "TARABUS": rgb_to_hex(0, 128, 0),
    "BOOSTER": rgb_to_hex(255, 0, 0),
    "TMAX": rgb_to_hex(255, 0, 0),
    "START": rgb_to_hex(128, 128, 128),
    "SPORISOL": rgb_to_hex(128, 128, 128),
    "NERA": rgb_to_hex(0, 0, 139),
}

DEFAULT_TRAIT_COLOR = "#000000"
DEFAULT_BAR_COLOR = "#FFFFFF"
DEFAULT_TEXT_COLOR = "#000000"

INTRO_DUREE_H = 2.3

# ============================================
# 3) FONCTIONS UTILITAIRES
# ============================================

def get_campagne_type(campagne_str):
    if not isinstance(campagne_str, str):
        return None
    campagne_upper = campagne_str.upper()
    for key in CAMPAGNE_TYPE_COLOR:
        if key in campagne_upper:
            return key
    return None

def get_trait_color(campagne_str):
    camp_type = get_campagne_type(campagne_str)
    if camp_type:
        return CAMPAGNE_TYPE_COLOR[camp_type]
    return DEFAULT_TRAIT_COLOR

def is_double_trait(campagne_str):
    if not isinstance(campagne_str, str):
        return False
    up = campagne_str.upper()
    return "TARABUS" in up or "TMAX" in up

def parse_horaire(jour, horaire_str):
    """Retourne start,end à partir d'un '12h30-20h30' """
    debut_str, fin_str = horaire_str.split("-")
    debut_str = debut_str.replace("h", ":")
    fin_str = fin_str.replace("h", ":")
    start = datetime.combine(jour.date(), datetime.strptime(debut_str, "%H:%M").time())
    end = datetime.combine(jour.date(), datetime.strptime(fin_str, "%H:%M").time())
    if end <= start:
        end += timedelta(days=1)
    return start, end
def build_open_slots_from_now(cal_df, horizon_days=14, now=None):
    """Retourne une liste ordonnée de créneaux OUVERT."""
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


def schedule_ofs_from_slots(ofs_df, slots):
    """Planifie les OFs dans les créneaux 'slots'. Pas d'INTRO à l'imprimerie."""
    from datetime import timedelta

    planning_rows = []
    if not slots:
        return pd.DataFrame()

    slot_idx = 0
    current_slot_start = slots[0]["start"]
    current_slot_end = slots[0]["end"]

    def consume_time(duration_h):
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
        of_id = row["ID_PLAN"]
        coloris = row["Coloris"]
        support = row["Support"]
        campagne = row.get("Campagne", "")
        ml = row["Ml"]
        duree_h = float(row["Temps en h"]) if pd.notna(row["Temps en h"]) else 0.0

        trait_color = get_trait_color(campagne)
        double_trait = is_double_trait(campagne)

        of_segments, ok = consume_time(duree_h)

        for seg_data in of_segments:
            planning_rows.append(
                {
                    "Ofs": of_id,
                    "Segment": seg_data["segment"],
                    "Coloris": coloris,
                    "Support": support,
                    "Campagne": campagne,
                    "Ml": ml,
                    "start": seg_data["start"],
                    "end": seg_data["end"],
                    "duree_h": seg_data["duree_h"],
                    "is_intro": False,
                    "trait_color": trait_color,
                    "double_trait": double_trait,
                }
            )

        if not ok:
            return pd.DataFrame(planning_rows)

    return pd.DataFrame(planning_rows)


# ============================================
# 4) CHARGEMENT DONNÉES
# ============================================

@st.cache_data(ttl=7200)
def load_data():
    ofs_imp = pd.read_excel(OFS_IMP_FILE, engine="openpyxl")
    cal = pd.read_excel(CAL_FILE, engine="openpyxl")
    return ofs_imp, cal


# ============================================
# 5) FONCTION PRINCIPALE D'AFFICHAGE IMPRIMERIE
# ============================================

def show_planning_imprimerie():

    st.title("📅 Planning IMPRIMERIE")

    # --- Bouton retour menu ---
    st.markdown(""" """, unsafe_allow_html=True)
    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            st.rerun()

    # Charger les données
    ofs_imp_df, cal_df = load_data()
    suivi_df = load_suivi_of()

    # Filtrer OFs terminés (<200 ML, statuts 60/61/99)
    ofs_exclus = get_ofs_exclus(suivi_df, "IMPRIMERIE|L09")
    ofs_imp_df = ofs_imp_df[~ofs_imp_df["Ofs"].isin(ofs_exclus)]

    # Statuts / stocks
    statut_dict = get_statut_dict(suivi_df, "IMPRIMERIE|L09")
    stock_supports = get_stock_supports(suivi_df)

    # ID unique
    ofs_imp_df["ID_PLAN"] = (
        ofs_imp_df.index.astype(str) + "_" + ofs_imp_df["Ofs"].astype(str)
    )

    # Label lisible
    def make_display_label(row):
        c = row["Coloris"]
        c_short = c[:25] if isinstance(c, str) else ""
        return f"{row['Ofs']} - {c_short}"

    ofs_imp_df["DISPLAY_LABEL"] = ofs_imp_df.apply(make_display_label, axis=1)

    with st.expander("Données OFs Imprimerie"):
        st.dataframe(ofs_imp_df)

    # Sauvegarde ordre original
    if "ofs_list_original_IMP" not in st.session_state:
        st.session_state.ofs_list_original_IMP = list(ofs_imp_df["ID_PLAN"].values)

    id_to_label = dict(zip(ofs_imp_df["ID_PLAN"], ofs_imp_df["DISPLAY_LABEL"]))
    id_to_campagne = dict(zip(ofs_imp_df["ID_PLAN"], ofs_imp_df["Campagne"]))

    # --- Réorganisation manuelle ---
    st.markdown("### 🔁 Réorganisation manuelle (OF ou Campagne)")

    if "ordre_ofs_IMP" not in st.session_state:
        st.session_state.ordre_ofs_IMP = list(ofs_imp_df["ID_PLAN"].values)
        st.session_state.ordre_ofs_origine_IMP = (
            st.session_state.ordre_ofs_IMP.copy()
        )

    ordre = st.session_state.ordre_ofs_IMP

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
        key="mode_deplacement_IMP",
    )
# ------------------------------
    # MODE : un OF
    # ------------------------------
    if mode_deplacement == "Un OF":
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])

        with col1:
            of_to_move_label = st.selectbox(
                "OF à déplacer",
                ordre_labels if ordre_labels else ["(aucun)"],
                index=0,
                key="of_to_move_select_IMP",
            )
            of_to_move = label_to_id.get(of_to_move_label, of_to_move_label)

        with col2:
            cibles_labels = [
                id_to_label.get(x, x) for x in ordre if x != of_to_move
            ]
            cible_label = st.selectbox(
                "Le placer par rapport à",
                cibles_labels if cibles_labels else ["(aucune cible)"],
                key="cible_select_IMP",
            )
            cible = label_to_id.get(cible_label, cible_label)

        with col3:
            position = st.radio(
                "Position", ["Avant", "Après"], index=1, horizontal=True, key="position_radio_IMP"
            )

        with col4:
            appliquer = st.button(
                "Appliquer", use_container_width=True, key="btn_appliquer_of_IMP"
            )

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
                st.session_state.ordre_ofs_IMP = new_order
                st.success(f"OF déplacé {position.lower()} la cible sélectionnée.")

            else:
                st.warning("Cible introuvable dans l'ordre courant.")

    # ------------------------------
    # MODE : une campagne
    # ------------------------------
    else:
        campagnes_ordre = get_campagnes_in_order(ordre)
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])

        with col1:
            campagne_to_move = st.selectbox(
                "Campagne à déplacer",
                campagnes_ordre if campagnes_ordre else ["(aucune)"],
                index=0,
                key="campagne_to_move_select_IMP",
            )

        with col2:
            cibles_campagnes = [c for c in campagnes_ordre if c != campagne_to_move]
            cible_campagne = st.selectbox(
                "La placer par rapport à",
                cibles_campagnes if cibles_campagnes else ["(aucune cible)"],
                key="cible_campagne_select_IMP",
            )

        with col3:
            position_camp = st.radio(
                "Position",
                ["Avant", "Après"],
                index=1,
                horizontal=True,
                key="position_campagne_radio_IMP",
            )

        with col4:
            appliquer_camp = st.button(
                "Appliquer", use_container_width=True, key="btn_appliquer_camp_IMP"
            )

        if (
            appliquer_camp
            and cibles_campagnes
            and campagne_to_move != "(aucune)"
            and cible_campagne != "(aucune cible)"
        ):
            ids_campagne_to_move = [
                id_plan
                for id_plan in ordre
                if id_to_campagne.get(id_plan) == campagne_to_move
            ]

            new_order = [
                id_plan
                for id_plan in ordre
                if id_to_campagne.get(id_plan) != campagne_to_move
            ]

            ids_cible = [
                id_plan
                for id_plan in new_order
                if id_to_campagne.get(id_plan) == cible_campagne
            ]

            if ids_cible:
                if position_camp == "Avant":
                    insert_idx = new_order.index(ids_cible[0])
                else:
                    insert_idx = new_order.index(ids_cible[-1]) + 1

                for i, id_plan in enumerate(ids_campagne_to_move):
                    new_order.insert(insert_idx + i, id_plan)

                st.session_state.ordre_ofs_IMP = new_order
                st.success(
                    f"Campagne '{campagne_to_move}' déplacée {position_camp.lower()} '{cible_campagne}'."
                )
            else:
                st.warning("Campagne cible introuvable.")

    # --------------------------------------------------------
    # Boutons de sauvegarde
    # --------------------------------------------------------
    col_reset, col_backup, col_save = st.columns([1, 2, 3])

    with col_reset:
        if st.button(
            "Réinitialiser l'ordre (Excel)", use_container_width=True, key="reset_IMP"
        ):
            st.session_state.ordre_ofs_IMP = (
                st.session_state.ordre_ofs_origine_IMP.copy()
            )
            st.info("Ordre réinitialisé à l'ordre Excel d'origine.")

    with col_backup:
        backup_before_save = st.checkbox(
            "Créer une sauvegarde avant d'écrire", value=True, key="backup_IMP"
        )

    with col_save:
        if st.button(
            "📥 Valider l'ordre et mettre à jour OFs_Imprimerie.xlsx",
            use_container_width=True,
            key="save_IMP",
        ):
            try:
                ordre_final = st.session_state.ordre_ofs_IMP

                if sorted(ordre_final) != sorted(
                    st.session_state.ofs_list_original_IMP
                ):
                    st.error(
                        "❌ L'ordre ne correspond pas exactement aux OFs d'origine."
                    )
                    st.stop()

                df_excel = pd.read_excel(OFS_IMP_FILE, engine="openpyxl")
                df_excel["ID_PLAN"] = (
                    df_excel.index.astype(str) + "_" + df_excel["Ofs"].astype(str)
                )

                if backup_before_save:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_path = (
                        BASE_PATH
                        / f"OFs_Imprimerie_backup_{timestamp}.xlsx"
                    )
                    df_excel.drop(columns=["ID_PLAN"]).to_excel(
                        backup_path, index=False, engine="openpyxl"
                    )

                df_excel = df_excel.set_index("ID_PLAN")
                df_excel = df_excel.loc[ordre_final].reset_index(drop=True)
                df_excel.to_excel(OFS_IMP_FILE, index=False, engine="openpyxl")

                st.success(
                    "✅ L'ordre a été appliqué et OFs_Imprimerie.xlsx a été mis à jour !"
                )

                st.session_state.ofs_list_original_IMP = ordre_final.copy()
                st.session_state.ordre_ofs_origine_IMP = ordre_final.copy()

            except Exception as e:
                st.error(f"❌ Erreur lors de la mise à jour du fichier : {e}")

    # --- ORDRE COURANT ---
    with st.expander("Ordre courant des OFs (appliqué au planning)"):
        ordre_lisible = [
            id_to_label.get(x, x) for x in st.session_state.ordre_ofs_IMP
        ]
        st.write(ordre_lisible)

    # -------------------------
    # Créneaux ouverts
    # -------------------------
    now = datetime.now()
    horizon_days = 14
    horizon_end = now + timedelta(days=horizon_days)

    st.markdown(f"**🕒 Maintenant : {now.strftime('%d/%m/%Y %H:%M')}**")

    slots = build_open_slots_from_now(cal_df, horizon_days=horizon_days, now=now)

    if not slots:
        st.error("Aucun créneau OUVERT trouvé dans les 14 prochains jours.")
        st.stop()

    # --------------------------------------------
    # Sécurisation ID_PLAN IMP (anti-KeyError)
    # --------------------------------------------
    ofs_imp_df["ID_PLAN"] = (
        ofs_imp_df.index.astype(str) + "_" + ofs_imp_df["Ofs"].astype(str)
    )

    if "ordre_ofs_IMP" not in st.session_state:
        st.session_state.ordre_ofs_IMP = list(ofs_imp_df["ID_PLAN"].values)

    current_ids = set(ofs_imp_df["ID_PLAN"])
    saved_order = st.session_state.ordre_ofs_IMP
    ids_missing = [x for x in saved_order if x not in current_ids]

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

    if ids_missing:
        try:
            repaired = rebuild_order_from_ofs(
                saved_order,
                list(ofs_imp_df["ID_PLAN"]),
                ofs_imp_df["Ofs"],
            )

            if len(repaired) == len(ofs_imp_df) and len(set(repaired)) == len(
                repaired
            ):
                st.session_state.ordre_ofs_IMP = repaired
                st.success("Ordre Imprimerie réparé automatiquement.")
            else:
                raise Exception("Reconstruction partielle")

        except Exception:
            st.warning(
                "⚠️ Ordre Imprimerie incompatible — réinitialisation totale."
            )
            st.session_state.ordre_ofs_IMP = list(ofs_imp_df["ID_PLAN"])

    ordre_final_imp = [
        x for x in st.session_state.ordre_ofs_IMP if x in current_ids
    ]
    st.session_state.ordre_ofs_IMP = ordre_final_imp

    # --------------------------------------------
    # PLANIFICATION DES OFs
    # --------------------------------------------
    ofs_imp_work = ofs_imp_df.set_index("ID_PLAN")
    ofs_imp_work = ofs_imp_work.loc[
        st.session_state.ordre_ofs_IMP
    ].reset_index()
    planning_df = schedule_ofs_from_slots(ofs_imp_work, slots)

    if planning_df.empty:
        st.warning("Planning Imprimerie vide : pas assez de créneaux.")
        st.stop()

    # ------------------------------------------------
    # LABEL IMPRIMERIE
    # ------------------------------------------------
    def build_label(row):
        if row.get("is_intro", False):
            return f"<b>INTRO</b><br>{row['duree_h']:.1f} h"

        of_display = row["Ofs"]
        if isinstance(of_display, str) and "_" in of_display:
            of_display = of_display.split("_")[-1]

        coloris = row.get("Coloris", "")
        if isinstance(coloris, str) and "-" in coloris:
            code, desc = coloris.split("-", 1)
        else:
            code, desc = coloris, ""

        return (
            f"<b>{code}</b><br>"
            f"{desc[:15]}<br>"
            f"{row['Support']} - {int(row['Ml']) if pd.notna(row['Ml']) else ''} ML<br>"
            f"OF {of_display} - {row['duree_h']:.1f} h"
        )

    # ------------------------------------------------
    # GANTT IMPRIMERIE — GO.FIGURE (avec traits)
    # ------------------------------------------------
    plot_df = planning_df.copy()
    plot_df["Ligne"] = LIGNE_NAME
    plot_df["Label"] = plot_df.apply(build_label, axis=1)

    st.markdown("### 🔎 Fenêtre d'affichage")

    largeur_heures = st.selectbox(
        "Largeur de vue",
        [12, 24, 36, 48, 72],
        index=1,
        format_func=lambda x: f"{x}h ({x//24}j {x%24}h)" if x >= 24 else f"{x}h",
        key="largeur_vue_IMP",
    )

    total_hours = (horizon_end - now).total_seconds() / 3600
    max_offset = max(0, total_hours - largeur_heures)

    if "offset_heures_IMP" not in st.session_state:
        st.session_state.offset_heures_IMP = 0.0

    offset_heures = st.session_state.offset_heures_IMP

    view_start = now + timedelta(hours=offset_heures)
    view_end = view_start + timedelta(hours=largeur_heures)

    if view_end > horizon_end:
        view_end = horizon_end

    fig = go.Figure()

    for _, row in plot_df.iterrows():
        start = row["start"]
        end = row["end"]
        label = row["Label"]
        is_intro = row.get("is_intro", False)
        trait_color = row.get("trait_color", DEFAULT_TRAIT_COLOR)
        double_trait = row.get("double_trait", False)

        bar_color = "#F0F0F0" if is_intro else "#FFFFFF"

        # BARRE PRINCIPALE
        fig.add_trace(
            go.Bar(
                x=[(end - start).total_seconds() * 1000],
                y=[LIGNE_NAME],
                base=[start],
                orientation="h",
                marker=dict(
                    color=bar_color, line=dict(color="#000000", width=1)
                ),
                text=label,
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(size=9, color="#000000", family="Arial"),
                hovertemplate=f"<b>Détail :</b><br>{label}<extra></extra>",
                showlegend=False,
            )
        )

        # PASTILLE VERTE STATUTS
        if not is_intro:
            of_num = row["Ofs"]
            if isinstance(of_num, str) and "_" in of_num:
                try:
                    of_num = int(of_num.split("_")[-1])
                except:
                    of_num = None

            if of_num and is_statut_actif(of_num, statut_dict):
                mid_time = start + (end - start) / 2
                fig.add_annotation(
                    x=mid_time,
                    y="Imprimerie",
                    yshift=-35,
                    text="🟢",
                    showarrow=False,
                    font=dict(size=12),
                    xref="x",
                    yref="y",
                )

        # TRAITS COLORES (simple ou double)
        if not is_intro:
            if double_trait:
                fig.add_shape(
                    type="line",
                    x0=start,
                    x1=end,
                    y0=0.33,
                    y1=0.33,
                    yref="paper",
                    line=dict(color=trait_color, width=3),
                )
                fig.add_shape(
                    type="line",
                    x0=start,
                    x1=end,
                    y0=0.67,
                    y1=0.67,
                    yref="paper",
                    line=dict(color=trait_color, width=3),
                )
            else:
                fig.add_shape(
                    type="line",
                    x0=start,
                    x1=end,
                    y0=0.5,
                    y1=0.5,
                    yref="paper",
                    line=dict(color=trait_color, width=3),
                )

    # ARRETS (ROUGE)
    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])
    mask_cal = (cal["Jour"].dt.date >= now.date()) & (
        cal["Jour"].dt.date <= horizon_end.date()
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
                        text="ARRET",
                        showarrow=False,
                        font=dict(color="white", size=12, family="Arial Black"),
                    )

    # AXES
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

    # JOURS EN HAUT
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

    while current_day <= end_day:
        day_middle = datetime.combine(
            current_day, datetime.min.time()
        ) + timedelta(hours=12)

        jour_semaine = JOURS_FR[current_day.weekday()]
        day_label = f"{jour_semaine} {current_day.strftime('%d/%m/%y')}"

        fig.add_annotation(
            x=day_middle,
            y=1.08,
            xref="x",
            yref="paper",
            text=f"<b>{day_label}</b>",
            showarrow=False,
            font=dict(color="white", size=11, family="Arial"),
            xanchor="center",
        )

        current_day += timedelta(days=1)

    fig.add_vline(
        x=now, line_color="white", line_width=2, line_dash="dot"
    )

    fig.update_layout(
        height=350,
        margin=dict(l=10, r=10, t=60, b=40),
        plot_bgcolor="#444444",
        paper_bgcolor="#444444",
        showlegend=False,
        dragmode=False,
        font=dict(color="white"),
        barmode="overlay",
    )

    st.plotly_chart(fig, use_container_width=True)

    # SLIDER DE DÉFILEMENT
    new_offset = st.slider(
        "⬅️ Faire défiler le planning ➡️",
        min_value=0.0,
        max_value=max(1.0, max_offset),
        value=float(st.session_state.offset_heures_IMP),
        step=1.0,
        format="%.0fh",
        key="slider_defilement_IMP",
    )

    if new_offset != st.session_state.offset_heures_IMP:
        st.session_state.offset_heures_IMP = new_offset
        st.rerun()

    view_start_display = now + timedelta(hours=new_offset)
    view_end_display = view_start_display + timedelta(hours=largeur_heures)

    st.caption(
        f"📍 Vue : {view_start_display.strftime('%d/%m %Hh')} → "
        f"{view_end_display.strftime('%d/%m %Hh')}"
    )

    # STOCK SUPPORTS L1
    st.markdown("### 📦 Stock Supports L1")
    stock_cols = st.columns(4)

    for i, support in enumerate(SUPPORTS_L1):
        stock_val = stock_supports.get(support)
        with stock_cols[i % 4]:
            if stock_val is not None:
                color = "#FF6B6B" if stock_val < 500 else "#4CAF50"
                st.markdown(
                    f"<div style='background:{color}; color:white; padding:8px; "
                    f"border-radius:5px; text-align:center; margin:2px;'>"
                    f"<b>{support}</b><br>{stock_val:,.0f} ML</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"<div style='background:#888; color:white; padding:8px; "
                    f"border-radius:5px; text-align:center; margin:2px;'>"
                    f"<b>{support}</b><br>N/A</div>",
                    unsafe_allow_html=True,
                )

    # DEBUG
    with st.expander("Segments planifiés Imprimerie (debug)"):
        df_debug = planning_df.copy()
        df_debug["Ml"] = pd.to_numeric(df_debug["Ml"], errors="coerce")
        st.dataframe(df_debug)