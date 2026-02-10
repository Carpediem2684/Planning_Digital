# ============================================
# Planning_Visitage.py — Planning Visitage (Streamlit + Plotly)
# Version corrigée avec étiquettes et pastilles - 2026-02-07
# ============================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import plotly.graph_objects as go
from pages.utils import load_suivi_of, get_ofs_exclus, get_statut_dict, is_statut_actif, STATUT_ACTIF

# -------- CONFIG --------
BASE_PATH = Path(r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main")
OFS_VIS_FILE = BASE_PATH / "OFs_Visitage.xlsx"
CAL_FILE = BASE_PATH / "Calendrier 2026.xlsx"

# Valeurs par défaut pour calcul durée
DEFAULT_ML_MIN = 15
DEFAULT_TRG = 0.8
OFFSET_TEMPS = 0.75

# -------- COULEURS --------
def rgb_to_hex(r, g, b):
    return f"#{r:02X}{g:02X}{b:02X}"

COULEUR_TOP_MAP = {
    "NERA": rgb_to_hex(0, 0, 139),
    "START": rgb_to_hex(211, 211, 211),
    "TARASTEP": rgb_to_hex(255, 255, 224),
    "PRIMETEX": rgb_to_hex(204, 204, 0),
    "GRIPX": rgb_to_hex(255, 182, 193),
    "TARABUS": rgb_to_hex(80, 200, 120),
    "TMAX": rgb_to_hex(101, 67, 33),
    "BOOSTER": rgb_to_hex(250, 128, 114),
    "LOFTEX": rgb_to_hex(255, 140, 0),
    "TEXLINE": rgb_to_hex(0, 128, 0),
    "SPORISOL": rgb_to_hex(128, 128, 128),
    "FUSION": rgb_to_hex(144, 238, 144),
}

LAISE_BOTTOM_MAP = {
    4: rgb_to_hex(119, 221, 119),
    3: rgb_to_hex(173, 216, 230),
    2: rgb_to_hex(255, 182, 193),
}

def get_top_color(campagne):
    """Retourne la couleur du haut basée sur la campagne."""
    if not campagne:
        return "#CCCCCC"
    c = str(campagne).upper()
    for key, color in COULEUR_TOP_MAP.items():
        if key in c:
            return color
    return "#CCCCCC"

def get_bottom_color(laise):
    try:
        return LAISE_BOTTOM_MAP.get(int(laise), "#77DD77")
    except:
        return "#77DD77"

def get_text_color(bg_color):
    darks = [rgb_to_hex(0, 0, 139), rgb_to_hex(101, 67, 33), rgb_to_hex(0, 128, 0)]
    return "#FFFFFF" if bg_color in darks else "#000000"

def calculate_duree(ml, ml_min=DEFAULT_ML_MIN, trg=DEFAULT_TRG):
    """Calcule la durée en heures."""
    if not ml or ml <= 0:
        return 0.5
    return (ml / ml_min / 60 * trg) + OFFSET_TEMPS

# -------- SLOTS --------
def parse_horaire(jour, h):
    d, f = h.split("-")
    d = d.replace("h", ":")
    f = f.replace("h", ":")
    start = datetime.combine(jour.date(), datetime.strptime(d, "%H:%M").time())
    end = datetime.combine(jour.date(), datetime.strptime(f, "%H:%M").time())
    if end <= start:
        end += timedelta(days=1)
    return start, end

def build_all_open_slots(cal_df, from_date=None):
    """Construit tous les créneaux OUVERT du calendrier."""
    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])
    
    if from_date is None:
        from_date = datetime.now()
    
    slots = []
    for _, row in cal.iterrows():
        day = row["Jour"]
        for i in [1, 2, 3]:
            e = row.get(f"Etat_{i}")
            h = row.get(f"Horaire_{i}")
            if e == "OUVERT" and isinstance(h, str) and "-" in h:
                start, end = parse_horaire(day, h)
                if end > from_date:
                    if start < from_date:
                        start = from_date
                    slots.append({"start": start, "end": end})
    
    return sorted(slots, key=lambda x: x["start"])

# -------- PLANIFICATION --------
def schedule_ofs(ofs_df, slots):
    """Planifie les OFs dans les créneaux."""
    if not slots or ofs_df.empty:
        return pd.DataFrame()

    planning = []
    slot_idx = 0
    cur_s = slots[0]["start"]
    cur_e = slots[0]["end"]

    def consume(hours):
        nonlocal slot_idx, cur_s, cur_e
        segs = []
        remain = timedelta(hours=hours)

        while remain > timedelta(0):
            if slot_idx >= len(slots):
                return segs, False

            dispo = cur_e - cur_s
            if dispo <= timedelta(0):
                slot_idx += 1
                if slot_idx >= len(slots):
                    return segs, False
                cur_s = slots[slot_idx]["start"]
                cur_e = slots[slot_idx]["end"]
                continue

            use = min(dispo, remain)
            segs.append({"start": cur_s, "end": cur_s + use, "duree_h": use.total_seconds() / 3600})
            remain -= use
            cur_s = cur_s + use

        return segs, True

    for _, row in ofs_df.iterrows():
        of_id = row.get("ID_PLAN", row.get("Ofs", ""))
        coloris = row.get("Coloris", "")
        laise = row.get("Laise", 4)
        campagne = row.get("Campagne", "")
        ml = float(row.get("Ml", 0)) if pd.notna(row.get("Ml")) else 0
        
        # Calculer durée
        duree = row.get("Temps en h")
        if pd.isna(duree) or duree is None or duree <= 0:
            duree = calculate_duree(ml)
        else:
            duree = float(duree)

        top_color = get_top_color(campagne)
        bottom_color = get_bottom_color(laise)
        text_color = get_text_color(top_color)

        segs, ok = consume(duree)
        for s in segs:
            planning.append({
                "Ofs": of_id,
                "Coloris": coloris,
                "Laise": laise,
                "Campagne": campagne,
                "Ml": ml,
                "start": s["start"],
                "end": s["end"],
                "duree_h": s["duree_h"],
                "top_color": top_color,
                "bottom_color": bottom_color,
                "text_color": text_color,
            })

        if not ok:
            break

    return pd.DataFrame(planning)

# -------- DATA --------
@st.cache_data(ttl=7200)
def load_data():
    ofs = pd.read_excel(OFS_VIS_FILE, sheet_name="Feuil1", engine="openpyxl")
    cal = pd.read_excel(CAL_FILE, engine="openpyxl")
    return ofs, cal

def show_planning_visitage():
    st.title("📅 Planning VISITAGE")

    # --- Bouton Retour Menu ---
    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            st.rerun()

    # --- CHARGEMENT ---
    try:
        ofs_df, cal_df = load_data()
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        return

    # Charger SUIVI_OF pour statuts
    suivi_df = load_suivi_of()
    
    # Filtrer les OFs terminés (60, 61, 99) et < 200 ML
    ofs_exclus = get_ofs_exclus(suivi_df, "VISITAGE|L10")
    ofs_df = ofs_df[~ofs_df["Ofs"].isin(ofs_exclus)]
    
    # Dict des statuts pour pastilles vertes
    statut_dict = get_statut_dict(suivi_df, "VISITAGE|L10")

    if ofs_df.empty:
        st.warning("Aucun OF trouvé.")
        return

    ofs_df["ID_PLAN"] = ofs_df.index.astype(str) + "_" + ofs_df["Ofs"].astype(str)

    with st.expander("📋 Données OFs Visitage"):
        st.dataframe(ofs_df)

    # ---- Créneaux ouverts ----
    now = datetime.now()
    slots = build_all_open_slots(cal_df, now)

    if not slots:
        st.error("Aucun créneau OUVERT trouvé dans le calendrier.")
        return

    first_slot = slots[0]
    st.info(f"⏰ Premier créneau : **{first_slot['start'].strftime('%d/%m à %Hh%M')}**")

    # ---- Planification ----
    planning = schedule_ofs(ofs_df, slots)
    
    if planning.empty:
        st.warning("Aucun OF planifié.")
        return

    st.success(f"✅ **{planning['Ofs'].nunique()}** OFs planifiés")

    # ---- Paramètres d'affichage ----
    display_start = first_slot["start"] - timedelta(hours=1)
    horizon_end = display_start + timedelta(days=7)

    largeur = st.selectbox("Largeur de vue :", [12, 24, 36, 48, 72], index=1, key="vue_vis")
    max_offset = max(0, int((horizon_end - display_start).total_seconds() / 3600) - largeur)

    if "offset_vis" not in st.session_state:
        st.session_state.offset_vis = 0

    offset = st.slider("Défilement", 0, max_offset, st.session_state.offset_vis, 1)
    st.session_state.offset_vis = offset

    view_start = display_start + timedelta(hours=offset)
    view_end = view_start + timedelta(hours=largeur)

    # ---- GANTT bicolore ----
    fig = go.Figure()

    visible = planning[(planning["end"] >= view_start) & (planning["start"] <= view_end)]

    for _, r in visible.iterrows():
        start = r["start"]
        end = r["end"]
        dur_ms = (end - start).total_seconds() * 1000
        of_num = str(r['Ofs']).split('_')[-1] if '_' in str(r['Ofs']) else str(r['Ofs'])
        
        # Étiquette du haut (coloris + ML + OF + durée)
        label_top = f"<b>{str(r['Coloris'])[:15]}</b><br>{int(r['Ml'])} ML<br>OF {of_num}<br>{r['duree_h']:.1f}h"
        
        fig.add_trace(go.Bar(
            x=[dur_ms], y=["Visitage"], base=[start], orientation="h",
            marker=dict(color=r["top_color"], line=dict(color="#000", width=1)),
            text=label_top, textposition="inside", insidetextanchor="middle",
            textfont=dict(size=9, color=r["text_color"]),
            hovertemplate=f"{r['Coloris']}<br>{int(r['Ml'])} ML<br>OF {of_num}<extra></extra>",
            showlegend=False,
        ))

        # Barre du bas (laise)
        fig.add_trace(go.Bar(
            x=[dur_ms], y=["Laise"], base=[start], orientation="h",
            marker=dict(color=r["bottom_color"], line=dict(color="#000", width=1)),
            text=f"<b>L{r['Laise']}</b>", textposition="inside",
            textfont=dict(size=12, color="#000"),
            hovertemplate=f"Laise {r['Laise']}<extra></extra>",
            showlegend=False,
        ))

    # ---- PASTILLES VERTES pour statuts 30, 40, 50 ----
    for _, r in visible.iterrows():
        of_num = r["Ofs"]
        if isinstance(of_num, str) and "_" in of_num:
            try:
                of_num = int(of_num.split("_")[-1])
            except:
                continue
        
        if is_statut_actif(of_num, statut_dict):
            mid_time = r["start"] + (r["end"] - r["start"]) / 2
            fig.add_annotation(
                x=mid_time, y="Visitage", yshift=-40,
                text="🟢", showarrow=False, font=dict(size=12),
                xref="x", yref="y",
            )

    # ---- Jours ----
    JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    current_day = view_start.date()
    while current_day <= view_end.date():
        day_middle = datetime.combine(current_day, datetime.min.time()) + timedelta(hours=12)
        if view_start <= day_middle <= view_end:
            fig.add_annotation(
                x=day_middle, y=1.08, xref="x", yref="paper",
                text=f"<b>{JOURS_FR[current_day.weekday()]} {current_day.strftime('%d/%m')}</b>",
                showarrow=False, font=dict(color="white", size=11),
            )
        current_day += timedelta(days=1)

    # ---- Ligne maintenant ----
    if view_start <= now <= view_end:
        fig.add_vline(x=now, line_color="yellow", line_dash="dot", line_width=2)

    fig.update_xaxes(type="date", range=[view_start, view_end], tickformat="%Hh", dtick=3600000)
    fig.update_yaxes(categoryorder='array', categoryarray=['Laise', 'Visitage'], tickfont=dict(color="white", size=12))
    fig.update_layout(
        height=300, margin=dict(l=80, r=20, t=50, b=40),
        plot_bgcolor="#444", paper_bgcolor="#444", font=dict(color="white"),
        barmode="overlay", showlegend=False,
    )

    st.plotly_chart(fig, use_container_width=True)

    # ---- Légendes ----
    st.markdown("### 🎨 Légendes")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Type de produit**")
        for t, cou in COULEUR_TOP_MAP.items():
            tc = "#FFF" if get_text_color(cou) == "#FFFFFF" else "#000"
            st.markdown(f"<div style='background:{cou}; color:{tc}; padding:4px 8px; display:inline-block; margin:2px; border-radius:4px; font-size:12px;'>{t}</div>", unsafe_allow_html=True)

    with col2:
        st.markdown("**Laise**")
        for la, cou in LAISE_BOTTOM_MAP.items():
            st.markdown(f"<div style='background:{cou}; color:#000; padding:4px 8px; display:inline-block; margin:2px; border-radius:4px; font-size:12px;'>Laise {la}M</div>", unsafe_allow_html=True)

    # ---- Résumé ----
    st.markdown("### 📊 Résumé")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("OFs planifiés", planning["Ofs"].nunique())
    with c2:
        st.metric("ML total", f"{planning.drop_duplicates('Ofs')['Ml'].sum():,.0f}")
    with c3:
        st.metric("Heures totales", f"{planning['duree_h'].sum():.1f}h")
