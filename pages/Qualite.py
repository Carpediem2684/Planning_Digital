# -*- coding: utf-8 -*-
"""
Page Streamlit : Qualite.py (v7 FINAL)

- KPI 100% FIXES (jamais déformés par Streamlit)
- Gauges avec valeurs SOUS la jauge : “95.22 % — Réalisé”
- Aucun texte au-dessus des jauges
- Patch dates/jours
- Colonnes Numéro OF + Dessin Coloris
- Thème dark modern
"""

import io
import re
import math
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime, date


def _safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.rerun()


# ----------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------
# st.set_page_config dans app.py

PLOTLY_TEMPLATE = "plotly_dark"

# ----------------------------------------------------------
# CSS
# ----------------------------------------------------------
CUSTOM_CSS = """
<style>

:root {
  --card-bg: #1f2630;
  --card-border: #2b3440;
}

.block-container {padding-top: 1.2rem;}

/* KPI FIXES (hauteur forcée + cadre stable) */
.kpi-box {
    height: 135px;
    width: 100%;
    background-color: #1f2630;
    border: 1px solid #2b3440;
    border-radius: 14px;
    padding: 20px 20px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    position: relative;
    overflow: hidden;
}

.kpi-title {
    font-size: 15px;
    color: #cfd8e3;
    margin-bottom: 6px;
    font-weight: 600;
}

.kpi-value {
    font-size: 31px;
    font-weight: 800;
    color: white;
    margin-top: -4px;
}

.hr {
    border-top: 1px solid var(--card-border);
    margin: 1rem 0 1.3rem 0;
}

</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------
# UTILITAIRES
# ----------------------------------------------------------
def _normalize(col: str) -> str:
    if col is None: return ""
    col = col.lower()
    repl = {
        "é":"e","è":"e","ê":"e","ë":"e",
        "à":"a","â":"a",
        "î":"i","ï":"i",
        "ô":"o",
        "ù":"u","û":"u","ü":"u",
        "ç":"c"
    }
    for k,v in repl.items(): col = col.replace(k,v)
    col = re.sub(r"[^a-z0-9]","", col)
    return col

def find_col(df, candidates):
    norm = {_normalize(c): c for c in df.columns}
    for cand in candidates:
        k = _normalize(cand)
        if k in norm: return norm[k]
    for cand in candidates:
        k = _normalize(cand)
        for key, real in norm.items():
            if k in key: return real
    raise KeyError(f"Colonne introuvable : {candidates}")

@st.cache_data(show_spinner=True)
def load_data(path="Qualite.xlsx"):
    df = pd.read_excel(path, engine="openpyxl").copy()

    col_date = find_col(df, ["Date début OF","datedebutof"])
    col_ligne = find_col(df, ["Libelle ligne","ligne"])
    col_im = find_col(df, ["Quantité mvt IM","Quantite mvt IM","IM"])
    col_ic = find_col(df, ["Quantité IC","Quantite IC","IC"])

    # optionnel
    try: col_of = find_col(df, ["Numéro OF","Numero OF"])
    except: col_of = None
    try: col_col = find_col(df, ["Dessin coloris"])
    except: col_col = None
    try: col_ecart = find_col(df, ["Rebuts en écart vs budget"])
    except: col_ecart = None

    df["DateDebutOF"] = pd.to_datetime(df[col_date], errors="coerce", dayfirst=True)
    df[col_im] = pd.to_numeric(df[col_im], errors="coerce")
    df[col_ic] = pd.to_numeric(df[col_ic], errors="coerce")

    df["PctRebut"] = np.where(df[col_im] > 0, (df[col_im] - df[col_ic]) / df[col_im] * 100, np.nan)
    df["PctRealise"] = np.where(df[col_im] > 0, df[col_ic] / df[col_im] * 100, np.nan)

    df.loc[df["PctRebut"] < 0, "PctRebut"] = 0
    df.loc[df["PctRealise"] < 0, "PctRealise"] = 0

    df["Jour"] = df["DateDebutOF"].dt.date

    df.rename(columns={col_ligne:"Ligne", col_im:"QteIM", col_ic:"QteIC"}, inplace=True)
    if col_of: df.rename(columns={col_of:"Numéro OF"}, inplace=True)
    if col_col: df.rename(columns={col_col:"Dessin coloris"}, inplace=True)
    if col_ecart: df.rename(columns={col_ecart:"RebutsEcartBudget"}, inplace=True)

    return df


# ----------------------------------------------------------
# GAUGE (sans titre + valeur SOUS la jauge)
# ----------------------------------------------------------
def gauge(value):
    value = 0 if value is None or (isinstance(value,float) and math.isnan(value)) else float(value)

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            number={"suffix":" %", "font":{"size":24,"color":"white"}},
            title={"text":""},
            gauge={
                "axis":{"range":[0,100], "tickcolor":"#A7B1C2"},
                "bar":{"color":"#19e5a6"},
                "borderwidth":1,
                "steps":[
                    {"range":[0,70],"color":"#3a3f4b"},
                    {"range":[70,90],"color":"#2b3340"},
                    {"range":[90,100],"color":"#1f2630"},
                ],
                "threshold":{
                    "line":{"color":"#19e5a6","width":2},
                    "thickness":0.8,
                    "value":value
                }
            }
        )
    )

    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        height=210,
        margin=dict(l=10,r=10,t=35,b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    return fig


# ----------------------------------------------------------
# MAIN PAGE
# ----------------------------------------------------------
def show_qualite():

    st.title("📊 Tableau de bord – Indicateurs Qualité")
    
    # --- Bouton retour menu (header, à droite)
    col_left, col_spacer, col_right = st.columns([1, 6, 1])
    with col_right:
        if st.button("⬅️ Retour menu", use_container_width=True):
            st.session_state["page"] = "menu"
            st.rerun()

    df = load_data()
    df["Jour"] = pd.to_datetime(df["Jour"], errors="coerce").dt.date
    df = df.dropna(subset=["Jour"])

    # --------------------------
    # SIDEBAR
    # --------------------------
    with st.sidebar:
        st.header("🎛️ Filtres")

        min_date = df["Jour"].min()
        max_date = df["Jour"].max()

        date_start, date_end = st.date_input(
            "Plage de dates", value=(min_date,max_date),
            min_value=min_date, max_value=max_date
        )

        lignes = sorted(df["Ligne"].dropna().unique())
        selected_lignes = st.multiselect("Lignes", lignes, default=lignes)

        st.subheader("Seuils d’alerte")
        seuil_rebut = st.slider("Seuil % Rebut",0.0,30.0,5.0,0.5)
        seuil_realise = st.slider("Seuil % Réalisé",70.0,100.0,95.0,0.5)

    # --------------------------
    # FILTER DF
    # --------------------------
    dff = df[
        (df["Jour"] >= date_start) &
        (df["Jour"] <= date_end) &
        (df["Ligne"].isin(selected_lignes))
    ].copy()

    # --------------------------
    # KPIs FIXES
    # --------------------------
    st.markdown("### 📌 KPIs")

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-title'>Total ML produits</div>
                <div class='kpi-value'>{dff['QteIM'].sum():,.0f}</div>
            </div>
            """, unsafe_allow_html=True
        )
    with c2:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-title'>% Réalisé moyen</div>
                <div class='kpi-value'>{dff['PctRealise'].mean():.2f} %</div>
            </div>
            """, unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-title'>% Rebut moyen</div>
                <div class='kpi-value'>{dff['PctRebut'].mean():.2f} %</div>
            </div>
            """, unsafe_allow_html=True
        )
    with c4:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-title'>Nombre d’OF</div>
                <div class='kpi-value'>{len(dff)}</div>
            </div>
            """, unsafe_allow_html=True
        )

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --------------------------
    # GAUGES
    # --------------------------
    st.subheader("🔧 Jauges par ligne")

    agg = dff.groupby("Ligne").agg(
        Realise=("PctRealise","mean"),
        Rebut=("PctRebut","mean"),
        ML=("QteIM","sum")
    ).reset_index().sort_values("ML",ascending=False)

    for start in range(0,min(6,len(agg)),3):
        cols = st.columns(3)
        for i in range(3):
            idx = start+i
            if idx >= len(agg): break
            row = agg.iloc[idx]

            with cols[i]:
                st.markdown(f"### 🔧 **{row['Ligne']}**")

                fig1 = gauge(row["Realise"])
                st.plotly_chart(fig1, use_container_width=True)
                st.markdown(
                    f"<div style='text-align:center;color:black;font-size:16px;'>{row['Realise']:.2f} % — Réalisé</div>",
                    unsafe_allow_html=True
                )

                fig2 = gauge(row["Rebut"])
                st.plotly_chart(fig2, use_container_width=True)
                st.markdown(
                    f"<div style='text-align:center;color:black;font-size:16px;'>{row['Rebut']:.2f} % — Rebut</div>",
                    unsafe_allow_html=True
                )

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --------------------------
    # GRAPH : Production / Réalisé / Donut
    # --------------------------
    st.subheader("📈 Production par jour (ML)")
    prod = dff.groupby("Jour")["QteIM"].sum().reset_index()
    st.plotly_chart(
        px.bar(prod, x="Jour", y="QteIM", text_auto=True, template=PLOTLY_TEMPLATE),
        use_container_width=True
    )

    st.subheader("📉 Tendance % Réalisé")
    prod2 = dff.groupby("Jour")["PctRealise"].mean().reset_index()
    fig2 = px.line(prod2, x="Jour", y="PctRealise", markers=True, template=PLOTLY_TEMPLATE)
    fig2.update_yaxes(range=[0,100])
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("🍩 Distribution par ligne")
    dist = dff.groupby("Ligne")["QteIM"].sum().reset_index()
    st.plotly_chart(
        px.pie(dist, values="QteIM", names="Ligne", hole=0.55, template=PLOTLY_TEMPLATE),
        use_container_width=True
    )

    if "RebutsEcartBudget" in dff.columns:
        st.subheader("📊 Histogramme – Écarts rebuts vs budget")
        st.plotly_chart(
            px.histogram(dff, x="RebutsEcartBudget", nbins=40, template=PLOTLY_TEMPLATE),
            use_container_width=True
        )

    # --------------------------
    # TABLEAU ALERTES
    # --------------------------
    st.subheader("🚨 OF en alerte qualité")

    crit = dff[(dff["PctRebut"] >= seuil_rebut) | (dff["PctRealise"] <= seuil_realise)]

    cols_show = [c for c in [
        "Numéro OF","Dessin coloris",
        "DateDebutOF","Ligne",
        "QteIM","QteIC",
        "PctRealise","PctRebut",
        "RdmtCalc","RdtBudget"
    ] if c in crit.columns]

    st.dataframe(
        crit.sort_values("PctRebut",ascending=False)[cols_show],
        use_container_width=True
    )


# ----------------------------------------------------------
# RUN DIRECT
# ----------------------------------------------------------
if __name__ == "__main__":
    show_qualite()