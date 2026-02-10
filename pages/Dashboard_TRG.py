import os
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# -------------------------------------------------
#  Mapping lignes (L1, L2, Imprimerie)
# -------------------------------------------------
LINE_MAP = {
    "U4M LIGNE 1": "L1",
    "U4M LIGNE 2 BASE": "L2",
    "U4M IMPRIMERIE BASE": "IMPRIMERIE",
}

# -------------------------------------------------
#  Style (fond dark + cards)
# -------------------------------------------------
# st.set_page_config dans app.py

st.markdown(
    """
    <style>
    body {
        background-color: #111827;
        color: #f9fafb;
    }
    .main, .block-container {
        background-color: #111827 !important;
    }
    .card {
        background-color: #1f2937;
        padding: 1.2rem 1.5rem;
        border-radius: 0.7rem;
        border: 1px solid #374151;
        box-shadow: 0 0 10px rgba(0,0,0,0.4);
    }
    .card-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #e5e7eb;
        margin-bottom: 0.4rem;
    }
    .kpi-row {
        display: flex;
        justify-content: space-between;
        margin-top: 0.25rem;
    }
    .kpi-label {
        font-size: 0.8rem;
        color: #9ca3af;
    }
    .kpi-value {
        font-size: 1.05rem;
        font-weight: 600;
        color: #f9fafb;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------
#  Chargement Qualite.xlsx + calculs de base
# -------------------------------------------------

def show_dashboard_trg():
    # Bouton retour
    col_back, _ = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"] = "menu"
            st.rerun()

    @st.cache_data
    def load_data(qualite_path: str) -> pd.DataFrame:
        df = pd.read_excel(qualite_path, sheet_name=0, engine="openpyxl")

        # On garde uniquement les ML
        df = df[df["Unité IM"] == "ML"].copy()

        # Normalisation lignes
        df["Ligne"] = df["Libelle ligne"].map(LINE_MAP).fillna(df["Libelle ligne"])

        # Dates
        df["Date début OF"] = pd.to_datetime(df["Date début OF"])
        df["Date fin OF"] = pd.to_datetime(df["Date fin OF"])
        df["duree_h"] = (df["Date fin OF"] - df["Date début OF"]).dt.total_seconds() / 3600
        df = df[df["duree_h"] > 0]

        # Semaine / année / jour
        iso = df["Date début OF"].dt.isocalendar()
        df["semaine"] = iso.week
        df["annee"] = iso.year
        df["jour_sem"] = df["Date début OF"].dt.dayofweek  # 0=lundi

        # Scrap & rendement qualité
        df["scrap_ml"] = df["Quantité mvt IM"] - df["Quantité IC"]
        df["rendement_qualite"] = df["Quantité IC"] / df["Quantité mvt IM"]

        # TRG réel / prévu basés uniquement sur quantités + durée
        df["TRG_reel_OF"] = df["Quantité mvt IM"] / df["duree_h"]          # ML/h réel
        df["TRG_prev_OF"] = df["Quantité demandée"] / df["duree_h"]        # ML/h prévu

        return df

    qualite_path = Path("Qualite.xlsx")

    if not os.path.exists(qualite_path):
        st.error(f"Fichier '{qualite_path}' introuvable dans : {os.getcwd()}")
        st.stop()

    df = load_data(qualite_path)

    # -------------------------------------------------
    #  Filtres
    # -------------------------------------------------
    st.sidebar.title("Filtres")

    annees = sorted(df["annee"].unique())
    annee_sel = st.sidebar.selectbox("Année", annees, index=len(annees) - 1)

    semaines = sorted(df[df["annee"] == annee_sel]["semaine"].unique())
    semaine_sel = st.sidebar.selectbox("Semaine ISO", semaines, index=len(semaines) - 1)

    df_week = df[(df["annee"] == annee_sel) & (df["semaine"] == semaine_sel)]
    if df_week.empty:
        st.warning("Pas de données pour cette semaine.")
        st.stop()

    # -------------------------------------------------
    #  Agrégations
    # -------------------------------------------------
    # Global par semaine (toutes lignes)
    agg_all = (
        df.groupby(["annee", "semaine"])
        .agg(
            prod_semaine_ml=("Quantité mvt IM", "sum"),
            scrap_ml=("scrap_ml", "sum"),
        )
        .reset_index()
    )

    # Par semaine & ligne (pour moyennes historiques)
    agg_all_line = (
        df.groupby(["annee", "semaine", "Ligne"])
        .agg(
            prod_semaine_ml=("Quantité mvt IM", "sum"),
            scrap_ml=("scrap_ml", "sum"),
        )
        .reset_index()
    )

    weekly_avg_by_line = (
        agg_all_line.groupby("Ligne")
        .agg(
            prod_moy=("prod_semaine_ml", "mean"),
            scrap_moy=("scrap_ml", "mean"),
        )
        .reset_index()
    )

    # Semaine sélectionnée : global
    global_week = agg_all[
        (agg_all["annee"] == annee_sel) & (agg_all["semaine"] == semaine_sel)
    ].iloc[0]
    total_prod_sem = global_week["prod_semaine_ml"]
    total_scrap_sem = global_week["scrap_ml"]

    weekly_avg_prod = agg_all["prod_semaine_ml"].mean()
    weekly_avg_scrap = agg_all["scrap_ml"].mean()

    # Semaine sélectionnée : par ligne
    agg_week_line = (
        df_week.groupby("Ligne")
        .agg(
            prod_semaine_ml=("Quantité mvt IM", "sum"),
            bon_ml=("Quantité IC", "sum"),
            scrap_ml=("scrap_ml", "sum"),
            rdt_budget_moy=("Rdt budget", "mean"),
            duree_h=("duree_h", "sum"),
            qte_dem=("Quantité demandée", "sum"),
            trg_reel_moy=("TRG_reel_OF", "mean"),
            trg_prev_moy=("TRG_prev_OF", "mean"),
            nb_of=("Numéro OF", "nunique"),
        )
        .reset_index()
    )

    # Rendement
    agg_week_line["rendement_reel"] = agg_week_line["bon_ml"] / agg_week_line["prod_semaine_ml"]
    agg_week_line["rendement_prev"] = agg_week_line["rdt_budget_moy"]

    # TRG réel / prévu (ML/h) : pondéré par la durée
    agg_week_line["TRG_reel_ml_h"] = agg_week_line["prod_semaine_ml"] / agg_week_line["duree_h"]
    agg_week_line["TRG_prev_ml_h"] = agg_week_line["qte_dem"] / agg_week_line["duree_h"]

    # Ratio TRG (pour le donut) : réel vs prévu
    agg_week_line["TRG_ratio"] = agg_week_line["TRG_reel_ml_h"] / agg_week_line["TRG_prev_ml_h"]
    agg_week_line.replace([float("inf"), -float("inf")], 0, inplace=True)

    # -------------------------------------------------
    #  Donut TRG (simple ratio)
    # -------------------------------------------------
    def donut_ratio(ratio, title="TRG"):
        ratio = max(0, min(ratio, 1.5))  # on cappe à 150%

        fig = go.Figure(
            go.Pie(
                values=[min(ratio, 1.0), max(0, 1 - ratio)],
                hole=0.7,
                marker_colors=["#10b981", "#374151"],
                textinfo="none",
            )
        )
        fig.update_layout(
            margin=dict(l=0, r=0, t=40, b=0),
            annotations=[
                dict(
                    text=f"{ratio*100:.1f}%",
                    x=0.5,
                    y=0.5,
                    font=dict(size=18, color="white"),
                    showarrow=False,
                ),
                dict(
                    text=title,
                    x=0.5,
                    y=1.1,
                    font=dict(size=13, color="#e5e7eb"),
                    showarrow=False,
                ),
            ],
            showlegend=False,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        return fig

    # -------------------------------------------------
    #  HEADER + SUMMARY
    # -------------------------------------------------
    st.markdown(
        f"""
        <h2 style="color:#f9fafb; margin-bottom:0;">Dashboard TRG Lignes 4m</h2>
        <h4 style="color:#9ca3af; margin-top:0.2rem;">Semaine {semaine_sel} - {annee_sel}</h4>
        """,
        unsafe_allow_html=True,
    )

    top_col_summary, top_col_L1, top_col_L2, top_col_IMP = st.columns([1.2, 1, 1, 1])

    # ---------- Summary global ----------
    with top_col_summary:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">Summary</div>', unsafe_allow_html=True)

        for label, value in [
            ("Total (ML)", total_prod_sem),
            ("Weekly Average (ML)", weekly_avg_prod),
            ("Scrap (ML)", total_scrap_sem),
            ("Scrap Weekly Average (ML)", weekly_avg_scrap),
        ]:
            st.markdown(
                f"""
                <div class="kpi-row">
                  <div class="kpi-label">{label}</div>
                  <div class="kpi-value">{value:,.0f}</div>
                </div>
                """.replace(",", " "),
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    avg_by_line = weekly_avg_by_line.set_index("Ligne")

    def render_line_card(col, ligne_name, titre_affiche):
        sub = agg_week_line[agg_week_line["Ligne"] == ligne_name]
        if sub.empty:
            with col:
                st.markdown(
                    f'<div class="card"><div class="card-title">{titre_affiche}</div>'
                    '<p style="color:#9ca3af;">Aucune donnée pour cette semaine.</p></div>',
                    unsafe_allow_html=True,
                )
            return

        row = sub.iloc[0]
        prod_moy = avg_by_line.loc[ligne_name, "prod_moy"] if ligne_name in avg_by_line.index else 0
        scrap_moy = avg_by_line.loc[ligne_name, "scrap_moy"] if ligne_name in avg_by_line.index else 0

        with col:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f'<div class="card-title">{titre_affiche}</div>', unsafe_allow_html=True)

            # KPIs type screenshot
            st.markdown(
                f"""
                <div class="kpi-row">
                  <div class="kpi-label">Total (ML)</div>
                  <div class="kpi-value">{row['prod_semaine_ml']:,.0f}</div>
                </div>
                <div class="kpi-row">
                  <div class="kpi-label">Weekly Average (ML)</div>
                  <div class="kpi-value">{prod_moy:,.0f}</div>
                </div>
                <div class="kpi-row">
                  <div class="kpi-label">Scrap (ML)</div>
                  <div class="kpi-value">{row['scrap_ml']:,.0f}</div>
                </div>
                <div class="kpi-row">
                  <div class="kpi-label">Scrap Average (ML)</div>
                  <div class="kpi-value">{scrap_moy:,.0f}</div>
                </div>
                """.replace(",", " "),
                unsafe_allow_html=True,
            )

            # Donut : TRG réel vs prévu (ratio en %)
            ratio = row["TRG_ratio"] if row["TRG_prev_ml_h"] > 0 else 0
            fig_trg = donut_ratio(ratio, title=f"Ligne {ligne_name} - TRG réel / prévu")
            st.plotly_chart(fig_trg, use_container_width=True)

            st.markdown("<hr style='border-color:#374151;'>", unsafe_allow_html=True)

            col_r1, col_r2 = st.columns(2)
            with col_r1:
                st.markdown("**TRG réel (ML/h)**")
                st.metric(
                    label="",
                    value=f"{row['TRG_reel_ml_h']:.1f}",
                    delta=f"{row['TRG_reel_ml_h'] - row['TRG_prev_ml_h']:.1f}",
                )
                st.markdown("**Rendement réel (%)**")
                st.metric(
                    label="",
                    value=f"{row['rendement_reel']*100:.1f}%",
                    delta=f"{(row['rendement_reel'] - row['rendement_prev'])*100:.1f} pts",
                )
            with col_r2:
                st.markdown("**TRG prévu (ML/h)**")
                st.metric(label="", value=f"{row['TRG_prev_ml_h']:.1f}")
                st.markdown("**Rendement prévu (%)**")
                st.metric(label="", value=f"{row['rendement_prev']*100:.1f}%")

            st.markdown("</div>", unsafe_allow_html=True)

    render_line_card(top_col_L1, "L1", "Line - L1")
    render_line_card(top_col_L2, "L2", "Line - L2")
    render_line_card(top_col_IMP, "IMPRIMERIE", "Line - Imprimerie")

    st.markdown("---")

    # -------------------------------------------------
    #  PRODUCTION THIS WEEK / TREND / DISTRIBUTION
    # -------------------------------------------------
    jour_labels = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}
    df_week["Jour"] = df_week["jour_sem"].map(jour_labels)

    prod_jour = (
        df_week.groupby(["Jour", "Ligne"])
        .agg(prod_ml=("Quantité mvt IM", "sum"))
        .reset_index()
    )

    fig_bar = px.bar(
        prod_jour,
        x="Jour",
        y="prod_ml",
        color="Ligne",
        barmode="group",
        title="Production this Week",
        color_discrete_sequence=px.colors.qualitative.Set2,
    )
    fig_bar.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(17,24,39,1)",
        font_color="white",
        xaxis_title="Day",
        yaxis_title="ML",
    )

    agg_all["annee_semaine"] = (
        agg_all["annee"].astype(str) + "-S" + agg_all["semaine"].astype(str)
    )
    fig_trend = px.line(
        agg_all.sort_values(["annee", "semaine"]),
        x="annee_semaine",
        y="prod_semaine_ml",
        title="Trend in Production",
    )
    fig_trend.update_traces(line_color="#22c55e")
    fig_trend.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(17,24,39,1)",
        font_color="white",
        xaxis_title="Week",
        yaxis_title="ML",
    )

    fig_dist = px.pie(
        agg_week_line,
        names="Ligne",
        values="prod_semaine_ml",
        hole=0.5,
        title="Production Line Distribution",
    )
    fig_dist.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(17,24,39,1)",
        font_color="white",
    )

    bottom_c1, bottom_c2, bottom_c3 = st.columns(3)
    with bottom_c1:
        st.plotly_chart(fig_bar, use_container_width=True)
    with bottom_c2:
        st.plotly_chart(fig_trend, use_container_width=True)
    with bottom_c3:
        st.plotly_chart(fig_dist, use_container_width=True)

    # -------------------------------------------------
    #  Imprimerie : Moyenne hebdo vs réalisé (ML)
    # -------------------------------------------------
    weekly_avg_by_line = weekly_avg_by_line.set_index("Ligne")

    if (
        "IMPRIMERIE" in weekly_avg_by_line.index
        and not agg_week_line[agg_week_line["Ligne"] == "IMPRIMERIE"].empty
    ):
        impr_avg = weekly_avg_by_line.loc["IMPRIMERIE", "prod_moy"]
        impr_real = agg_week_line[agg_week_line["Ligne"] == "IMPRIMERIE"]["prod_semaine_ml"].iloc[0]

        df_impr = pd.DataFrame(
            {"Type": ["Moyenne hebdo", f"Semaine {semaine_sel}"], "ML": [impr_avg, impr_real]}
        )

        fig_impr = px.bar(
            df_impr,
            x="Type",
            y="ML",
            title="Imprimerie - Moyenne hebdo vs Réalisé (ML)",
            color="Type",
            color_discrete_sequence=["#6b7280", "#22c55e"],
        )
        fig_impr.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(17,24,39,1)",
            font_color="white",
            showlegend=False,
        )
        st.plotly_chart(fig_impr, use_container_width=True)
