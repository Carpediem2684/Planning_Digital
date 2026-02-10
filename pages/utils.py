# ============================================
# utils.py — Fonctions utilitaires partagées
# Version: 2026-02-07
# ============================================

import pandas as pd
from pathlib import Path
import streamlit as st

BASE_PATH = Path(r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main")
SUIVI_OF_FILE = BASE_PATH / "SUIVI_OF.xlsx"

# Statuts
STATUT_ACTIF = [30, 40, 50]          # Pastille verte
STATUT_TERMINE = [60, 61, 99]        # À exclure
ML_MINIMUM = 200

# Supports L1 pour stock Imprimerie
SUPPORTS_L1 = ["S1013", "S2012", "S2009", "S2010", "S1162", "S1166", "S2019", "S2020", 
               "S2003", "S2004", "S2005", "S2006", "S2011", "S2015", "S2014", "S1016"]


@st.cache_data(ttl=7200)  # Cache de 2h
def load_suivi_of():
    """Charge SUIVI_OF avec cache de 2h."""
    return pd.read_excel(SUIVI_OF_FILE, engine="openpyxl")


def get_ofs_exclus(suivi_df, ligne_pattern):
    """Retourne les OFs à exclure (terminés + < 200ml)."""
    df = suivi_df[suivi_df["LIB_LIGNE"].str.contains(ligne_pattern, case=False, na=False)]
    ofs_termines = set(df[df["STATUT"].isin(STATUT_TERMINE)]["NUM_OF"].tolist())
    ofs_petits = set(df[df["COMMANDE"] < ML_MINIMUM]["NUM_OF"].tolist())
    return ofs_termines | ofs_petits


def get_statut_dict(suivi_df, ligne_pattern):
    """Retourne un dict NUM_OF -> STATUT pour une ligne."""
    df = suivi_df[suivi_df["LIB_LIGNE"].str.contains(ligne_pattern, case=False, na=False)]
    return dict(zip(df["NUM_OF"], df["STATUT"]))


def get_fabrique_dict(suivi_df):
    """Retourne dict COLORIS -> FABRIQUE pour OFs Imprimerie avec FABRIQUE > 0."""
    df = suivi_df[suivi_df["LIB_LIGNE"].str.contains("IMPRIMERIE", case=False, na=False)]
    fabrique = {}
    for _, row in df.iterrows():
        coloris = row["COLORIS"]
        fab = row["FABRIQUE"]
        if coloris and pd.notna(fab) and fab > 0:
            if coloris not in fabrique or fab > fabrique[coloris]:
                fabrique[coloris] = fab
    return fabrique


def get_stock_supports(suivi_df):
    """Récupère le stock des supports L1 (une seule valeur par support)."""
    stocks = {}
    seen = set()
    
    for _, row in suivi_df.iterrows():
        composant = str(row.get("COMPOSANT", ""))
        stock = row.get("STOCK_COMPOSANT", 0)
        
        if pd.isna(composant) or composant == "nan":
            continue
            
        for support in SUPPORTS_L1:
            if support in composant and support not in seen:
                stocks[support] = stock if pd.notna(stock) else 0
                seen.add(support)
    
    return stocks


def is_statut_actif(of_num, statut_dict):
    """Vérifie si un OF a un statut actif (30, 40, 50)."""
    statut = statut_dict.get(of_num)
    return statut in STATUT_ACTIF
