# ============================================
# Menu.py — Page Menu principal
# ============================================

import streamlit as st

def show_menu():
    
    uap = st.session_state.get("uap_selection", "4M")

    st.markdown(
        """
        <style>
        .menu-title {
            text-align: center;
            font-size: 38px;
            font-weight: 800;
            color: #1B263B;
            margin-top: 10px;
            margin-bottom: 5px;
        }
        .menu-subtitle {
            font-size: 18px;
            color: #555;
            margin-bottom: 30px;
        }
        .menu-section-label {
            font-size: 18px;
            font-weight: 600;
            color: #2C3E50;
            margin-top: 10px;
            margin-bottom: 15px;
        }
        div.stButton > button {
            border-radius: 8px;
            border: 1px solid #2C3E50;
            padding-top: 10px;
            padding-bottom: 10px;
            font-size: 16px;
            font-weight: 600;
            color: #2C3E50;
            background: #F8F9FA;
        }
        div.stButton > button:hover {
            color: #FFFFFF;
            border-color: #1F8FFF;
            background: linear-gradient(90deg, #1F8FFF, #6EC6FF);
            box-shadow: 0 0 10px rgba(31,143,255,0.4);
        }
        .back-btn-container {
            margin-top: 30px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f"<div class='menu-title'>Menu – UAP {uap}</div>", unsafe_allow_html=True)
    st.markdown("<div class='menu-subtitle' style='text-align:center;'>Sélectionne un module de planification ou de suivi.</div>", unsafe_allow_html=True)

    st.markdown("<div class='menu-section-label'>Choisir un module :</div>", unsafe_allow_html=True)

    # Dashboard PIC
    _, col_pic, _ = st.columns([1, 2, 1])
    with col_pic:
        if st.button("📊  Dashboard PIC", use_container_width=True):
            st.session_state["page"] = "dashboard_pic"
            st.rerun()

    st.write("")

    # Qualité
    _, col_qual, _ = st.columns([1, 2, 1])
    with col_qual:
        if st.button("📈  Qualité", use_container_width=True):
            st.session_state["page"] = "qualite"
            st.rerun()

    st.write("")

    # Dashboard TRG (NOUVEAU)
    _, col_trg, _ = st.columns([1, 2, 1])
    with col_trg:
        if st.button("📉  Dashboard TRG", use_container_width=True):
            st.session_state["page"] = "dashboard_trg"
            st.rerun()

    st.write("")

    # Planning Global
    _, col_global, _ = st.columns([1, 2, 1])
    with col_global:
        if st.button("🗺️  Planning Global", use_container_width=True):
            st.session_state["page"] = "planning_global"
            st.rerun()

    st.write("")

    # 4 boutons plannings
    col_l1, col_imp, col_l2, col_vis = st.columns(4)

    with col_l1:
        if st.button("🏭  Planning L1", use_container_width=True):
            st.session_state["page"] = "planning_l1"
            st.rerun()

    with col_imp:
        if st.button("🖨️  Planning Imprimerie", use_container_width=True):
            st.session_state["page"] = "planning_imprimerie"
            st.rerun()

    with col_l2:
        if st.button("⚙️  Planning L2", use_container_width=True):
            st.session_state["page"] = "planning_l2"
            st.rerun()

    with col_vis:
        if st.button("👁  Planning Visitage", use_container_width=True):
            st.session_state["page"] = "planning_visitage"
            st.rerun()

    st.write("")

    # SETTINGS (NOUVEAU)
    _, col_set, _ = st.columns([1, 2, 1])
    with col_set:
        if st.button("⚙️  SETTINGS - Gestion des Plannings", use_container_width=True):
            st.session_state["page"] = "settings"
            st.rerun()

    st.markdown("<div class='back-btn-container'></div>", unsafe_allow_html=True)

    # Retour accueil
    _, col_back, _ = st.columns([1, 2, 1])
    with col_back:
        if st.button("⬅️  Retour à l'accueil", use_container_width=True):
            st.session_state["page"] = "home"
            st.rerun()
