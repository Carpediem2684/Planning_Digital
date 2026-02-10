# ============================================
# app.py — Fichier principal Streamlit
# Planification – Tetart.Y
# ============================================

import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Planification – Tetart.Y", layout="wide")

# ----- INITIALISATION -----
if "page" not in st.session_state:
    st.session_state["page"] = "home"

page = st.session_state["page"]

# ============================
#     PAGE D'ACCUEIL DESIGN
# ============================

if page == "home":

    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
        
        .main-container {
            max-width: 900px;
            margin: 0 auto;
            padding: 2rem;
        }
        
        .hero-title {
            font-family: 'Inter', sans-serif;
            font-size: 52px;
            font-weight: 800;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
            margin-bottom: 0.5rem;
            letter-spacing: -1px;
        }
        
        .hero-subtitle {
            font-family: 'Inter', sans-serif;
            text-align: center;
            font-size: 18px;
            color: #6B7280;
            margin-bottom: 3rem;
            font-weight: 400;
        }
        
        .card-container {
            background: linear-gradient(145deg, #ffffff 0%, #f8fafc 100%);
            border-radius: 20px;
            padding: 2.5rem;
            box-shadow: 0 10px 40px rgba(0,0,0,0.08);
            border: 1px solid #e5e7eb;
            margin-bottom: 2rem;
        }
        
        .section-title {
            font-family: 'Inter', sans-serif;
            font-size: 16px;
            font-weight: 600;
            color: #374151;
            margin-bottom: 1.5rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .uap-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 1rem;
        }
        
        div.stButton > button {
            font-family: 'Inter', sans-serif;
            width: 100%;
            padding: 1rem 1.5rem;
            border-radius: 12px;
            font-size: 18px;
            font-weight: 600;
            transition: all 0.3s ease;
            border: 2px solid transparent;
        }
        
        /* Style pour les boutons UAP */
        .stButton > button[kind="secondary"] {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
        }
        
        .stButton > button[kind="secondary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
        }
        
        /* Bouton Menu principal */
        .menu-btn-wrapper {
            margin-top: 2rem;
            text-align: center;
        }
        
        .footer-info {
            text-align: center;
            color: #9CA3AF;
            font-size: 13px;
            margin-top: 2rem;
            padding-top: 1.5rem;
            border-top: 1px solid #E5E7EB;
        }
        
        .status-badge {
            display: inline-block;
            background: #DEF7EC;
            color: #03543F;
            padding: 0.25rem 0.75rem;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 0.5rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # Titre principal
    st.markdown('<h1 class="hero-title">🏭 Planification Production</h1>', unsafe_allow_html=True)
    st.markdown('<p class="hero-subtitle">Outil de gestion et optimisation des plannings 4M<span class="status-badge">✓ En ligne</span></p>', unsafe_allow_html=True)

    # Container principal
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    st.markdown('<p class="section-title">📍 Sélectionner une UAP</p>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        if st.button("🔵 4M", key="uap_4m", use_container_width=True):
            st.session_state["uap_selection"] = "4M"
            st.session_state["page"] = "menu"
            st.rerun()
        
        if st.button("🟢 P2000", key="uap_p2000", use_container_width=True):
            st.session_state["uap_selection"] = "P2000"
            st.session_state["page"] = "menu"
            st.rerun()

    with col2:
        if st.button("🟡 2M", key="uap_2m", use_container_width=True):
            st.session_state["uap_selection"] = "2M"
            st.session_state["page"] = "menu"
            st.rerun()
        
        if st.button("🔴 KLAM", key="uap_klam", use_container_width=True):
            st.session_state["uap_selection"] = "KLAM"
            st.session_state["page"] = "menu"
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # Bouton accès direct au menu
    st.markdown('<div class="menu-btn-wrapper">', unsafe_allow_html=True)
    col_left, col_center, col_right = st.columns([1, 2, 1])
    with col_center:
        if st.button("🚀 ACCÈS DIRECT AU MENU", key="big_menu_btn", use_container_width=True, type="primary"):
            st.session_state["page"] = "menu"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown(f'''
        <div class="footer-info">
            <strong>Tetart.Y</strong> • Version 2.0 • {datetime.now().strftime("%d/%m/%Y %H:%M")}<br>
            Optimisation de la production 4M
        </div>
    ''', unsafe_allow_html=True)

# ============================
#     ROUTING DES PAGES
# ============================
elif page == "menu":
    from pages.Menu import show_menu
    show_menu()

elif page == "planning_global":
    from pages.Planning_Global import show_planning_global
    show_planning_global()

elif page == "dashboard_pic":
    from pages.Dashboard_PIC import show_dashboard_pic
    show_dashboard_pic()

elif page == "dashboard_trg":
    from pages.Dashboard_TRG import show_dashboard_trg
    show_dashboard_trg()

elif page == "planning_l1":
    from pages.Planning_L1 import show_planning_l1
    show_planning_l1()

elif page == "planning_imprimerie":
    from pages.Planning_Imprimerie import show_planning_imprimerie
    show_planning_imprimerie()

elif page == "planning_l2":
    from pages.Planning_L2 import show_planning_l2
    show_planning_l2()

elif page == "planning_visitage":
    from pages.Planning_Visitage import show_planning_visitage
    show_planning_visitage()

elif page == "qualite":
    from pages.Qualite import show_qualite
    show_qualite()

elif page == "settings":
    from pages.Settings import show_settings
    show_settings()
