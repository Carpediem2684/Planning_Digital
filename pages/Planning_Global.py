# ============================================================
# PLANNING GLOBAL HARMONISÉ — VERSION FINALE
# ============================================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import plotly.graph_objects as go

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
BASE_PATH = Path(
    r"C:\Users\yannick.tetard\OneDrive - GERFLOR\Desktop\Planning Streamlit\xarpediem2684-repo-main"
)

FILES = {
    "Ligne 1": {
        "ofs": BASE_PATH / "OFs_L1.xlsx",
        "cal": BASE_PATH / "Calendrier 2026 L1.xlsx",
        "sheet": "Feuil1",
    },
    "Imprimerie": {
        "ofs": BASE_PATH / "OFs_Imprimerie.xlsx",
        "cal": BASE_PATH / "Calendrier 2026 imprimerie.xlsx",
        "sheet": "Feuil1",
    },
    "Ligne 2": {
        "ofs": BASE_PATH / "OFs_L2.xlsx",
        "cal": BASE_PATH / "Calendrier 2026.xlsx",
        "sheet": "Sheet1",
    },
    "Visitage": {
        "ofs": BASE_PATH / "OFs_Visitage.xlsx",
        "cal": BASE_PATH / "Calendrier 2026.xlsx",
        "sheet": "Feuil1",
    }
}

INTRO_DUREE = {
    "Ligne 1": 2.3,
    "Imprimerie": 0,
    "Ligne 2": 2.3,
    "Visitage": 0,
}

# ------------------------------------------------------------
# COULEURS / UTILITAIRES
# ------------------------------------------------------------
def rgb_to_hex(r,g,b): return f"#{r:02X}{g:02X}{b:02X}"

# ===== L1 - couleurs produits  (source Planning_L1) [2](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_L2.py) =====
PRODUIT_COLOR_L1 = {
    "CICD03 3M": rgb_to_hex(204,204,0),
    "CICD03 4M": rgb_to_hex(255,255,0),
    "CICD05 4M": rgb_to_hex(205,133,63),
    "CICD06 3M": rgb_to_hex(255,182,193),
    "CICD06 4M": rgb_to_hex(255,119,255),
    "CICD02 4M": rgb_to_hex(211,211,211),
    "CICDMD01 4M": rgb_to_hex(169,169,169),
    "CIMD02 3M": rgb_to_hex(0,128,0),
    "CIMD02 4M": rgb_to_hex(144,238,144),
    "CIMD03 3M": rgb_to_hex(0,255,255),
    "CIMD03 4M": rgb_to_hex(0,100,0),
    "CICD01 4M": rgb_to_hex(0,139,139),
    "CICD04 4M": rgb_to_hex(173,216,230),
}

# ===== Imprimerie - couleurs traits campagnes (Planning_Imprimerie) [1](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_Visitage.py)
CAMPAGNE_COLOR_IMP = {
    "PRIMETEX": rgb_to_hex(204,204,0),
    "TEXLINE": rgb_to_hex(0,128,0),
    "TARABUS": rgb_to_hex(0,128,0),
    "BOOSTER": rgb_to_hex(255,0,0),
    "TMAX": rgb_to_hex(255,0,0),
    "START": rgb_to_hex(128,128,128),
    "SPORISOL": rgb_to_hex(128,128,128),
    "NERA": rgb_to_hex(0,0,139),
}
DOUBLE_TRAIT = ["TARABUS","TMAX"]

# ===== L2 - mapping complet famille (Planning_L2) [2](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_L2.py)
FAMILLE_COLOR_L2 = {
    "INTRO": ("#FFFFFF","#000000"),
    "TIMBERLINE/NEROK 3M": ("#00008B","#FFFFFF"),
    "START 4M": ("#D3D3D3","#000000"),
    "SPORISOL 4M": ("#C0C0C0","#000000"),
    "TARASTEP PRO": ("#FFFF66","#000000"),
    "PRIMETEX GRT 4M": ("#CCCC00","#000000"),
    "TEXLINE GRIP'X 3M": ("#FFB6C1","#000000"),
    "TEXLINE GRIP'X 4M": ("#8B008B","#FFFFFF"),
    "NEROK TEX": ("#000000","#FFFFFF"),
    "NEROK 50 TEX": ("#000000","#FFFFFF"),
    "BAGNOSTAR 4M": ("#FF6666","#FFFFFF"),
    "BAGNOSTAR METAL 4M": ("#FF6666","#FFFFFF"),
    "BAGNOSTAR 3M": ("#FF6666","#FFFFFF"),
    "SRA ACOUSTIC 3M": ("#808080","#FFFFFF"),
    "FUSION 3M": ("#006400","#FFFFFF"),
    "TARABUS HARMONIA 1/2": ("#50C878","#000000"),
    "RECONCEPTION TARABUS H 1/2": ("#50C878","#000000"),
    "TRADIFLOR 2S2 3M": ("#483C32","#FFFFFF"),
    "TRADIFLOR 2S2 4M": ("#483C32","#FFFFFF"),
    "TRANSIT-TEX MAX 33-43 2/2 4M": ("#654321","#FFFFFF"),
    "TARABUS HARMONIA INTER": ("#ADD8E6","#000000"),
    "RECONCEPTION TARABUS H 2/2": ("#ADD8E6","#000000"),
    "GERBAD EVOLUTION 3M": ("#E0FFFF","#000000"),
    "GERBAD EVOLUTION 4M": ("#E0FFFF","#000000"),
    "TIMBERLINE 4M": ("#E0FFFF","#000000"),
    "SRA ACOUSTIC 4M": ("#F5F5F5","#000000"),
    "TRANSIT-TEX MAX 33-43 1/2 4M": ("#CD853F","#FFFFFF"),
    "PRIMETEX GRT 3M": ("#C3B091","#000000"),
    "BAGNOSTAR 2.5 3M": ("#E6E6FA","#000000"),
    "BAGNOSTAR 2.5 4M": ("#E6E6FA","#000000"),
    "BAGNOSTAR 2.5 METAL 4M": ("#E6E6FA","#000000"),
    "TRANSIT TEX MAX 2S3 1/2 4M": ("#D2B48C","#000000"),
    "BOOSTER 2.6 DIAM 4M": ("#FA8072","#000000"),
    "MELODY 4M": ("#FA8072","#000000"),
    "LOFTEX NATURE 4M": ("#FF8C00","#000000"),
    "BAGNOSTAR MATT 4M": ("#FF6961","#FFFFFF"),
    "SRA ACOUSTIC PU 4M": ("#FF6961","#FFFFFF"),
    "TRANSIT-TEX 4M": ("#FF77FF","#000000"),
    "TRANSIT-TEX 2/2 4M": ("#FF77FF","#000000"),
    "TEXLINE NATURE 4M": ("#FFA500","#000000"),
    "TEXLINE NATURE 3M": ("#FFA500","#000000"),
    "LOFTEX GRT 3M": ("#FFA500","#000000"),
    "LOFTEX GRT 4M": ("#FFA500","#000000"),
    "LOFTEX NATURE 3M": ("#FFA500","#000000"),
    "BOOSTER 2.6 3M": ("#FFA500","#000000"),
    "PRIMETEX 4M": ("#FDFD96","#000000"),
    "PRIMETEX 3M": ("#FDFD96","#000000"),
    "PRIMETEX MATT 3M": ("#FDFD96","#000000"),
    "PRIMETEX MATT 4M": ("#FDFD96","#000000"),
    "START 3M": ("#2F4F4F","#FFFFFF"),
    "GRAFIC PU 3M": ("#2F4F4F","#FFFFFF"),
    "GRAFIC PU 4M": ("#2F4F4F","#FFFFFF"),
    "TARABUS HARMONIA": ("#000080","#FFFFFF"),
    "TIMBERLINE/NEROK 4M": ("#000080","#FFFFFF"),
    "TRANSIT TEX MAX 2S3 2/2 4M": ("#483C32","#FFFFFF"),
    "TEXLINE HQR 3M": ("#008000","#FFFFFF"),
    "TEXLINE HQR 4M": ("#008000","#FFFFFF"),
    "TEXLINE GRT 3M": ("#008000","#FFFFFF"),
    "FUSION 4M": ("#CCFFCC","#000000"),
    "TEXLINE GRT 4M": ("#006400","#FFFFFF"),
    "NEROK TEX NATURE": ("#FFFFFF","#000000"),
    "NERA FIRST 4M": ("#FFFFFF","#000000"),
    "TRANSIT-TEX PLUS 2/2 4M": ("#FFFFFF","#000000"),
    "BOOSTER 2.6 GRAIN 3M PUR BLANC": ("#FFFFFF","#000000"),
    "TRANSIT-TEX PLUS 1/2 4M": ("#FFFFFF","#000000"),
    "ESSAI UAP4M L2 1/2": ("#FFFFFF","#000000"),
    "BOOSTER PUR BLANC 4M": ("#FFFFFF","#000000"),
    "BOOSTER 2.6 GRAIN 3M": ("#FFC0CB","#000000"),
    "ESSAI UAP4M L2 1/1": ("#FFFFFF","#000000"),
    "BOOSTER 2.6 4M": ("#FF0000","#FFFFFF")
}

# ===== Visitage (Planning_Visitage) [1](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_Visitage.py)
VIS_TOP_COLOR = {
    "NERA": rgb_to_hex(0,0,139),
    "START": rgb_to_hex(211,211,211),
    "TARASTEP": rgb_to_hex(255,255,224),
    "PRIMETEX": rgb_to_hex(204,204,0),
    "GRIPX": rgb_to_hex(255,182,193),
    "TARABUS": rgb_to_hex(80,200,120),
    "TMAX": rgb_to_hex(101,67,33),
    "BOOSTER": rgb_to_hex(250,128,114),
    "LOFTEX": rgb_to_hex(255,140,0),
    "TEXLINE": rgb_to_hex(0,128,0),
    "SPORISOL": rgb_to_hex(128,128,128),
    "FUSION": rgb_to_hex(144,238,144),
}
VIS_BOTTOM_COLOR = {
    4: rgb_to_hex(119,221,119),
    3: rgb_to_hex(173,216,230),
    2: rgb_to_hex(255,182,193),
}

# ------------------------------------------------------------
# LOAD DATA
# ------------------------------------------------------------
@st.cache_data
def load_all():
    data = {}
    for ligne,f in FILES.items():
        try:
            ofs = pd.read_excel(f["ofs"], sheet_name=f["sheet"], engine="openpyxl")
            cal = pd.read_excel(f["cal"], engine="openpyxl")
            data[ligne] = {"ofs": ofs, "cal": cal}
        except Exception as e:
            data[ligne] = {"ofs": pd.DataFrame(), "cal": pd.DataFrame(), "error": str(e)}
    return data

# ------------------------------------------------------------
# SLOTS
# ------------------------------------------------------------
def parse_horaire(jour, text):
    if not isinstance(text,str) or "-" not in text: return None,None
    d,f = text.split("-")
    d = d.replace("h",":"); f = f.replace("h",":")
    start = datetime.combine(jour.date(), datetime.strptime(d,"%H:%M").time())
    end   = datetime.combine(jour.date(), datetime.strptime(f,"%H:%M").time())
    if end <= start: end += timedelta(days=1)
    return start,end

def build_slots(cal_df, start, end):
    cal = cal_df.copy()
    cal["Jour"] = pd.to_datetime(cal["Jour"])
    mask = (cal["Jour"].dt.date >= start.date()) & (cal["Jour"].dt.date <= end.date())
    slots = []
    for _,row in cal.loc[mask].iterrows():
        jour = row["Jour"]
        for i in [1,2,3]:
            if row.get(f"Etat_{i}")=="OUVERT":
                h=row.get(f"Horaire_{i}")
                if isinstance(h,str) and "-" in h:
                    s,e = parse_horaire(jour,h)
                    if s and e and e>start and s<end:
                        slots.append({"start":max(s,start),"end":min(e,end)})
    return sorted(slots,key=lambda x:x["start"])


# ------------------------------------------------------------
# RÈGLES INTRO L1 (Planning_L1) [2](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_L2.py)
# ------------------------------------------------------------
def l1_needs_intro(prev_prod,curr_prod):
    if not prev_prod or not curr_prod: return False
    p=prev_prod.upper(); c=curr_prod.upper()

    def t(x):
        if "CICDMD" in x: return "CICDMD"
        if "CIMD"   in x: return "CIMD"
        if "CICD"   in x: return "CICD"
        return "OTHER"

    def larg(x):
        if "3M" in x: return "3M"
        if "4M" in x: return "4M"
        return "?"

    pt,ct = t(p),t(c)
    pl,cl = larg(p),larg(c)

    if (pt=="CICD" and ct=="CIMD") or (pt=="CIMD" and ct=="CICD"): return True
    if (pt=="CICDMD" and ct=="CICD") or (pt=="CICD" and ct=="CICDMD"): return True
    if (pt=="CICDMD" and ct=="CIMD") or (pt=="CIMD" and ct=="CICDMD"): return True
    if pl=="3M" and cl=="4M": return True

    return False

# ------------------------------------------------------------
# RÈGLES INTRO L2 (Planning_L2) [2](https://gerflorgroup-my.sharepoint.com/personal/yannick_tetart_gerflor_com/Documents/Fichiers%20de%20conversation%20Microsoft%20Copilot/Planning_L2.py)
# ------------------------------------------------------------
def l2_needs_intro(prev_camp,curr_camp):
    if not prev_camp: return False
    return prev_camp != curr_camp

# ------------------------------------------------------------
# LABELS harmonisés
# ------------------------------------------------------------
def lbl_l1(r):
    prod=str(r.get("Produit",""))[:18]
    ml=r.get("Ml","")
    of=str(r.get("Ofs","")).split("_")[-1]
    return f"<b>{prod}</b><br>{ml} ML<br>OF {of} — {r['duree_h']:.1f}h"

def lbl_imp(r):
    col=str(r.get("Coloris",""))
    if "-" in col: code,desc = col.split("-",1)
    else: code,desc=col,""
    ml=r.get("Ml","")
    sup=r.get("Support","")
    of=str(r.get("Ofs","")).split("_")[-1]
    return f"<b>{code}</b><br>{desc[:15]}<br>{sup} — {ml} ML<br>OF {of} — {r['duree_h']:.1f}h"

def lbl_l2(r):
    col=r.get("COLORIS","")
    if "-" in str(col): code,desc=col.split("-",1)
    else: code,desc=col,""
    fam=r.get("FAMILLE","")
    gr=r.get("GRAIN","")
    ml=r.get("Ml","")
    of=str(r.get("Ofs","")).split("_")[-1]
    return f"{code}<br>{desc[:15]}<br>{fam} {gr}<br>{ml} ML — OF {of}<br>{r['duree_h']:.1f}h"

def lbl_vis(r):
    col=str(r.get("Coloris",""))[:12]
    ml =r.get("Ml","")
    la =r.get("Laise","")
    of=str(r.get("Ofs","")).split("_")[-1]
    return f"<b>{col}</b><br>{ml} ML — OF {of}<br>{r['duree_h']:.1f}h<br>Laise {la}"

LABEL_FUN = {
    "Ligne 1": lbl_l1,
    "Imprimerie": lbl_imp,
    "Ligne 2": lbl_l2,
    "Visitage": lbl_vis
}

# ------------------------------------------------------------
# PLANIFICATION GÉNÉRIQUE
# ------------------------------------------------------------
def schedule_generic(ofs_df,slots,ligne):
    rows=[]
    if ofs_df.empty or not slots: return pd.DataFrame()

    intro_d=INTRO_DUREE[ligne]
    slot_idx=0
    cur_s=slots[0]["start"]; cur_e=slots[0]["end"]
    prev_key=None

    def consume(h):
        nonlocal slot_idx,cur_s,cur_e
        segs=[]; rem=timedelta(hours=h)
        while rem>timedelta(0):
            if slot_idx>=len(slots): return segs,False
            avail=cur_e-cur_s
            if avail<=timedelta(0):
                slot_idx+=1
                if slot_idx>=len(slots): return segs,False
                cur_s=slots[slot_idx]["start"]; cur_e=slots[slot_idx]["end"]
                continue
            use=min(avail,rem)
            segs.append({"start":cur_s,"end":cur_s+use,"duree_h":use.total_seconds()/3600})
            rem-=use
            cur_s+=use
        return segs,True

    def key_of(r):
        if ligne=="Ligne 1" : return r.get("Produit","")
        if ligne=="Ligne 2" : return r.get("Campagne","")
        if ligne=="Imprimerie": return r.get("Campagne","")
        return None

    def color_of(r):
        if ligne=="Ligne 1":
            prod=str(r.get("Produit","")).upper()
            for k,v in PRODUIT_COLOR_L1.items():
                if k.upper() in prod: return v
            return "#CCCCCC"

        if ligne=="Ligne 2":
            fam=r.get("FAMILLE","")
            c,_=FAMILLE_COLOR_L2.get(fam,("#FFFFFF","#000000"))
            return c

        if ligne=="Visitage":
            camp=str(r.get("Campagne","")).upper()
            for k,v in VIS_TOP_COLOR.items():
                if k in camp: return v
            return "#CCCCCC"

        if ligne=="Imprimerie":
            return "#FFFFFF"

        return "#FFFFFF"

    # --- boucle OF ---
    for _,r0 in ofs_df.iterrows():
        curr_key=key_of(r0)

        need_intro=False
        if ligne=="Ligne 1": need_intro=l1_needs_intro(prev_key,curr_key)
        elif ligne=="Ligne 2": need_intro=l2_needs_intro(prev_key,curr_key)

        if need_intro and intro_d>0:
            intro_segs,_=consume(intro_d)
            for s in intro_segs:
                rows.append({
                    "Ligne":ligne,"Ofs":"INTRO","is_intro":True,
                    "start":s["start"],"end":s["end"],"duree_h":s["duree_h"],
                    "Label":"<b>INTRO</b>",
                    "color":"#FFFFFF","text_color":"#000000",
                })

        prev_key=curr_key

        duree=r0.get("Temps en h")
        if pd.isna(duree) or duree in [None,"",0]:
            if ligne=="Visitage":
                ml=float(r0.get("Ml",0))
                duree=(ml/15/60*0.8)+0.75
            else: duree=1.0
        duree=float(duree)

        segs,_=consume(duree)
        col=color_of(r0)
        labelf=LABEL_FUN[ligne]

        for s in segs:
            r=r0.copy()
            r["start"]=s["start"]; r["end"]=s["end"]; r["duree_h"]=s["duree_h"]

            ml = r.get("Ml", r.get("ML",""))
            laise = r.get("Laise","")
            colU = r.get("COLORIS",""); colL = r.get("Coloris","")
            if colU=="": colAny=colL
            else: colAny=colU

            rows.append({
                "Ligne":ligne,
                "Ofs":r.get("Ofs",""),
                "is_intro":False,
                "start":r["start"],"end":r["end"],"duree_h":r["duree_h"],
                "color":col,"text_color":"#000000",
                "Label":labelf(r),
                "Produit":r.get("Produit",""),
                "Ml":ml,
                "Laise":laise,
                "COLORIS":colU,
                "Coloris":colAny,
                "Support":r.get("Support",""),
                "FAMILLE":r.get("FAMILLE",""),
                "GRAIN":r.get("GRAIN",""),
                "Campagne":r.get("Campagne",""),
            })

    return pd.DataFrame(rows)



# ============================================================
#              AFFICHAGE GLOBAL STREAMLIT
# ============================================================
def show_planning_global():

    st.title("📅 Planning Global Harmonisé — Vue Semaine")

    cb,_=st.columns([1,5])
    with cb:
        if st.button("⬅️ Retour Menu"):
            st.session_state["page"]="menu"
            st.rerun()

    data=load_all()
    now=datetime.now()

    st.markdown("### 📆 Sélection de la semaine")
    c1,c2,_=st.columns([1,2,1])

    with c1:
        offset=st.number_input("Décalage semaine",0,10,0)

    def week_bounds(ref,off):
        mon=ref - timedelta(days=ref.weekday())
        mon=datetime.combine(mon.date(),datetime.min.time())+timedelta(weeks=off)
        sun=mon+timedelta(days=6, hours=23, minutes=59)
        return mon,sun

    ws,we = week_bounds(now,offset)

    with c2:
        st.info(f"**{ws.strftime('%d/%m')} → {we.strftime('%d/%m/%Y')}**")

    lignes=["Ligne 1","Imprimerie","Ligne 2","Visitage"]
    planning={}

    # -----------------------------------------------------
    # PLANIFICATION PAR LIGNE
    # -----------------------------------------------------
    for ligne in lignes:
        if ligne not in data or "error" in data[ligne] or data[ligne]["ofs"].empty:
            continue

        if offset==0:
            slots=build_slots(data[ligne]["cal"], now, we)
            planning[ligne]=schedule_generic(data[ligne]["ofs"],slots,ligne)
        else:
            slots_before=build_slots(data[ligne]["cal"], now, ws)
            p_before=schedule_generic(data[ligne]["ofs"],slots_before,ligne)

            done=set(p_before[~p_before["is_intro"]]["Ofs"].unique()) if not p_before.empty else set()

            remaining=data[ligne]["ofs"][~data[ligne]["ofs"]["Ofs"].isin(done)]
            slots=build_slots(data[ligne]["cal"], ws, we)
            planning[ligne]=schedule_generic(remaining,slots,ligne)

    # -----------------------------------------------------
    # GANTT
    # -----------------------------------------------------
    fig=go.Figure()
    display_start = now if offset==0 else ws

    for ligne in lignes:
        if ligne not in planning or planning[ligne].empty:
            continue

        df=planning[ligne]

        # --- BARRES ---
        for _,r in df.iterrows():
            if r["end"] < display_start or r["start"] > we:
                continue
            dur_ms=(r["end"]-r["start"]).total_seconds()*1000

            bar_col = r["color"]
            if ligne=="Imprimerie":
                bar_col="#FFFFFF"

            fig.add_trace(go.Bar(
                x=[dur_ms], y=[ligne], base=[r["start"]],
                orientation='h',
                marker=dict(color=bar_col, line=dict(color="#000",width=1)),
                text=r["Label"], textposition="inside",
                insidetextanchor="middle",
                textfont=dict(color=r["text_color"],size=9),
                hovertemplate=f"{r['Label']}<extra></extra>",
                showlegend=False
            ))

        # --- TRAITS IMPRIMERIE ---
        if ligne=="Imprimerie":
            for _,r in df.iterrows():
                if r["is_intro"]: continue
                s,e = r["start"],r["end"]
                if e<display_start or s>we: continue

                camp=str(r["Campagne"]).upper()
                tcol="#000000"
                double=False

                for k,v in CAMPAGNE_COLOR_IMP.items():
                    if k in camp: tcol=v

                for dt in DOUBLE_TRAIT:
                    if dt in camp: double=True

                if double:
                    fig.add_shape(type="line", x0=s,x1=e, y0=0.33,y1=0.33,
                        xref="x",yref="paper", line=dict(color=tcol,width=3))
                    fig.add_shape(type="line", x0=s,x1=e, y0=0.67,y1=0.67,
                        xref="x",yref="paper", line=dict(color=tcol,width=3))
                else:
                    fig.add_shape(type="line", x0=s,x1=e, y0=0.5,y1=0.5,
                        xref="x",yref="paper", line=dict(color=tcol,width=3))

        # --- VISITAGE : LAISE ---
        if ligne=="Visitage":
            for _,r in df.iterrows():
                if r["is_intro"]: continue
                s,e=r["start"],r["end"]
                if e<display_start or s>we: continue

                dur_ms=(e-s).total_seconds()*1000
                la=int(r.get("Laise",4))
                bot=VIS_BOTTOM_COLOR.get(la,"#55CC55")

                fig.add_trace(go.Bar(
                    x=[dur_ms], y=["Laise"], base=[s],
                    orientation='h',
                    marker=dict(color=bot, line=dict(color="#000",width=1)),
                    text=f"L{la}",
                    textposition="inside",
                    textfont=dict(color="#000",size=10),
                    hovertemplate=f"Laise {la}<extra></extra>",
                    showlegend=False
                ))

        # --- ARRETS FERMÉS ---
        cal=data[ligne]["cal"].copy()
        cal["Jour"]=pd.to_datetime(cal["Jour"])
        mask=(cal["Jour"].dt.date>=display_start.date())&(cal["Jour"].dt.date<=we.date())
        for _,row in cal.loc[mask].iterrows():
            j=row["Jour"]
            for i in [1,2,3]:
                if row.get(f"Etat_{i}")=="FERME":
                    h=row.get(f"Horaire_{i}")
                    if isinstance(h,str) and "-" in h:
                        s,e=parse_horaire(j,h)
                        if s<we and e>display_start:
                            s=max(s,display_start); e=min(e,we)
                            dur_ms=(e-s).total_seconds()*1000
                            fig.add_trace(go.Bar(
                                x=[dur_ms], y=[ligne], base=[s],
                                orientation='h',
                                marker=dict(color="rgba(255,0,0,0.75)"),
                                text="ARRÊT", textposition="inside",
                                textfont=dict(size=9,color="white",family="Arial Black"),
                                hovertemplate="ARRÊT<extra></extra>",
                                showlegend=False
                            ))

    # --- JOURS ---
    JFR=["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"]
    d=display_start.date()
    while d<=we.date():
        mid=datetime.combine(d,datetime.min.time())+timedelta(hours=12)
        fig.add_annotation(
            x=mid,y=1.05,xref="x",yref="paper",
            text=f"<b>{JFR[d.weekday()]} {d.strftime('%d/%m')}</b>",
            showarrow=False,font=dict(color="white",size=11)
        )
        d+=timedelta(days=1)

    if offset==0:
        fig.add_vline(x=now,line_color="yellow",line_width=2,line_dash="dot")

    fig.update_xaxes(type="date",range=[display_start,we],tickformat="%Hh",dtick=3600000*4)
    fig.update_yaxes(categoryorder="array",
                     categoryarray=list(reversed(lignes))+["Laise"],
                     tickfont=dict(size=12,color="white"))

    fig.update_layout(height=520,margin=dict(l=100,r=20,t=80,b=40),
                      plot_bgcolor="#444",paper_bgcolor="#444",
                      font=dict(color="white"),barmode="overlay",showlegend=False)

    st.plotly_chart(fig,use_container_width=True)

    # -----------------------------------------------------
    # RÉSUMÉ PAR LIGNE
    # -----------------------------------------------------
    st.markdown("### 📊 Résumé par ligne")
    c=st.columns(4)

    for i,ligne in enumerate(lignes):
        with c[i]:
            if ligne not in planning or planning[ligne].empty:
                st.metric(ligne,"0 OF","—")
            else:
                df=planning[ligne]
                tot=df["duree_h"].sum()
                nb=df[~df["is_intro"]]["Ofs"].nunique()
                ni=len(df[df["is_intro"]])
                st.metric(ligne,f"{nb} OFs",f"{tot:.1f}h total")
                if ni>0: st.caption(f"+ {ni} INTRO")

    st.divider()

    # -----------------------------------------------------
    # KPIs
    # -----------------------------------------------------
    st.markdown("### 📈 Indicateurs clés de la semaine")
    c1,c2,c3,c4=st.columns(4)

    # ML MOYEN IMPRIMERIE
    with c1:
        if "Imprimerie" in planning and not planning["Imprimerie"].empty:
            df=planning["Imprimerie"]
            uniq=df[~df["is_intro"]].drop_duplicates(subset=["Ofs"])
            ml = pd.to_numeric(uniq["Ml"],errors="coerce").sum()
            nb=len(uniq)
            st.metric("📏 ML moyen Imprimerie",f"{(ml/nb) if nb>0 else 0:,.0f} ML")
        else:
            st.metric("📏 ML moyen Imprimerie","N/A")

    # GRAINEURS L2
    with c2:
        if "Ligne 2" in planning and not planning["Ligne 2"].empty:
            df=planning["Ligne 2"]
            grains=[g for g in df["GRAIN"].unique() if str(g).strip()]
            st.metric("🔧 Graineurs L2",f"{len(grains)}")
        else:
            st.metric("🔧 Graineurs L2","N/A")

    # CICD01 L1
    with c3:
        if "Ligne 1" in planning and not planning["Ligne 1"].empty:
            df=planning["Ligne 1"]
            d1=df[df["Produit"].str.contains("CICD01",na=False)]
            uniq=d1.drop_duplicates(subset=["Ofs"])
            ml=pd.to_numeric(uniq["Ml"],errors="coerce").sum()
            st.metric("🏭 CICD01 posé (L1)",f"{ml:,.0f} ML")
        else:
            st.metric("🏭 CICD01 posé (L1)","N/A")

    # CICD04 L1
    with c4:
        if "Ligne 1" in planning and not planning["Ligne 1"].empty:
            df=planning["Ligne 1"]
            d1=df[df["Produit"].str.contains("CICD04",na=False)]
            uniq=d1.drop_duplicates(subset=["Ofs"])
            ml=pd.to_numeric(uniq["Ml"],errors="coerce").sum()
            st.metric("🏭 CICD04 posé (L1)",f"{ml:,.0f} ML")
        else:
            st.metric("🏭 CICD04 posé (L1)","N/A")

    # -----------------------------------------------------
    # DÉTAILS
    # -----------------------------------------------------
    with st.expander("📋 Détails par ligne"):
        for l in lignes:
            if l in planning and not planning[l].empty:
                st.markdown(f"**{l}** — {len(planning[l])} segments")
                st.dataframe(
                    planning[l][["Ofs","Produit","Campagne","Ml","start","end","duree_h"]],
                    use_container_width=True
                )