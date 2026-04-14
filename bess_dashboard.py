"""
================================================================================
BESS VALORISATION — Dashboard Streamlit v2
================================================================================
Lancement local :
    streamlit run bess_dashboard.py

Déploiement public (Streamlit Cloud) :
    1. Pousser ce fichier + bess_engine.py + spot_data.csv sur GitHub
    2. https://share.streamlit.io → connecter le repo → déployer
    3. URL publique : https://[pseudo]-bess.streamlit.app
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import io, sys, json
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from bess_engine import (
    load_spot, simulate_arbitrage, aggregate_arbitrage,
    simulate_lissage, HOUR_COLS
)

# ──────────────────────────────────────────────────────────────────────────────
# EXPORT PDF
# ──────────────────────────────────────────────────────────────────────────────

def fig_to_png(fig, width=860, height=340):
    return fig.to_image(format="png", width=width, height=height, scale=2)

def build_pdf(title, subtitle, kpis, figs, table_df):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, Image, HRFlowable)
    from reportlab.lib.enums import TA_CENTER

    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(buf, pagesize=A4,
                              leftMargin=1.8*cm, rightMargin=1.8*cm,
                              topMargin=2*cm, bottomMargin=2*cm)
    W    = A4[0] - 3.6*cm
    BLUE = colors.HexColor("#1F4E79")
    LB   = colors.HexColor("#BDD7EE")
    GRAY = colors.HexColor("#F0F4FA")

    def S(name, **kw):
        return ParagraphStyle(name, **kw)

    st_title   = S("t",  fontSize=18, fontName="Helvetica-Bold",
                   textColor=BLUE, spaceAfter=4)
    st_sub     = S("s",  fontSize=9,  fontName="Helvetica",
                   textColor=colors.HexColor("#555"), spaceAfter=14)
    st_sec     = S("sc", fontSize=12, fontName="Helvetica-Bold",
                   textColor=BLUE, spaceBefore=14, spaceAfter=6)
    st_kpi_v   = S("kv", fontSize=15, fontName="Helvetica-Bold",
                   textColor=BLUE, alignment=TA_CENTER)
    st_kpi_l   = S("kl", fontSize=8,  fontName="Helvetica",
                   textColor=colors.HexColor("#666"), alignment=TA_CENTER)
    st_foot    = S("f",  fontSize=7,  fontName="Helvetica",
                   textColor=colors.HexColor("#999"), alignment=TA_CENTER,
                   spaceBefore=4)

    story = []
    story.append(Paragraph(title, st_title))
    story.append(HRFlowable(width="100%", thickness=2, color=BLUE))
    story.append(Spacer(1, 6))
    story.append(Paragraph(subtitle, st_sub))

    # KPIs
    story.append(Paragraph("Résultats clés", st_sec))
    n = len(kpis)
    kpi_tbl = Table(
        [[Paragraph(v, st_kpi_v) for v in kpis.values()],
         [Paragraph(k, st_kpi_l) for k in kpis.keys()]],
        colWidths=[W/n]*n
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), GRAY),
        ("BOX", (0,0), (-1,-1), 1, LB),
        ("INNERGRID", (0,0), (-1,-1), 0.5, LB),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 10))

    # Graphiques
    for sec, fig in figs:
        story.append(Paragraph(sec, st_sec))
        png = fig_to_png(fig)
        story.append(Image(io.BytesIO(png), width=W, height=W*340/860))
        story.append(Spacer(1, 6))

    # Tableau
    if table_df is not None and len(table_df):
        story.append(Paragraph("Récapitulatif annuel", st_sec))
        cols = table_df.columns.tolist()
        data = [cols] + table_df.values.tolist()
        cw   = W / len(cols)
        tbl  = Table(data, colWidths=[cw]*len(cols))
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), BLUE),
            ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
            ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",   (0,0), (-1,-1), 8),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, GRAY]),
            ("BOX",        (0,0), (-1,-1), 1, LB),
            ("INNERGRID",  (0,0), (-1,-1), 0.3, LB),
            ("ALIGN",      (0,0), (-1,-1), "CENTER"),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(tbl)

    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=LB))
    story.append(Paragraph(
        f"Rapport généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} "
        f"— Plénitude B-Charge | BESS Valorisation v2.0",
        st_foot
    ))
    doc.build(story)
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="BESS Valorisation DA",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  .main-title { font-size:1.85rem; font-weight:700; color:#1F4E79; margin-bottom:0; }
  .sub-title  { font-size:0.9rem; color:#666; margin-top:2px; margin-bottom:1.2rem; }
  .section    { font-size:1.05rem; font-weight:600; color:#1F4E79;
                border-bottom:2px solid #BDD7EE; padding-bottom:3px;
                margin-top:1.2rem; margin-bottom:0.7rem; }
  div[data-testid="stMetricValue"] { font-size:1.5rem !important; }
  div[data-testid="stMetricLabel"] { font-size:0.8rem !important; color:#444 !important; }
  .stDownloadButton > button {
    background-color:#1F4E79 !important; color:white !important;
    border-radius:6px; font-weight:600; border:none; }
  .stDownloadButton > button:hover { background-color:#2E75B6 !important; }
</style>
""", unsafe_allow_html=True)

COLORS = ["#1F4E79","#2E75B6","#ED7D31","#A9D18E","#FF6B6B","#7030A0"]

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## BESS Valorisation")
    st.markdown("---")
    st.markdown("### Données")

    # Détection mode cloud (fichier CSV embarqué) vs local (Excel)
    CSV_PATH = Path(__file__).parent / "spot_data.csv"
    EXCEL_DEFAULT = r"C:\Users\nicolas.pissavin\Documents\Duplication projet batteries\BESS_valorisation_RESULTATS - Copie.xlsx"

    if CSV_PATH.exists():
        st.info("Mode cloud — données intégrées")
        fichier  = str(CSV_PATH)
        use_csv  = True
    else:
        fichier  = st.text_input("Chemin fichier Excel", value=EXCEL_DEFAULT)
        use_csv  = False

    st.markdown("### Paramètres batterie")
    energy_MWh = st.number_input("Capacité (MWh)",     0.1, 100.0, 2.0,  0.1)
    power_MW   = st.number_input("Puissance max (MW)", 0.1,  50.0, 0.43, 0.01, format="%.3f")
    efficiency = st.slider("Rendement (%)", 70, 100, 92) / 100
    max_cycles = st.number_input("Max cycles/an (usure)", 0, 365, 300)
    max_cycles = max_cycles if max_cycles > 0 else None

    st.markdown("### Mode de valorisation")
    mode = st.radio("", ["Arbitrage Day-Ahead", "Lissage de charge"],
                    label_visibility="collapsed")
    st.markdown("---")
    st.caption("Plénitude B-Charge | BESS Valorisation v2.0")

# ──────────────────────────────────────────────────────────────────────────────
# CHARGEMENT DONNÉES
# ──────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Chargement des données...")
def get_pivot(path, is_csv):
    if is_csv:
        df = pd.read_csv(path)
        for c in ["annee","mois","jour"]:
            df[c] = df[c].astype(int)
        df["date"]         = pd.to_datetime(df["date"])
        df["weekday"]      = df["date"].dt.weekday
        df["weekday_name"] = df["date"].dt.strftime("%A")
        return df.sort_values("date").reset_index(drop=True)
    return load_spot(path)

try:
    pivot  = get_pivot(fichier, use_csv)
    annees = sorted(pivot["annee"].unique())
except Exception as e:
    st.error(f"Impossible de lire le fichier : {e}")
    st.info("Vérifiez le chemin dans la sidebar.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# EN-TETE
# ──────────────────────────────────────────────────────────────────────────────

st.markdown('<p class="main-title">BESS Valorisation — Marché Day-Ahead</p>',
            unsafe_allow_html=True)
st.markdown(f'<p class="sub-title">{len(pivot):,} jours | {annees[0]}–{annees[-1]} | '
            f'Batterie {energy_MWh} MWh / {power_MW} MW | '
            f'Rendement {efficiency*100:.0f}%</p>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MODE 1 — ARBITRAGE DA
# ══════════════════════════════════════════════════════════════════════════════

if mode == "Arbitrage Day-Ahead":

    st.markdown('<p class="section">Paramètres du scénario</p>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        duration_h = st.selectbox("Durée du cycle", [1,2],
                                   format_func=lambda x: f"{x}h", index=1)
    with c2:
        jours_excl = st.multiselect("Jours avec restriction",
            ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"],
            default=["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"])
    with c3:
        h_debut = st.number_input("Heure début restriction", 0, 23, 10)
    with c4:
        h_fin   = st.number_input("Heure fin restriction",   0, 23, 12)

    day_map  = {"Lundi":0,"Mardi":1,"Mercredi":2,"Jeudi":3,
                "Vendredi":4,"Samedi":5,"Dimanche":6}
    days_num = [day_map[d] for d in jours_excl]
    excluded = ({"days":days_num,"hours":list(range(h_debut, h_fin))}
                if days_num and h_debut < h_fin else {})

    @st.cache_data(show_spinner="Simulation en cours...")
    def run_arb(fichier, is_csv, power_MW, duration_h, excl_str,
                efficiency, max_cycles, energy_MWh):
        pv = get_pivot(fichier, is_csv)
        p  = {"power_MW":power_MW, "duration_h":duration_h,
              "excluded_hours":json.loads(excl_str),
              "efficiency":efficiency, "max_cycles_year":max_cycles}
        d  = simulate_arbitrage(pv, p)
        y  = aggregate_arbitrage(d, power_MW)
        return d, y

    daily, yearly = run_arb(fichier, use_csv, power_MW, duration_h,
                             json.dumps(excluded), efficiency, max_cycles, energy_MWh)

    # KPIs
    st.markdown('<p class="section">Résultats globaux</p>', unsafe_allow_html=True)
    total_pnl    = daily["pnl"].sum()
    total_libre  = daily["pnl_libre"].sum()
    jours_actifs = int(daily["valid"].sum())
    jours_tot    = len(daily)
    spread_moy   = daily.loc[daily["valid"],"spread"].mean() if daily["valid"].any() else 0
    jours_usure  = int(daily["usure_bloque"].sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("PnL total (contraintes)", f"{total_pnl:,.0f} €")
    k2.metric("PnL borne max (libre)",   f"{total_libre:,.0f} €",
              delta=f"{total_pnl/total_libre*100:.1f}% du max" if total_libre else None)
    k3.metric("Spread moyen",            f"{spread_moy:.1f} €/MWh")
    k4.metric("Jours actifs",            f"{jours_actifs} / {jours_tot}",
              delta=f"{jours_actifs/jours_tot*100:.0f}% taux d'activation")
    k5.metric("Jours bloqués (usure)",   f"{jours_usure}",
              delta="Aucun" if not jours_usure else f"{jours_usure/jours_tot*100:.1f}%",
              delta_color="inverse")

    # Graphique 1 — PnL annuel
    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=yearly["annee"].astype(str), y=yearly["pnl_libre_total"],
                          name="Borne max", marker_color="#BDD7EE",
                          text=yearly["pnl_libre_total"].apply(lambda x: f"{x:,.0f} €"),
                          textposition="outside"))
    fig1.add_trace(go.Bar(x=yearly["annee"].astype(str), y=yearly["pnl_total"],
                          name="PnL réel", marker_color="#1F4E79",
                          text=yearly["pnl_total"].apply(lambda x: f"{x:,.0f} €"),
                          textposition="inside", textfont_color="white"))
    fig1.update_layout(barmode="group", height=340, yaxis_title="PnL (€)",
                       xaxis_title="Année", legend=dict(orientation="h", y=1.1),
                       margin=dict(t=20,b=40), yaxis_tickformat=",",
                       plot_bgcolor="white", paper_bgcolor="white")
    fig1.update_yaxes(gridcolor="#f0f0f0")
    st.plotly_chart(fig1, width='stretch')

    # Graphiques 2 & 3
    ca, cb = st.columns(2)
    with ca:
        st.markdown('<p class="section">Spread moyen mensuel (€/MWh)</p>', unsafe_allow_html=True)
        monthly = (daily[daily["valid"]]
                   .groupby(["annee","mois"])["spread"].mean().reset_index())
        monthly["label"] = (monthly["annee"].astype(str)+"-"
                            +monthly["mois"].astype(str).str.zfill(2))
        fig2 = go.Figure()
        for i, yr in enumerate(sorted(monthly["annee"].unique())):
            d = monthly[monthly["annee"]==yr]
            fig2.add_trace(go.Scatter(x=d["label"], y=d["spread"],
                                      mode="lines+markers", name=str(int(yr)),
                                      line=dict(width=2, color=COLORS[i]),
                                      marker=dict(size=5)))
        fig2.update_layout(height=300, margin=dict(t=10,b=40),
                           yaxis_title="Spread (€/MWh)",
                           plot_bgcolor="white", paper_bgcolor="white",
                           legend=dict(orientation="h", y=1.1))
        fig2.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig2, width='stretch')

    with cb:
        st.markdown('<p class="section">Profil horaire — fréquence charge/décharge</p>',
                    unsafe_allow_html=True)
        h_ch  = {h:0 for h in range(24)}
        h_dch = {h:0 for h in range(24)}
        for hc, hd in zip(daily.loc[daily["valid"],"h_charge"],
                          daily.loc[daily["valid"],"h_decharge"]):
            for h in hc: h_ch[h]  += 1
            for h in hd: h_dch[h] += 1
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(x=list(range(24)), y=list(h_ch.values()),
                              name="Charge (achat)", marker_color="#2E75B6"))
        fig3.add_trace(go.Bar(x=list(range(24)), y=list(h_dch.values()),
                              name="Décharge (vente)", marker_color="#ED7D31"))
        fig3.update_layout(height=300, barmode="group", margin=dict(t=10,b=40),
                           xaxis_title="Heure", yaxis_title="Nb jours",
                           plot_bgcolor="white", paper_bgcolor="white",
                           legend=dict(orientation="h", y=1.1),
                           xaxis=dict(tickmode="linear", tick0=0, dtick=2))
        fig3.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig3, width='stretch')

    # Graphiques 4 & 5
    cc, cd = st.columns(2)
    with cc:
        st.markdown('<p class="section">Distribution des spreads journaliers</p>',
                    unsafe_allow_html=True)
        spreads = daily.loc[daily["valid"],"spread"]
        fig4 = go.Figure()
        fig4.add_trace(go.Histogram(x=spreads, nbinsx=40,
                                    marker_color="#1F4E79", opacity=0.8))
        fig4.add_vline(x=spreads.mean(), line_dash="dash", line_color="#ED7D31",
                       annotation_text=f"Moy: {spreads.mean():.1f} €/MWh")
        fig4.update_layout(height=300, margin=dict(t=10,b=40),
                           xaxis_title="Spread (€/MWh)", yaxis_title="Nb jours",
                           plot_bgcolor="white", paper_bgcolor="white", showlegend=False)
        fig4.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig4, width='stretch')

    with cd:
        st.markdown('<p class="section">PnL cumulé dans le temps</p>', unsafe_allow_html=True)
        ds = daily.sort_values("date")
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(x=ds["date"], y=ds["pnl"].cumsum(),
                                  fill="tozeroy", mode="lines",
                                  line=dict(color="#1F4E79", width=2),
                                  fillcolor="rgba(31,78,121,0.10)", name="PnL cumulé"))
        fig5.add_trace(go.Scatter(x=ds["date"], y=ds["pnl_libre"].cumsum(),
                                  mode="lines",
                                  line=dict(color="#BDD7EE", width=1.5, dash="dot"),
                                  name="Borne max cumulée"))
        fig5.update_layout(height=300, margin=dict(t=10,b=40),
                           yaxis_title="PnL cumulé (€)", yaxis_tickformat=",",
                           plot_bgcolor="white", paper_bgcolor="white",
                           legend=dict(orientation="h", y=1.1))
        fig5.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig5, width='stretch')

    # Tableau récap
    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
    df_show = yearly.copy()
    df_show["pnl_libre_total"] = df_show["pnl_libre_total"].apply(lambda x: f"{x:,.0f} €")
    df_show["pnl_total"]       = df_show["pnl_total"].apply(lambda x: f"{x:,.0f} €")
    df_show["pnl_par_MW"]      = df_show["pnl_par_MW"].apply(lambda x: f"{x:.1f} €/MW")
    df_show["taux_activation"] = df_show["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
    df_show["spread_libre_moy"]= df_show["spread_libre_moy"].apply(lambda x: f"{x:.1f}")
    df_show["spread_moy"]      = df_show["spread_moy"].apply(lambda x: f"{x:.1f}")
    df_show.columns = ["Année","Jours simulés","Jours actifs","Bloqués usure",
                       "Taux activation","Spread libre","PnL libre",
                       "Spread moy.","PnL réel","PnL/MW"]
    st.dataframe(df_show, width='stretch', hide_index=True)

    # Export
    st.markdown('<p class="section">Export</p>', unsafe_allow_html=True)
    e1, e2, _ = st.columns([1,1,4])

    with e1:
        st.download_button("Télécharger CSV",
                           data=daily.to_csv(index=False).encode(),
                           file_name="BESS_detail_journalier.csv",
                           mime="text/csv")
    with e2:
        if st.button("Générer rapport PDF"):
            with st.spinner("Génération du PDF..."):
                kpis_pdf = {
                    "PnL total"       : f"{total_pnl:,.0f} €",
                    "PnL borne max"   : f"{total_libre:,.0f} €",
                    "Spread moyen"    : f"{spread_moy:.1f} €/MWh",
                    "Jours actifs"    : f"{jours_actifs}/{jours_tot}",
                    "Taux activation" : f"{jours_actifs/jours_tot*100:.0f}%",
                }
                sub = (f"{len(pivot):,} jours | {annees[0]}–{annees[-1]} | "
                       f"{energy_MWh} MWh / {power_MW} MW | "
                       f"Rendement {efficiency*100:.0f}% | Cycle {duration_h}h")
                pdf = build_pdf("BESS Valorisation — Marché Day-Ahead", sub,
                                kpis_pdf,
                                [("PnL annuel — borne max vs réel", fig1),
                                 ("Spread moyen mensuel", fig2),
                                 ("Profil horaire charge/décharge", fig3),
                                 ("Distribution des spreads", fig4),
                                 ("PnL cumulé", fig5)],
                                df_show)
            st.download_button("Télécharger le PDF",
                               data=pdf, file_name="BESS_rapport.pdf",
                               mime="application/pdf")

    # Explorateur journalier
    with st.expander("Explorer un jour spécifique"):
        ce1, ce2 = st.columns(2)
        with ce1:
            date_sel = st.date_input("Date",
                                      value=pd.Timestamp(pivot["date"].iloc[0]).date())
        row_s = daily[daily["date"].dt.date == date_sel]
        if not row_s.empty:
            r = row_s.iloc[0]
            prix_j = pivot[pivot["date"].dt.date == date_sel][HOUR_COLS].values.flatten()
            with ce2:
                st.metric("Spread", f"{r['spread']:.2f} €/MWh")
                st.metric("PnL",    f"{r['pnl']:.2f} €")
                st.metric("Trade valide", "Oui" if r["valid"] else "Non")
            fig_d = go.Figure()
            fig_d.add_trace(go.Bar(x=list(range(24)), y=prix_j.tolist(),
                                   marker_color="#BDD7EE", name="Prix spot"))
            if r["valid"]:
                fig_d.add_trace(go.Scatter(
                    x=r["h_charge"], y=prix_j[r["h_charge"]], mode="markers",
                    marker=dict(color="#2E75B6", size=12, symbol="triangle-up"),
                    name=f"Charge ({duration_h}h — achat)"))
                fig_d.add_trace(go.Scatter(
                    x=r["h_decharge"], y=prix_j[r["h_decharge"]], mode="markers",
                    marker=dict(color="#ED7D31", size=12, symbol="triangle-down"),
                    name=f"Décharge ({duration_h}h — vente)"))
            fig_d.update_layout(height=320, margin=dict(t=10,b=40),
                                xaxis_title="Heure", yaxis_title="Prix (€/MWh)",
                                plot_bgcolor="white", paper_bgcolor="white",
                                legend=dict(orientation="h", y=1.1),
                                xaxis=dict(tickmode="linear", tick0=0, dtick=2))
            fig_d.update_yaxes(gridcolor="#f0f0f0")
            st.plotly_chart(fig_d, width='stretch')

# ══════════════════════════════════════════════════════════════════════════════
# MODE 2 — LISSAGE
# ══════════════════════════════════════════════════════════════════════════════

else:
    st.markdown('<p class="section">Paramètres du lissage</p>', unsafe_allow_html=True)
    cl1, cl2 = st.columns([2,1])

    with cl2:
        st.markdown("**Paramètres économiques**")
        seuil_pct = st.slider("Seuil d'écrêtage (percentile)", 50, 95, 75)
        tarif_kw  = st.number_input("Tarif puissance souscrite (€/MW/an)",
                                     1000, 100000, 12000, 1000)
        soc_init  = st.slider("SOC initial (%)", 10, 90, 50)

    with cl1:
        st.markdown("**Profil de consommation client (MW par heure)**")
        profil_def = [0.3,0.3,0.3,0.3,0.3,0.4,0.6,0.9,1.1,1.2,
                      1.3,1.2,1.0,1.1,1.2,1.3,1.2,1.0,0.9,1.4,
                      1.5,1.2,0.7,0.4]
        cols_in = st.columns(8)
        profil  = []
        for h in range(24):
            with cols_in[h % 8]:
                profil.append(st.number_input(f"H{h:02d}", 0.0, 100.0,
                                               float(profil_def[h]), 0.05,
                                               format="%.2f"))

    profil_arr = np.array(profil)
    params_l   = {"power_MW":power_MW, "energy_MWh":energy_MWh,
                  "soc_min_pct":0.10, "soc_max_pct":0.90,
                  "efficiency":efficiency, "seuil_percentile":seuil_pct,
                  "tarif_puissance_souscrite":tarif_kw, "soc_init":soc_init/100}
    res  = simulate_lissage(pivot, profil_arr, params_l)
    jour = res["jour_type"]

    # KPIs
    st.markdown('<p class="section">Résultats du lissage</p>', unsafe_allow_html=True)
    lk1,lk2,lk3,lk4 = st.columns(4)
    lk1.metric("Réduction de pointe", f"{res['reduction_MW']:.3f} MW",
               delta=f"-{res['reduction_MW']/jour['pointe_avant']*100:.1f}%")
    lk2.metric("Économie annuelle estimée", f"{res['economie_an']:,.0f} €")
    lk3.metric("Pointe avant / après",
               f"{jour['pointe_avant']:.2f} → {jour['pointe_apres']:.2f} MW")
    lk4.metric("Seuil d'écrêtage", f"{res['seuil_MW']:.2f} MW",
               delta=f"Percentile {seuil_pct}%")

    # Graphique principal
    st.markdown('<p class="section">Profil de consommation — avant et après lissage</p>',
                unsafe_allow_html=True)
    actions_bess = [v if a=="charge" else (-v if a=="decharge" else 0)
                    for a, v in jour["actions"]]
    fig_l = make_subplots(rows=2, cols=1, shared_xaxes=True,
                          row_heights=[0.65,0.35],
                          subplot_titles=["Consommation client (MW)",
                                          "Action BESS — + charge / − décharge (MW)"])
    fig_l.add_trace(go.Scatter(x=list(range(24)), y=profil_arr.tolist(),
                               mode="lines+markers", name="Consommation originale",
                               line=dict(color="#ED7D31", width=2.5, dash="dot"),
                               marker=dict(size=5)), row=1, col=1)
    fig_l.add_trace(go.Scatter(x=list(range(24)), y=jour["profil_lisse"].tolist(),
                               mode="lines+markers", name="Consommation lissée",
                               line=dict(color="#1F4E79", width=2.5),
                               fill="tozeroy", fillcolor="rgba(31,78,121,0.08)",
                               marker=dict(size=5)), row=1, col=1)
    fig_l.add_hline(y=res["seuil_MW"], line_dash="dash", line_color="#A9D18E",
                    line_width=2, annotation_text=f"Seuil {res['seuil_MW']:.2f} MW",
                    annotation_position="top right", row=1, col=1)
    fig_l.add_trace(go.Bar(x=list(range(24)), y=actions_bess,
                           marker_color=["#2E75B6" if v>=0 else "#ED7D31"
                                         for v in actions_bess],
                           name="Action BESS"), row=2, col=1)
    fig_l.update_layout(height=500, margin=dict(t=40,b=40),
                        plot_bgcolor="white", paper_bgcolor="white",
                        legend=dict(orientation="h", y=1.05),
                        xaxis2=dict(tickmode="linear", tick0=0, dtick=2, title="Heure"),
                        yaxis2=dict(title="MW"))
    fig_l.update_yaxes(gridcolor="#f0f0f0")
    st.plotly_chart(fig_l, width='stretch')

    # SOC + Projection
    ls1, ls2 = st.columns(2)
    with ls1:
        st.markdown('<p class="section">État de charge (SOC) batterie</p>', unsafe_allow_html=True)
        fig_soc = go.Figure()
        fig_soc.add_trace(go.Scatter(x=list(range(25)), y=jour["soc_hist"],
                                     mode="lines+markers",
                                     line=dict(color="#2E75B6", width=2),
                                     fill="tozeroy",
                                     fillcolor="rgba(46,117,182,0.12)"))
        fig_soc.add_hline(y=energy_MWh*0.9, line_dash="dash",
                          line_color="#A9D18E", annotation_text="SOC max")
        fig_soc.add_hline(y=energy_MWh*0.1, line_dash="dash",
                          line_color="#FF6B6B", annotation_text="SOC min")
        fig_soc.update_layout(height=280, margin=dict(t=10,b=40),
                              xaxis_title="Heure", yaxis_title="SOC (MWh)",
                              plot_bgcolor="white", paper_bgcolor="white",
                              xaxis=dict(tickmode="linear", tick0=0, dtick=2),
                              yaxis=dict(range=[0, energy_MWh*1.1]), showlegend=False)
        fig_soc.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig_soc, width='stretch')

    with ls2:
        st.markdown('<p class="section">Projection économique annuelle</p>', unsafe_allow_html=True)
        yl = res["yearly"]
        fig_eco = go.Figure()
        fig_eco.add_trace(go.Bar(x=yl["annee"].astype(str), y=yl["economie_an"],
                                 marker_color="#1F4E79",
                                 text=yl["economie_an"].apply(lambda x: f"{x:,.0f} €"),
                                 textposition="outside"))
        fig_eco.update_layout(height=280, margin=dict(t=10,b=40),
                              xaxis_title="Année", yaxis_title="Économie (€)",
                              plot_bgcolor="white", paper_bgcolor="white",
                              yaxis_tickformat=",", showlegend=False)
        fig_eco.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig_eco, width='stretch')

    with st.expander("Détail heure par heure"):
        detail = pd.DataFrame({
            "Heure"                : [f"H{h:02d}" for h in range(24)],
            "Conso originale (MW)" : profil_arr.round(3),
            "Action BESS"          : [f"{a} {v:.3f} MW" for a,v in jour["actions"]],
            "Conso lissée (MW)"    : jour["profil_lisse"].round(3),
            "SOC après (MWh)"      : jour["soc_hist"][1:],
        })
        st.dataframe(detail, width='stretch', hide_index=True)
        st.download_button("Télécharger CSV",
                           data=detail.to_csv(index=False).encode(),
                           file_name="BESS_lissage_detail.csv",
                           mime="text/csv")

    st.info(
        f"**Méthodologie :** Seuil au percentile {seuil_pct}% = {res['seuil_MW']:.2f} MW. "
        f"La batterie décharge quand la conso dépasse ce seuil et charge dans les creux. "
        f"Économie = {res['reduction_MW']:.3f} MW × {tarif_kw:,} €/MW/an "
        f"= **{res['economie_an']:,.0f} €/an**."
    )