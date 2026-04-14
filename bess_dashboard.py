"""
================================================================================
BESS VALORISATION — Dashboard Streamlit
================================================================================
Déploiement public : streamlit run bess_dashboard.py
Streamlit Cloud   : pointer vers ce fichier sur GitHub

Modes :
  1. Arbitrage DA  — achat heures creuses / vente heures de pointe
  2. Lissage       — écrêtage des pics de consommation client
================================================================================
"""

import io
import json
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

from bess_engine import (
    HOUR_COLS, aggregate_arbitrage, load_spot,
    simulate_arbitrage, simulate_lissage,
)

# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="BESS Valorisation DA",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #f8f9fb; }
  [data-testid="stSidebar"]          { background: #ffffff; border-right: 1px solid #e8ecf0; }
  .main-title  { font-size:1.9rem; font-weight:700; color:#1a3a5c; margin-bottom:0; letter-spacing:-0.5px; }
  .sub-title   { font-size:0.92rem; color:#6b7a8d; margin-top:2px; margin-bottom:1.4rem; }
  .section     { font-size:1.05rem; font-weight:600; color:#1a3a5c;
                 border-bottom:2px solid #d0dff0; padding-bottom:4px;
                 margin-top:1.4rem; margin-bottom:0.7rem; }
  div[data-testid="stMetricValue"] { font-size:1.55rem !important; font-weight:700; color:#1a3a5c; }
  div[data-testid="stMetricLabel"] { font-size:0.8rem !important; color:#6b7a8d; font-weight:500; }
  .stButton > button { background:#1a3a5c; color:white; border:none;
                        border-radius:6px; padding:8px 20px; font-weight:600; }
  .stButton > button:hover { background:#2e5d8e; }
  .upload-box { border:2px dashed #b0c4de; border-radius:10px;
                padding:20px; text-align:center; background:#f0f5fb;
                margin-bottom:12px; }
</style>
""", unsafe_allow_html=True)

BLUE    = "#1a3a5c"
LBLUE   = "#2e75b6"
ORANGE  = "#d46b1a"
GREEN   = "#3a8a5c"
LGREY   = "#e8ecf0"
COLORS  = [BLUE, LBLUE, ORANGE, GREEN, "#7b3fa0", "#b05050"]


# ──────────────────────────────────────────────────────────────────────────────
# EXPORT PDF
# ──────────────────────────────────────────────────────────────────────────────

def _make_table(data, col_widths, col_colors=None):
    """Helper : crée un Table ReportLab stylé."""
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle

    t = Table(data, colWidths=col_widths)
    style = [
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#1a3a5c")),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1),
         [colors.HexColor("#f0f5fb"), colors.white]),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#c0d0e0")),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("ALIGN",         (1, 1), (-1, -1), "RIGHT"),
    ]
    t.setStyle(TableStyle(style))
    return t


def build_pdf_arbitrage(params_txt, kpis, yearly_df, daily_df,
                         h_charge_freq, h_decharge_freq, date_str):
    """Génère un rapport PDF complet pour le mode Arbitrage."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    S = {
        "title": ParagraphStyle("T",  parent=styles["Heading1"], fontSize=17,
                                 textColor=colors.HexColor("#1a3a5c"), spaceAfter=2),
        "sub":   ParagraphStyle("S",  parent=styles["Normal"],   fontSize=9,
                                 textColor=colors.HexColor("#6b7a8d"), spaceAfter=10),
        "h2":    ParagraphStyle("H2", parent=styles["Heading2"], fontSize=12,
                                 textColor=colors.HexColor("#1a3a5c"),
                                 spaceBefore=12, spaceAfter=4),
        "body":  ParagraphStyle("B",  parent=styles["Normal"],   fontSize=8,
                                 textColor=colors.HexColor("#333"), spaceAfter=4),
        "foot":  ParagraphStyle("F",  parent=styles["Normal"],   fontSize=7,
                                 textColor=colors.HexColor("#aaa")),
    }
    story = []

    # ── En-tête ───────────────────────────────────────────────────────────────
    story.append(Paragraph("BESS Valorisation — Marché Day-Ahead", S["title"]))
    story.append(Paragraph(
        f"Rapport Arbitrage Day-Ahead | Généré le {date_str}", S["sub"]))
    story.append(Spacer(1, 0.2*cm))

    # ── Paramètres ────────────────────────────────────────────────────────────
    story.append(Paragraph("1. Paramètres de simulation", S["h2"]))
    for line in params_txt:
        story.append(Paragraph(f"• {line}", S["body"]))
    story.append(Spacer(1, 0.3*cm))

    # ── KPIs ─────────────────────────────────────────────────────────────────
    story.append(Paragraph("2. Résultats clés (toutes années)", S["h2"]))
    kpi_data = [["Indicateur", "Valeur"]] + [[k, v] for k, v in kpis]
    story.append(_make_table(kpi_data, [10*cm, 5*cm]))
    story.append(Spacer(1, 0.4*cm))

    # ── Récapitulatif annuel ──────────────────────────────────────────────────
    story.append(Paragraph("3. Récapitulatif annuel", S["h2"]))
    yr = yearly_df.copy()
    yr_data = [["Année", "Jours actifs", "Taux activ.", "Spread libre\n(€/MWh)",
                 "PnL borne max\n(€)", "Spread contraint\n(€/MWh)",
                 "PnL réel\n(€)", "PnL/MW\n(€/MW)"]]
    for _, r in yr.iterrows():
        yr_data.append([
            str(int(r["annee"])),
            f"{int(r['jours_actifs'])} / {int(r['jours_simules'])}",
            f"{r['taux_activation']*100:.0f}%",
            f"{r['spread_libre_moy']:.1f}",
            f"{r['pnl_libre_total']:,.0f}",
            f"{r['spread_moy']:.1f}",
            f"{r['pnl_total']:,.0f}",
            f"{r['pnl_par_MW']:.1f}",
        ])
    story.append(_make_table(yr_data,
                              [1.5*cm, 2.5*cm, 1.8*cm, 2*cm, 2.5*cm, 2.5*cm, 2.3*cm, 2*cm]))
    story.append(Spacer(1, 0.4*cm))

    # ── Profil horaire charge/décharge ────────────────────────────────────────
    story.append(Paragraph("4. Profil horaire — fréquence charge / décharge", S["h2"]))
    story.append(Paragraph(
        "Nombre de jours où chaque heure a été utilisée pour charger ou décharger.",
        S["body"]))
    h_data = [["Heure"] + [str(h) for h in range(24)],
              ["Charge (jours)"] + [str(h_charge_freq.get(h, 0)) for h in range(24)],
              ["Décharge (jours)"] + [str(h_decharge_freq.get(h, 0)) for h in range(24)]]
    col_w = [2.5*cm] + [0.6*cm]*24
    story.append(_make_table(h_data, col_w))
    story.append(Spacer(1, 0.4*cm))

    # ── Distribution des spreads par tranche ──────────────────────────────────
    story.append(Paragraph("5. Distribution des spreads journaliers", S["h2"]))
    spreads = daily_df.loc[daily_df["valid"], "spread"].dropna()
    bins = [0, 20, 40, 60, 80, 100, 150, 200, float("inf")]
    labels = ["0-20", "20-40", "40-60", "60-80", "80-100",
              "100-150", "150-200", ">200"]
    dist_data = [["Tranche (€/MWh)", "Nb jours", "% du total"]]
    total_v = len(spreads)
    for i, (lo, hi) in enumerate(zip(bins[:-1], bins[1:])):
        cnt = ((spreads >= lo) & (spreads < hi)).sum()
        dist_data.append([labels[i], str(cnt),
                           f"{cnt/total_v*100:.1f}%" if total_v else "0%"])
    story.append(_make_table(dist_data, [6*cm, 4*cm, 4*cm]))
    story.append(Spacer(1, 0.4*cm))

    # ── PnL cumulé par mois ───────────────────────────────────────────────────
    story.append(Paragraph("6. PnL mensuel", S["h2"]))
    monthly = daily_df.groupby(["annee", "mois"]).agg(
        pnl=("pnl", "sum"), jours_actifs=("valid", "sum")).reset_index()
    months_fr = ["Jan","Fév","Mar","Avr","Mai","Jun",
                  "Jul","Aoû","Sep","Oct","Nov","Déc"]
    m_data = [["Période", "Jours actifs", "PnL (€)", "PnL cumulé (€)"]]
    cumul = 0
    for _, r in monthly.iterrows():
        cumul += r["pnl"]
        m_data.append([
            f"{months_fr[int(r['mois'])-1]} {int(r['annee'])}",
            str(int(r["jours_actifs"])),
            f"{r['pnl']:,.0f}",
            f"{cumul:,.0f}",
        ])
    story.append(_make_table(m_data, [3.5*cm, 3*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.5*cm))

    # ── Pied de page ─────────────────────────────────────────────────────────
    story.append(Paragraph(
        "Plénitude B-Charge — BESS Valorisation v2.0 — Document confidentiel",
        S["foot"]))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def build_pdf_lissage(params_txt, kpis, detail_df, yearly_df, date_str):
    """Génère un rapport PDF complet pour le mode Lissage."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    S = {
        "title": ParagraphStyle("T",  parent=styles["Heading1"], fontSize=17,
                                 textColor=colors.HexColor("#1a3a5c"), spaceAfter=2),
        "sub":   ParagraphStyle("S",  parent=styles["Normal"],   fontSize=9,
                                 textColor=colors.HexColor("#6b7a8d"), spaceAfter=10),
        "h2":    ParagraphStyle("H2", parent=styles["Heading2"], fontSize=12,
                                 textColor=colors.HexColor("#1a3a5c"),
                                 spaceBefore=12, spaceAfter=4),
        "body":  ParagraphStyle("B",  parent=styles["Normal"],   fontSize=8,
                                 textColor=colors.HexColor("#333"), spaceAfter=4),
        "foot":  ParagraphStyle("F",  parent=styles["Normal"],   fontSize=7,
                                 textColor=colors.HexColor("#aaa")),
    }
    story = []

    story.append(Paragraph("BESS Valorisation — Lissage de charge", S["title"]))
    story.append(Paragraph(f"Rapport Lissage | Généré le {date_str}", S["sub"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("1. Paramètres", S["h2"]))
    for line in params_txt:
        story.append(Paragraph(f"• {line}", S["body"]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("2. Résultats clés", S["h2"]))
    kpi_data = [["Indicateur", "Valeur"]] + [[k, v] for k, v in kpis]
    story.append(_make_table(kpi_data, [10*cm, 5*cm]))
    story.append(Spacer(1, 0.4*cm))

    story.append(Paragraph("3. Projection économique annuelle", S["h2"]))
    yr_data = [["Année", "Réduction pointe (MW)",
                 "Pointe avant (MW)", "Pointe après (MW)", "Économie (€)"]]
    for _, r in yearly_df.iterrows():
        yr_data.append([
            str(int(r["annee"])),
            f"{r['reduction_pointe']:.3f}",
            f"{r['pointe_avant_MW']:.2f}",
            f"{r['pointe_apres_MW']:.2f}",
            f"{r['economie_an']:,.0f}",
        ])
    story.append(_make_table(yr_data, [2.5*cm, 3.5*cm, 3.5*cm, 3.5*cm, 3*cm]))
    story.append(Spacer(1, 0.4*cm))

    story.append(Paragraph("4. Profil heure par heure (jour type)", S["h2"]))
    det_data = [["Heure", "Conso originale\n(MW)", "Action BESS",
                  "Puissance BESS\n(MW)", "Conso lissée\n(MW)", "SOC après\n(MWh)"]]
    for _, r in detail_df.iterrows():
        action_parts = str(r["Action BESS"]).split()
        det_data.append([
            str(r["Heure"]),
            f"{float(r['Conso originale (MW)']):.3f}",
            action_parts[0] if action_parts else "idle",
            action_parts[1] if len(action_parts) > 1 else "0",
            f"{float(r['Conso lissée (MW)']):.3f}",
            f"{float(r['SOC après (MWh)']):.3f}",
        ])
    story.append(_make_table(det_data,
                              [1.5*cm, 3*cm, 2.5*cm, 3*cm, 3*cm, 3*cm]))
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph(
        "Plénitude B-Charge — BESS Valorisation v2.0 — Document confidentiel",
        S["foot"]))
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## BESS Valorisation")
    st.markdown("---")

    # ── Upload fichier ────────────────────────────────────────────────────────
    st.markdown("### Données spot")
    uploaded = st.file_uploader(
        "Fichier Excel (Spot_input)",
        type=["xlsx"],
        help="Glisser-déposer le fichier BESS_valorisation_RESULTATS.xlsx"
    )

    # ── Paramètres batterie ───────────────────────────────────────────────────
    st.markdown("### Paramètres batterie")
    energy_MWh = st.number_input("Capacité (MWh)",     0.1, 100.0, 2.0,  0.1)
    power_MW   = st.number_input("Puissance max (MW)", 0.05, 50.0, 0.43, 0.01, format="%.3f")
    efficiency = st.slider("Rendement (%)", 70, 100, 92) / 100
    max_cycles = st.number_input("Max cycles/an (0 = illimité)", 0, 365, 300)
    max_cycles = max_cycles if max_cycles > 0 else None

    # ── Mode ─────────────────────────────────────────────────────────────────
    st.markdown("### Mode de valorisation")
    mode = st.radio("", ["Arbitrage Day-Ahead", "Lissage de charge"],
                    label_visibility="collapsed")

    st.markdown("---")
    st.caption("Plénitude B-Charge | BESS Valorisation v2.0")


# ──────────────────────────────────────────────────────────────────────────────
# CHARGEMENT DONNÉES
# ──────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Lecture du fichier...")
def get_pivot(file_bytes: bytes) -> pd.DataFrame:
    return load_spot(io.BytesIO(file_bytes))


# Écran d'accueil si pas de fichier
if uploaded is None:
    st.markdown('<p class="main-title">BESS Valorisation — Marché Day-Ahead</p>',
                unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Outil de valorisation d\'un système de stockage par batteries</p>',
                unsafe_allow_html=True)

    st.markdown("""
    <div class="upload-box">
        <h3 style="color:#1a3a5c; margin:0">Importez votre fichier de données</h3>
        <p style="color:#6b7a8d; margin:8px 0 0 0">
            Glissez-déposez le fichier Excel <strong>BESS_valorisation_RESULTATS.xlsx</strong>
            dans la sidebar à gauche, ou cliquez sur le bouton d'upload.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        **Mode Arbitrage Day-Ahead**
        - Achat aux heures creuses, vente aux heures de pointe
        - Calcul du spread et PnL journalier
        - Comparaison borne max vs scénario contraint
        - Contraintes : heures exclues, jours, usure batterie
        - Durée de cycle : 1h ou 2h
        """)
    with col2:
        st.markdown("""
        **Mode Lissage de charge**
        - Ecrêtage des pics de consommation client
        - Réduction de la puissance souscrite au réseau
        - Calcul de l'économie annuelle estimée
        - Profil SOC batterie heure par heure
        - Paramétrable : seuil, tarif réseau, profil client
        """)
    st.stop()

# Chargement
try:
    file_bytes = uploaded.read()
    pivot = get_pivot(file_bytes)
    annees = sorted(pivot["annee"].unique())
except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

# En-tête principal
st.markdown('<p class="main-title">BESS Valorisation — Marché Day-Ahead</p>',
            unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-title">Fichier : {uploaded.name} &nbsp;|&nbsp; '
    f'{len(pivot):,} jours &nbsp;|&nbsp; {annees[0]}–{annees[-1]} &nbsp;|&nbsp; '
    f'Batterie {energy_MWh} MWh / {power_MW} MW &nbsp;|&nbsp; '
    f'Rendement {efficiency*100:.0f}%</p>',
    unsafe_allow_html=True
)


# ══════════════════════════════════════════════════════════════════════════════
# MODE 1 — ARBITRAGE DA
# ══════════════════════════════════════════════════════════════════════════════

if mode == "Arbitrage Day-Ahead":

    # ── Paramètres scénario ──────────────────────────────────────────────────
    st.markdown('<p class="section">Paramètres du scénario</p>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        duration_h = st.selectbox(
            "Durée du cycle", [1, 2],
            format_func=lambda x: f"{x}h — {x} heure{'s' if x>1 else ''} achat+vente",
            index=1
        )
    with c2:
        jours_excl = st.multiselect(
            "Jours avec restriction horaire",
            ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"],
            default=["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
        )
    with c3:
        h_debut = st.number_input("Heure début restriction", 0, 23, 10)
    with c4:
        h_fin = st.number_input("Heure fin restriction",   0, 23, 12)

    day_map = {"Lundi":0,"Mardi":1,"Mercredi":2,"Jeudi":3,
               "Vendredi":4,"Samedi":5,"Dimanche":6}
    days_num = [day_map[d] for d in jours_excl]
    excluded = ({"days": days_num, "hours": list(range(h_debut, h_fin))}
                if days_num and h_debut < h_fin else {})

    # ── Simulation ───────────────────────────────────────────────────────────
    @st.cache_data(show_spinner="Simulation...")
    def run_arb(file_bytes, power_MW, duration_h, excl_str, efficiency, max_cycles, energy_MWh):
        pv = get_pivot(file_bytes)
        excl = json.loads(excl_str)
        p = {"power_MW": power_MW, "duration_h": duration_h,
             "excluded_hours": excl, "efficiency": efficiency,
             "max_cycles_year": max_cycles}
        daily = simulate_arbitrage(pv, p)
        yearly = aggregate_arbitrage(daily, power_MW)
        return daily, yearly

    daily, yearly = run_arb(
        file_bytes, power_MW, duration_h,
        json.dumps(excluded), efficiency, max_cycles, energy_MWh
    )

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Résultats globaux</p>', unsafe_allow_html=True)

    total_pnl       = daily["pnl"].sum()
    total_pnl_libre = daily["pnl_libre"].sum()
    spread_moy      = daily.loc[daily["valid"], "spread"].mean() if daily["valid"].any() else 0
    jours_actifs    = int(daily["valid"].sum())
    jours_total     = len(daily)
    jours_usure     = int(daily["usure_bloque"].sum())
    ratio           = total_pnl / total_pnl_libre * 100 if total_pnl_libre else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("PnL total (avec contraintes)",   f"{total_pnl:,.0f} €")
    m2.metric("PnL borne max (sans contrainte)",f"{total_pnl_libre:,.0f} €",
              delta=f"{ratio:.1f}% du max")
    m3.metric("Spread moyen",                   f"{spread_moy:.1f} €/MWh")
    m4.metric("Jours actifs",                   f"{jours_actifs} / {jours_total}",
              delta=f"{jours_actifs/jours_total*100:.0f}% taux activation")
    m5.metric("Jours bloqués (usure)",          f"{jours_usure}",
              delta="Aucun" if jours_usure == 0 else f"{jours_usure/jours_total*100:.1f}%",
              delta_color="off" if jours_usure == 0 else "inverse")

    # Config commune pour tous les graphiques — active toolbar complète
    PLOTLY_CFG = dict(
        scrollZoom=True,
        displayModeBar=True,
        modeBarButtonsToAdd=["drawrect", "eraseshape"],
        modeBarButtonsToRemove=["lasso2d"],
        displaylogo=False,
        toImageButtonOptions=dict(format="png", width=1400, height=700, scale=2),
    )

    # ── Graphique 1 : PnL annuel ─────────────────────────────────────────────
    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=yearly["annee"].astype(str), y=yearly["pnl_libre_total"],
        name="Borne max", marker_color="#c8d8ec",
        text=yearly["pnl_libre_total"].apply(lambda x: f"{x:,.0f} €"),
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Borne max : %{y:,.0f} €<extra></extra>",
    ))
    fig1.add_trace(go.Bar(
        x=yearly["annee"].astype(str), y=yearly["pnl_total"],
        name="PnL réel", marker_color=BLUE,
        text=yearly["pnl_total"].apply(lambda x: f"{x:,.0f} €"),
        textposition="inside", textfont_color="white",
        hovertemplate="<b>%{x}</b><br>PnL réel : %{y:,.0f} €<extra></extra>",
    ))
    fig1.update_layout(
        barmode="group", height=380,
        yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis_title="Année",
        legend=dict(orientation="h", y=1.08, x=0),
        margin=dict(t=10, b=40, l=60, r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    st.plotly_chart(fig1, width="stretch", config=PLOTLY_CFG)

    # ── Graphiques 2×2 ───────────────────────────────────────────────────────
    col_a, col_b = st.columns(2)

    with col_a:
        # Spread HEBDOMADAIRE (plus de points) + tendance
        st.markdown('<p class="section">Spread moyen — par semaine (€/MWh)</p>',
                    unsafe_allow_html=True)
        d_valid = daily[daily["valid"]].copy()
        d_valid["semaine"] = pd.to_datetime(d_valid["date"]).dt.to_period("W").apply(
            lambda r: r.start_time)
        weekly = d_valid.groupby(["annee", "semaine"])["spread"].agg(
            ["mean", "min", "max", "count"]).reset_index()
        weekly.columns = ["annee", "semaine", "spread_moy", "spread_min",
                           "spread_max", "nb_jours"]

        fig2 = go.Figure()
        for i, yr in enumerate(sorted(weekly["annee"].unique())):
            w = weekly[weekly["annee"] == yr].sort_values("semaine")
            # Bande min-max
            fig2.add_trace(go.Scatter(
                x=pd.concat([w["semaine"], w["semaine"].iloc[::-1]]),
                y=pd.concat([w["spread_max"], w["spread_min"].iloc[::-1]]),
                fill="toself",
                fillcolor=f"rgba({int(COLORS[i % len(COLORS)][1:3], 16)},"
                          f"{int(COLORS[i % len(COLORS)][3:5], 16)},"
                          f"{int(COLORS[i % len(COLORS)][5:7], 16)},0.1)",
                line=dict(color="rgba(0,0,0,0)"),
                showlegend=False, hoverinfo="skip",
            ))
            # Courbe moyenne
            fig2.add_trace(go.Scatter(
                x=w["semaine"], y=w["spread_moy"],
                mode="lines+markers", name=str(int(yr)),
                line=dict(width=2, color=COLORS[i % len(COLORS)]),
                marker=dict(size=4),
                hovertemplate=(
                    "<b>Semaine du %{x|%d/%m/%Y}</b><br>"
                    f"Année {int(yr)}<br>"
                    "Spread moy : <b>%{y:.1f} €/MWh</b><br>"
                    "Min : %{customdata[0]:.1f} | Max : %{customdata[1]:.1f}<br>"
                    "Jours actifs : %{customdata[2]}<extra></extra>"
                ),
                customdata=w[["spread_min", "spread_max", "nb_jours"]].values,
            ))
        fig2.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            yaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0"),
            xaxis=dict(title="", tickformat="%b %Y", nticks=12),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", y=1.1),
            hoverlabel=dict(bgcolor="white", font_size=12),
            hovermode="x unified",
        )
        st.plotly_chart(fig2, width="stretch", config=PLOTLY_CFG)

    with col_b:
        st.markdown('<p class="section">Profil horaire charge / décharge</p>',
                    unsafe_allow_html=True)
        h_ch  = {h: 0 for h in range(24)}
        h_dch = {h: 0 for h in range(24)}
        for hc_list, hd_list in zip(daily.loc[daily["valid"], "h_charge"],
                                     daily.loc[daily["valid"], "h_decharge"]):
            for h in hc_list:  h_ch[h]  += 1
            for h in hd_list:  h_dch[h] += 1
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_ch.values()),
            name="Charge (achat)", marker_color=LBLUE,
            hovertemplate="H%{x:02d} — Charge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_ch.values()],
        ))
        fig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_dch.values()),
            name="Décharge (vente)", marker_color=ORANGE,
            hovertemplate="H%{x:02d} — Décharge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_dch.values()],
        ))
        fig3.update_layout(
            height=320, barmode="group", margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Heure", tickmode="linear", dtick=1,
                       ticktext=[f"H{h:02d}" for h in range(24)],
                       tickvals=list(range(24))),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", y=1.1),
            hoverlabel=dict(bgcolor="white", font_size=12),
        )
        st.plotly_chart(fig3, width="stretch", config=PLOTLY_CFG)

    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown('<p class="section">Distribution des spreads journaliers</p>',
                    unsafe_allow_html=True)
        sv = daily.loc[daily["valid"], "spread"]
        fig4 = go.Figure()
        fig4.add_trace(go.Histogram(
            x=sv, nbinsx=50, marker_color=BLUE, opacity=0.8,
            hovertemplate="Spread : %{x:.0f}–%{x:.0f} €/MWh<br>"
                          "Nb jours : <b>%{y}</b><extra></extra>",
        ))
        fig4.add_vline(
            x=sv.mean(), line_dash="dash", line_color=ORANGE, line_width=2,
            annotation_text=f"Moy: {sv.mean():.1f} €/MWh",
            annotation_position="top right",
            annotation_font=dict(size=12, color=ORANGE),
        )
        fig4.add_vline(
            x=sv.median(), line_dash="dot", line_color=GREEN, line_width=1.5,
            annotation_text=f"Méd: {sv.median():.1f}",
            annotation_position="top left",
            annotation_font=dict(size=11, color=GREEN),
        )
        fig4.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0"),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12),
        )
        st.plotly_chart(fig4, width="stretch", config=PLOTLY_CFG)

    with col_d:
        st.markdown('<p class="section">PnL cumulé dans le temps</p>',
                    unsafe_allow_html=True)
        ds = daily.sort_values("date").copy()
        ds["pnl_cum"]       = ds["pnl"].cumsum()
        ds["pnl_libre_cum"] = ds["pnl_libre"].cumsum()
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_cum"],
            fill="tozeroy", mode="lines", name="PnL cumulé",
            line=dict(color=BLUE, width=2),
            fillcolor="rgba(26,58,92,0.12)",
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "PnL cumulé : <b>%{y:,.0f} €</b><br>"
                "PnL du jour : %{customdata:,.0f} €<extra></extra>"
            ),
            customdata=ds["pnl"].values,
        ))
        fig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_libre_cum"],
            mode="lines", name="Borne max",
            line=dict(color="#b0c8e0", width=1.5, dash="dot"),
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "Borne max cumulée : %{y:,.0f} €<extra></extra>"
            ),
        ))
        fig5.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            yaxis=dict(title="PnL cumulé (€)", tickformat=",", gridcolor="#f0f0f0"),
            xaxis=dict(title="", tickformat="%b %Y"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", y=1.1),
            hoverlabel=dict(bgcolor="white", font_size=12),
            hovermode="x unified",
        )
        st.plotly_chart(fig5, width="stretch", config=PLOTLY_CFG)

    # ── Tableau récap ────────────────────────────────────────────────────────
    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
    show = yearly.copy()
    show["pnl_libre_total"] = show["pnl_libre_total"].apply(lambda x: f"{x:,.0f} €")
    show["pnl_total"]       = show["pnl_total"].apply(lambda x: f"{x:,.0f} €")
    show["pnl_par_MW"]      = show["pnl_par_MW"].apply(lambda x: f"{x:.1f} €/MW")
    show["taux_activation"] = show["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
    show["spread_libre_moy"]= show["spread_libre_moy"].apply(lambda x: f"{x:.1f}")
    show["spread_moy"]      = show["spread_moy"].apply(lambda x: f"{x:.1f}")
    show.columns = ["Année","Jours simulés","Jours actifs","Bloqués usure",
                    "Taux activation","Spread libre moy.","PnL borne max",
                    "Spread contraint moy.","PnL réel","PnL/MW"]
    st.dataframe(show, hide_index=True, width="stretch")

    # ── Explorer un jour ─────────────────────────────────────────────────────
    with st.expander("Explorer un jour spécifique"):
        date_sel = st.date_input("Date", value=pd.Timestamp(pivot["date"].iloc[0]).date())
        row_sel  = daily[daily["date"].dt.date == date_sel]
        if not row_sel.empty:
            r = row_sel.iloc[0]
            prix_j = pivot[pivot["date"].dt.date == date_sel][HOUR_COLS].values.flatten()
            ca, cb = st.columns([3, 1])
            with cb:
                st.metric("Spread", f"{r['spread']:.2f} €/MWh")
                st.metric("PnL", f"{r['pnl']:.2f} €")
                st.metric("Trade valide", "Oui" if r["valid"] else "Non")
            with ca:
                fig_d = go.Figure()
                fig_d.add_trace(go.Bar(
                    x=list(range(24)), y=prix_j.tolist(),
                    marker_color="#c8d8ec", name="Prix spot",
                    hovertemplate="H%{x:02d} — Prix : <b>%{y:.2f} €/MWh</b><extra></extra>",
                ))
                if r["valid"]:
                    fig_d.add_trace(go.Scatter(
                        x=r["h_charge"], y=prix_j[r["h_charge"]],
                        mode="markers", name=f"Charge ({duration_h}h)",
                        marker=dict(color=LBLUE, size=14, symbol="triangle-up"),
                        hovertemplate="H%{x:02d} — Charge : <b>%{y:.2f} €/MWh</b><extra></extra>",
                    ))
                    fig_d.add_trace(go.Scatter(
                        x=r["h_decharge"], y=prix_j[r["h_decharge"]],
                        mode="markers", name=f"Décharge ({duration_h}h)",
                        marker=dict(color=ORANGE, size=14, symbol="triangle-down"),
                        hovertemplate="H%{x:02d} — Décharge : <b>%{y:.2f} €/MWh</b><extra></extra>",
                    ))
                    # Ligne spread
                    fig_d.add_hline(y=r["prix_charge"], line_dash="dot",
                                    line_color=LBLUE, line_width=1,
                                    annotation_text=f"Achat moy: {r['prix_charge']:.1f}€")
                    fig_d.add_hline(y=r["prix_decharge"], line_dash="dot",
                                    line_color=ORANGE, line_width=1,
                                    annotation_text=f"Vente moy: {r['prix_decharge']:.1f}€")
                fig_d.update_layout(
                    height=360, margin=dict(t=10, b=40, l=60, r=10),
                    xaxis=dict(title="Heure", tickmode="linear", dtick=1,
                               ticktext=[f"H{h:02d}" for h in range(24)],
                               tickvals=list(range(24))),
                    yaxis=dict(title="Prix (€/MWh)", gridcolor="#f0f0f0"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", y=1.1),
                    hoverlabel=dict(bgcolor="white", font_size=12),
                )
                st.plotly_chart(fig_d, width="stretch", config=PLOTLY_CFG)

    # ── Export PDF ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Export du rapport</p>', unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns([2, 3])
    with col_exp1:
        if st.button("Générer le rapport PDF"):
            with st.spinner("Génération du PDF..."):
                params_txt = [
                    f"Fichier : {uploaded.name}",
                    f"Puissance : {power_MW} MW | Capacité : {energy_MWh} MWh",
                    f"Durée cycle : {duration_h}h | Rendement : {efficiency*100:.0f}%",
                    f"Max cycles/an : {max_cycles or 'illimité'}",
                    f"Restriction : {', '.join(jours_excl) or 'aucune'} "
                    f"{'H'+str(h_debut)+'-H'+str(h_fin) if jours_excl else ''}",
                    f"Données : {annees[0]}–{annees[-1]} ({jours_total} jours)",
                ]
                kpis = [
                    ("PnL total (avec contraintes)",    f"{total_pnl:,.0f} €"),
                    ("PnL borne max (sans contrainte)", f"{total_pnl_libre:,.0f} €"),
                    ("Ratio réel / max",                f"{ratio:.1f}%"),
                    ("Spread moyen",                    f"{spread_moy:.2f} €/MWh"),
                    ("Jours actifs",                    f"{jours_actifs} / {jours_total}"),
                    ("Taux d'activation",               f"{jours_actifs/jours_total*100:.0f}%"),
                    ("Jours bloqués (usure)",           str(jours_usure)),
                ]
                pdf_bytes = build_pdf_arbitrage(
                    params_txt, kpis, yearly, daily,
                    h_ch, h_dch,
                    datetime.now().strftime("%d/%m/%Y %H:%M")
                )
                st.download_button(
                    label="Télécharger le PDF",
                    data=pdf_bytes,
                    file_name=f"BESS_rapport_arbitrage_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf"
                )

    with col_exp2:
        # Export CSV données journalières
        csv_buf = daily[["date","annee","mois","weekday_name",
                          "spread_libre","pnl_libre","spread","pnl",
                          "valid","h_charge","h_decharge"]].copy()
        csv_buf["date"] = csv_buf["date"].dt.strftime("%Y-%m-%d")
        st.download_button(
            label="Exporter les données (CSV)",
            data=csv_buf.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_donnees_journalieres_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )


# ══════════════════════════════════════════════════════════════════════════════
# MODE 2 — LISSAGE DE COURBE DE CHARGE
# ══════════════════════════════════════════════════════════════════════════════

else:
    st.markdown('<p class="section">Profil de consommation client (MW / heure)</p>',
                unsafe_allow_html=True)

    PROFIL_DEFAUT = [0.3,0.3,0.3,0.3,0.3,0.4,0.6,0.9,1.1,1.2,
                     1.3,1.2,1.0,1.1,1.2,1.3,1.2,1.0,0.9,1.4,
                     1.5,1.2,0.7,0.4]

    col_profil, col_params = st.columns([3, 1])

    with col_params:
        st.markdown("**Paramètres**")
        seuil_pct   = st.slider("Seuil d'écrêtage (percentile)", 50, 95, 75,
                                 help="Percentile de la courbe utilisé comme seuil")
        tarif_kw    = st.number_input("Tarif puissance souscrite (€/MW/an)",
                                       1000, 100000, 12000, 1000)
        soc_init_pct = st.slider("SOC initial (%)", 10, 90, 50)

    with col_profil:
        st.caption("Saisir la consommation en MW pour chaque heure H00 à H23 :")
        cols8 = st.columns(8)
        profil = []
        for h in range(24):
            with cols8[h % 8]:
                v = st.number_input(f"H{h:02d}", 0.0, 100.0,
                                     float(PROFIL_DEFAUT[h]), 0.05,
                                     format="%.2f", label_visibility="visible")
                profil.append(v)

    profil_arr = np.array(profil)

    # ── Simulation ───────────────────────────────────────────────────────────
    params_l = {
        "power_MW": power_MW, "energy_MWh": energy_MWh,
        "soc_min_pct": 0.10, "soc_max_pct": 0.90,
        "efficiency": efficiency, "seuil_percentile": seuil_pct,
        "tarif_puissance_souscrite": tarif_kw,
        "soc_init": soc_init_pct / 100,
    }
    res  = simulate_lissage(pivot, profil_arr, params_l)
    jour = res["jour_type"]

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Résultats</p>', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Réduction de pointe",
              f"{res['reduction_MW']:.3f} MW",
              delta=f"−{res['reduction_MW']/jour['pointe_avant']*100:.1f}%")
    m2.metric("Économie annuelle estimée",
              f"{res['economie_an']:,.0f} €")
    m3.metric("Pointe avant / après",
              f"{jour['pointe_avant']:.2f} → {jour['pointe_apres']:.2f} MW")
    m4.metric("Seuil d'écrêtage appliqué",
              f"{res['seuil_MW']:.2f} MW",
              delta=f"Percentile {seuil_pct}%")

    # ── Graphique principal ──────────────────────────────────────────────────
    st.markdown('<p class="section">Profil de consommation — avant et après lissage</p>',
                unsafe_allow_html=True)

    actions_bess = [v if a == "charge" else -v if a == "decharge" else 0
                    for a, v in jour["actions"]]

    fig_l = make_subplots(rows=2, cols=1, shared_xaxes=True,
                           row_heights=[0.65, 0.35],
                           subplot_titles=["Consommation (MW)",
                                           "Action BESS (+ charge / − décharge)"])
    fig_l.add_trace(go.Scatter(
        x=list(range(24)), y=profil_arr.tolist(),
        mode="lines+markers", name="Avant lissage",
        line=dict(color=ORANGE, width=2.5, dash="dot"), marker=dict(size=6),
        hovertemplate="H%{x:02d} — Avant : <b>%{y:.3f} MW</b><extra></extra>",
    ), row=1, col=1)
    fig_l.add_trace(go.Scatter(
        x=list(range(24)), y=jour["profil_lisse"].tolist(),
        mode="lines+markers", name="Après lissage",
        line=dict(color=BLUE, width=2.5),
        fill="tozeroy", fillcolor="rgba(26,58,92,0.07)", marker=dict(size=6),
        hovertemplate="H%{x:02d} — Après : <b>%{y:.3f} MW</b><extra></extra>",
    ), row=1, col=1)
    fig_l.add_hline(y=res["seuil_MW"], line_dash="dash", line_color=GREEN,
                    line_width=2, annotation_text=f"Seuil {res['seuil_MW']:.2f} MW",
                    annotation_position="top right", row=1, col=1)
    colors_bar = [LBLUE if v >= 0 else ORANGE for v in actions_bess]
    fig_l.add_trace(go.Bar(
        x=list(range(24)), y=actions_bess,
        marker_color=colors_bar, name="BESS",
        hovertemplate="H%{x:02d} — BESS : <b>%{y:.3f} MW</b>"
                      " (%{customdata})<extra></extra>",
        customdata=[a for a, v in jour["actions"]],
    ), row=2, col=1)
    fig_l.add_hline(y=0, line_color="#333", line_width=0.5, row=2, col=1)
    fig_l.update_layout(
        height=520, margin=dict(t=40, b=40, l=60, r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=1.04),
        xaxis2=dict(tickmode="linear", dtick=1, title="Heure",
                    ticktext=[f"H{h:02d}" for h in range(24)],
                    tickvals=list(range(24))),
        yaxis2=dict(title="MW BESS", gridcolor="#f0f0f0"),
        yaxis=dict(gridcolor="#f0f0f0"),
        hoverlabel=dict(bgcolor="white", font_size=12),
        hovermode="x unified",
    )
    st.plotly_chart(fig_l, width="stretch", config=PLOTLY_CFG)

    # ── SOC + Projection économique ──────────────────────────────────────────
    col_s1, col_s2 = st.columns(2)

    with col_s1:
        st.markdown('<p class="section">État de charge batterie (SOC)</p>',
                    unsafe_allow_html=True)
        fig_soc = go.Figure()
        fig_soc.add_trace(go.Scatter(
            x=list(range(25)), y=jour["soc_hist"],
            mode="lines+markers", name="SOC",
            line=dict(color=LBLUE, width=2),
            fill="tozeroy", fillcolor="rgba(46,117,182,0.12)",
            marker=dict(size=6),
            hovertemplate="Après H%{x:02d} — SOC : <b>%{y:.3f} MWh</b><extra></extra>",
        ))
        fig_soc.add_hline(y=energy_MWh * 0.9, line_dash="dash",
                          line_color=GREEN, annotation_text=f"SOC max ({energy_MWh*0.9:.1f} MWh)")
        fig_soc.add_hline(y=energy_MWh * 0.1, line_dash="dash",
                          line_color=ORANGE, annotation_text=f"SOC min ({energy_MWh*0.1:.1f} MWh)")
        fig_soc.update_layout(
            height=300, margin=dict(t=10, b=40, l=60, r=20),
            xaxis=dict(title="Heure", tickmode="linear", dtick=2),
            yaxis=dict(title="SOC (MWh)", gridcolor="#f0f0f0",
                       range=[0, energy_MWh * 1.1]),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12),
        )
        st.plotly_chart(fig_soc, width="stretch", config=PLOTLY_CFG)

    with col_s2:
        st.markdown('<p class="section">Projection économique annuelle</p>',
                    unsafe_allow_html=True)
        yl = res["yearly"]
        fig_eco = go.Figure()
        fig_eco.add_trace(go.Bar(
            x=yl["annee"].astype(str), y=yl["economie_an"],
            marker_color=BLUE,
            text=yl["economie_an"].apply(lambda x: f"{x:,.0f} €"),
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Économie : <b>%{y:,.0f} €</b><extra></extra>",
        ))
        fig_eco.update_layout(
            height=300, margin=dict(t=10, b=40, l=60, r=20),
            xaxis_title="Année",
            yaxis=dict(title="Économie (€)", tickformat=",", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12),
        )
        st.plotly_chart(fig_eco, width="stretch", config=PLOTLY_CFG)

    # ── Détail heure par heure ───────────────────────────────────────────────
    with st.expander("Détail heure par heure"):
        detail = pd.DataFrame({
            "Heure"              : [f"H{h:02d}" for h in range(24)],
            "Conso originale (MW)": profil_arr.round(3),
            "Action BESS"        : [f"{a} {v:.3f} MW" for a, v in jour["actions"]],
            "Conso lissée (MW)"  : jour["profil_lisse"].round(3),
            "SOC après (MWh)"    : jour["soc_hist"][1:],
        })
        st.dataframe(detail, hide_index=True, width="stretch")

    st.info(
        f"**Méthode :** Seuil au percentile {seuil_pct}% de la courbe "
        f"({res['seuil_MW']:.2f} MW). La batterie décharge au-dessus de ce seuil "
        f"et charge en dessous. Économie = réduction de puissance × tarif réseau "
        f"({res['reduction_MW']:.3f} MW × {tarif_kw:,} €/MW = "
        f"**{res['economie_an']:,.0f} €/an**)."
    )

    # ── Export ───────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Export du rapport</p>', unsafe_allow_html=True)
    col_e1, col_e2 = st.columns([2, 3])
    with col_e1:
        if st.button("Générer le rapport PDF"):
            with st.spinner("Génération du PDF..."):
                params_txt = [
                    f"Fichier : {uploaded.name}",
                    f"Puissance BESS : {power_MW} MW | Capacité : {energy_MWh} MWh",
                    f"Rendement : {efficiency*100:.0f}% | SOC init : {soc_init_pct}%",
                    f"Seuil écrêtage : percentile {seuil_pct}% = {res['seuil_MW']:.2f} MW",
                    f"Tarif puissance souscrite : {tarif_kw:,} €/MW/an",
                    f"Données : {annees[0]}–{annees[-1]}",
                ]
                kpis = [
                    ("Réduction de pointe",      f"{res['reduction_MW']:.3f} MW"),
                    ("Pointe avant lissage",      f"{jour['pointe_avant']:.2f} MW"),
                    ("Pointe après lissage",      f"{jour['pointe_apres']:.2f} MW"),
                    ("Économie annuelle estimée", f"{res['economie_an']:,.0f} €"),
                    ("Seuil appliqué",            f"{res['seuil_MW']:.2f} MW"),
                ]
                detail_df = pd.DataFrame({
                    "Heure": [f"H{h:02d}" for h in range(24)],
                    "Conso originale (MW)": profil_arr.round(3),
                    "Action BESS": [f"{a} {v:.3f} MW" for a, v in jour["actions"]],
                    "Conso lissée (MW)": jour["profil_lisse"].round(3),
                    "SOC après (MWh)": jour["soc_hist"][1:],
                })
                pdf_bytes = build_pdf_lissage(
                    params_txt, kpis, detail_df, res["yearly"],
                    datetime.now().strftime("%d/%m/%Y %H:%M")
                )
                st.download_button(
                    label="Télécharger le PDF",
                    data=pdf_bytes,
                    file_name=f"BESS_rapport_lissage_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf"
                )
    with col_e2:
        detail_csv = pd.DataFrame({
            "Heure": [f"H{h:02d}" for h in range(24)],
            "Conso_MW": profil_arr.round(3),
            "Action": [a for a, v in jour["actions"]],
            "BESS_MW": [v for a, v in jour["actions"]],
            "Lisse_MW": jour["profil_lisse"].round(3),
            "SOC_MWh": jour["soc_hist"][1:],
        })
        st.download_button(
            label="Exporter profil lissé (CSV)",
            data=detail_csv.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_lissage_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )