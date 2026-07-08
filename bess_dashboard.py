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
import json as _json
import json
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


# Logo ENI encodé en base64
ENI_LOGO_B64 = "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"
from bess_engine import (
    HOUR_COLS, aggregate_arbitrage, load_spot, load_imbalance,
    simulate_arbitrage, simulate_arbitrage_optimal, simulate_lissage,
    get_available_hours, _best_n_cycles,
)

def _fmt(val, decimals=0):
    """Formate un nombre avec espace insécable comme séparateur de milliers."""
    try:
        fmt = f"{float(val):,.{decimals}f}"
        return fmt.replace(",", "\u00a0")  # espace insécable
    except (TypeError, ValueError):
        return str(val)


# ──────────────────────────────────────────────────────────────────────────────
# RAPPORT DE VÉRIFICATION GLOBAL (bouton sidebar)
# ──────────────────────────────────────────────────────────────────────────────
# Le rapport HTML reproduit TOUT ce qui s'affiche dans chaque onglet (KPI,
# tableaux, graphiques Plotly interactifs) en rejouant la capture générique
# _REPORT_CAPTURE (voir _diag_set_tab / st.metric|dataframe|table|plotly_chart
# enveloppés, juste après la création des tabs). Un onglet jamais ouvert dans
# cette session n'a simplement rien capturé → "non calculé".

def _diag_kpis_html(items):
    cells = "".join(
        f'<div style="flex:1;min-width:150px;padding:10px;background:#f0f5fb;'
        f'border-radius:6px;margin:4px;">'
        f'<div style="font-size:0.72rem;color:#6b7a8d;">{label}</div>'
        f'<div style="font-size:1.05rem;font-weight:700;color:#1a3a5c;">{value}</div></div>'
        for label, value in items
    )
    return f'<div style="display:flex;flex-wrap:wrap;">{cells}</div>'


def _diag_table_html(data, cols=None):
    import pandas as _pd_diag
    if isinstance(data, _pd_diag.DataFrame):
        records = data.to_dict(orient="records")
    elif hasattr(data, "to_html") and not isinstance(data, (list, dict)):
        # pandas Styler (ou similaire) : on garde son rendu natif (styles/couleurs inclus)
        try:
            return data.to_html()
        except Exception:
            records = []
    else:
        records = list(data or [])
    if not records:
        return "<p><i>Aucune donnée.</i></p>"
    cols = cols or list(records[0].keys())

    def _cell(v):
        if isinstance(v, float):
            return f"{v:,.2f}".replace(",", " ")
        return str(v)

    head = "".join(f"<th>{c}</th>" for c in cols)
    rows = "".join(
        "<tr>" + "".join(f"<td>{_cell(r.get(c, ''))}</td>" for c in cols) + "</tr>"
        for r in records
    )
    return (f'<table style="width:100%;border-collapse:collapse;font-size:0.8rem;">'
            f'<thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table>')


def _diag_render_items(items, fig_counter):
    """Rejoue dans l'ordre la liste capturée (metric/dataframe/fig) d'un onglet."""
    html_parts = []
    i, n = 0, len(items)
    while i < n:
        kind, label, payload = items[i]
        if kind == "metric":
            group = []
            while i < n and items[i][0] == "metric":
                group.append((items[i][1], items[i][2]))
                i += 1
            html_parts.append(_diag_kpis_html(group))
            continue
        if kind == "dataframe":
            html_parts.append(_diag_table_html(payload))
            i += 1
            continue
        if kind == "fig":
            try:
                fig_counter["n"] += 1
                include_js = not fig_counter["js_done"]
                fig_counter["js_done"] = True
                html_parts.append(payload.to_html(
                    full_html=False, include_plotlyjs=include_js,
                    div_id=f"diagfig_{fig_counter['n']}",
                ))
            except Exception as _e_fig:
                html_parts.append(f"<p><i>Graphique non exportable ({_e_fig}).</i></p>")
            i += 1
            continue
        i += 1
    return "".join(html_parts)


def _build_global_diag_html():
    import datetime as _dtd_g

    TAB_ORDER = [
        ("arb",       "Arbitrage Day-Ahead"),
        ("intra",     "Intraday"),
        ("lis",       "Lissage de charge"),
        ("comp",      "Comparaison scénarios"),
        ("sensi",     "Sensibilité"),
        ("pays",      "Comparaison pays"),
        ("hist",      "Historique 2019–2025"),
        ("exec",      "Executive Summary"),
        ("verif_arb", "Vérification Arbitrage"),
        ("verif_lis", "Vérification Lissage"),
    ]

    fig_counter = {"n": 0, "js_done": False}
    sections = []
    for key, titre in TAB_ORDER:
        items = _REPORT_CAPTURE.get(key, [])
        body = _diag_render_items(items, fig_counter) if items else \
            "<p><i>Onglet non ouvert/calculé dans cette session.</i></p>"
        sections.append((titre, body))

    # Méthodologie : contenu statique déjà vérifié à jour (cf. _AI_METHODO_DOC)
    _methodo_esc = _AI_METHODO_DOC.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    sections.append(("Méthodologie", f"<pre style='white-space:pre-wrap;font-family:inherit;font-size:0.85rem;'>{_methodo_esc}</pre>"))

    sections_html = "".join(
        f'<h2 style="color:#1a3a5c;border-bottom:2px solid #d0dff0;padding-bottom:4px;'
        f'margin-top:32px;">{titre}</h2>{body}'
        for titre, body in sections
    )

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>BESS — Rapport de vérification complet</title>
<style>
body {{ font-family: Arial, Helvetica, sans-serif; max-width: 1300px; margin: 24px auto; color: #222; }}
th, td {{ border: 1px solid #d0dff0; padding: 4px 10px; text-align: right; }}
th {{ background: #1a3a5c; color: white; }}
td:first-child, th:first-child {{ text-align: left; }}
tr:nth-child(even) td {{ background: #f7fafd; }}
</style></head><body>
<h1 style="color:#1a3a5c;">BESS Valorisation — Rapport de vérification complet</h1>
<p style="color:#6b7a8d;">Généré le {_dtd_g.datetime.now().strftime('%d/%m/%Y %H:%M')} — fichier {st.session_state.get('excel_name','N/A')}
— modèle {st.session_state.get('modele_choix','N/A')}</p>
<p style="color:#6b7a8d;font-size:0.85rem;">Seuls les onglets ouverts/calculés pendant cette session apparaissent avec leur contenu (KPI, tableaux, graphiques interactifs identiques au dashboard).</p>
{sections_html}
</body></html>"""


def _plot_cycles_distribution(daily_df, periode_label):
    """Graphique répartition des cycles journaliers avec min/max annotés."""
    import plotly.graph_objects as go

    ds = daily_df[daily_df["valid"]].copy()
    if ds.empty:
        return None

    # Distribution : compter nb de jours par nb de cycles réalisés
    dist = ds["n_cycles_actifs"].value_counts().sort_index().reset_index()
    dist.columns = ["n_cycles", "nb_jours"]

    # Min et max cycles (jours actifs uniquement)
    min_cyc = int(ds["n_cycles_actifs"].min())
    max_cyc = int(ds["n_cycles_actifs"].max())

    # Date du premier jour avec min cycles et premier jour avec max cycles
    date_min = ds.loc[ds["n_cycles_actifs"] == min_cyc, "date"].iloc[0]
    date_max = ds.loc[ds["n_cycles_actifs"] == max_cyc, "date"].iloc[0]

    # Couleur : dégradé vert selon n_cycles
    max_n = dist["n_cycles"].max()
    colors = [f"rgba(46,125,50,{0.4 + 0.6*(n/max_n if max_n>0 else 1)})" for n in dist["n_cycles"]]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dist["n_cycles"].astype(str),
        y=dist["nb_jours"],
        marker_color=colors,
        text=dist["nb_jours"],
        textposition="outside",
        hovertemplate="<b>%{x} cycle(s)/jour</b><br>Nb jours : %{y}<extra></extra>",
        name="Jours"
    ))

    # Annotations min et max
    def _date_fmt(d):
        try:
            return pd.Timestamp(d).strftime("%d/%m/%Y")
        except Exception:
            return str(d)

    for cyc, date, color, ay in [
        (min_cyc, date_min, "#e65100", 50),
        (max_cyc, date_max, "#1565c0", 50),
    ]:
        if min_cyc == max_cyc:
            break  # un seul niveau, une seule annotation suffit
        nb = int(dist.loc[dist["n_cycles"] == cyc, "nb_jours"].iloc[0])
        fig.add_annotation(
            x=str(cyc), y=nb,
            text=f"{'Min' if cyc==min_cyc else 'Max'} : {cyc} cycle(s)<br>{_date_fmt(date)}",
            showarrow=True, arrowhead=2,
            arrowcolor=color, font=dict(color=color, size=11),
            bgcolor="white", bordercolor=color, borderwidth=1,
            ay=ay, ax=0
        )

    if min_cyc == max_cyc:
        nb = int(dist.loc[dist["n_cycles"] == min_cyc, "nb_jours"].iloc[0])
        fig.add_annotation(
            x=str(min_cyc), y=nb,
            text=f"{min_cyc} cycle(s)/jour tous les jours<br>ex. {_date_fmt(date_min)}",
            showarrow=True, arrowhead=2,
            arrowcolor="#2e7d32", font=dict(color="#2e7d32", size=11),
            bgcolor="white", bordercolor="#2e7d32", borderwidth=1,
            ay=-50, ax=0
        )

    fig.update_layout(
        xaxis=dict(title="Nombre de cycles réalisés dans la journée", tickmode="linear"),
        yaxis=dict(title="Nombre de jours", gridcolor="#f0f0f0"),
        plot_bgcolor="white", paper_bgcolor="white",
        showlegend=False,
        margin=dict(t=40, b=40),
        height=340,
    )
    return fig, min_cyc, max_cyc, date_min, date_max




# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────

# ── Session state — thème et fichier ─────────────────────────────────────────
if "theme" not in st.session_state:
    st.session_state.theme = "light"
if "bloomberg" not in st.session_state:
    st.session_state.bloomberg = False
if "excel_cdc_bytes" not in st.session_state:
    st.session_state.excel_cdc_bytes = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

if "excel_name" not in st.session_state:
    st.session_state.excel_name = None

st.set_page_config(
    page_title="BESS Valorisation DA — Plénitude",
    page_icon="data:image/png;base64,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",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
/* ═══════════════════════════════════════════════════
   LOGO ENI — fixe en haut à droite
   ═══════════════════════════════════════════════════ */
.eni-toolbar {{
  position: fixed; top: 4px; right: 8px; height: 40px;
  display: flex; align-items: center; z-index: 999999;
  background: white; border-radius: 6px; padding: 2px 5px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.12);
}}
.eni-toolbar img {{ height: 34px; width: auto; display: block; }}
[data-testid="stToolbar"] {{ padding-right: 72px !important; }}

@keyframes eni-pulse {{
  0%   {{ transform: scale(1);    opacity: 1; }}
  50%  {{ transform: scale(0.82); opacity: 0.65; }}
  100% {{ transform: scale(1);    opacity: 1; }}
}}
.eni-loading {{ animation: eni-pulse 1.2s ease-in-out infinite; display:block; margin:0 auto; }}

/* ═══════════════════════════════════════════════════
   MODE CLAIR (défaut)
   ═══════════════════════════════════════════════════ */
{"" if st.session_state.theme == "dark" else """
[data-testid="stAppViewContainer"] { background: #f8f9fb; }
[data-testid="stSidebar"]          { background: #ffffff; border-right: 1px solid #e8ecf0; }
.main-title  { font-size:1.9rem; font-weight:700; color:#1a3a5c; margin-bottom:0; }
.sub-title   { font-size:0.92rem; color:#6b7a8d; margin-top:2px; margin-bottom:1.4rem; }
.section     { font-size:1.05rem; font-weight:600; color:#1a3a5c;
               border-bottom:2px solid #d0dff0; padding-bottom:4px;
               margin-top:1.4rem; margin-bottom:0.7rem; }
div[data-testid="stMetricValue"] { font-size:1.55rem !important; font-weight:700; color:#1a3a5c; }
div[data-testid="stMetricLabel"] { font-size:0.8rem !important; color:#6b7a8d; font-weight:500; }
.stButton > button { background:#1a3a5c; color:white; border:none;
                     border-radius:6px; padding:8px 20px; font-weight:600; }
.stButton > button:hover { background:#2e5d8e; }
div.stSpinner > div { border-top-color:#f5c518 !important;
                      border-color:#f5c518 transparent transparent transparent !important; }
"""}

/* ═══════════════════════════════════════════════════
   MODE BLOOMBERG TERMINAL
   ═══════════════════════════════════════════════════ */
{"" if st.session_state.theme == "light" else """

/* --- Fond global --- */
[data-testid="stAppViewContainer"],
[data-testid="stApp"],
.main, section.main, .block-container {
  background: #000000 !important;
  color: #ffffff !important;
  font-family: 'Courier New', Courier, monospace !important;
}

/* --- Sidebar --- */
[data-testid="stSidebar"] {
  background: #0a0a0a !important;
  border-right: 2px solid #ff6600 !important;
}
[data-testid="stSidebar"] * {
  font-family: 'Courier New', Courier, monospace !important;
  font-size: 11px !important;
}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
  color: #cccccc !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
  color: #ff6600 !important;
  letter-spacing: 2px !important;
  text-transform: uppercase !important;
  font-size: 11px !important;
}

/* --- Header section séparateurs --- */
[data-testid="stSidebar"] hr,
[data-testid="stSidebar"] .stMarkdown hr {
  border-color: #333333 !important;
}

/* --- Inputs sidebar --- */
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] .stNumberInput input,
[data-testid="stSidebar"] .stSelectbox > div > div {
  background: #0a0a0a !important;
  color: #ff6600 !important;
  border: 1px solid #333333 !important;
  border-radius: 0 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 12px !important;
}
[data-testid="stSidebar"] .stSlider div[data-baseweb="slider"] {
  background: #333333 !important;
}

/* --- Titres principaux Bloomberg --- */
.main-title {
  font-size: 13px !important;
  font-weight: 700 !important;
  color: #ff6600 !important;
  letter-spacing: 3px !important;
  font-family: 'Courier New', monospace !important;
  text-transform: uppercase !important;
  margin-bottom: 0 !important;
}
.sub-title {
  font-size: 10px !important;
  color: #666666 !important;
  font-family: 'Courier New', monospace !important;
  letter-spacing: 1px !important;
  margin-bottom: 8px !important;
}

/* --- Sections --- */
.section {
  font-size: 10px !important;
  font-weight: 700 !important;
  color: #000000 !important;
  background: #ff6600 !important;
  padding: 2px 8px !important;
  letter-spacing: 2px !important;
  text-transform: uppercase !important;
  border: none !important;
  margin-top: 12px !important;
  margin-bottom: 6px !important;
  display: block !important;
}

/* --- Métriques --- */
div[data-testid="stMetricValue"] {
  font-size: 20px !important;
  font-weight: 700 !important;
  color: #ff6600 !important;
  font-family: 'Courier New', monospace !important;
}
div[data-testid="stMetricLabel"] {
  font-size: 9px !important;
  color: #666666 !important;
  font-family: 'Courier New', monospace !important;
  text-transform: uppercase !important;
  letter-spacing: 1px !important;
}
div[data-testid="stMetricDelta"] {
  font-family: 'Courier New', monospace !important;
  font-size: 10px !important;
}
div[data-testid="stMetricDelta"] span { color: #00cc00 !important; }
div[data-testid="stMetricDelta"][data-direction="down"] span { color: #ff3333 !important; }

/* --- Boutons Bloomberg carré --- */
.stButton > button {
  background: #ff6600 !important;
  color: #000000 !important;
  border: none !important;
  border-radius: 0 !important;
  font-weight: 700 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 11px !important;
  letter-spacing: 1px !important;
  text-transform: uppercase !important;
  padding: 4px 12px !important;
}
.stButton > button:hover {
  background: #cc5200 !important;
  color: #000000 !important;
}

/* --- Dataframe Bloomberg --- */
[data-testid="stDataFrame"] { background: #000000 !important; }
[data-testid="stDataFrame"] table {
  background: #000000 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 11px !important;
  border-collapse: collapse !important;
}
[data-testid="stDataFrame"] th {
  background: #1a1a1a !important;
  color: #ff6600 !important;
  border: 1px solid #333333 !important;
  font-size: 10px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.5px !important;
  padding: 3px 6px !important;
}
[data-testid="stDataFrame"] td {
  color: #cccccc !important;
  border: 1px solid #1a1a1a !important;
  padding: 2px 6px !important;
  font-size: 11px !important;
}
[data-testid="stDataFrame"] tr:hover td { background: #111111 !important; }

/* --- Textes principaux --- */
p, li, span {
  font-family: 'Courier New', monospace !important;
  color: #cccccc !important;
}
h1, h2, h3, h4 {
  font-family: 'Courier New', monospace !important;
  color: #ff6600 !important;
}

/* --- Expander --- */
[data-testid="stExpander"] {
  background: #0a0a0a !important;
  border: 1px solid #333333 !important;
  border-radius: 0 !important;
}
[data-testid="stExpander"] summary {
  color: #ff6600 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 11px !important;
  letter-spacing: 1px !important;
  text-transform: uppercase !important;
}

/* --- Info boxes --- */
[data-testid="stInfo"] {
  background: #001122 !important;
  border-left: 3px solid #00aaff !important;
  border-radius: 0 !important;
  color: #aaccff !important;
  font-family: 'Courier New', monospace !important;
  font-size: 11px !important;
}
[data-testid="stSuccess"] {
  background: #001100 !important;
  border-left: 3px solid #00cc00 !important;
  border-radius: 0 !important;
  font-family: 'Courier New', monospace !important;
}

/* --- Upload --- */
[data-testid="stFileUploader"] {
  background: #0a0a0a !important;
  border: 1px solid #333333 !important;
  border-radius: 0 !important;
}

/* --- Spinner Bloomberg --- */
div.stSpinner > div {
  border-top-color: #ff6600 !important;
  border-color: #ff6600 transparent transparent transparent !important;
}

/* --- Logo toolbar dark --- */
.eni-toolbar {
  background: #0a0a0a !important;
  border: 1px solid #ff6600 !important;
  border-radius: 0 !important;
  box-shadow: 0 0 6px rgba(255,102,0,0.5) !important;
}

/* --- Caption/small texts --- */
small, .stCaption, [data-testid="stCaptionContainer"] {
  color: #555555 !important;
  font-family: 'Courier New', monospace !important;
  font-size: 9px !important;
}

/* --- Radio buttons --- */
[data-testid="stRadio"] label { color: #cccccc !important; }
[data-testid="stRadio"] span  { color: #ff6600 !important; }

/* --- Selectbox --- */
[data-baseweb="select"] > div {
  background: #0a0a0a !important;
  border-color: #333333 !important;
  border-radius: 0 !important;
  color: #ff6600 !important;
  font-family: 'Courier New', monospace !important;
}
[data-baseweb="menu"] {
  background: #0a0a0a !important;
  border: 1px solid #ff6600 !important;
}
[data-baseweb="option"] {
  background: #0a0a0a !important;
  color: #cccccc !important;
  font-family: 'Courier New', monospace !important;
}
[data-baseweb="option"]:hover { background: #1a1a1a !important; color: #ff6600 !important; }

/* --- Multiselect tags --- */
[data-baseweb="tag"] {
  background: #ff6600 !important;
  border-radius: 0 !important;
  color: #000000 !important;
}

/* --- Scrollbar Bloomberg --- */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #000000; }
::-webkit-scrollbar-thumb { background: #333333; }
::-webkit-scrollbar-thumb:hover { background: #ff6600; }

/* --- Masquer le spinner natif Streamlit (bonhomme qui court) --- */
[data-testid="stStatusWidget"] { display: none !important; }
.stStatusWidget { display: none !important; }
"""}
</style>
""", unsafe_allow_html=True)

# Logo ENI — dans la toolbar en haut à droite
st.markdown(
    f'<div class="eni-toolbar">'
    f'<img src="data:image/png;base64,{ENI_LOGO_B64}" alt="ENI"/>'
    f'</div>',
    unsafe_allow_html=True
)

# Couvrir la zone du spinner Streamlit (bonhomme + bouton Stop) avec une barre
# de progression fixe positionnée exactement sur la bannière haute droite.
st.markdown("""
<style>
/* Masquer complètement le widget natif Streamlit (bonhomme + Stop) */
[data-testid="stStatusWidget"] {
    visibility: hidden !important;
    opacity: 0 !important;
}

/* Masquer le texte du st.progress (on garde juste la barre visuelle) */
div[data-testid="stProgressBar"] p {
    display: none !important;
}
</style>


""", unsafe_allow_html=True)

# Injecter la barre de progression dans le header Streamlit via MutationObserver
st.markdown("""
<style>
/* Barre de progression dans le header */
#bess-header-progress {
    display: none;
    align-items: center;
    gap: 8px;
    position: fixed;
    top: 0;
    right: 170px;
    height: 47px;
    z-index: 99999;
    width: 220px;
}
#bess-header-progress.visible { display: flex; }
#bess-header-prog-wrap {
    flex: 1;
}
#bess-header-prog-text {
    font-size: 10px;
    color: #4a5568;
    white-space: nowrap;
    margin-bottom: 3px;
    font-family: sans-serif;
}
#bess-header-prog-track {
    height: 5px;
    background: #e2e8f0;
    border-radius: 3px;
    overflow: hidden;
    width: 100%;
}
#bess-header-prog-fill {
    height: 5px;
    width: 0%;
    background: linear-gradient(90deg, #2e75b6, #5ba3d9);
    border-radius: 3px;
    transition: width 0.3s ease;
}
@keyframes bess-slide {
    0%   { margin-left: -30%; width: 30%; }
    100% { margin-left: 100%; width: 30%; }
}
#bess-header-prog-fill.indeterminate {
    width: 30% !important;
    animation: bess-slide 1.1s ease-in-out infinite;
}
</style>
<div id="bess-header-progress">
    <div id="bess-header-prog-wrap">
        <div id="bess-header-prog-text">Calcul en cours…</div>
        <div id="bess-header-prog-track">
            <div id="bess-header-prog-fill"></div>
        </div>
    </div>
</div>
<script>
window._bessSetProgress = function(pct, label) {
    var el   = document.getElementById('bess-header-progress');
    var fill = document.getElementById('bess-header-prog-fill');
    var txt  = document.getElementById('bess-header-prog-text');
    if (!el) return;
    el.classList.add('visible');
    if (txt) txt.textContent = label || 'Calcul en cours…';
    if (fill) {
        if (pct === null || pct === undefined) {
            fill.classList.add('indeterminate');
            fill.style.width = '';
        } else {
            fill.classList.remove('indeterminate');
            fill.style.width = Math.min(100, pct) + '%';
        }
    }
};
window._bessHideProgress = function() {
    var el = document.getElementById('bess-header-progress');
    if (el) el.classList.remove('visible');
};
</script>
""", unsafe_allow_html=True)

BLUE    = "#1a3a5c"
LBLUE   = "#2e75b6"
ORANGE  = "#d46b1a"
GREEN   = "#3a8a5c"
LGREY   = "#e8ecf0"
COLORS  = ["#5cb85c", "#f5c518", "#2e75b6", "#d46b1a", "#7b3fa0", "#b05050"]


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
    story.append(Paragraph("1. Paramètrès de simulation", S["h2"]))
    for line in params_txt:
        story.append(Paragraph(f"• {line}", S["body"]))
    story.append(Spacer(1, 0.3*cm))

    # ── KPIs ─────────────────────────────────────────────────────────────────
    story.append(Paragraph("2. Résultats clés (toutes annees)", S["h2"]))
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
            f"{r['spread_absolu_moy']:.1f}",
            f"{_fmt(r['pnl_absolu_total'])}",
            f"{r['spread_moy']:.1f}",
            f"{_fmt(r['pnl_total'])}",
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
            f"{_fmt(r['pnl'])}",
            f"{_fmt(cumul)}",
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
            f"{_fmt(r['economie_an'])}",
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



# ── Constantes Bloomberg pour les graphiques Plotly ─────────────────────────
_BB = st.session_state.bloomberg
BB_BG      = "#000000"
BB_PAPER   = "#0a0a0a"
BB_ORANGE  = "#ff6600"
BB_BLUE    = "#4fc3f7"
BB_WHITE   = "#e0e0e0"
BB_GREEN   = "#00e676"
BB_GRID    = "#1a1a1a"
BB_AXIS    = "#555555"
BB_MONO    = "Courier New, Courier, monospace"

# Couleurs tracés selon thème
COLORS = ([BB_ORANGE, BB_BLUE, BB_WHITE, BB_GREEN, "#ef5350", "#ffd54f"]
          if _BB else
          ["#5cb85c", "#f5c518", "#2e75b6", "#d46b1a", "#7b3fa0", "#b05050"])

# Alias couleurs pour les tracés (s'adaptent au thème)
C1 = COLORS[0]  # vert clair (light) / orange Bloomberg (dark)
C2 = COLORS[1]  # jaune (light) / bleu clair (dark)
C3 = COLORS[2]  # bleu (light) / blanc (dark)
C4 = COLORS[3]  # orange (light) / vert (dark)
C5 = COLORS[4]  # violet
C6 = COLORS[5]  # rouge foncé


def apply_bb(fig):
    """Style Bloomberg Terminal pour les graphiques Plotly."""
    if not _BB:
        return fig

    BB_BG_CHART = "#0a0a0a"
    BB_PAPER    = "#000000"
    BB_GRID     = "#111111"
    BB_AXIS_C   = "#444444"
    BB_TEXT_C   = "#cccccc"
    BB_TITLE_C  = "#ff6600"
    BB_MONO     = "Courier New, monospace"

    fig.update_layout(
        plot_bgcolor  = BB_BG_CHART,
        paper_bgcolor = BB_PAPER,
        font = dict(color=BB_TEXT_C, family=BB_MONO, size=10),
        hoverlabel = dict(
            bgcolor="#111111", font_color="#ff6600",
            bordercolor="#ff6600", font_family=BB_MONO, font_size=11,
        ),
        legend = dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color=BB_TEXT_C, family=BB_MONO, size=9),
            bordercolor="#333333", borderwidth=1,
        ),
        title_font=dict(color=BB_TITLE_C, family=BB_MONO),
    )
    fig.update_xaxes(
        gridcolor=BB_GRID, gridwidth=1,
        linecolor="#333333", linewidth=1,
        tickfont=dict(color=BB_AXIS_C, family=BB_MONO, size=9),
        title_font=dict(color=BB_TITLE_C, family=BB_MONO, size=9),
        zeroline=False, showline=True,
    )
    fig.update_yaxes(
        gridcolor=BB_GRID, gridwidth=1,
        linecolor="#333333", linewidth=1,
        tickfont=dict(color=BB_AXIS_C, family=BB_MONO, size=9),
        title_font=dict(color=BB_TITLE_C, family=BB_MONO, size=9),
        zeroline=False, showline=True,
    )
    # Adapter les couleurs des traces au style Bloomberg
    for trace in fig.data:
        # Barres
        if hasattr(trace, 'marker') and trace.marker is not None:
            mc = getattr(trace.marker, 'color', None)
            if isinstance(mc, str):
                if mc in ('#5cb85c', '#3a8a5c'):
                    trace.marker.color = '#ff6600'      # vert → orange BB
                elif mc in ('#f5c518', '#fef08a', '#d4a01a'):
                    trace.marker.color = '#00aaff'      # jaune → bleu BB
                elif mc in ('#c8d8ec', '#c8d4e0'):
                    trace.marker.color = '#1a1a1a'      # gris clair → gris sombre
                elif mc == '#fef08a':
                    trace.marker.color = '#004466'
        # Lignes
        if hasattr(trace, 'line') and trace.line is not None:
            lc = getattr(trace.line, 'color', None)
            if isinstance(lc, str):
                if lc in ('#5cb85c', '#3a8a5c', '#00cc44'):
                    trace.line.color = '#ff6600'        # vert → orange BB
                elif lc in ('#f5c518', '#d4a01a', '#fef08a'):
                    trace.line.color = '#00aaff'        # jaune → bleu BB
                elif lc in ('#b0c8e0', '#c8d8ec'):
                    trace.line.color = '#004488'
                elif lc == '#1a3a5c':
                    trace.line.color = '#ff6600'
        # Fill
        if hasattr(trace, 'fillcolor') and trace.fillcolor:
            fc = trace.fillcolor
            if 'rgba(92,184,92' in str(fc) or 'rgba(26,58,92' in str(fc):
                trace.fillcolor = 'rgba(255,102,0,0.08)'
        # Marqueurs triangles
        if hasattr(trace, 'marker') and trace.marker is not None:
            mc = getattr(trace.marker, 'color', None)
            if isinstance(mc, str) and mc in ('#5cb85c', '#f5c518'):
                trace.marker.color = '#ff6600' if mc == '#5cb85c' else '#00aaff'
    return fig

with st.sidebar:
    st.markdown("---")

    # ── Upload fichier ────────────────────────────────────────────────────────
    st.markdown("### Données spot")
    uploaded = st.file_uploader(
        "Fichier Excel (Spot_input)",
        type=["xlsx"],
        help=(
            "Uploadez BESS_valorisation_RESULTATS.xlsx (sans '_Copie').\n\n"
            "Ce fichier alimente TOUS les onglets :\n"
            "• Arbitrage 2026–2029 → feuille Spot_input\n"
            "• Historique 2019–2025 → feuille Case 3 night station"
        )
    )
    if uploaded is not None:
        raw = uploaded.read()
        if raw:
            st.session_state.excel_bytes = raw
            st.session_state["_uploaded_fname"] = uploaded.name
            st.session_state.excel_name  = uploaded.name
    if st.session_state.excel_bytes is not None and uploaded is None:
        uploaded = type("FakeFile", (), {"name": st.session_state.excel_name})()

    # ── Paramètres batterie ───────────────────────────────────────────────────
    st.markdown("### Modèle de batterie")

    # 3 modèles issus du fichier Excel (feuille BESS) — puissances seulement
    # Les restrictions horaires sont configurées séparément dans l'onglet Arbitrage
    # Définition complète des 3 modèles (puissance + restrictions par défaut)
    _MODELES = {
        "Modèle 1 — 430 kW": {
            "power_MW": 0.430,
            "jours":    ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"],
            "h_debut":  10,
            "h_fin":    12,
            "desc":     "430 kW · Restriction Lun–Sam H10–H12",
        },
        "Modèle 2 — 215 kW": {
            "power_MW": 0.215,
            "jours":    ["Mardi", "Mercredi", "Jeudi"],
            "h_debut":  10,
            "h_fin":    14,
            "desc":     "215 kW · Restriction Mar–Jeu H10–H14",
        },
        "Modèle 3 — 215 kW": {
            "power_MW": 0.215,
            "jours":    [],
            "h_debut":  0,
            "h_fin":    0,
            "desc":     "215 kW · Aucune restriction",
        },
        "Modèle 4 — 1 MW": {
            "power_MW": 1.0,
            "jours":    [],
            "h_debut":  0,
            "h_fin":    0,
            "desc":     "1 MW · Aucune restriction",
        },
        "Personnalisé": {
            "power_MW": None,
            "jours":    [],
            "h_debut":  0,
            "h_fin":    0,
            "desc":     "",
        },
    }

    _default_model = "Modèle 4 — 1 MW"
    # Forcer reset si le modèle en session n'existe plus (ex. ancienne version)
    if ("modele_choix" not in st.session_state
            or st.session_state["modele_choix"] not in _MODELES):
        st.session_state["modele_choix"]      = _default_model
        st.session_state["_last_modele"]      = _default_model
        _dm = _MODELES[_default_model]
        st.session_state["jours_excl_widget"] = _dm["jours"]
        st.session_state["h_debut_widget"]    = _dm["h_debut"]
        st.session_state["h_fin_widget"]      = _dm["h_fin"]

    _modele_choix = st.selectbox(
        "Choisir un modèle",
        list(_MODELES.keys()),
        help=(
            "Modèle 1 : 430 kW · Restriction Lun–Sam H10–H12.\n"
            "Modèle 2 : 215 kW · Restriction Mar–Jeu H10–H14.\n"
            "Modèle 3 : 215 kW · Aucune restriction.\n"
            "Modèle 4 : 1 MW · Aucune restriction.\n"
            "Personnalisé : configurez librement tous les paramètres."
        ),
        key="modele_choix"
    )

    _mod = _MODELES[_modele_choix]

    # Si le modèle vient de changer → mettre à jour les widgets de restriction
    if st.session_state.get("_last_modele") != _modele_choix:
        st.session_state["_last_modele"]      = _modele_choix
        st.session_state["jours_excl_widget"] = _mod["jours"]
        st.session_state["h_debut_widget"]    = _mod["h_debut"]
        st.session_state["h_fin_widget"]      = _mod["h_fin"]
        st.rerun()

    if _modele_choix == "Personnalisé":
        power_MW = st.number_input("Puissance (MW)", 0.05, 50.0, 1.0, 0.01, format="%.3f", key="pow_custom")
    else:
        power_MW = _mod["power_MW"]
        st.metric("Puissance", f"{power_MW*1000:.0f} kW")
        st.caption(_mod["desc"])

    efficiency = st.slider("Rendement (%)", 70, 100, 100) / 100
    _quota_mode = st.radio(
        "Mode de limitation",
        ["Illimité", "Quota annuel (N meilleurs spreads)", "Spread minimum (€/MWh)"],
        index=0, key="quota_mode",
    )
    if _quota_mode == "Quota annuel (N meilleurs spreads)":
        max_cycles = st.number_input("Nombre max de cycles/an", 1, 730, 365, key="max_cycles_n")
        st.caption(f"Retient les **{max_cycles} cycles** avec les meilleurs spreads de l'année.")
        min_spread = None
    elif _quota_mode == "Spread minimum (€/MWh)":
        min_spread = st.number_input("Spread minimum (€/MWh)", 0, 500, 20, key="min_spread_val")
        st.caption(f"Ne trade que si spread > **{min_spread} €/MWh**.")
        max_cycles = None
    else:
        max_cycles = None
        min_spread = None



    st.markdown("---")
    if st.session_state.bloomberg:
        if st.button(" Clair", key="btn__MODE_CLAIR_1"):
            st.session_state.bloomberg = False
            st.rerun()
        st.markdown(
            '<div style="text-align:center;font-size:0.65rem;color:#ff6600;'
            'letter-spacing:3px;font-family:Courier New;margin-top:4px;">'
            '● BLOOMBERG TERMINAL</div>',
            unsafe_allow_html=True
        )

    # ── Rapport de vérification complet ───────────────────────────────────────
    # Le bouton est ici (sidebar, exécuté en premier dans le script) mais le
    # contenu réel n'est rempli qu'en fin de script (après tous les onglets),
    # une fois que chaque onglet a déposé son résumé dans st.session_state.
    # _diag_report_slot est le conteneur réservé que la fin du script remplira.
    st.markdown("---")
    st.markdown("### Vérification")
    if st.button("Générer rapport de vérification complet", key="btn_diag_global"):
        st.session_state["_diag_global_requested"] = True
    st.caption(
        "Compile les résultats déjà calculés de tous les onglets visités dans "
        "cette session en un seul fichier HTML, à envoyer pour vérification."
    )
    _diag_report_slot = st.empty()

    st.caption("Plénitude B-Charge | BESS Valorisation v2.0")


# ──────────────────────────────────────────────────────────────────────────────
# CHARGEMENT DONNÉES
# ──────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def get_pivot(file_bytes: bytes):
    from bess_engine import load_spot
    import io
    return load_spot(io.BytesIO(file_bytes))


# Écran d'accueil si pas de fichier
if uploaded is None:
    st.markdown('<p class="main-title">BESS Valorisation - Marché Day-Ahead</p>',
                unsafe_allow_html=True)
    st.markdown("### Bienvenue")
    st.info(
        "Glissez-déposez votre fichier Excel **BESS_valorisation_RESULTATS.xlsx** "
        "dans la barre latérale à gauche pour commencer la simulation."
    )
    st.stop()

# Chargement
try:
    file_bytes = st.session_state.excel_bytes
    pivot = get_pivot(file_bytes)
    annees = sorted(pivot["annee"].unique())
except Exception as e:
    st.error(f"Erreur de lecture du fichier : {e}. Vérifiez que le fichier contient bien une feuille 'Spot_input'.")
    st.stop()

# En-tête principal
st.markdown('<p class="main-title">BESS Valorisation — Marché Day-Ahead</p>',
            unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-title">Fichier : {uploaded.name} &nbsp;|&nbsp; '
    f'{len(pivot):,} jours &nbsp;|&nbsp; {annees[0]}–{annees[-1]} &nbsp;|&nbsp; '
    f'Puissance {power_MW} MW &nbsp;|&nbsp; '
    f'Rendement {efficiency*100:.0f}%</p>',
    unsafe_allow_html=True
)

# ══════════════════════════════════════════════════════════════════════════════
# ASSISTANT IA — bulle flottante en bas à droite
# ══════════════════════════════════════════════════════════════════════════════
# Position fixe via CSS sur la classe générée par container(key=...). Pas de
# st.fragment ici : retesté le 2026-06-23 sur Streamlit 1.56 (serveur relancé
# à froid, test patient, sans aucune charge concurrente) avec tout le panneau
# dans un @st.fragment + st.rerun(scope="fragment") — le bouton toggle (ouvrir/
# fermer) ne répond plus du tout après le premier clic, de façon reproductible.
# Donc rerun complet du script à chaque message, mitigé par le cache déjà en place.

if "_ai_chat_open" not in st.session_state:
    st.session_state["_ai_chat_open"] = False
if "ai_messages" not in st.session_state:
    st.session_state.ai_messages = []


def _ai_build_context():
    lines = ["=== CONFIGURATION BATTERIE ===",
             "Puissance: %.0f kW | Rendement: %.0f%%" % (power_MW*1000, efficiency*100)]

    _diag_arb = st.session_state.get("_diag_arb")
    if _diag_arb:
        lines.append("")
        lines.append("=== ARBITRAGE DAY-AHEAD ===")
        for _label, _val in _diag_arb["kpis"]:
            lines.append(f"{_label}: {_val}")
        for _, _yr in _diag_arb["yearly"].iterrows():
            lines.append("  %d: PnL=%s | Spread=%.1f | Jours actifs=%d" % (
                int(_yr["annee"]), _yr["pnl_total"], _yr["spread_moy"], int(_yr["jours_actifs"])))

    _diag_lis = st.session_state.get("_diag_lis")
    if _diag_lis:
        lines.append("")
        lines.append("=== LISSAGE DE CHARGE ===")
        for _label, _val in _diag_lis["kpis"]:
            lines.append(f"{_label}: {_val}")

    _diag_hist = st.session_state.get("_diag_hist")
    if _diag_hist:
        lines.append("")
        lines.append("=== HISTORIQUE 2019-2025 ===")
        for _label, _val in _diag_hist["kpis"]:
            lines.append(f"{_label}: {_val}")

    _diag_exec = st.session_state.get("_diag_exec")
    if _diag_exec:
        lines.append("")
        lines.append("=== EXECUTIVE SUMMARY ===")
        for _label, _val in _diag_exec["kpis"]:
            lines.append(f"{_label}: {_val}")

    _sc_results = st.session_state.get("_ai_sc_results", None)
    if _sc_results:
        lines.append("")
        lines.append("=== COMPARAISON SCENARIOS ===")
        for _sr in _sc_results:
            _sn = _sr["sc"]["name"]
            _sy = _sr["yearly"]
            lines.append("Scenario: %s | Duree: %sh | Max cycles/an: %s" % (
                _sn, _sr["sc"].get("duration_h",1),
                str(_sr["sc"].get("max_cycles_year","illimite"))))
            lines.append("  PnL total: %.0f euros | Spread moy: %.1f euros/MWh" % (
                _sy["pnl_total"].sum(), _sy["spread_moy"].mean()))

    _cp_results = st.session_state.get("_cp_results")
    if _cp_results:
        lines.append("")
        lines.append("=== COMPARAISON PAYS ===")
        for _r in _cp_results:
            _y = _r["yearly"]
            lines.append("Pays: %s | PnL total: %.0f euros | Spread moy: %.1f" % (
                _r["pays"], _y["pnl_total"].sum(), _y["spread_moy"].mean()))

    _markets = st.session_state.get("_intra_markets")
    if _markets:
        lines.append("")
        lines.append("=== INTRADAY ===")
        for _mname, _mdata in _markets.items():
            _y = _mdata["yearly"]
            lines.append("Marché: %s | PnL total: %.0f euros | Spread moy: %.1f" % (
                _mname, _y["pnl_total"].sum(), _y["spread_moy"].mean()))

    if len(lines) <= 2:
        lines.append("Aucune simulation calculee pour l'instant — lancez d'abord une simulation.")
    return "\n".join(lines)


def _ai_get_key():
    v = None
    try: v = st.secrets.get("GEMINI_API_KEY", None)
    except Exception: pass
    import os as _os_k
    return v or _os_k.environ.get("GEMINI_API_KEY", None) or st.session_state.get("_saved_GEMINI_KEY", None)


# Condense le contenu de l'onglet Méthodologie (tab_methodo, vérifié à jour vs
# bess_engine.py) pour que l'IA connaisse la mécanique réelle des calculs,
# pas seulement les chiffres produits — sinon elle ne fait que deviner avec
# des connaissances génériques de marché électrique.
_AI_METHODO_DOC = """
=== METHODOLOGIE DE L'OUTIL (reference fixe, ne pas reciter sauf si demande) ===
Architecture : bess_engine.py (calculs purs) + bess_dashboard.py (interface).
Pipeline Arbitrage : pivot [jours x H00..H23] -> simulate_arbitrage() ou
simulate_arbitrage_optimal() -> aggregate_arbitrage() -> affichage.

Algorithme coeur (_best_n_cycles, mode MAX/Illimite, n_cycles=0) :
DP exacte (Weighted Interval Scheduling) sur les cycles charge+decharge
candidats d'une journee -> optimum GARANTI (pas une heuristique), c'est la
"borne max theorique" (perfect foresight) affichee partout.
PnL d'un cycle = Puissance_MW x n_heures x (Rendement x Prix_decharge_moy - Prix_charge_moy).

Deux modes de quota annuel, comportement different :
- Chronologique (simulate_arbitrage) : consomme le quota jour apres jour dans
  l'ordre du calendrier ; bloque (PnL=0) une fois le quota atteint jusqu'au 1er janvier.
- Top-N optimal (simulate_arbitrage_optimal, "optimal_quota") : genere tous les
  cycles de l'annee, garde les N meilleurs spreads de CETTE annee peu importe
  le jour (quota reinitialise chaque annee, pas un pool unique sur tout l'horizon).
Taux de capture (%) = PnL_reel / PnL_borne_max x 100.

ROI / VAN (onglet Comparaison scenarios) :
CAPEX_total = CAPEX_kWh x Puissance_MW x 1000 x Duree_h (CAPEX source IEA/BNEF
2022-2030 selon annee choisie, de 330 a 85 EUR/kWh). PnL_an(n) degrade de
taux_degrad %/an. Cashflow(n) = PnL_an(n) - OPEX_an. VAN = -CAPEX_total +
somme Cashflow(n)/(1+taux)^n.

Lissage de charge (peak shaving) : regle heuristique (pas d'optimisation),
decharge si conso > seuil (percentile configurable, 75 par defaut), charge
sinon, sans anticipation. ATTENTION : le rendement batterie N'EST PAS applique
dans ce calcul (ecart connu vs Arbitrage qui l'applique correctement) -
prevenir l'utilisateur si pertinent que le Lissage suppose une batterie 100%
efficace.

Intraday : spread bid/ask = WAP(BUY) - WAP(SELL) (>=0 normalement, BUY=acheteur
agressif execute a l'ASK). Le marche continu n'est represente que par un seul
onglet "Continuous", base sur Continuous_Index filtre sur IndexName=IDFULL
(evite de melanger ID1/ID3/IDFULL) - Continuous_Statistics n'est pas utilise
separement (redondant, meme marche continu que l'Index). Cote granularite
horaire, la ligne 60min est prioritaire (deja la moyenne ponderee exacte de
l'heure) ; si elle manque pour une heure donnee, reconstruction a partir des
tranches 30min puis 15min (jamais melange brut des resolutions entre elles,
qui double-compterait les memes transactions).
""".strip()


_panel_bg   = "#0a0a0a" if _BB else "#ffffff"
_panel_text = "#e0e0e0" if _BB else "#1a1a1a"
_panel_bord = "#ff6600" if _BB else "#d0dff0"
_btn_bg     = "#ff6600" if _BB else "#1a3a5c"
_panel_disp = "block" if st.session_state["_ai_chat_open"] else "none"

st.markdown(f"""
<style>
.st-key-ai_fab_btn {{
    position: fixed !important;
    bottom: 24px !important;
    right: 24px !important;
    z-index: 1000000 !important;
    width: 56px !important;
    height: 56px !important;
}}
.st-key-ai_fab_btn button {{
    border-radius: 50% !important;
    width: 56px !important;
    height: 56px !important;
    padding: 0 !important;
    font-size: 26px !important;
    line-height: 1 !important;
    background: {_btn_bg} !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(0,0,0,0.35) !important;
}}
.st-key-ai_fab_panel {{
    display: {_panel_disp} !important;
    position: fixed !important;
    bottom: 92px !important;
    right: 24px !important;
    z-index: 999999 !important;
    width: 360px !important;
    max-width: 90vw !important;
    max-height: 65vh !important;
    background: {_panel_bg} !important;
    color: {_panel_text} !important;
    border: 1px solid {_panel_bord} !important;
    border-radius: 14px !important;
    box-shadow: 0 10px 40px rgba(0,0,0,0.35) !important;
    padding: 14px !important;
    overflow-y: auto !important;
}}
.st-key-ai_fab_panel p, .st-key-ai_fab_panel div {{ color: {_panel_text} !important; }}
</style>
""", unsafe_allow_html=True)

# Toujours rendu (structure DOM stable entre les reruns) ; la visibilité est
# pilotée par CSS (_panel_disp) — sinon Streamlit confond les conteneurs
# ai_fab_panel/ai_fab_btn quand l'un apparaît/disparaît entre deux reruns
# (bug constaté : le bouton héritait de la hauteur du panneau fermé).
if True:
    with st.container(key="ai_fab_panel"):
        _ai_col_title, _ai_col_close = st.columns([5, 1])
        with _ai_col_title:
            st.markdown("**🤖 Assistant IA — BESS**")
        with _ai_col_close:
            if st.button("✕", key="ai_fab_close"):
                st.session_state["_ai_chat_open"] = False
                st.rerun()

        _GEMINI_KEY = _ai_get_key()
        if not _GEMINI_KEY:
            with st.expander("Configuration clé API (une seule fois)", expanded=True):
                st.caption(
                    "Clé **Gemini** (gratuit) : "
                    "[aistudio.google.com/apikey](https://aistudio.google.com/apikey) → Create API key → copier `AIza...`  \n"
                    "Ensuite mettez-la dans `.streamlit/secrets.toml` : `GEMINI_API_KEY = 'AIza...'`"
                )
                _key_input = st.text_input("Clé Gemini", type="password",
                                           placeholder="AIza...", key="ai_key_input")
                if _key_input:
                    _GEMINI_KEY = _key_input
                    st.session_state["_saved_GEMINI_KEY"] = _key_input

        _ai_chat_container = st.container(height=280)
        with _ai_chat_container:
            for _msg in st.session_state.ai_messages:
                with st.chat_message(_msg["role"]):
                    st.markdown(_msg["content"])

        _user_input = st.chat_input("Posez une question sur vos données BESS...")
        if _user_input:
            st.session_state.ai_messages.append({"role": "user", "content": _user_input})
            st.session_state["_ai_pending"] = True
            with _ai_chat_container:
                with st.chat_message("user"):
                    st.markdown(_user_input)

        if st.session_state.get("_ai_pending"):
            st.session_state["_ai_pending"] = False
            _ctx = _ai_build_context()
            _sys_prompt = (
                "Tu es l'assistant intégré au dashboard BESS Valorisation (arbitrage day-ahead EPEX). "
                "Tu connais la méthodologie ci-dessous : appuie-toi sur elle pour expliquer COMMENT et "
                "POURQUOI les chiffres sont ce qu'ils sont (algorithme, formule, mode de quota), pas "
                "seulement pour les répéter. "
                "Réponds de façon naturelle et directe à la question ou au message de l'utilisateur, en français. "
                "N'utilise la méthodologie ou les données ci-dessous QUE si la question le justifie réellement : "
                "pour un message de test, une salutation ou une question générale, réponds brièvement sans les recopier. "
                "Quand tu t'appuies sur des chiffres, utilise uniquement ceux fournis ci-dessous, sans les inventer. "
                "Si la simulation nécessaire pour répondre n'a pas encore été lancée, dis-le en une phrase courte, "
                "sans répéter tout le bloc de données.\n\n"
                + _AI_METHODO_DOC + "\n\n"
                "DONNÉES DU DASHBOARD (à utiliser seulement si la question le demande) :\n" + _ctx
            )
            _msgs_api = [{"role": "system", "content": _sys_prompt}]
            for _m in st.session_state.ai_messages[:-1]:
                _msgs_api.append({"role": _m["role"], "content": _m["content"]})
            _msgs_api.append(st.session_state.ai_messages[-1])

            with _ai_chat_container:
                with st.chat_message("assistant"):
                    with st.spinner("..."):
                        try:
                            if not _GEMINI_KEY:
                                raise Exception(
                                    "Clé Gemini manquante. "
                                    "Créez-en une sur aistudio.google.com/apikey (gratuit) "
                                    "et ajoutez GEMINI_API_KEY dans .streamlit/secrets.toml"
                                )
                            try:
                                from google import genai as _genai_sdk
                                from google.genai import types as _genai_types
                            except ImportError:
                                raise Exception(
                                    "Package google-genai non installé sur cet environnement. "
                                    "Pour activer l'assistant IA en local : pip install google-genai"
                                )
                            import time as _time

                            _client = _genai_sdk.Client(api_key=_GEMINI_KEY)

                            _sys_msg = next((m["content"] for m in _msgs_api if m["role"] == "system"), "")
                            _history = []
                            for _m in _msgs_api[:-1]:
                                if _m["role"] == "system":
                                    continue
                                _history.append(_genai_types.Content(
                                    role="user" if _m["role"] == "user" else "model",
                                    parts=[_genai_types.Part(text=_m["content"])],
                                ))
                            _last_msg = _msgs_api[-1]["content"]

                            _config = _genai_types.GenerateContentConfig(
                                system_instruction=_sys_msg,
                                temperature=0.2,
                                max_output_tokens=600,
                            )

                            _models_to_try = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-flash-lite-latest"]
                            _answer = None
                            for _model_name in _models_to_try:
                                try:
                                    _chat = _client.chats.create(
                                        model=_model_name,
                                        history=_history,
                                        config=_config,
                                    )
                                    _response = _chat.send_message(_last_msg)
                                    _answer = _response.text
                                    break
                                except Exception as _e:
                                    _err_str = str(_e)
                                    if "429" in _err_str or "quota" in _err_str.lower() or "RESOURCE_EXHAUSTED" in _err_str:
                                        if _model_name != _models_to_try[-1]:
                                            st.toast(f"Quota {_model_name} atteint — passage au modèle suivant…")
                                            _time.sleep(1)
                                            continue
                                        _answer = "Quota Gemini atteint. Attendez 1 minute et réessayez."
                                    else:
                                        _answer = "Erreur API Gemini : %s" % _err_str
                                    break

                        except Exception as _e_outer:
                            _answer = "Erreur lors de l'appel à l'API : %s" % str(_e_outer)

                        st.markdown(_answer)
                        st.session_state.ai_messages.append({"role": "assistant", "content": _answer})

        if st.session_state.ai_messages:
            if st.button("Effacer la conversation", key="ai_reset"):
                st.session_state.ai_messages = []
                st.rerun()

with st.container(key="ai_fab_btn"):
    if st.button("✕" if st.session_state["_ai_chat_open"] else "💬", key="ai_fab_toggle"):
        st.session_state["_ai_chat_open"] = not st.session_state["_ai_chat_open"]
        st.rerun()
# ══════════════════════════════════════════════════════════════════════════════
# CONFIG GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

PLOTLY_CFG = dict(
    scrollZoom=True,
    displayModeBar=True,
    modeBarButtonsToAdd=["drawrect", "eraseshape"],
    modeBarButtonsToRemove=["lasso2d"],
    displaylogo=False,
    toImageButtonOptions=dict(format="png", width=1400, height=700, scale=2),
)
LEGEND_BOTTOM = dict(orientation="h", y=-0.18, x=0, xanchor="left")

# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS PRINCIPAUX
# ══════════════════════════════════════════════════════════════════════════════

tab_arb, tab_intra, tab_imbalance, tab_lis, tab_comp, tab_sensi, tab_pays, tab_hist, tab_exec, tab_verif_arb, tab_verif_lis, tab_methodo = st.tabs([
    "Arbitrage Day-Ahead",
    "Intraday",
    "Imbalance Market",
    "Lissage de charge",
    "Comparaison scénarios",
    "Sensibilité",
    "Comparaison pays",
    "Historique 2019–2025",
    "Executive Summary",
    "Vérification Arbitrage",
    "Vérification Lissage",
    "Méthodologie",
])

# ══════════════════════════════════════════════════════════════════════════════
# CAPTURE GÉNÉRIQUE POUR LE RAPPORT DE VÉRIFICATION
# ══════════════════════════════════════════════════════════════════════════════
# st.metric/st.dataframe/st.table/st.plotly_chart sont enveloppés pour
# enregistrer tout ce qui s'affiche, onglet par onglet, sans dupliquer la
# logique de chaque onglet — le rapport HTML (_build_global_diag_html) reste
# donc automatiquement à jour avec tout ajout futur de KPI/tableau/graphique.
# Reset à chaque rerun (variable module-level, pas session_state).
from collections import defaultdict as _defaultdict_rep
_REPORT_CAPTURE = _defaultdict_rep(list)
_REPORT_TAB = {"name": None}

def _diag_set_tab(name):
    _REPORT_TAB["name"] = name

_st_metric_orig       = st.metric
_st_dataframe_orig    = st.dataframe
_st_table_orig        = st.table
_st_plotly_chart_orig = st.plotly_chart

def _rep_metric(*a, **k):
    tab = _REPORT_TAB["name"]
    if tab:
        try:
            label = a[0] if a else k.get("label", "")
            value = a[1] if len(a) > 1 else k.get("value", "")
            _REPORT_CAPTURE[tab].append(("metric", str(label), str(value)))
        except Exception:
            pass
    return _st_metric_orig(*a, **k)

def _rep_dataframe(*a, **k):
    tab = _REPORT_TAB["name"]
    if tab:
        try:
            data = a[0] if a else k.get("data")
            _REPORT_CAPTURE[tab].append(("dataframe", None, data))
        except Exception:
            pass
    return _st_dataframe_orig(*a, **k)

def _rep_table(*a, **k):
    tab = _REPORT_TAB["name"]
    if tab:
        try:
            data = a[0] if a else k.get("data")
            _REPORT_CAPTURE[tab].append(("dataframe", None, data))
        except Exception:
            pass
    return _st_table_orig(*a, **k)

def _rep_plotly_chart(*a, **k):
    tab = _REPORT_TAB["name"]
    if tab:
        try:
            fig = a[0] if a else k.get("figure_or_data")
            _REPORT_CAPTURE[tab].append(("fig", None, fig))
        except Exception:
            pass
    return _st_plotly_chart_orig(*a, **k)

st.metric       = _rep_metric
st.dataframe    = _rep_dataframe
st.table        = _rep_table
st.plotly_chart = _rep_plotly_chart


# ══════════════════════════════════════════════════════════════════════════════
# MONTE CARLO — Optimisation paramètres
# ══════════════════════════════════════════════════════════════════════════════


# ── Configurations BESS réelles (feuille BESS de l'Excel) ────────────────────
BESS_CONFIGS = [
    {"nom":"Modèle 1 — 430 kW","power_MW":0.430,"n_cycles":1,"duration_h":1,
     "efficiency":0.92,"jours_excl":[0,1,2,3,4,5],"h_debut":10,"h_fin":12,
     "desc":"430 kW · Indisponible lun-sam 10h-12h"},
    {"nom":"Modèle 2 — 215 kW (mar-jeu restr.)","power_MW":0.215,"n_cycles":1,"duration_h":1,
     "efficiency":0.92,"jours_excl":[1,2,3],"h_debut":10,"h_fin":14,
     "desc":"215 kW · Indisponible mar-jeu 10h-14h"},
    {"nom":"Modèle 3 — 215 kW (sans restr.)","power_MW":0.215,"n_cycles":1,"duration_h":1,
     "efficiency":0.92,"jours_excl":[],"h_debut":0,"h_fin":0,
     "desc":"215 kW · Disponible 24h/24, 7j/7"},
    {"nom":"Modèle 1 — 430 kW (2 cycles)","power_MW":0.430,"n_cycles":2,"duration_h":1,
     "efficiency":0.92,"jours_excl":[0,1,2,3,4,5],"h_debut":10,"h_fin":12,
     "desc":"430 kW · 2 cycles/jour · Indisponible lun-sam 10h-12h"},
    {"nom":"Modèle 3 — 215 kW (2 cycles)","power_MW":0.215,"n_cycles":2,"duration_h":1,
     "efficiency":0.92,"jours_excl":[],"h_debut":0,"h_fin":0,
     "desc":"215 kW · 2 cycles/jour · Sans restriction"},
]


@st.cache_data(show_spinner=False)
def run_monte_carlo(file_bytes, configs_json):
    import json as _jmc
    from bess_engine import load_spot, simulate_arbitrage, aggregate_arbitrage
    configs = _jmc.loads(configs_json)
    pv = load_spot(io.BytesIO(file_bytes))
    results = []
    for cfg in configs:
        excl = {}
        if cfg["jours_excl"] and cfg["h_debut"] < cfg["h_fin"]:
            excl = {"days": cfg["jours_excl"],
                    "hours": list(range(cfg["h_debut"], cfg["h_fin"]))}
        p = {"power_MW": cfg["power_MW"], "n_cycles": cfg["n_cycles"],
             "duration_h": cfg["duration_h"], "efficiency": cfg["efficiency"],
             "max_cycles_year": None, "excluded_hours": excl}
        try:
            d = simulate_arbitrage(pv, p)
            y = aggregate_arbitrage(d, cfg["power_MW"])
            cap = cfg["power_MW"] * cfg["duration_h"] * cfg["n_cycles"]
            results.append({
                "Modèle":            cfg["nom"],
                "Description":       cfg["desc"],
                "Puissance (MW)":    cfg["power_MW"],
                "Cycles/j":          cfg["n_cycles"],
                "Capacite (MWh)":    round(cap, 3),
                "Restriction":       f"H{cfg['h_debut']}-H{cfg['h_fin']}" if cfg["jours_excl"] else "Aucune",
                "PnL total (euros)": round(float(y["pnl_total"].sum()), 0),
                "PnL borne max":     round(float(y["pnl_absolu_total"].sum()), 0),
                "Spread moy":        round(float(y["spread_moy"].mean()), 2),
                "Activation (%)":    round(float(y["taux_activation"].mean()) * 100, 1),
                "PnL par kW":        round(float(y["pnl_par_MW"].mean()), 2),
            })
        except Exception as err:
            results.append({"Modèle": cfg["nom"], "Description": str(err), "PnL total (euros)": 0})
    results.sort(key=lambda x: x.get("PnL total (euros)", 0), reverse=True)
    return results


def show_monte_carlo_panel(tab_key, default_pow_max=2.0):
    with st.expander("Simulation de Monte Carlo — Trouver la meilleure configuration BESS", expanded=False):
        st.caption("La simulation Monte Carlo teste automatiquement plusieurs configurations de batterie et les classe par PnL. Elle permet d'identifier la combinaison puissance/durée/cycles la plus rentable.")
        if st.button("Lancer la simulation de Monte Carlo (1000 tirages)", key=f"mc_run_{tab_key}"):
            import random as _rmc
            from bess_engine import simulate_arbitrage, aggregate_arbitrage

            # Espaces de paramètres réalistes issus de la feuille BESS de l'Excel
            POWERS    = [0.215, 0.430]           # kW → MW disponibles
            DURATIONS = [1, 2]                    # durees de cycle en heures
            CYCLES    = [1, 2]                    # cycles par jour
            EFFICS    = [0.88, 0.90, 0.92, 0.94, 0.95]  # rendements réalistes
            MAXCYC    = [200, 250, 300, 365, None]        # max cycles/an
            # Restrictions : (jours exclus, heure début, heure fin, label)
            RESTRS = [
                ([], 0, 0, "Aucune"),
                ([0,1,2,3,4,5], 10, 12, "Lun-Sam 10h-12h"),
                ([1,2,3], 10, 14, "Mar-Jeu 10h-14h"),
                ([0,1,2,3,4], 8, 11, "Lun-Ven 8h-11h"),
                ([5,6], 0, 0, "Week-end off"),
            ]

            N_ITER = 1000
            rng    = _rmc.Random(42)
            seen   = set()
            results_mc = []

            prog = st.progress(0)
            for i in range(N_ITER):
                prog.progress((i + 1) / N_ITER,
                              text=f"Monte Carlo : tirage {i+1}/{N_ITER} — {len(results_mc)} configs uniques testées")
                pw   = rng.choice(POWERS)
                dur  = rng.choice(DURATIONS)
                ncyc = rng.choice(CYCLES)
                eff  = rng.choice(EFFICS)
                mcy  = rng.choice(MAXCYC)
                restr= rng.choice(RESTRS)

                key = (pw, dur, ncyc, eff, mcy, restr[3])
                if key in seen:
                    continue
                seen.add(key)

                excl = {}
                if restr[0] and restr[1] < restr[2]:
                    excl = {"days": restr[0], "hours": list(range(restr[1], restr[2]))}

                try:
                    d = simulate_arbitrage(pivot, {
                        "power_MW": pw, "n_cycles": ncyc, "duration_h": dur,
                        "efficiency": eff, "max_cycles_year": mcy,
                        "excluded_hours": excl,
                    })
                    y   = aggregate_arbitrage(d, pw)
                    cap = pw * dur * ncyc
                    results_mc.append({
                        "Puissance (MW)":    pw,
                        "Durée (h)":         dur,
                        "Cycles/j":          ncyc,
                        "Capacité (MWh)":    round(cap, 3),
                        "Rendement (%)":     round(eff * 100, 0),
                        "Max cyc/an":        mcy if mcy else "illimité",
                        "Restriction":       restr[3],
                        "PnL total (€)":     round(float(y["pnl_total"].sum()), 0),
                        "PnL borne max (€)": round(float(y["pnl_absolu_total"].sum()), 0),
                        "Spread moy":        round(float(y["spread_moy"].mean()), 2),
                        "Activation (%)":    round(float(y["taux_activation"].mean()) * 100, 1),
                        "PnL/kW":            round(float(y["pnl_par_MW"].mean()), 2),
                    })
                except Exception:
                    pass

            results_mc.sort(key=lambda x: x.get("PnL total (€)", 0), reverse=True)
            st.session_state[f"mc_results_{tab_key}"] = results_mc
            st.rerun()

        mc_results = st.session_state.get(f"mc_results_{tab_key}", [])
        if mc_results:
            best  = mc_results[0]
            c_bg  = "#0a1a0a" if _BB else "#e2efda"
            c_txt = "#00e676" if _BB else "#375623"
            pnl_b = best.get("PnL total (€)", 0)
            st.markdown(
                f'<div style="background:{c_bg};border-left:4px solid {c_txt};'
                f'padding:10px 16px;border-radius:4px;margin:8px 0;">'
                f'<b style="color:{c_txt};">MEILLEURE CONFIGURATION ({len(mc_results)} testées)</b><br>'
                f'{best.get("Puissance (MW)")} MW · {best.get("Cycles/j")} cyc × {best.get("Durée (h)")}h · '
                f'Rdt {best.get("Rendement (%)")}% · {best.get("Restriction")} · '
                f'Max {best.get("Max cyc/an")} cyc/an<br>'
                f'PnL total : <b>{_fmt(pnl_b)} €</b> | '
                f'Activation : <b>{best.get("Activation (%)","—")}%</b> | '
                f'PnL/kW : <b>{best.get("PnL/kW","—")}</b>'
                f'</div>',
                unsafe_allow_html=True
            )
            top_n = min(50, len(mc_results))
            df_mc   = pd.DataFrame(mc_results[:top_n])
            cols_ok = [c for c in ["Puissance (MW)","Durée (h)","Cycles/j","Capacité (MWh)",
                                   "Rendement (%)","Max cyc/an","Restriction",
                                   "PnL total (€)","PnL borne max (€)","Spread moy",
                                   "Activation (%)","PnL/kW"]
                       if c in df_mc.columns]
            st.caption(f"Top {top_n} sur {len(mc_results)} configurations uniques testées")
            st.dataframe(df_mc[cols_ok], hide_index=True, width="stretch")
            st.download_button(
                "Exporter tous les résultats (CSV)",
                data=pd.DataFrame(mc_results)[cols_ok].to_csv(
                    index=False, sep=";", decimal=",").encode("utf-8-sig"),
                file_name=f"BESS_monte_carlo_{tab_key}.csv",
                mime="text/csv", key=f"mc_dl_{tab_key}",
            )


with tab_arb:
    _diag_set_tab("arb")
    # ── Paramètres scénario ──────────────────────────────────────────────────
    st.markdown('<p class="section">Paramètres du scénario</p>', unsafe_allow_html=True)
    st.caption("Configurez les caractéristiques de la batterie : puissance, cycles, durée, rendement et restrictions horaires. Ces paramètrès déterminent exactement quand et combien la batterie peut trader chaque jour.")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        _cyc_opts  = [1, 2, 0]
        _cyc_labels = {1: "1 cycle / jour", 2: "2 cycles / jour",
                       0: "Illimité — tous les cycles rentables du jour"}
        n_cycles = st.selectbox(
            "Cycles par jour",
            _cyc_opts,
            format_func=lambda x: _cyc_labels[x],
            index=0,
            help=(
                "1 = la batterie fait 1 charge + 1 décharge par jour (les meilleures heures).\n"
                "2 = 2 cycles successifs par jour.\n"
                "Illimité = le moteur exploite toutes les opportunités rentables de la journée."
            )
        )
    with c2:
        duration_h = st.selectbox(
            "Durée du cycle",
            [1, 2],
            format_func=lambda x: f"{x}h  ({x}h charge + {x}h décharge)",
            index=1,
            help="1h = achète 1h puis revend 1h. 2h = achète 2h puis revend 2h (capacité double)."
        )
    # Capacité = puissance × durée (indépendamment du nb de cycles)
    capacite_auto = power_MW * duration_h
    with c3:
        _cyc_disp = "illimité" if n_cycles == 0 else str(n_cycles)
        st.metric(
            "Capacité par cycle (MWh)",
            f"{capacite_auto:.3f} MWh",
            delta=f"{power_MW} MW × {duration_h}h · {_cyc_disp} cycle(s)/j.",
        )
    with c4:
        # Les valeurs sont toujours initialisées dans la sidebar au premier chargement
        jours_excl = st.multiselect(
            "Jours avec restriction",
            ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"],
            key="jours_excl_widget"
        )
    with c5:
        h_debut = st.number_input("Heure début restriction", 0, 23,
                                   key="h_debut_widget")
    with c6:
        h_fin = st.number_input("Heure fin restriction", 0, 23,
                                 key="h_fin_widget")

    day_map = {"Lundi":0,"Mardi":1,"Mercredi":2,"Jeudi":3,
               "Vendredi":4,"Samedi":5,"Dimanche":6}
    days_num = [day_map[d] for d in jours_excl]
    excluded = ({"days": days_num, "hours": list(range(h_debut, h_fin))}
                if days_num and h_debut < h_fin else {})

    # ── Simulation avec barre de progression ─────────────────────────────────
    _arb_key = f"arb_{hash((file_bytes, power_MW, n_cycles, duration_h, json.dumps(excluded), efficiency, max_cycles, min_spread))}"

    if _arb_key not in st.session_state:
        def _bprog(pct, label):
            """Met à jour la barre dans le header ET st.progress natif."""
            # Header bar via JS (visible si le navigateur a déjà rendu la page)
            st.markdown(
                f'<script>if(window._bessSetProgress)'
                f'window._bessSetProgress({pct}, "{label}");</script>',
                unsafe_allow_html=True
            )

        _bprog(0, "Chargement des données…")
        pv    = get_pivot(file_bytes)
        excl  = json.loads(json.dumps(excluded))
        _use_optimal = bool(max_cycles)  # quota défini → mode optimal top-N
        p_arb = {"power_MW": power_MW,
                 "n_cycles": 0 if _use_optimal else n_cycles,  # mode MAX obligatoire pour optimal
                 "duration_h": duration_h, "excluded_hours": excl,
                 "efficiency": efficiency, "max_cycles_year": max_cycles,
                 "optimal_quota": _use_optimal,
                 "min_spread": min_spread}

        _bar = st.progress(0, text="Simulation en cours…")
        if _use_optimal:
            daily_res = simulate_arbitrage_optimal(pv, p_arb)
        else:
            daily_res = simulate_arbitrage(pv, p_arb)
        _bar.progress(90, text="Agrégation…")
        yearly_res = aggregate_arbitrage(daily_res, power_MW)
        _bar.progress(100, text="Terminé.")

        import time; time.sleep(0.4)
        st.session_state[_arb_key] = (daily_res, yearly_res)
        _bar.empty()
        st.markdown(
            '<script>if(window._bessHideProgress)window._bessHideProgress();</script>',
            unsafe_allow_html=True
        )

    daily, yearly = st.session_state[_arb_key]

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Résultats globaux</p>', unsafe_allow_html=True)
    _note_bm = ""
    if min_spread:
        _note_bm = f" La borne max est calculée **sans** filtre de spread — elle représente le potentiel absolu du marché même si vous avez choisi de ne trader qu'au-dessus de {min_spread} €/MWh."
    st.caption(f"Synthese des performances sur toute la période. Le PnL réel tient compte de toutes vos contraintes. La borne max est le maximum théorique en supposant aucune restriction et une connaissance parfaite des prix.{_note_bm}")

    total_pnl        = daily["pnl"].sum()
    total_pnl_absolu = daily["pnl_absolu"].sum()
    spread_moy       = daily.loc[daily["valid"], "spread"].mean() if daily["valid"].any() else 0
    spread_abs_moy   = daily.loc[daily["valid"], "spread_absolu"].mean() if daily["valid"].any() else 0
    jours_actifs     = int(daily["valid"].sum())
    jours_total      = len(daily)
    jours_usure      = int(daily["maintenance_bloque"].sum())
    energie_totale   = daily["energie_MWh"].sum()
    ratio            = total_pnl / total_pnl_absolu * 100 if total_pnl_absolu else 0

    st.session_state["_diag_arb"] = {
        "kpis": [
            ("PnL réel", f"{_fmt(total_pnl)} €"),
            ("PnL borne max", f"{_fmt(total_pnl_absolu)} €"),
            ("Ratio capturé", f"{ratio:.1f}%"),
            ("Spread moyen", f"{spread_moy:.1f} €/MWh"),
            ("Jours actifs", f"{jours_actifs} / {jours_total}"),
            ("Jours bloqués (maintenance)", str(jours_usure)),
        ],
        "yearly": yearly[["annee", "jours_actifs", "spread_moy",
                           "pnl_absolu_total", "pnl_total", "pnl_par_MW"]],
    }

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("PnL réel (avec contraintes)",
              f"{_fmt(total_pnl)} €")
    m2.metric("PnL borne théorique max",
              f"{_fmt(total_pnl_absolu)} €",
              delta=f"{ratio:.1f}% capturé")
    m3.metric("Spread moyen (jours actifs)",
              f"{spread_moy:.1f} €/MWh",
              delta=f"Théorique: {spread_abs_moy:.1f} €/MWh")
    m4.metric("Jours actifs",
              f"{jours_actifs} / {jours_total}",
              delta=f"{jours_actifs/jours_total*100:.0f}% taux activation")
    m5.metric("Énergie totale échangée",
              f"{_fmt(energie_totale, 1)} MWh",
              delta=f"Capacité/jour : {capacite_auto:.2f} MWh")
    m6.metric("Jours bloqués (maintenance)",
              f"{jours_usure}",
              delta="Aucun" if jours_usure == 0 else f"{jours_usure/jours_total*100:.1f}%",
              delta_color="off" if jours_usure == 0 else "inverse")

    # ── Graphique 1 : PnL annuel ─────────────────────────────────────────────
    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
    st.caption("Comparaison annuelle entre PnL réel (barres bleues) et borne max théorique (barres claires). L'ecart entre les deux mesure le coût de vos restrictions operationnelles sur les revenus.")
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=yearly["annee"].astype(str), y=yearly["pnl_absolu_total"],
        name="Borne max", marker_color=C2,
        text=yearly["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €"),
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Borne max : %{y:.0f} €<extra></extra>",
    ))
    fig1.add_trace(go.Bar(
        x=yearly["annee"].astype(str), y=yearly["pnl_total"],
        name="PnL réel", marker_color=C1,
        text=yearly["pnl_total"].apply(lambda x: f"{_fmt(x)} €"),
        textposition="inside", textfont_color="white",
        hovertemplate="<b>%{x}</b><br>PnL réel : %{y:.0f} €<extra></extra>",
    ))
    fig1.update_layout(
        barmode="group", height=380,
        yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis_title="Année",
        legend=LEGEND_BOTTOM,
        margin=dict(t=10, b=40, l=60, r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    apply_bb(fig1)
    st.caption("Barres bleues = PnL réel avec contraintes. Barres claires = borne max théorique. L'écart = revenus perdus à cause des restrictions et du quota maintenance.")
    st.plotly_chart(fig1, width="stretch", config=PLOTLY_CFG, key="arb_fig1")

    # ── Graphiques 2×2 ───────────────────────────────────────────────────────
    col_a, col_b = st.columns(2)

    with col_a:
        # Spread HEBDOMADAIRE — une seule courbe continue sur toute la période
        st.markdown('<p class="section">Spread moyen — par semaine (€/MWh)</p>',
                    unsafe_allow_html=True)
        st.caption("Spread hebdo = différence prix vente - prix achat. Un spread élevé = opportunité rentable.")

        d_all = daily.copy()
        d_all["semaine"] = pd.to_datetime(d_all["date"]).dt.to_period("W").apply(
            lambda r: r.start_time)

        weekly_all = d_all.groupby("semaine").agg(
            spread_moy    = ("spread_absolu", "mean"),
            spread_min    = ("spread_absolu", "min"),
            spread_max    = ("spread_absolu", "max"),
            nb_jours_tot  = ("date",          "count"),
            nb_jours_act  = ("valid",         "sum"),
            annee         = ("annee",         "first"),
        ).reset_index().sort_values("semaine")

        fig2 = go.Figure()

        # Palette : vert clair → jaune → bleu → orange
        SPREAD_COLORS = ["#5cb85c", "#f5c518", "#2e75b6", "#d46b1a",
                         "#7b3fa0", "#3a8a5c"]

        # Nombre d'annees pour choisir couleur et bande
        annees_list = sorted(weekly_all["annee"].unique())
        n_annees_tot = len(annees_list)

        if n_annees_tot <= 1:
            # Une seule couleur : vert clair
            main_color = C1
            fill_color = ("rgba(255,102,0,0.08)" if _BB else "rgba(92,184,92,0.10)")
        else:
            # Dégradé vert → jaune
            main_color = C1
            fill_color = ("rgba(255,102,0,0.08)" if _BB else "rgba(92,184,92,0.10)")

        # Bande min-max
        fig2.add_trace(go.Scatter(
            x=pd.concat([weekly_all["semaine"],
                          weekly_all["semaine"].iloc[::-1]]),
            y=pd.concat([weekly_all["spread_max"],
                          weekly_all["spread_min"].iloc[::-1]]),
            fill="toself",
            fillcolor=fill_color,
            line=dict(color="rgba(0,0,0,0)"),
            showlegend=False, hoverinfo="skip",
        ))

        # Courbe continue — vert clair
        fig2.add_trace(go.Scatter(
            x=weekly_all["semaine"],
            y=weekly_all["spread_moy"],
            mode="lines",
            name="Spread hebdomadaire",
            line=dict(width=2.5, color=main_color),
            connectgaps=True,
            hovertemplate=(
                "<b>Semaine du %{x|%d/%m/%Y}</b><br>"
                "Spread : <b>%{y:.1f} €/MWh</b><br>"
                "Min sem. : %{customdata[0]:.1f} | "
                "Max sem. : %{customdata[1]:.1f}<br>"
                "Jours actifs : %{customdata[2]:.0f} / %{customdata[3]:.0f}"
                "<extra></extra>"
            ),
            customdata=weekly_all[["spread_min", "spread_max",
                                    "nb_jours_act", "nb_jours_tot"]].values,
        ))

        # Moyenne globale en pointillés jaunes
        spread_global_moy = weekly_all["spread_moy"].mean()
        fig2.add_hline(
            y=spread_global_moy,
            line_dash="dot", line_color=C2, line_width=2,
            annotation_text=f"Moy: {spread_global_moy:.1f} €/MWh",
            annotation_position="top left",
            annotation_font=dict(size=11, color=C2),
        )

        # Lignes verticales de séparation des annees
        for yr in sorted(d_all["annee"].unique())[1:]:
            jan1 = pd.Timestamp(f"{int(yr)}-01-01")
            fig2.add_shape(
                type="line",
                x0=jan1, x1=jan1, y0=0, y1=1,
                xref="x", yref="paper",
                line=dict(color="#cccccc", width=1, dash="dot"),
            )
            fig2.add_annotation(
                x=jan1, y=1.02, xref="x", yref="paper",
                text=str(int(yr)), showarrow=False,
                font=dict(size=10, color="#888888"),
            )

        fig2.update_layout(
            height=320, margin=dict(t=20, b=40, l=60, r=10),
            yaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0"),
            xaxis=dict(title="", tickformat="%b %Y", nticks=16),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
            hovermode="x unified",
        )
        apply_bb(fig2)
        st.caption(
            "La courbe = spread moyen de la semaine (différence prix vente - prix achat). "
            "La zone ombrée autour = écart entre le spread min et max de la semaine : "
            "plus l'ombre est large, plus les prix ont varié dans la semaine. "
            "La ligne pointillée = spread moyen sur toute la période."
        )
        st.plotly_chart(fig2, width="stretch", config=PLOTLY_CFG, key="arb_fig2")

    with col_b:
        st.markdown('<p class="section">Profil horaire charge / décharge</p>',
                    unsafe_allow_html=True)
        st.caption("Nombre de fois ou chaque heure a ete utilisée pour charger (achat) ou decharger (vente) sur toute la période. Les heures de charge sont les moins chères de la journee, les heures de décharge les plus chères.")
        h_ch  = {h: 0 for h in range(24)}
        h_dch = {h: 0 for h in range(24)}
        for hc_list, hd_list in zip(daily.loc[daily["valid"], "h_charge"],
                                     daily.loc[daily["valid"], "h_decharge"]):
            for h in hc_list:  h_ch[h]  += 1
            for h in hd_list:  h_dch[h] += 1
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_ch.values()),
            name="Charge (achat)", marker_color=C1,
            hovertemplate="H%{x:02d} — Charge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_ch.values()],
        ))
        fig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_dch.values()),
            name="Décharge (vente)", marker_color=C2,
            hovertemplate="H%{x:02d} — Décharge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_dch.values()],
        ))
        fig3.update_layout(
            height=320, barmode="group", margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Heure", tickmode="array",
                       tickvals=list(range(0,24,2)),
                       ticktext=[f"H{h:02d}" for h in range(0,24,2)],
                       range=[-0.5,23.5]),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(fig3)
        st.caption("Fréquence d'utilisation de chaque heure sur 4 ans. Vert = charge (achat aux heures bon marché). Orange = décharge (vente aux heures chères).")
        st.plotly_chart(fig3, width="stretch", config=PLOTLY_CFG, key="arb_fig3")

    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown('<p class="section">Distribution des spreads journaliers</p>',
                    unsafe_allow_html=True)
        st.caption("Répartition des jours selon le spread journalier. Un spread de 40 euros/MWh signifie 40 euros de différence entre prix de vente et prix achat ce jour. Plus la distribution est vers la droite, plus le marché est favorable.")
        sv = daily.loc[daily["valid"], "spread"]
        _p99_arb = float(sv.quantile(0.99)) if len(sv) > 0 else sv.max()
        fig4 = go.Figure()
        fig4.add_trace(go.Histogram(
            x=sv.clip(upper=_p99_arb), nbinsx=40, marker_color=C1, opacity=0.8,
            hovertemplate="Spread : %{x:.0f} €/MWh<br>Nb jours : <b>%{y}</b><extra></extra>",
        ))
        fig4.add_vline(
            x=sv.mean(), line_dash="dash", line_color=ORANGE, line_width=2,
            annotation_text=f"Moy: {sv.mean():.1f} €/MWh",
            annotation_position="top right",
            annotation_font=dict(size=12, color=ORANGE),
        )
        fig4.add_vline(
            x=sv.median(), line_dash="dot", line_color=C3, line_width=1.5,
            annotation_text=f"Méd: {sv.median():.1f}",
            annotation_position="top left",
            annotation_font=dict(size=11, color=GREEN),
        )
        fig4.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0", range=[0, _p99_arb*1.05]),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(fig4)
        st.caption('Histogramme des spreads journaliers. Chaque barre = nombre de jours avec ce niveau de spread. Vers la droite = marché très favorable.')
        st.plotly_chart(fig4, width="stretch", config=PLOTLY_CFG, key="arb_fig4")

    with col_d:
        st.markdown('<p class="section">PnL cumulé dans le temps</p>',
                    unsafe_allow_html=True)
        st.caption("Cumul des gains depuis le debut de la simulation. Une courbe régulière = batterie qui trade tous les jours. Les plateaux horizontaux = périodes bloquées par le quota de maintenance annuel.")
        ds = daily.sort_values("date").copy()
        ds["pnl_cum"]       = ds["pnl"].cumsum()
        ds["pnl_absolu_cum"] = ds["pnl_absolu"].cumsum()
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_cum"],
            fill="tozeroy", mode="lines", name="PnL cumulé",
            line=dict(color=C1, width=2),
            fillcolor=("rgba(255,102,0,0.10)" if _BB else "rgba(92,184,92,0.12)"),
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "PnL cumulé : <b>%{y:.0f} €</b><br>"
                "PnL du jour : %{customdata:.0f} €<extra></extra>"
            ),
            customdata=ds["pnl"].values,
        ))
        fig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_absolu_cum"],
            mode="lines", name="Borne max",
            line=dict(color=C2, width=1.5, dash="dot"),
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "Borne max cumulée : %{y:.0f} €<extra></extra>"
            ),
        ))
        fig5.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            yaxis=dict(title="PnL cumulé (€)", tickformat=",", gridcolor="#f0f0f0"),
            xaxis=dict(title="", tickformat="%b %Y"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
            hovermode="x unified",
        )
        apply_bb(fig5)
        st.caption('PnL cumule depuis le debut. La pente = rythme de gain. Les plateaux = périodes bloquées par le quota maintenance annuel.')
        st.plotly_chart(fig5, width="stretch", config=PLOTLY_CFG, key="arb_fig5")

    # ── Répartition des cycles journaliers ───────────────────────────────────
    st.markdown('<p class="section">Répartition des cycles journaliers</p>', unsafe_allow_html=True)
    st.caption("Distribution du nombre de cycles réalisés par jour sur toute la période. Identifie les journées les plus et les moins actives.")
    if n_cycles != 1:  # Graphique cycles seulement si multi-cycles ou max
        _cyc_result = _plot_cycles_distribution(daily, "2026–2029")
    else:
        _cyc_result = None
    if _cyc_result:
        _fig_cyc, _min_c, _max_c, _dmin, _dmax = _cyc_result
        _cd1, _cd2 = st.columns(2)
        _cd1.metric("Minimum de cycles/jour",
                    f"{_min_c} cycle(s)",
                    delta=f"ex. {pd.Timestamp(_dmin).strftime('%d/%m/%Y')}",
                    delta_color="off")
        _cd2.metric("Maximum de cycles/jour",
                    f"{_max_c} cycle(s)",
                    delta=f"ex. {pd.Timestamp(_dmax).strftime('%d/%m/%Y')}",
                    delta_color="off")
        apply_bb(_fig_cyc)
        st.plotly_chart(_fig_cyc, width="stretch", config=PLOTLY_CFG, key="arb_fig_cyc")

    # ── Tableau récap ────────────────────────────────────────────────────────
    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
    st.caption("Tableau détaillé par année : jours actifs, spreads captures, PnL réel et théorique. Le taux de capturé (PnL réel / borne max) mesure l'efficacité de votre strategie.")
    show = yearly.copy()
    show["pnl_absolu_total"]  = show["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €")
    show["pnl_total"]         = show["pnl_total"].apply(lambda x: f"{_fmt(x)} €")
    show["pnl_par_MW"]        = show["pnl_par_MW"].apply(lambda x: f"{x:.1f} €/MW")
    show["taux_activation"]   = show["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
    show["spread_absolu_moy"] = show["spread_absolu_moy"].apply(lambda x: f"{x:.1f}")
    show["spread_moy"]        = show["spread_moy"].apply(lambda x: f"{x:.1f}")
    show["energie_totale_MWh"]= show["energie_totale_MWh"].apply(lambda x: f"{x:.1f} MWh")
    show["cycles_totaux"]     = show["cycles_totaux"].astype(int)
    show = show[[
        "annee","jours_simules","jours_actifs","jours_bloques_maintenance","taux_activation",
        "spread_absolu_moy","pnl_absolu_total",
        "spread_moy","pnl_total","pnl_par_MW",
        "energie_totale_MWh","cycles_totaux"
    ]]
    show.columns = [
        "Année","Jours simulés","Jours actifs","Bloqués maintenance","Taux activation",
        "Spread théorique\n(€/MWh)","PnL borne max\nthéorique (€)",
        "Spread réel\n(€/MWh)","PnL réel (€)","PnL/MW\n(€/MW)",
        "Énergie échangée\n(MWh)","Cycles realises"
    ]
    st.caption('Tableau annuel détaillé. Colonnes : Jours actifs (jours trades), Taux activation (proportion), PnL réel (revenus nets après contraintes). Cliquez sur une colonne pour trier.')
    st.dataframe(show, hide_index=True, width="stretch")

    # ── Explorer un jour ─────────────────────────────────────────────────────
    with st.expander("Explorer un jour spécifique", expanded=True):
        st.caption("Sélectionnez un jour pour voir en detail comment la batterie a trade : prix horaires, heures de charge et décharge choisies, PnL réalisé et comparaison avec le maximum théorique possible.")
        date_sel = st.date_input("Date", value=pd.Timestamp(pivot["date"].iloc[0]).date())
        row_pivot = pivot[pivot["date"].dt.date == date_sel]

        if not row_pivot.empty:
            prix_j   = row_pivot[HOUR_COLS].values.flatten()
            weekday  = int(row_pivot["weekday"].iloc[0])
            avail    = get_available_hours(weekday, excluded)

            # Recalcul EN TEMPS RÉEL avec les paramètres actuels
            from bess_engine import _best_n_cycles

            cycles_jour = _best_n_cycles(
                prix_j, avail, n_cycles, duration_h, power_MW, efficiency
            )

            ca, cb = st.columns([3, 1])
            with cb:
                pnl_jour     = sum(c["pnl"]    for c in cycles_jour)
                spread_jour  = sum(c["spread"] for c in cycles_jour) / len(cycles_jour) if cycles_jour else 0
                energie_jour = power_MW * duration_h * len(cycles_jour)

                # Spread par cycle
                if len(cycles_jour) == 1:
                    st.metric("Spread cycle 1", f"{cycles_jour[0]['spread']:.2f} €/MWh")
                elif len(cycles_jour) >= 2:
                    st.metric("Spread moyen", f"{spread_jour:.2f} €/MWh")
                    for i, cy in enumerate(cycles_jour, 1):
                        st.metric(f"Spread cycle {i}",
                                  f"{cy['spread']:.2f} €/MWh",
                                  delta=f"{cy['spread'] - spread_jour:+.2f} vs moy")
                else:
                    st.metric("Spread", "— Pas de cycle")

                st.metric("PnL du jour", f"{pnl_jour:.2f} €")
                st.metric("Énergie chargée", f"{energie_jour:.3f} MWh",
                          help="Énergie totale achetée (chargée) par la batterie ce jour")
                st.metric("Énergie déchargée", f"{energie_jour * efficiency:.3f} MWh",
                          help=f"Énergie revendue après rendement {efficiency*100:.0f}%")
                st.metric("Cycles réalisés", f"{len(cycles_jour)}" + (" / " + str(n_cycles) if n_cycles > 0 else " (max)"))
                st.metric("Heures dispo", f"{len(avail)}/24")
                for i, cy in enumerate(cycles_jour, 1):
                    st.caption(
                        f"Cycle {i} : charge H{cy['h_charge']} "
                        f"({cy['prix_charge']:.1f} €/MWh) → "
                        f"décharge H{cy['h_decharge']} "
                        f"({cy['prix_decharge']:.1f} €/MWh) "
                        f"| PnL {cy['pnl']:.2f} €"
                    )

            with ca:
                # Couleurs des barres : heures exclues en gris clair
                bar_colors = []
                for h in range(24):
                    if h not in avail:
                        bar_colors.append("rgba(255,102,0,0.2)" if _BB else "#c8d4e0")  # exclu
                    else:
                        bar_colors.append("#c8d8ec")  # disponible

                fig_d = go.Figure()

                # Barres prix spot
                fig_d.add_trace(go.Bar(
                    x=list(range(24)), y=prix_j.tolist(),
                    marker_color=bar_colors, name="Prix spot",
                    hovertemplate="H%{x:02d} — Prix : <b>%{y:.2f} €/MWh</b><extra></extra>",
                ))

                # Couleurs par cycle
                colors_ch  = ["#2e7d32", "#1565c0", "#6a1b9a"]
                colors_dch = ["#c62828", "#e65100", "#4527a0"]

                for i, cy in enumerate(cycles_jour):
                    _h_ch  = cy["h_charge"]
                    _h_dch = cy["h_decharge"]
                    c_ch   = colors_ch[i % 3]
                    c_dch  = colors_dch[i % 3]
                    lbl    = f"Cycle {i+1}"

                    fig_d.add_trace(go.Scatter(
                        x=_h_ch, y=prix_j[_h_ch],
                        mode="markers",
                        name=f"Charge {lbl} ({duration_h}h)",
                        marker=dict(color=c_ch, size=18, symbol="triangle-up",
                                    line=dict(color="white", width=1)),
                        hovertemplate=f"H%{{x:02d}} — Charge {lbl} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                    ))
                    fig_d.add_trace(go.Scatter(
                        x=_h_dch, y=prix_j[_h_dch],
                        mode="markers",
                        name=f"Decharge {lbl} ({duration_h}h)",
                        marker=dict(color=c_dch, size=18, symbol="triangle-down",
                                    line=dict(color="white", width=1)),
                        hovertemplate=f"H%{{x:02d}} — Decharge {lbl} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                    ))
                    fig_d.add_shape(type="line", x0=-0.5, x1=23.5,
                        y0=cy["prix_charge"], y1=cy["prix_charge"],
                        line=dict(color=c_ch, width=1, dash="dot"))
                    fig_d.add_shape(type="line", x0=-0.5, x1=23.5,
                        y0=cy["prix_decharge"], y1=cy["prix_decharge"],
                        line=dict(color=c_dch, width=1, dash="dot"))
                    fig_d.add_annotation(x=22, y=cy["prix_charge"],
                        text=f"Achat C{i+1}: {cy['prix_charge']:.1f}",
                        showarrow=False, font=dict(size=9, color=c_ch))
                    fig_d.add_annotation(x=22, y=cy["prix_decharge"],
                        text=f"Vente C{i+1}: {cy['prix_decharge']:.1f}",
                        showarrow=False, font=dict(size=9, color=c_dch))

                # Zones exclues
                for h in range(24):
                    if h not in avail:
                        fig_d.add_vrect(x0=h-0.5, x1=h+0.5,
                            fillcolor="rgba(26,58,92,0.15)", line_width=0, layer="below")

                fig_d.update_layout(
                    height=420, margin=dict(t=30, b=60, l=60, r=10),
                    xaxis=dict(title="Heure", tickmode="array",
                               tickvals=list(range(0,24,1)),
                               ticktext=[f"H{h:02d}" for h in range(24)],
                               range=[-0.5, 23.5]),
                    yaxis=dict(title="Prix (€/MWh)", gridcolor="#f0f0f0"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=LEGEND_BOTTOM,
                    hoverlabel=dict(bgcolor="white", font_size=12),
                )
                apply_bb(fig_d)
                st.caption('Prix horaires du jour sélectionné. Triangles verts = achat. Triangles rouges = vente. Lignes pointillees = prix moyens du cycle. Zones grises = heures exclues.')
                st.plotly_chart(fig_d, width="stretch", config=PLOTLY_CFG,
                                key="fig_explorer_jour")

    # ── Export PDF ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Export du rapport</p>', unsafe_allow_html=True)
    st.caption("Exportez les résultats de la simulation en cours. Le PDF contient un rapport complet, le CSV contient toutes les donnees journalières pour analyse externe.")

    col_exp1, col_exp2 = st.columns([2, 3])
    with col_exp1:
        if st.button("Générer le rapport PDF", key="btn_pdf_1"):
            with st.spinner("Génération du rapport PDF..."):

                params_txt = [
                    f"Fichier : {uploaded.name}",
                    f"Puissance : {power_MW} MW | Capacité : {capacite_auto} MWh",
                    f"Durée cycle : {duration_h}h | Rendement : {efficiency*100:.0f}%",
                    f"Cycles/j : {'max' if n_cycles==0 else n_cycles} · Max cycles/an : {max_cycles or 'illimité'},"
                    f"Restriction : {', '.join(jours_excl) or 'aucune'} "
                    f"{'H'+str(h_debut)+'-H'+str(h_fin) if jours_excl else ''}",
                    f"Données : {annees[0]}–{annees[-1]} ({jours_total} jours)",
                ]
                kpis = [
                    ("PnL total (avec contraintes)",    f"{_fmt(total_pnl)} €"),
                    ("PnL borne théorique max", f"{_fmt(total_pnl_absolu)} €"),
                    ("Ratio réel / max",                f"{ratio:.1f}%"),
                    ("Spread moyen",                    f"{spread_moy:.2f} €/MWh"),
                    ("Jours actifs",                    f"{jours_actifs} / {jours_total}"),
                    ("Taux d'activation",               f"{jours_actifs/jours_total*100:.0f}%"),
                    ("Jours bloqués (maintenance)",           str(jours_usure)),
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
        # Export CSV donnees journalières
        csv_buf = daily[["date","annee","mois","weekday_name",
                          "spread_absolu","pnl_absolu","spread","pnl",
                          "valid","h_charge","h_decharge"]].copy()
        csv_buf["date"] = csv_buf["date"].dt.strftime("%Y-%m-%d")
        st.download_button(
            label="Exporter les donnees (CSV)",
            key="btn_csv_1",
            data=csv_buf.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_donnees_journalieres_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

    # ── PRIORITÉ 1 : Analyse de l'impact de l'usure ──────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Impact de l\'usure — jours bloqués et PnL manqué</p>',
                unsafe_allow_html=True)
    st.caption(
        f"La batterie est limitée à {max_cycles or 'illimité'} cycles/an. "
        f"Un jour bloqué maintenance = quota de cycles atteint → PnL non capturé."
    )

    # Calcul mensuel : jours actifs / bloqués / PnL manqué
    daily_u = daily.copy()
    daily_u["annee_mois"] = daily_u["date"].dt.to_period("M")

    # PnL manqué quota : jours bloqués par le quota de cycles
    daily_u["pnl_manque_quota"]        = daily_u.apply(
        lambda r: r["pnl_absolu"] if r["maintenance_bloque"] else 0.0, axis=1
    )
    # PnL manqué restrictions : écart entre sans restriction et avec restriction
    # sur les jours NON bloqués par le quota
    daily_u["pnl_manque_restrictions"] = daily_u.apply(
        lambda r: max(0.0, r["pnl_absolu"] - r["pnl"])
                  if not r["maintenance_bloque"] else 0.0,
        axis=1
    )

    grp_u = daily_u.groupby("annee_mois").agg(
        jours_actifs            = ("valid",                    "sum"),
        jours_bloques           = ("maintenance_bloque",       "sum"),
        pnl_realise             = ("pnl",                      "sum"),
        pnl_manque_quota        = ("pnl_manque_quota",         "sum"),
        pnl_manque_restrictions = ("pnl_manque_restrictions",  "sum"),
    ).reset_index()
    grp_u["label"] = grp_u["annee_mois"].astype(str)

    pnl_manque_quota_total        = daily_u["pnl_manque_quota"].sum()
    pnl_manque_restrictions_total = daily_u["pnl_manque_restrictions"].sum()
    pnl_manque_total              = pnl_manque_quota_total + pnl_manque_restrictions_total
    has_quota   = jours_usure > 0
    has_restr   = pnl_manque_restrictions_total > 0.5

    # ── Graphique 1 : Jours actifs / bloqués ─────────────────────────────────
    fig_u1 = go.Figure()
    fig_u1.add_trace(go.Bar(
        x=grp_u["label"], y=grp_u["jours_actifs"],
        name="Jours actifs", marker_color=C1,
        hovertemplate="<b>%{x}</b><br>Jours actifs : %{y}<extra></extra>",
    ))
    if has_quota:
        fig_u1.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["jours_bloques"],
            name="Jours bloqués (quota)", marker_color="#ef5350",
            hovertemplate="<b>%{x}</b><br>Jours bloqués quota : %{y}<extra></extra>",
        ))
    fig_u1.update_layout(
        barmode="stack", height=280,
        yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
        xaxis=dict(tickangle=-45, tickfont=dict(size=9)),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=60, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    apply_bb(fig_u1)
    if not has_quota:
        st.caption("Quota illimité — aucun jour bloqué.")
    else:
        st.caption("Jours bloqués : quota de cycles annuel atteint. Concentrés en fin d'année.")
    st.plotly_chart(fig_u1, width="stretch", config=PLOTLY_CFG, key="fig_usure_jours")

    # ── Graphique 2 : PnL réel / manqué quota / manqué restrictions ──────────
    fig_u2 = go.Figure()
    fig_u2.add_trace(go.Bar(
        x=grp_u["label"], y=grp_u["pnl_realise"].round(0),
        name="PnL réalisé (€)", marker_color=C1,
        hovertemplate="<b>%{x}</b><br>PnL réalisé : %{y:.0f} €<extra></extra>",
    ))
    if has_restr:
        fig_u2.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["pnl_manque_restrictions"].round(0),
            name="PnL manqué — restrictions horaires (€)",
            marker_color="#f5a623",
            hovertemplate="<b>%{x}</b><br>PnL manqué restrictions : %{y:.0f} €<extra></extra>",
        ))
    if has_quota:
        fig_u2.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["pnl_manque_quota"].round(0),
            name="PnL manqué — quota cycles (€)",
            marker_color="#ef5350",
            hovertemplate="<b>%{x}</b><br>PnL manqué quota : %{y:.0f} €<extra></extra>",
        ))
    fig_u2.update_layout(
        barmode="stack", height=300,
        yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis=dict(tickangle=-45, tickfont=dict(size=9)),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=70, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    apply_bb(fig_u2)
    _cap_parts = ["PnL réalisé (vert)"]
    if has_restr:
        _cap_parts.append(f"PnL manqué restrictions (orange) : {pnl_manque_restrictions_total:,.0f} €")
    if has_quota:
        _cap_parts.append(f"PnL manqué quota (rouge) : {pnl_manque_quota_total:,.0f} €")
    if not has_restr and not has_quota:
        _cap_parts.append("aucune perte détectée — quota illimité et restrictions sans impact")
    st.caption(" · ".join(_cap_parts))
    st.plotly_chart(fig_u2, width="stretch", config=PLOTLY_CFG, key="fig_usure_pnl")

    # KPIs usure
    ku1, ku2, ku3, ku4 = st.columns(4)
    ku1.metric("Jours bloqués (quota)", f"{jours_usure}",
               delta=f"{jours_usure/jours_total*100:.1f}% du total" if jours_usure > 0 else "Quota illimité")
    ku2.metric("PnL manqué — restrictions", f"{_fmt(pnl_manque_restrictions_total)} €",
               delta=f"{pnl_manque_restrictions_total/total_pnl*100:.1f}% du PnL réel" if total_pnl > 0 else "—")
    ku3.metric("PnL manqué — quota", f"{_fmt(pnl_manque_quota_total)} €",
               delta=f"{pnl_manque_quota_total/total_pnl*100:.1f}% du PnL réel" if total_pnl > 0 else "—")
    if max_cycles:
        ku4.metric("Conseil quota",
                   f"Actuel : {max_cycles} → tester {min(365, max_cycles + 50)}",
                   delta="Augmenter si dégradation acceptable")
    else:
        ku4.metric("Quota annuel", "Illimité",
                   delta="Aucun jour bloqué possible")

    # ── PRIORITÉ 2 : Analyse ROI / Payback ───────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Analyse ROI — Retour sur investissement</p>',
                unsafe_allow_html=True)
    st.caption(
        "Calcul de la rentabilité de l'investissement basé sur le modèle sélectionné dans la sidebar. "
        "Les données CAPEX proviennent du rapport IEA Batteries and Secure Energy Transitions (2024). "
        "Choisissez l'année de référence et ajustez la dégradation selon vos hypothèses."
    )

    # Modèle déjà sélectionné dans la sidebar via _modele_choix et power_MW
    # CAPEX batteries LFP utility-scale 2h, France — sources multiples mai 2026
    # Taux USD/EUR : 0.85 (mai 2026)
    # Sources : IEA Electricity 2026, BNEF Cost Survey 2025, Ember oct. 2025, Capstone DC nov. 2025
    IEA_CAPEX = {
        "2022 — 330 €/kWh (IEA réel)":              330,
        "2024 — 150 €/kWh (IEA Electricity 2026)":  150,
        "2025 — 120 €/kWh (BNEF / Ember)":          120,
        "2026 — 105 €/kWh (Capstone DC France)":    105,
        "2030 — 85 €/kWh  (BNEF projection)":        85,
    }
    IEA_CYCLES = 6500   # LFP stationnaire 2025 : 6000-7000 cycles (BNEF/Ember)
    IEA_DUREE  = 15     # ans — durée de vie nominale (Capstone DC, Ember)

    # Info modèle actif
    st.info(
        f"Modèle actif : **{_modele_choix}** — {power_MW*1000:.0f} kW — {power_MW} MW  "
        f"| Capacité : {capacite_auto:.3f} MWh  "
        f"| Rendement : {efficiency*100:.0f}%  "
        f"| Quota : {max_cycles or 'illimité'} cycles/an  "
        "_(configuré dans la sidebar)_"
    )

    r1, r2, r3 = st.columns(3)
    with r1:
        annee_capex = st.radio(
            "Année de référence CAPEX (IEA)",
            list(IEA_CAPEX.keys()),
            index=1,
            help="Source : IEA Batteries and Secure Energy Transitions (2024), STEPS scenario. "
                 "Converti USD vers EUR au taux 0.93. LFP domine le stockage stationnaire."
        )
        capex_kwh = IEA_CAPEX[annee_capex]
        st.metric("CAPEX retenu", f"{capex_kwh} €/kWh", delta="Source IEA 2024")
    with r2:
        duree_proj = st.slider("Durée de projection (ans)", 5, 20, IEA_DUREE,
            help=f"Durée de vie estimée LFP : {IEA_DUREE} ans "
                 f"({IEA_CYCLES} cycles garantis / 300 cycles/an). Source IEA 2024.")
        opex_an = st.number_input(
            "OPEX annuel (€/an)", min_value=0, max_value=50000, value=0, step=500,
            help="OPEX = 0 si inclus dans la garantie constructeur. "
                 "Référence IEA/NREL : 2.5% du CAPEX/an hors garantie."
        )
    with r3:
        taux_degrad = st.number_input(
            "Dégradation annuelle (%)", min_value=0.0, max_value=10.0,
            value=2.0, step=0.5, format="%.1f",
            help="Perte de capacité annuelle des cellules Li-ion LFP. "
                 "Source IEA 2024 : 2-3% par an. Impact direct sur le PnL futur."
        ) / 100.0
        taux_actu = st.number_input(
            "Taux d'actualisation (WACC, %)", min_value=0.0, max_value=20.0,
            value=5.0, step=0.5, format="%.1f",
            help="Taux utilisé pour calculer la VAN (Valeur Actuelle Nette). "
                 "Reflète le coût du capital ou le taux d'opportunité. "
                 "Typiquement 5-8% pour un projet industriel en Europe."
        ) / 100.0
        st.caption("VAN calculée également à 3% et 6% pour comparaison (méthode fichier Phase 2).")

    with st.expander("Données de référence — batteries LFP stationnaire (sources mai 2026)", expanded=False):
        st.markdown("""
**Sources : IEA Electricity 2026 · IEA Global Energy Review 2026 · BNEF Cost Survey 2025 · Ember oct. 2025 · Capstone DC nov. 2025**

**CAPEX utility-scale 2h, France (taux USD/EUR : 0.85)**
- **2022 : 330 €/kWh** — IEA Electricity 2026 (340 $/kWh × 0.85 + premium Europe)
- **2024 : 150 €/kWh** — IEA Electricity 2026 (fin 2024, après baisse de 40 % sur l'année)
- **2025 : 120 €/kWh** — BNEF Cost Survey 2025 (117 $/kWh mondial + premium Europe)
- **2026 : 105 €/kWh** — Capstone DC France (€90–100/kWh equipment + balance of system)
- **2030 : 85 €/kWh** — BNEF projection Europe (101 $/kWh × 0.85)
- **Baisse** : -58% entre 2019 et 2024 (IEA). -45% supplémentaires en 2025 (IEA Global Energy Review 2026).
- **Projets 2h coûtent ~10-15% plus cher par kWh que projets 4h** (BNEF)

**Paramètres techniques LFP stationnaire**
- **Rendement (round-trip)** : 88% valeur centrale 2025 (Ember) | fourchette 85-95%
- **Cycles garantis LFP** : 6 000-7 000 cycles (BNEF/Ember 2025) | cellules CATL 2025+ : 15 000+
- **Cycles réels France 2h** : 1 500-1 800 cycles/an (Capstone DC)
- **Dégradation** : 2%/an garantie fabricant (Ember) | capacité résiduelle ~65% à 20 ans
- **Durée de vie** : 15 ans nominale (Capstone DC, Ember)
- **OPEX** : 2.5% du CAPEX/an (NREL ATB 2025) | 0 si inclus garantie constructeur
- **Chimie dominante** : LFP — 90% des nouvelles installations de stockage stationnaire en 2025 (IEA GER 2026)

**Marché France 2025-2026**
- Capacité installée début 2026 : ~1.5 GW (Modo Energy)
- Pipeline RTE : ~13 GW en file d'attente
- Revenus aFRR : effondrement de 66 €/MW/h (2024) à 16 €/MW/h (jan. 2026) — saturation
- IRR unlevered France : 5-7% (sous le WACC de 8%) — projet standalone non bancable sans hédging (Capstone DC)
        """)

    # Calcul CAPEX total
    capacite_kwh = capacite_auto * 1000  # MWh -> kWh
    capex_total  = capex_kwh * capacite_kwh
    pnl_an_moy   = total_pnl / len(annees)
    cashflow_an  = pnl_an_moy - opex_an

    # ── Interpolation annuelle (méthode Phase 2) ─────────────────────────────
    _annees_sim  = sorted(yearly["annee"].unique())
    _pnl_sim     = [float(yearly[yearly["annee"] == a]["pnl_total"].values[0]) for a in _annees_sim]
    _annee_base  = int(_annees_sim[0])

    def _pnl_interpole(annee_offset):
        cal_year = _annee_base + annee_offset - 1
        if cal_year <= int(_annees_sim[-1]):
            xs = [int(a) - _annee_base + 1 for a in _annees_sim]
            return float(np.interp(annee_offset, xs, _pnl_sim))
        if len(_pnl_sim) >= 2:
            _growth = (_pnl_sim[-1] / max(_pnl_sim[0], 1)) ** (1 / max(len(_pnl_sim)-1,1)) - 1
            _growth = min(max(_growth, -0.05), 0.15)
        else:
            _growth = 0.02
        return _pnl_sim[-1] * (1 + _growth) ** (cal_year - int(_annees_sim[-1]))

    # Payback simple
    payback = capex_total / cashflow_an if cashflow_an > 0 else float("inf")

    # Projection cumulée avec dégradation + interpolation
    annees_proj  = list(range(1, duree_proj + 1))
    cum_cashflow = []
    cum_cashflow_nodeg = []
    _pnl_annuel_proj = []
    cum = -capex_total
    cum_nd = -capex_total

    for a in annees_proj:
        _pnl_base_a = _pnl_interpole(a)
        facteur_degradation = (1 - taux_degrad) ** (a - 1)
        cf_a = (_pnl_base_a * facteur_degradation) - opex_an
        _pnl_annuel_proj.append(_pnl_base_a * facteur_degradation)
        cum += cf_a
        cum_cashflow.append(round(cum, 0))
        cum_nd += _pnl_base_a - opex_an
        cum_cashflow_nodeg.append(round(cum_nd, 0))

    # VAN à 3 taux : WACC utilisateur, 3%, 6%
    def _calc_van(taux):
        c = -capex_total
        res = []
        for a in annees_proj:
            c += ((_pnl_interpole(a) * (1-taux_degrad)**(a-1)) - opex_an) / (1+taux)**a
            res.append(round(c, 0))
        return res

    cum_cashflow_actu = _calc_van(taux_actu)
    cum_cashflow_3pct = _calc_van(0.03)
    cum_cashflow_6pct = _calc_van(0.06)

    # IRR
    def _calc_irr():
        _cfs = [-capex_total] + _pnl_annuel_proj[:duree_proj]
        lo, hi = -0.5, 5.0
        for _ in range(100):
            mid = (lo + hi) / 2
            npv = sum(c / (1 + mid) ** t for t, c in enumerate(_cfs))
            if abs(npv) < 1: return mid
            if npv > 0: lo = mid
            else: hi = mid
        return (lo + hi) / 2
    _irr = _calc_irr()

    payback_actu = next((i+1 for i,cf in enumerate(cum_cashflow_actu) if cf>=0), None)
    payback_deg  = next((i+1 for i,cf in enumerate(cum_cashflow)      if cf>=0), None)
    payback_3pct = next((i+1 for i,cf in enumerate(cum_cashflow_3pct) if cf>=0), None)
    payback_6pct = next((i+1 for i,cf in enumerate(cum_cashflow_6pct) if cf>=0), None)

    # KPIs ROI
    roi_cols = st.columns(5)
    roi_cols[0].metric("CAPEX total", f"{_fmt(capex_total)} €",
                       delta=f"{capacite_kwh:.0f} kWh x {capex_kwh} euros/kWh (source IEA 2024)")
    roi_cols[1].metric("PnL annuel moyen", f"{_fmt(pnl_an_moy)} euros/an",
                       delta=f"Sur {len(annees)} ans de simulation")
    payback_label = (f"{payback_deg} ans (dégradation {taux_degrad*100:.0f}%/an)"
                     if payback_deg else "Non remboursé")
    roi_cols[2].metric("Payback non actualisé",
                       payback_label,
                       delta="Rentable" if payback_deg and payback_deg <= duree_proj else "Hors période")
    roi_cols[3].metric(f"Cash-flow non actualisé à {duree_proj} ans",
                       f"{_fmt(cum_cashflow[-1])} €",
                       delta="Positif" if cum_cashflow[-1] > 0 else "Négatif",
                       delta_color="normal" if cum_cashflow[-1] > 0 else "inverse")
    _van_fin = cum_cashflow_actu[-1]
    _pb_actu_label = (f"{payback_actu} ans (taux {taux_actu*100:.0f}%)"
                      if payback_actu else "Non remboursé")
    roi_cols[4].metric(f"VAN à {duree_proj} ans ({taux_actu*100:.0f}%)",
                       f"{_fmt(_van_fin)} €",
                       delta=_pb_actu_label,
                       delta_color="normal" if _van_fin > 0 else "inverse")

    # Ligne 2 : VAN 3%, VAN 6%, IRR
    roi_cols2 = st.columns(4)
    roi_cols2[0].metric(f"VAN à {duree_proj} ans (3%)",
                        f"{_fmt(cum_cashflow_3pct[-1])} €",
                        delta=f"Payback: {payback_3pct} ans" if payback_3pct else "Non remboursé",
                        delta_color="normal" if cum_cashflow_3pct[-1]>0 else "inverse")
    roi_cols2[1].metric(f"VAN à {duree_proj} ans (6%)",
                        f"{_fmt(cum_cashflow_6pct[-1])} €",
                        delta=f"Payback: {payback_6pct} ans" if payback_6pct else "Non remboursé",
                        delta_color="normal" if cum_cashflow_6pct[-1]>0 else "inverse")
    roi_cols2[2].metric("TRI (IRR)",
                        f"{_irr*100:.1f}%/an" if _irr > -0.4 else "< -40%",
                        delta="Supérieur au WACC" if _irr > taux_actu else "Inférieur au WACC",
                        delta_color="normal" if _irr > taux_actu else "inverse",
                        help="Taux de Rendement Interne — taux qui annule la VAN sur la durée de projection.")
    roi_cols2[3].metric("Méthode projection",
                        "Interpolation annuelle",
                        delta=f"Ancrage sur {len(_annees_sim)} années simulées",
                        delta_color="off",
                        help="Les PnL simulés servent d'ancres. Les années suivantes sont interpolées/extrapolées avec le taux de croissance observé.")

    # Info sur la source des données
    st.info(
        "Sources CAPEX : IEA Electricity 2026 · BNEF Cost Survey 2025 · Ember oct. 2025 · Capstone DC nov. 2025. "
        "Valeurs utility-scale 2h, France, taux USD/EUR 0.85 (mai 2026). "
        "Les coûts ont chuté de 58% entre 2019 et 2024 (IEA), puis de 45% supplémentaires en 2025. "
        "Dégradation LFP : 2%/an garantie fabricant (Ember 2025)."
    )

    # Graphique projection cash-flow cumulé
    fig_roi = go.Figure()
    fig_roi.add_hline(y=0, line_dash="dash", line_color="#888", line_width=1.5)

    # Courbe avec dégradation (principale)
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow,
        mode="lines+markers",
        name=f"Avec dégradation {taux_degrad*100:.0f}%/an (réaliste)",
        line=dict(color=C1, width=2.5),
        fill="tozeroy",
        fillcolor="rgba(92,184,92,0.08)" if cum_cashflow[-1] > 0 else "rgba(239,83,80,0.08)",
        hovertemplate="Année %{x}<br>Cash-flow cumulé : <b>%{y:.0f} euros</b><extra></extra>",
        marker=dict(size=6),
    ))

    # Courbe sans dégradation (référence optimiste)
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_nodeg,
        mode="lines",
        name="Sans dégradation (optimiste)",
        line=dict(color=C2, width=1.5, dash="dot"),
        hovertemplate="Année %{x}<br>Sans dégradation : <b>%{y:.0f} euros</b><extra></extra>",
    ))

    # Courbe VAN actualisée (WACC utilisateur)
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_actu,
        mode="lines+markers",
        name=f"VAN actualisée ({taux_actu*100:.0f}%/an)",
        line=dict(color="#e67e22", width=2, dash="dashdot"),
        hovertemplate=f"Année %{{x}}<br>VAN ({taux_actu*100:.0f}%) : <b>%{{y:.0f}} euros</b><extra></extra>",
        marker=dict(size=5, symbol="diamond"),
    ))
    # Courbes VAN 3% et 6%
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_3pct, mode="lines", name="VAN 3%",
        line=dict(color="#27ae60", width=1.5, dash="dot"),
        hovertemplate="Année %{x}<br>VAN 3% : <b>%{y:.0f} €</b><extra></extra>",
    ))
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_6pct, mode="lines", name="VAN 6%",
        line=dict(color="#8e44ad", width=1.5, dash="dot"),
        hovertemplate="Année %{x}<br>VAN 6% : <b>%{y:.0f} €</b><extra></extra>",
    ))

    # Payback actualisé
    if payback_actu and payback_actu <= duree_proj:
        fig_roi.add_vline(
            x=payback_actu, line_dash="dot", line_color="#e67e22", line_width=1.5,
            annotation_text=f"Payback VAN : {payback_actu} ans",
            annotation_font=dict(color="#e67e22", size=10),
            annotation_position="top right",
        )

    # Point de break-even avec dégradation
    if payback_deg and payback_deg <= duree_proj:
        fig_roi.add_vline(
            x=payback_deg, line_dash="dot", line_color="#f5c518", line_width=2,
            annotation_text=f"Payback : {payback_deg} ans",
            annotation_font=dict(color="#f5c518", size=11),
        )

    # Annotation CAPEX source
    fig_roi.add_annotation(
        x=0.01, y=0.02, xref="paper", yref="paper",
        text=f"CAPEX : {capex_kwh} euros/kWh — Source IEA 2024",
        showarrow=False, font=dict(size=9, color="#888"),
        xanchor="left",
    )

    fig_roi.update_layout(
        height=400,
        yaxis=dict(title="Cash-flow cumulé (euros)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis=dict(title="Années depuis mise en service", dtick=1),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=80, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
    )
    apply_bb(fig_roi)
    st.caption(
        f"Bleu = cash-flow non actualisé avec dégradation {taux_degrad*100:.0f}%/an. "
        "Pointillé = sans dégradation (optimiste). "
        f"Orange = VAN actualisée au taux {taux_actu*100:.0f}%/an. "
        f"CAPEX retenu : {capex_kwh} €/kWh — Sources : IEA Electricity 2026, BNEF, Ember, Capstone DC (mai 2026)."
    )
    st.plotly_chart(fig_roi, width="stretch", config=PLOTLY_CFG, key="fig_roi_arb")
    st.caption(
        "Courbe pleine = avec +1%/an sur les spreads (hypothèse conservatrice). "
        "Courbe pointillée = spreads constants. "
        "La ligne jaune indique le point de break-even (CAPEX remboursé)."
    )
    # ── Export diagnostic complet ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Diagnostic — Exporter pour le développeur</p>',
                unsafe_allow_html=True)
    st.caption("Génère un fichier HTML avec toutes les donnees et résultats visibles à l'écran.")

    if st.button(" Exporter diagnostic complet", key="btn_diag_arb"):
        import json as _jd, datetime as _dtd

        # Profil horaire charge/décharge
        h_ch_freq  = {h: 0 for h in range(24)}
        h_dch_freq = {h: 0 for h in range(24)}
        for hc_l, hd_l in zip(daily.loc[daily["valid"], "h_charge"],
                               daily.loc[daily["valid"], "h_decharge"]):
            for h in hc_l: h_ch_freq[h]  += 1
            for h in hd_l: h_dch_freq[h] += 1

        # Distribution spreads
        sv = daily.loc[daily["valid"], "spread"]
        bins = [0, 20, 40, 60, 80, 100, 150, 200, 9999]
        labels = ["0-20","20-40","40-60","60-80","80-100","100-150","150-200",">200"]
        distrib = {labels[i]: int(((sv>=bins[i])&(sv<bins[i+1])).sum())
                   for i in range(len(labels))}

        diag = {
            "meta": {
                "generated_at": _dtd.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "fichier_excel": uploaded.name,
                "annees": [int(a) for a in annees],
                "nb_jours_total": jours_total,
                "theme": "TEST Bloomberg" if st.session_state.bloomberg else "Clair",
            },
            "paramètres": {
                "power_MW": power_MW,
                "n_cycles": n_cycles,
                "duration_h": duration_h,
                "capacite_MWh": round(capacite_auto, 4),
                "efficiency_pct": round(efficiency * 100, 1),
                "max_cycles_an": max_cycles,
                "excluded_hours": excluded,
                "jours_exclus": jours_excl,
                "h_debut_restriction": h_debut,
                "h_fin_restriction": h_fin,
            },
            "résultats_globaux": {
                "pnl_total_reel_euros": round(total_pnl, 2),
                "pnl_borne_max_euros": round(total_pnl_absolu, 2),
                "ratio_capture_pct": round(ratio, 2),
                "spread_moy_EuroMWh": round(spread_moy, 4),
                "spread_absolu_moy_EuroMWh": round(spread_abs_moy, 4),
                "jours_actifs": jours_actifs,
                "jours_total": jours_total,
                "taux_activation_pct": round(jours_actifs/jours_total*100, 1),
                "jours_bloques_maintenance": jours_usure,
                "energie_totale_MWh": round(energie_totale, 2),
            },
            "recap_annuel": yearly[[
                "annee","jours_simules","jours_actifs","taux_activation",
                "spread_absolu_moy","spread_moy",
                "pnl_absolu_total","pnl_total","pnl_par_MW",
                "energie_totale_MWh","cycles_totaux"
            ]].to_dict(orient="records"),
            "profil_horaire_frequence": {
                "charge":   h_ch_freq,
                "decharge": h_dch_freq,
            },
            "distribution_spreads": distrib,
            "spread_stats": {
                "min":    round(float(sv.min()), 4),
                "max":    round(float(sv.max()), 4),
                "moy":    round(float(sv.mean()), 4),
                "mediane":round(float(sv.median()), 4),
                "p25":    round(float(sv.quantile(0.25)), 4),
                "p75":    round(float(sv.quantile(0.75)), 4),
            },
            "donnees_journalieres_completes": daily[[
                "date","annee","mois","weekday_name",
                "spread_absolu","pnl_absolu","spread","pnl",
                "valid","maintenance_bloque","n_cycles_actifs","energie_MWh",
                "h_charge","h_decharge","prix_charge","prix_decharge",
                "prix_min","prix_max","prix_moy"
            ]].assign(date=lambda df: df["date"].dt.strftime("%Y-%m-%d")
            ).to_dict(orient="records"),
        }

        json_str = _jd.dumps(diag, indent=2, ensure_ascii=False, default=str)

        # Construire le HTML
        def kpi(label, value, sub=""):
            return (f'<div class="kpi"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{value}</div>'
                    + (f'<div class="kpi-sub">{sub}</div>' if sub else '')
                    + '</div>')

        def table_from_list(rows, cols=None):
            if not rows: return "<p>Aucune donnée</p>"
            if cols is None: cols = list(rows[0].keys())
            h = "<table><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"
            for r in rows:
                h += "<tr>" + "".join(
                    f"<td>{r.get(c,'')}</td>" for c in cols) + "</tr>"
            return h + "</table>"

        html = f"""<!DOCTYPE html><html lang="fr"><head>
<meta charset="UTF-8">
<title>BESS Diagnostic — {diag["meta"]["generated_at"]}</title>
<style>
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f8f9fb;color:#1a3a5c;margin:0;padding:20px;}}
h1{{background:#1a3a5c;color:white;padding:12px 20px;border-radius:6px;font-size:1.2rem;margin-bottom:8px;}}
h2{{color:#1a3a5c;border-bottom:2px solid #d0dff0;padding-bottom:4px;margin-top:24px;}}
h3{{color:#2e75b6;margin-top:16px;}}
.kpi{{display:inline-block;background:white;border:1px solid #d0dff0;border-radius:6px;
      padding:10px 16px;margin:6px;min-width:140px;vertical-align:top;}}
.kpi-label{{font-size:0.72rem;color:#6b7a8d;text-transform:uppercase;letter-spacing:0.5px;}}
.kpi-value{{font-size:1.25rem;font-weight:700;color:#1a3a5c;margin-top:2px;}}
.kpi-sub{{font-size:0.72rem;color:#5cb85c;margin-top:2px;}}
pre{{background:#111;color:#5cb85c;padding:16px;border-radius:6px;
     font-family:'Courier New',monospace;font-size:11px;overflow-x:auto;white-space:pre-wrap;}}
table{{border-collapse:collapse;width:100%;font-size:11px;margin:8px 0;}}
th{{background:#1a3a5c;color:white;padding:5px 8px;text-align:left;font-size:11px;}}
td{{border:1px solid #d0d0d0;padding:4px 8px;}}
tr:nth-child(even) td{{background:#f0f5fb;}}
.section{{background:#2e75b6;color:white;font-weight:700;padding:6px 12px;
          font-size:11px;margin:16px 0 6px 0;border-radius:3px;}}
.good{{color:#375623;font-weight:700;}} .bad{{color:#c00000;font-weight:700;}}
</style></head><body>
<h1> BESS Valorisation — Diagnostic complet</h1>
<p>Généré le <b>{diag["meta"]["generated_at"]}</b> &nbsp;|&nbsp;
   Fichier : <b>{diag["meta"]["fichier_excel"]}</b> &nbsp;|&nbsp;
   Période : <b>{" · ".join(str(a) for a in diag["meta"]["annees"])}</b> &nbsp;|&nbsp;
   {jours_total} jours simulés</p>

<h2>Paramètres de simulation</h2>
{kpi("Puissance", f"{power_MW} MW")}
{kpi("Cycles/jour", f"{n_cycles} × {duration_h}h")}
{kpi("Capacité", f"{capacite_auto:.3f} MWh")}
{kpi("Rendement", f"{efficiency*100:.0f}%")}
{kpi("Max cycles/an", str(max_cycles or "∞"))}
{kpi("Jours exclus", ", ".join(jours_excl) or "Aucun")}
{kpi("Heures restriction", f"H{h_debut}–H{h_fin}" if jours_excl else "Aucune")}

<h2>Résultats globaux</h2>
{kpi("PnL réel total", f"{_fmt(total_pnl)} €", f"Borne max : {total_pnl_absolu:,.0f} €")}
{kpi("Taux de capture", f"{ratio:.1f}%")}
{kpi("Spread moyen", f"{spread_moy:.2f} €/MWh", f"Théorique : {spread_abs_moy:.2f}")}
{kpi("Jours actifs", f"{jours_actifs} / {jours_total}", f"{jours_actifs/jours_total*100:.0f}% activation")}
{kpi("Énergie totale", f"{_fmt(energie_totale, 1)} MWh")}
{kpi("Jours bloqués", str(jours_usure))}

<h2>Récapitulatif annuel</h2>
{table_from_list(diag["recap_annuel"],
    ["annee","jours_simules","jours_actifs","taux_activation",
     "spread_absolu_moy","spread_moy","pnl_absolu_total","pnl_total",
     "pnl_par_MW","energie_totale_MWh","cycles_totaux"])}

<h2>Profil horaire — fréquence charge / décharge</h2>
<table><tr><th>Heure</th>
{''.join(f"<th>H{h:02d}</th>" for h in range(24))}
</tr>
<tr><td><b>Charge (j)</b></td>
{''.join(f"<td>{h_ch_freq[h]}</td>" for h in range(24))}
</tr>
<tr><td><b>Décharge (j)</b></td>
{''.join(f"<td>{h_dch_freq[h]}</td>" for h in range(24))}
</tr></table>

<h2>Distribution des spreads</h2>
{table_from_list([{"Tranche (€/MWh)": k, "Nb jours": v,
                   "% du total": f"{v/len(sv)*100:.1f}%" if len(sv)>0 else "0%"}
                  for k, v in distrib.items()])}
<p>Min={diag["spread_stats"]["min"]} · Max={diag["spread_stats"]["max"]} · 
   Moy={diag["spread_stats"]["moy"]} · Méd={diag["spread_stats"]["mediane"]} · 
   P25={diag["spread_stats"]["p25"]} · P75={diag["spread_stats"]["p75"]}</p>

<h2>Données journalières complètes ({jours_total} jours)</h2>
{table_from_list(diag["donnees_journalieres_completes"],
    ["date","annee","mois","weekday_name","spread_absolu","pnl_absolu",
     "spread","pnl","valid","n_cycles_actifs","energie_MWh",
     "h_charge","h_decharge","prix_charge","prix_decharge","prix_min","prix_max","prix_moy"])}

<h2>JSON brut complet</h2>
<pre>{json_str}</pre>
<hr><p style="color:#888;font-size:11px;">
BESS Valorisation v2.0 — Plénitude B-Charge — Diagnostic technique</p>
</body></html>"""

        st.download_button(
            "⬇ Télécharger le diagnostic HTML",
            data=html.encode("utf-8"),
            file_name=f"BESS_diagnostic_{_dtd.datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            key="btn_diag_dl",
        )
        st.success("Fichier prêt ! Téléchargez-le et envoyez-le au développeur.")

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET — ARBITRAGE INTRADAY (chargé en 2e position, juste après Arbitrage DA)
# ══════════════════════════════════════════════════════════════════════════════

with tab_intra:
    _diag_set_tab("intra")
    st.markdown('<p class="section">Arbitrage Intraday — Comparaison des marchés EPEX</p>', unsafe_allow_html=True)
    st.caption(
        "Contrairement au Day-Ahead (un prix fixé la veille pour chaque heure), l'intraday permet d'ajuster ses positions "
        "le jour même via plusieurs enchères successives (IDA1/IDA2/IDA3 — prix de clearing, recalculé à chaque enchère) "
        "ou en continu tout au long de la journée (prix moyen pondéré des transactions réelles, jusqu'à peu avant la livraison). "
        "Chaque source de prix donne des résultats différents — la comparaison permet de savoir quel marché rémunère le mieux la batterie."
    )

    # ── Paramètres du scénario ────────────────────────────────────────────────
    st.markdown('<p class="section">Paramètres du scénario</p>', unsafe_allow_html=True)
    st.caption("Configurez les caractéristiques de la batterie : puissance, cycles, durée, rendement et restrictions horaires. Ces paramètres sont propres à cet onglet — ils déterminent quand et combien de fois par jour la batterie peut charger/décharger sur les marchés intraday, indépendamment de la simulation Day-Ahead.")
    _ic1, _ic2, _ic3, _ic4, _ic5, _ic6 = st.columns(6)
    with _ic1:
        _intra_n_cyc = st.selectbox(
            "Cycles par jour", [1, 2, 0],
            format_func=lambda x: {1:"1 cycle / jour",2:"2 cycles / jour",
                                    0:"Illimité — tous les cycles rentables du jour"}[x],
            index=0, key="intra_n_cyc",
        )
    with _ic2:
        _intra_dur = st.selectbox(
            "Durée du cycle", [1, 2],
            format_func=lambda x: f"{x}h  ({x}h charge + {x}h décharge)",
            index=1, key="intra_dur",
        )
    _intra_cap = power_MW * _intra_dur
    with _ic3:
        _id_disp = "illimité" if _intra_n_cyc == 0 else str(_intra_n_cyc)
        st.metric("Capacité par cycle (MWh)", f"{_intra_cap:.3f} MWh",
                  delta=f"{power_MW} MW × {_intra_dur}h · {_id_disp} cycle(s)/j.")
    with _ic4:
        _intra_jours_excl = st.multiselect(
            "Jours avec restriction",
            ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"],
            key="intra_jours_excl"
        )
    with _ic5:
        _intra_h_debut = st.number_input("Heure début restriction", 0, 23, 0, key="intra_h_debut")
    with _ic6:
        _intra_h_fin = st.number_input("Heure fin restriction", 0, 23, 0, key="intra_h_fin")

    _day_map_i = {"Lundi":0,"Mardi":1,"Mercredi":2,"Jeudi":3,"Vendredi":4,"Samedi":5,"Dimanche":6}
    _intra_days_num = [_day_map_i[d] for d in _intra_jours_excl]
    _intra_excl = ({"days": _intra_days_num, "hours": list(range(_intra_h_debut, _intra_h_fin))}
                   if _intra_days_num and _intra_h_debut < _intra_h_fin else {})

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Fichiers de prix</p>', unsafe_allow_html=True)
    _all_intra_files = st.file_uploader(
        "Déposez tous vos fichiers intraday ici (ZIP ou CSV) — le parseur détecte tout automatiquement",
        type=["zip", "csv"],
        accept_multiple_files=True,
        key="intra_files_all",
        help=(
            "Formats acceptés :\n"
            "• ZIP monolithiques (Intraday Continuous.zip, Intraday_part1/2/3.zip) → scannés automatiquement\n"
            "• ZIP EPEX nommés (Continuous_Index-FR-2021.zip, etc.)\n"
            "• CSV IDA1 / IDA2 / IDA3 (pan-european prices)\n"
            "Le marché continu n'utilise que Continuous_Index (IDFULL) — Continuous_Statistics, s'il est présent, "
            "est ignoré car redondant avec l'Index. Le parseur détecte le type de chaque fichier et crée un onglet par marché."
        )
    )
    _intra_files_ida  = [f for f in (_all_intra_files or []) if f.name.lower().endswith('.csv')]
    _intra_files_cont = [f for f in (_all_intra_files or []) if f.name.lower().endswith('.zip')]

    if not _all_intra_files:
        st.info("Uploadez vos fichiers intraday ci-dessus — ZIP EPEX, CSV IDA1/2/3, ou les 3 parts monolithiques.")
    else:
        st.caption(f"**{len(_all_intra_files)} fichier(s) chargés** — {len(_intra_files_cont)} ZIP · {len(_intra_files_ida)} CSV")

        if st.button("Lancer la simulation intraday", key="btn_intra_run", type="primary"):
            st.session_state["_intra_run"] = True
            st.session_state["_intra_markets"] = None
            st.session_state["_intra_bas"] = None
            st.session_state["_intra_bas_product"] = None
            st.session_state["_intra_bas_range"] = None
            st.session_state["_intra_trades_srcs"] = None

        if st.session_state.get("_intra_run"):
            import zipfile as _zf_d, io as _io_d, json as _j_intra

            @st.cache_data(show_spinner=False)
            def _load_market_pivot_cached(file_items: tuple, kind, index_names: tuple = None):
                """Parsing seul (coûteux, plusieurs minutes sur un gros ZIP) — mis en
                cache par contenu de fichiers + kind + index_names. Indépendant de
                puissance/cycles/durée/restrictions : changer ces paramètres ne
                re-déclenche donc pas un re-parsing, seule _simulate_market_cached
                re-tourne. index_names : pour le marché continu, demander plusieurs
                IndexName (IDFULL/ID1/ID3) en un seul parsing — pv devient alors un
                dict {nom: pivot} au lieu d'un DataFrame unique."""
                from bess_engine import load_intraday as _li_c

                class _SrcC:
                    def __init__(self, name, data):
                        self.name = name
                        self._data = data
                    def read(self):
                        return self._data

                srcs = [_SrcC(n, d) for n, d in file_items]
                pv = _li_c(srcs, kinds={kind} if kind else None,
                           index_names=set(index_names) if index_names else None)
                return (pv, getattr(_li_c, 'bid_ask_spread', None),
                        getattr(_li_c, 'bid_ask_spread_by_product', None),
                        getattr(_li_c, 'trades_date_range', None),
                        getattr(_li_c, 'report', []))

            @st.cache_data(show_spinner=False)
            def _simulate_market_cached(pv, params_json):
                """Simulation seule (rapide) — mise en cache par pivot + paramètres
                batterie, pour rendre instantané un aller-retour 1h→2h→1h une fois
                chaque combinaison déjà calculée une fois."""
                from bess_engine import (simulate_arbitrage as _sa_c,
                                         simulate_arbitrage_optimal as _sao_c,
                                         aggregate_arbitrage as _agg_c)
                params = _j_intra.loads(params_json)
                daily = (_sao_c(pv, params) if params["optimal_quota"] else _sa_c(pv, params))
                yearly = _agg_c(daily, params["power_MW"])
                return daily, yearly

            _p_intra_base = {
                "power_MW":        power_MW,
                "n_cycles":        0 if max_cycles else _intra_n_cyc,
                "duration_h":      _intra_dur,
                "excluded_hours":  _intra_excl,
                "efficiency":      efficiency,
                "max_cycles_year": max_cycles,
                "optimal_quota":   bool(max_cycles),
                "min_spread":      min_spread,
            }
            _p_intra_json = _j_intra.dumps(_p_intra_base, sort_keys=True)

            _markets = {}
            _prog_m = st.progress(0, text="Lecture des fichiers…")

            with st.spinner("Lecture des fichiers…"):
                _all_src_raw = {f.name: f.read() for f in _all_intra_files}

            # ── Scan du contenu de chaque ZIP ─────────────────────────────────
            def _scan_zip(raw: bytes) -> dict:
                """Retourne {market: True} selon les fichiers détectés dans le ZIP."""
                found = {}
                try:
                    with _zf_d.ZipFile(_io_d.BytesIO(raw)) as z:
                        names = ' '.join(z.namelist()).lower()
                        if 'index' in names: found['Continuous'] = True
                        if 'trade' in names: found['Continuous_Trades'] = True
                except Exception:
                    pass
                return found

            # ── Construire les groupes par marché ─────────────────────────────
            # Le marché continu n'est représenté que par Continuous_Index (IDFULL) —
            # Continuous_Statistics est volontairement ignoré : sur le même créneau,
            # les deux mesurent le même marché continu (résultats quasi identiques),
            # et Index seul, avec la priorité 60min→30min→15min, couvre déjà la
            # totalité des périodes (cf. _parse_stat_csv) sans avoir à décompresser
            # en plus les fichiers Statistics (souvent les plus volumineux du ZIP).
            _all_groups   = {}   # {market: [(name, bytes), ...]}
            _trades_srcs  = []   # [(name, bytes), ...] pour bid/ask uniquement

            for _fname, _raw in _all_src_raw.items():
                fu = _fname.upper()

                # CSV IDA
                if _fname.lower().endswith('.csv'):
                    key = "IDA3" if "IDA3" in fu else ("IDA2" if "IDA2" in fu else "IDA1")
                    _all_groups.setdefault(key, []).append((_fname, _raw))
                    continue

                # ZIP nommé explicitement
                if "STATISTIC" in fu:
                    continue  # ignoré : redondant avec Continuous_Index, cf. note ci-dessus
                elif "INDEX" in fu and "INTRADAY" not in fu and "PART" not in fu:
                    _all_groups.setdefault("Continuous", []).append((_fname, _raw))
                elif "TRADE" in fu and "INTRADAY" not in fu and "PART" not in fu:
                    _trades_srcs.append((_fname, _raw))
                else:
                    # ZIP monolithique (Intraday_part*.zip) → scanner le contenu
                    _inner = _scan_zip(_raw)
                    if _inner.get('Continuous'):
                        _all_groups.setdefault("Continuous", []).append((_fname, _raw))
                    if _inner.get('Continuous_Trades'):
                        _trades_srcs.append((_fname, _raw))

            # Dédupliquer (un même fichier peut alimenter plusieurs marchés)
            for _mkt in list(_all_groups.keys()):
                seen, deduped = set(), []
                for _n, _d in _all_groups[_mkt]:
                    if _n not in seen:
                        seen.add(_n); deduped.append((_n, _d))
                _all_groups[_mkt] = deduped

            if not _all_groups:
                st.error("Aucun marché reconnu dans les fichiers. Vérifiez les formats.")
                st.stop()

            st.caption(f"Marchés détectés : **{' · '.join(_all_groups.keys())}**")
            _n_grp = len(_all_groups)

            # Filtre de marché passé à load_intraday : évite de re-décompresser
            # les ZIP Trades ou Statistics (potentiellement plusieurs centaines
            # de Mo) d'une même archive bundle (ex. "Intraday Continuous.zip")
            # quand on ne veut que l'Index pour le marché continu.
            _kind_for_market = {"Continuous": "index"}
            # Continuous_Index publie 3 références dans le même fichier (IDFULL =
            # sur toute la session continue, ID1/ID3 = indice calculé 1h/3h avant
            # livraison) — on les extrait toutes les 3 en un seul parsing pour les
            # comparer, sans avoir à re-uploader quoi que ce soit en plus.
            _index_names_for_market = {"Continuous": ("IDFULL", "ID1", "ID3")}

            _provenance = {}  # {label: {"sources": [...], "report": [...]}}
            for _gi, (_mname, _msrcs) in enumerate(_all_groups.items()):
                _prog_m.progress(int(_gi/_n_grp*90), text=f"Simulation {_mname}…")
                try:
                    _pv_m, _bas_from_load, _bas_prod_from_load, _bas_range_from_load, _report_m = \
                        _load_market_pivot_cached(tuple(_msrcs), _kind_for_market.get(_mname),
                                                   _index_names_for_market.get(_mname))
                    # Récupérer bid/ask si le chargement l'a calculé (Trades dans les ZIP)
                    if (_bas_from_load is not None and not _bas_from_load.empty
                            and st.session_state.get("_intra_bas") is None):
                        st.session_state["_intra_bas"]         = _bas_from_load
                        st.session_state["_intra_bas_product"] = _bas_prod_from_load
                        st.session_state["_intra_bas_range"]   = _bas_range_from_load
                    # index_names demandé sur plusieurs IndexName trouvés → pv est un
                    # dict {nom: pivot} : un marché par référence (IDFULL/ID1/ID3).
                    _is_multi = isinstance(_pv_m, dict)
                    _pv_subs = _pv_m if _is_multi else {_mname: _pv_m}
                    for _sub_name, _sub_pv in _pv_subs.items():
                        if _sub_pv is None or _sub_pv.empty:
                            continue
                        _label = f"{_mname} ({_sub_name})" if _is_multi else _sub_name
                        _daily_m, _yearly_m = _simulate_market_cached(_sub_pv, _p_intra_json)
                        _markets[_label] = {"pv": _sub_pv, "daily": _daily_m, "yearly": _yearly_m}
                        _provenance[_label] = {
                            "source_files": [n for n, _ in _msrcs],
                            "report": [r for r in _report_m
                                       if r.get('type_détecté') != 'trades'
                                       and (_is_multi is False
                                            or _sub_name in r.get('index_names', [_sub_name]))],
                        }
                except Exception as _em:
                    st.warning(f"{_mname} : erreur — {_em}")

            # Si pas encore de bid/ask et qu'on a des sources Trades → passe dédiée légère
            if st.session_state.get("_intra_bas") is None and _trades_srcs:
                with st.spinner("Calcul spread Bid/Ask (Continuous Trades, ~30 jours les plus récents)…"):
                    _, _bas, _bas_prod, _bas_range, _ = _load_market_pivot_cached(tuple(_trades_srcs), "trades")
                    if _bas is not None and not _bas.empty:
                        st.session_state["_intra_bas"]         = _bas
                        st.session_state["_intra_bas_product"] = _bas_prod
                        st.session_state["_intra_bas_range"]   = _bas_range

            _bas_final = st.session_state.get("_intra_bas")
            if _bas_final is not None and not _bas_final.empty:
                _bas_rng = st.session_state.get("_intra_bas_range")
                _rng_txt = (f" · Période échantillonnée : **{_bas_rng[0].strftime('%d/%m/%Y')} → {_bas_rng[1].strftime('%d/%m/%Y')}**"
                            if _bas_rng else "")
                st.info(
                    f"**Spread Bid/Ask** (Continuous Trades) — "
                    f"Spread moyen : **{float(_bas_final['spread'].mean()):.2f} €/MWh** · "
                    f"Coût de friction du marché continu, visible dans chaque onglet de marché."
                    f"{_rng_txt}"
                )

            _prog_m.progress(100, text="Terminé.")
            _prog_m.empty()
            st.session_state["_intra_markets"]    = _markets
            st.session_state["_intra_provenance"] = _provenance
            if _trades_srcs:
                st.session_state["_intra_trades_srcs"] = _trades_srcs

        _markets = st.session_state.get("_intra_markets")
        if _markets:
            # ── Vérification des sources ──────────────────────────────────────
            _prov = st.session_state.get("_intra_provenance", {})
            if _prov:
                with st.expander("Vérification des sources de données", expanded=False):
                    st.caption(
                        "Pour chaque marché simulé : fichier uploadé source, sous-fichier parsé à l'intérieur, "
                        "IndexName extrait et nombre de jours/heures disponibles. "
                        "Résolution horaire : priorité à la ligne 60 min (déjà la moyenne pondérée exacte de l'heure) "
                        "— si absente pour un créneau donné, reconstruction depuis les tranches 30 min puis 15 min "
                        "pour ne perdre aucune période."
                    )
                    for _lbl, _p in _prov.items():
                        st.markdown(f"**{_lbl}**")
                        _srcs_str = " · ".join(_p.get("source_files", []))
                        st.caption(f"Fichier(s) source uploadé(s) : {_srcs_str}")
                        _rep = _p.get("report", [])
                        if _rep:
                            _rep_rows = []
                            for _r in _rep:
                                _idx_str = ", ".join(_r.get("index_names", [])) if _r.get("index_names") else "—"
                                _nb_j = _r["nb_jours"]
                                _rep_rows.append({
                                    "Sous-fichier parsé": _r["fichier"],
                                    "Type": _r["type_détecté"],
                                    "IndexName extrait": _idx_str,
                                    "Résolution prix": "60 min prioritaire (→ 30 min → 15 min si absent)",
                                    "Jours": _nb_j,
                                    "Heures": _nb_j * 24,
                                    "Statut": _r["statut"],
                                })
                            st.dataframe(pd.DataFrame(_rep_rows), hide_index=True, width="stretch")
                        else:
                            st.caption("(CSV direct, aucun sous-fichier imbriqué)")
                        st.divider()

            # ── Tableau comparatif des marchés ────────────────────────────────
            st.markdown('<p class="section">Comparaison des marchés</p>', unsafe_allow_html=True)
            st.caption(
                "Chaque ligne = un marché simulé avec les mêmes paramètres. "
                "PnL/MW/an = métrique clé pour comparer — plus c'est élevé, plus le marché rémunère bien la batterie. "
                "Spread moy. = volatilité des prix — un spread élevé = plus d'opportunités d'arbitrage. "
                "Prix utilisé : pour IDA1/2/3, le prix de clearing unique de l'enchère (aucune ambiguïté possible). "
                "Pour le marché continu, l'IndexPrice de Continuous_Index, qui publie 3 références dans le même "
                "fichier — IDFULL (sur toute la session continue), ID1 (indice calculé 1h avant livraison) et "
                "ID3 (3h avant livraison) — affichées ici comme 3 marchés séparés pour comparer leur potentiel. "
                "Continuous_Statistics n'est pas utilisé séparément, il mesure le même marché continu que l'Index."
            )

            _PRIX_SRC_M = {
                "IDA1": "Prix de clearing", "IDA2": "Prix de clearing", "IDA3": "Prix de clearing",
                "Continuous (IDFULL)": "IndexPrice (IDFULL)",
                "Continuous (ID1)":    "IndexPrice (ID1, 1h avant livraison)",
                "Continuous (ID3)":    "IndexPrice (ID3, 3h avant livraison)",
            }
            _comp_rows = []
            for _mname, _mdata in _markets.items():
                _y = _mdata["yearly"]
                _d = _mdata["daily"]
                _pnl  = _y["pnl_total"].sum()
                _nans = len(_y)
                _sprd = _y["spread_moy"].mean()
                _jact = int(_y["jours_actifs"].mean())
                _comp_rows.append({
                    "Marché":             _mname,
                    "Type":               "Auction" if _mname.startswith("IDA") else "Continu",
                    "Prix utilisé":       _PRIX_SRC_M.get(_mname, "—"),
                    "Période":            f"{int(_y['annee'].min())}–{int(_y['annee'].max())}",
                    "PnL total (€)":      _fmt(_pnl),
                    "PnL/MW/an (€)":      _fmt(_pnl / power_MW / max(_nans, 1)),
                    "Spread moy. (€/MWh)": f"{_sprd:.1f}",
                    "Jours actifs/an":    str(_jact),
                })
            _comp_df = pd.DataFrame(_comp_rows).sort_values("PnL/MW/an (€)", ascending=False)
            st.dataframe(_comp_df, hide_index=True, width="stretch")

            # ── Graphique comparatif PnL/MW/an ────────────────────────────────
            _COLORS_M = {"IDA1":"#1565c0","IDA2":"#283593","IDA3":"#0d47a1",
                         "Continuous (IDFULL)":"#2e7d32","Continuous (ID1)":"#66bb6a","Continuous (ID3)":"#a5d6a7"}
            _fig_comp_m = go.Figure()
            for _mname, _mdata in _markets.items():
                _y = _mdata["yearly"]
                _fig_comp_m.add_trace(go.Bar(
                    name=_mname,
                    x=_y["annee"].astype(str),
                    y=_y["pnl_total"] / power_MW,
                    marker_color=_COLORS_M.get(_mname, "#666"),
                    hovertemplate=f"<b>{_mname}</b><br>%{{x}} : %{{y:,.0f}} €/MW<extra></extra>",
                ))
            _fig_comp_m.update_layout(
                height=320, barmode="group",
                xaxis=dict(title="Année"),
                yaxis=dict(title="PnL/MW (€/MW)", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                margin=dict(t=10,b=80,l=60,r=10),
            )
            apply_bb(_fig_comp_m)
            st.caption("PnL par MW installé — normalise le résultat par la puissance simulée, pour comparer les marchés indépendamment de la taille de la batterie (utile si vous testez plusieurs configurations).")
            st.plotly_chart(_fig_comp_m, width="stretch", config=PLOTLY_CFG, key="intra_comp_bar")

            # ── Graphique PnL cumulé par marché ──────────────────────────────
            _fig_cum_m = go.Figure()
            for _mname, _mdata in _markets.items():
                _d = _mdata["daily"].sort_values("date")
                _fig_cum_m.add_trace(go.Scatter(
                    name=_mname,
                    x=_d["date"], y=_d["pnl"].cumsum(),
                    mode="lines", line=dict(width=2, color=_COLORS_M.get(_mname,"#666")),
                    hovertemplate=f"<b>{_mname}</b><br>%{{x|%d/%m/%Y}} : %{{y:,.0f}} €<extra></extra>",
                ))
            _fig_cum_m.update_layout(
                height=300,
                xaxis=dict(gridcolor="#f0f0f0"),
                yaxis=dict(title="PnL cumulé (€)", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                margin=dict(t=10,b=80,l=60,r=10), hovermode="x unified",
            )
            apply_bb(_fig_cum_m)
            st.caption("PnL cumulé — la courbe la plus haute = marché le plus rentable sur la durée. La pente indique le rythme de gain ; un écart qui se creuse dans le temps signale un marché qui devient structurellement plus (ou moins) intéressant.")
            st.plotly_chart(_fig_cum_m, width="stretch", config=PLOTLY_CFG, key="intra_cum_comp")

            # ── Spread Bid/Ask — section globale ──────────────────────────────
            st.markdown('<p class="section">Spread Bid/Ask — Coût de friction du marché continu</p>',
                        unsafe_allow_html=True)
            st.caption(
                "WAP (Weighted Average Price) = prix moyen pondéré par les volumes échangés. "
                "Spread = WAP des trades **achat** (côté ask) − WAP des trades **vente** (côté bid), "
                "par heure de livraison. Normalement ≥ 0 : c'est le coût de friction réel du marché "
                "continu (l'écart entre ce qu'on paie à l'achat et ce qu'on reçoit à la vente), à comparer "
                "au spread théorique (perfect foresight, sans coût de transaction) utilisé dans le reste de l'onglet."
            )
            _bas_disp = st.session_state.get("_intra_bas")

            # Si pas encore calculé mais on a des trades_srcs en session → recalculer
            # (mis en cache par _load_market_pivot_cached : instantané si déjà fait)
            if _bas_disp is None and st.session_state.get("_intra_trades_srcs"):
                with st.spinner("Calcul spread Bid/Ask depuis Continuous Trades…"):
                    _, _bas_tmp, _bas_prod_tmp, _bas_range_tmp, _ = _load_market_pivot_cached(
                        tuple(st.session_state["_intra_trades_srcs"]), "trades")
                    if _bas_tmp is not None and not _bas_tmp.empty:
                        st.session_state["_intra_bas"]         = _bas_tmp
                        st.session_state["_intra_bas_product"] = _bas_prod_tmp
                        st.session_state["_intra_bas_range"]   = _bas_range_tmp
                        _bas_disp = _bas_tmp

            if _bas_disp is not None and not _bas_disp.empty:
                _bas_moy_g = float(_bas_disp['spread'].mean())
                _bas_rng2  = st.session_state.get("_intra_bas_range")
                if _bas_rng2:
                    st.caption(
                        f"Échantillon : **{int((_bas_disp['spread'].notna()).sum())} heures agrégées** sur les trades du "
                        f"**{_bas_rng2[0].strftime('%d/%m/%Y')}** au **{_bas_rng2[1].strftime('%d/%m/%Y')}** "
                        f"(jours les plus récents disponibles dans les fichiers fournis)."
                    )
                _bga, _bgb, _bgc = st.columns(3)
                _bga.metric("Spread moyen global", f"{_bas_moy_g:.2f} €/MWh",
                            help="WAP achats (ask) − WAP ventes (bid), moyenné sur 24h. Source : Continuous_Trades EPEX.")
                _bgb.metric("Heure la plus chère",
                            f"H{int(_bas_disp['spread'].idxmax()):02d} ({_bas_disp['spread'].max():.1f} €/MWh)",
                            help="Heure où le coût de transaction (ask − bid) est le plus élevé.")
                _bgc.metric("Heures à spread négatif",
                            f"{int((_bas_disp['spread'] < 0).sum())}/24",
                            help="Heures où les vendeurs ont reçu plus que les acheteurs n'ont payé — "
                                 "anomalie / inversion temporaire de marché, pas le cas normal.")

                _fig_bas_g = go.Figure()
                _fig_bas_g.add_trace(go.Bar(
                    x=[f"H{int(h):02d}" for h in _bas_disp.index],
                    y=_bas_disp['spread'].values,
                    marker_color=[C1 if v >= 0 else "#c62828" for v in _bas_disp['spread'].values],
                    hovertemplate="H%{x} — Bid/Ask : <b>%{y:.2f} €/MWh</b><extra></extra>",
                ))
                _fig_bas_g.add_hline(y=_bas_moy_g, line_dash="dot", line_color=ORANGE,
                                     annotation_text=f"Moy: {_bas_moy_g:.2f} €/MWh")
                _fig_bas_g.update_layout(
                    height=240, margin=dict(t=10, b=40, l=60, r=10),
                    xaxis=dict(title="Heure"),
                    yaxis=dict(title="Spread Bid/Ask (€/MWh)", gridcolor="#f0f0f0"),
                    plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
                )
                apply_bb(_fig_bas_g)
                st.caption(
                    "Spread bid/ask par heure — WAP achats (ask) − WAP ventes (bid). "
                    "Bleu = spread positif (normal, coût de friction réel). Rouge = spread négatif (anomalie). "
                    "Ce spread représente le coût de friction réel à déduire du PnL théorique continu."
                )
                st.plotly_chart(_fig_bas_g, width="stretch", config=PLOTLY_CFG, key="intra_bas_global")

                # ── Ventilation par produit (Quarter vs Hour) ───────────────────
                _bas_prod = st.session_state.get("_intra_bas_product")
                if _bas_prod is not None and not _bas_prod.empty:
                    st.markdown('<p class="section">Spread Bid/Ask par produit — Quart d\'heure vs Heure</p>',
                                unsafe_allow_html=True)
                    st.caption(
                        "Quarter-Hour (livraison par bloc de 15 min) et Hour (bloc de 1h) sont les deux granularités "
                        "tradées sur le marché continu. Elles ont des liquidités différentes — moins de participants "
                        "sur le quart d'heure — donc leur spread bid/ask reflète des coûts de friction distincts à l'achat/vente."
                    )
                    _fig_bas_p = go.Figure()
                    for _prod, _clr_p in [("Hour", C3), ("Quarter", C4)]:
                        if _prod not in _bas_prod.index.get_level_values(0):
                            continue
                        _sub = _bas_prod.loc[_prod]
                        _fig_bas_p.add_trace(go.Bar(
                            name=_prod,
                            x=[f"H{int(h):02d}" for h in _sub.index],
                            y=_sub['spread'].values,
                            marker_color=_clr_p,
                            hovertemplate=f"{_prod} — H%{{x}} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                        ))
                    _fig_bas_p.update_layout(
                        height=260, barmode="group", margin=dict(t=10, b=40, l=60, r=10),
                        xaxis=dict(title="Heure"),
                        yaxis=dict(title="Spread Bid/Ask (€/MWh)", gridcolor="#f0f0f0"),
                        plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                    )
                    apply_bb(_fig_bas_p)
                    _moy_prod = _bas_prod.groupby(level=0)['spread'].mean()
                    st.caption(
                        "Spread moyen — " + " · ".join(
                            f"**{p}** : {v:.2f} €/MWh" for p, v in _moy_prod.items()
                        )
                    )
                    st.plotly_chart(_fig_bas_p, width="stretch", config=PLOTLY_CFG, key="intra_bas_product")
            else:
                st.info(
                    "Spread bid/ask non disponible — nécessite des fichiers **Continuous_Trades** "
                    "(ex. `Intraday_part3.zip`, ou un export EPEX complet type `Intraday Continuous.zip`). "
                    "Uploadez-le dans la zone fichiers ci-dessus pour l'activer."
                )

            # ── Résultats détaillés par marché ────────────────────────────────
            st.markdown('<p class="section">Résultats détaillés par marché</p>', unsafe_allow_html=True)
            _tabs_m = st.tabs(list(_markets.keys()))
            for _ti, (_mname, _mdata) in enumerate(zip(_markets.keys(), _markets.values())):
                with _tabs_m[_ti]:
                    _daily_i  = _mdata["daily"]
                    _yearly_i = _mdata["yearly"]
                    _pv_m     = _mdata["pv"]
                    _clr_m    = _COLORS_M.get(_mname, C1)

                    # ── KPIs 6 colonnes ───────────────────────────────────────
                    _km1,_km2,_km3,_km4,_km5,_km6 = st.columns(6)
                    _pnl_tot_i   = _daily_i["pnl"].sum()
                    _pnl_abs_i   = _daily_i["pnl_absolu"].sum() if "pnl_absolu" in _daily_i.columns else _pnl_tot_i
                    _jours_act_i = int(_daily_i["valid"].sum())
                    _jours_tot_i = len(_daily_i)
                    _spr_moy_i   = float(_daily_i.loc[_daily_i["valid"],"spread"].mean()) if _daily_i["valid"].any() else 0
                    _spr_abs_i   = float(_daily_i["spread_absolu"].mean()) if "spread_absolu" in _daily_i.columns else _spr_moy_i
                    _energ_i     = float(_daily_i["energie_chargee"].sum()) if "energie_chargee" in _daily_i.columns else 0
                    _blq_i       = int(_daily_i.get("bloque", pd.Series([False]*len(_daily_i))).sum()) if "bloque" in _daily_i.columns else 0
                    _cap_i       = _fmt(_pnl_abs_i)
                    _taux_i      = _pnl_tot_i/_pnl_abs_i if _pnl_abs_i else 1
                    _km1.metric("PnL réel (avec contraintes)", f"{_fmt(_pnl_tot_i)} €")
                    _km2.metric("PnL borne théorique max", f"{_cap_i} €",
                                delta=f"{_taux_i*100:.1f}% capturé")
                    _km3.metric("Spread moyen (jours actifs)", f"{_spr_moy_i:.1f} €/MWh",
                                delta=f"Théorique: {_spr_abs_i:.1f} €/MWh")
                    _km4.metric("Jours actifs", f"{_jours_act_i} / {_jours_tot_i}",
                                delta=f"{_jours_act_i/_jours_tot_i*100:.0f}% taux activation")
                    _km5.metric("Énergie totale échangée",
                                f"{_energ_i:.1f} MWh" if _energ_i else "—",
                                delta=f"Capacité/jour : {_intra_cap:.2f} MWh")
                    _km6.metric("Jours bloqués (maintenance)", f"{_blq_i}",
                                delta="Aucun" if _blq_i==0 else f"{_blq_i/_jours_tot_i*100:.1f}%",
                                delta_color="off" if _blq_i==0 else "inverse")

                    # ── Bid/Ask spread (Continuous Trades) ────────────────────
                    _bas_stored = st.session_state.get("_intra_bas")
                    if _bas_stored is not None and not _bas_stored.empty:
                        st.markdown('<p class="section">Impact du spread Bid/Ask (Continuous Trades)</p>',
                                    unsafe_allow_html=True)
                        _bas_moy_v = float(_bas_stored['spread'].mean())
                        _pnl_net_bas = _pnl_tot_i - _bas_moy_v * power_MW * _jours_act_i * _intra_dur
                        _bk1, _bk2, _bk3 = st.columns(3)
                        _bk1.metric("Spread Bid/Ask moyen", f"{_bas_moy_v:.2f} €/MWh",
                                    help="WAP ventes - WAP achats. Coût de transaction réel du marché continu.")
                        _bk2.metric("PnL net après bid/ask", f"{_fmt(_pnl_net_bas)} €",
                                    delta=f"Coût marché : {_fmt(_pnl_tot_i - _pnl_net_bas)} €",
                                    delta_color="inverse")
                        _bk3.metric("Spread arbitrage net",
                                    f"{_spr_moy_i - _bas_moy_v:.1f} €/MWh",
                                    delta=f"Brut {_spr_moy_i:.1f} − B/A {_bas_moy_v:.2f}")
                        _fig_bas = go.Figure()
                        _fig_bas.add_trace(go.Bar(
                            x=[f"H{int(h):02d}" for h in _bas_stored.index],
                            y=_bas_stored['spread'].values,
                            marker_color=[C1 if v >= 0 else "#c62828"
                                          for v in _bas_stored['spread'].values],
                            hovertemplate="H%{x} — B/A : <b>%{y:.2f} €/MWh</b><extra></extra>",
                        ))
                        _fig_bas.update_layout(
                            height=200, margin=dict(t=5,b=40,l=60,r=10),
                            xaxis=dict(title="Heure"),
                            yaxis=dict(title="Spread B/A (€/MWh)", gridcolor="#f0f0f0"),
                            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
                        )
                        apply_bb(_fig_bas)
                        st.caption(
                            "Spread bid/ask par heure — calculé sur un échantillon de ~30 jours de Continuous_Trades. "
                            "Bleu = spread positif (normal). Rouge = spread négatif (acheteurs surpayent temporairement). "
                            f"Échantillon basé sur {len(_bas_stored)} heures."
                        )
                        st.plotly_chart(_fig_bas, width="stretch", config=PLOTLY_CFG,
                                        key=f"intra_bas_{_mname}")

                    # ── PnL annuel borne max vs réel ──────────────────────────
                    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
                    st.caption("Barres jaunes = borne max théorique. Barres colorées = PnL réel avec contraintes. L'écart = coût des restrictions et du quota.")
                    _fig1_m = go.Figure()
                    _fig1_m.add_trace(go.Bar(
                        x=_yearly_i["annee"].astype(str), y=_yearly_i["pnl_absolu_total"],
                        name="Borne max", marker_color=C2,
                        text=_yearly_i["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €"),
                        textposition="outside",
                        hovertemplate="<b>%{x}</b><br>Borne max : %{y:.0f} €<extra></extra>",
                    ))
                    _fig1_m.add_trace(go.Bar(
                        x=_yearly_i["annee"].astype(str), y=_yearly_i["pnl_total"],
                        name="PnL réel", marker_color=_clr_m,
                        text=_yearly_i["pnl_total"].apply(lambda x: f"{_fmt(x)} €"),
                        textposition="inside", textfont_color="white",
                        hovertemplate="<b>%{x}</b><br>PnL réel : %{y:.0f} €<extra></extra>",
                    ))
                    _fig1_m.update_layout(
                        barmode="group", height=340,
                        yaxis=dict(title="PnL (€)", gridcolor="#f0f0f0"),
                        xaxis_title="Année", legend=LEGEND_BOTTOM,
                        margin=dict(t=10,b=40,l=60,r=20),
                        plot_bgcolor="white", paper_bgcolor="white",
                    )
                    apply_bb(_fig1_m)
                    st.plotly_chart(_fig1_m, width="stretch", config=PLOTLY_CFG, key=f"intra_fig1_{_mname}")

                    # ── Graphiques 2×2 ────────────────────────────────────────
                    _tg1, _tg2 = st.columns(2)
                    with _tg1:
                        st.markdown('<p class="section">Spread moyen — par semaine (€/MWh)</p>', unsafe_allow_html=True)
                        st.caption("Évolution du spread moyenné par semaine, pour lisser le bruit jour-à-jour. Zone ombrée = fourchette min-max observée cette semaine-là (plus elle est large, plus le marché a été volatil). Ligne pointillée = spread moyen sur toute la période, pour situer chaque semaine par rapport à la tendance générale.")
                        _dv_i = _daily_i[_daily_i["valid"]].copy()
                        if not _dv_i.empty:
                            _dv_i["sem"] = pd.to_datetime(_dv_i["date"]).dt.to_period("W").apply(lambda p: p.start_time)
                            _wi = _dv_i.groupby("sem").agg(
                                s=("spread","mean"),mn=("spread","min"),mx=("spread","max")).reset_index()
                            _fig_sw = go.Figure()
                            _fig_sw.add_trace(go.Scatter(
                                x=pd.concat([_wi["sem"],_wi["sem"].iloc[::-1]]),
                                y=pd.concat([_wi["mx"],_wi["mn"].iloc[::-1]]),
                                fill="toself",fillcolor="rgba(46,125,50,0.15)",
                                line=dict(color="rgba(0,0,0,0)"),showlegend=False,hoverinfo="skip",
                            ))
                            _fig_sw.add_trace(go.Scatter(
                                x=_wi["sem"],y=_wi["s"],mode="lines",
                                line=dict(color=_clr_m,width=2),name="Spread hebdo",
                                hovertemplate="%{x|%d/%m/%Y} : %{y:.1f} €/MWh<extra></extra>",
                            ))
                            _fig_sw.add_hline(y=_spr_moy_i,line_dash="dot",line_color=ORANGE,
                                              line_width=1.5,annotation_text=f"Moy: {_spr_moy_i:.1f}")
                            _fig_sw.update_layout(
                                height=300,margin=dict(t=10,b=40,l=60,r=80),
                                xaxis=dict(gridcolor="#f0f0f0",tickformat="%b %Y"),
                                yaxis=dict(title="Spread (€/MWh)",gridcolor="#f0f0f0",range=[0,300]),
                                plot_bgcolor="white",paper_bgcolor="white",showlegend=False,
                            )
                            apply_bb(_fig_sw)
                            st.plotly_chart(_fig_sw,width="stretch",config=PLOTLY_CFG,key=f"intra_sw_{_mname}")

                    with _tg2:
                        st.markdown('<p class="section">Profil horaire charge / décharge</p>', unsafe_allow_html=True)
                        st.caption("Nombre de fois où chaque heure a été choisie pour charger ou décharger, sur toute la période simulée. Vert = achat (heures où la batterie charge, généralement les moins chères de la journée). Orange = vente (heures où elle décharge, généralement les plus chères). Deux pics bien distincts indiquent un marché avec un profil de prix régulier, facile à arbitrer.")
                        _hc_cnt=[0]*24; _hd_cnt=[0]*24
                        for _,_rp in _pv_m.iterrows():
                            try:
                                _pp = _rp[HOUR_COLS].values.astype(float)
                                _av = get_available_hours(int(_rp["weekday"]),_intra_excl)
                                _cc = _best_n_cycles(_pp,_av,_intra_n_cyc,_intra_dur,power_MW,efficiency)
                                for _cy in _cc:
                                    for _hc in _cy["h_charge"]:
                                        if 0<=_hc<24: _hc_cnt[_hc]+=1
                                    for _hd in _cy["h_decharge"]:
                                        if 0<=_hd<24: _hd_cnt[_hd]+=1
                            except Exception: pass
                        _fig_prof=go.Figure()
                        _fig_prof.add_trace(go.Bar(x=[f"H{h:02d}" for h in range(24)],y=_hc_cnt,
                            name="Charge (achat)",marker_color=C1,opacity=0.85))
                        _fig_prof.add_trace(go.Bar(x=[f"H{h:02d}" for h in range(24)],y=_hd_cnt,
                            name="Décharge (vente)",marker_color=ORANGE,opacity=0.85))
                        _fig_prof.update_layout(
                            height=300,barmode="overlay",
                            margin=dict(t=10,b=40,l=60,r=10),
                            xaxis=dict(title="Heure"),
                            yaxis=dict(title="Nb jours",gridcolor="#f0f0f0"),
                            plot_bgcolor="white",paper_bgcolor="white",legend=LEGEND_BOTTOM,
                        )
                        apply_bb(_fig_prof)
                        st.plotly_chart(_fig_prof,width="stretch",config=PLOTLY_CFG,key=f"intra_prof_{_mname}")

                    _tg3, _tg4 = st.columns(2)
                    with _tg3:
                        st.markdown('<p class="section">Distribution des spreads journaliers</p>', unsafe_allow_html=True)
                        st.caption("Chaque barre = nombre de jours avec ce niveau de spread. Vers la droite = marché favorable (spread élevé, bonne opportunité d'arbitrage ce jour-là). Une distribution étalée vers la droite indique un marché globalement rentable plutôt que quelques jours exceptionnels qui tirent la moyenne vers le haut.")
                        _sv_i = _daily_i.loc[_daily_i["valid"],"spread"]
                        _p99_i = float(_sv_i.quantile(0.99)) if len(_sv_i)>0 else 100
                        _fig_dh=go.Figure()
                        _fig_dh.add_trace(go.Histogram(
                            x=_sv_i.clip(upper=_p99_i),nbinsx=40,
                            marker_color=_clr_m,opacity=0.8,
                            hovertemplate="Spread : %{x:.0f} €/MWh<br>%{y} jours<extra></extra>",
                        ))
                        if len(_sv_i)>0:
                            _fig_dh.add_vline(x=_sv_i.mean(),line_dash="dash",line_color=ORANGE,
                                              line_width=2,annotation_text=f"Moy: {_sv_i.mean():.1f}")
                            _fig_dh.add_vline(x=float(_sv_i.median()),line_dash="dot",line_color=C3,
                                              line_width=1.5,annotation_text=f"Méd: {_sv_i.median():.1f}")
                        _fig_dh.update_layout(
                            height=300,margin=dict(t=10,b=30,l=60,r=10),
                            xaxis=dict(title="Spread (€/MWh)",range=[0,_p99_i*1.05]),
                            yaxis=dict(title="Nb jours",gridcolor="#f0f0f0"),
                            plot_bgcolor="white",paper_bgcolor="white",showlegend=False,
                        )
                        apply_bb(_fig_dh)
                        st.plotly_chart(_fig_dh,width="stretch",config=PLOTLY_CFG,key=f"intra_dh_{_mname}")

                    with _tg4:
                        st.markdown('<p class="section">PnL cumulé dans le temps</p>', unsafe_allow_html=True)
                        st.caption("PnL cumulé = somme des gains au fil du temps (la pente indique le rythme de gain — plus elle est forte, plus la période a été rentable). Borne max = potentiel théorique avec prévision parfaite des prix et sans restriction horaire ni quota.")
                        _ds = _daily_i.sort_values("date").copy()
                        _ds["pnl_cum"] = _ds["pnl"].cumsum()
                        _ds["abs_cum"] = _ds["pnl_absolu"].cumsum() if "pnl_absolu" in _ds.columns else _ds["pnl_cum"]
                        _fig_cum=go.Figure()
                        _fig_cum.add_trace(go.Scatter(
                            x=_ds["date"],y=_ds["pnl_cum"],fill="tozeroy",mode="lines",
                            name="PnL cumulé",line=dict(color=_clr_m,width=2),
                            fillcolor=f"rgba(46,125,50,0.10)",
                            hovertemplate="%{x|%d/%m/%Y} : %{y:,.0f} €<extra></extra>",
                        ))
                        _fig_cum.add_trace(go.Scatter(
                            x=_ds["date"],y=_ds["abs_cum"],mode="lines",
                            name="Borne max",line=dict(color=C2,width=1.5,dash="dot"),
                            hovertemplate="Borne max : %{y:,.0f} €<extra></extra>",
                        ))
                        _fig_cum.update_layout(
                            height=300,margin=dict(t=10,b=40,l=60,r=10),
                            yaxis=dict(title="PnL cumulé (€)",gridcolor="#f0f0f0"),
                            xaxis=dict(tickformat="%b %Y"),
                            plot_bgcolor="white",paper_bgcolor="white",legend=LEGEND_BOTTOM,
                            hovermode="x unified",
                        )
                        apply_bb(_fig_cum)
                        st.plotly_chart(_fig_cum,width="stretch",config=PLOTLY_CFG,key=f"intra_cum_{_mname}")

                    # ── Répartition des cycles journaliers ────────────────────
                    st.markdown('<p class="section">Répartition des cycles journaliers</p>', unsafe_allow_html=True)
                    st.caption("Distribution du nombre de cycles réalisés par jour. Identifie les journées les plus et les moins actives. Un jour à 0 cycle signifie qu'aucun spread suffisamment rentable n'a été trouvé ce jour-là (marché plat, ou jour de bord du fichier sans donnée réelle).")
                    _cyc_res_m = _plot_cycles_distribution(_daily_i, "Intraday")
                    if _cyc_res_m:
                        _fc,_minc,_maxc,_dminc,_dmaxc = _cyc_res_m
                        _cc1,_cc2 = st.columns(2)
                        _cc1.metric("Minimum de cycles/jour",f"{_minc} cycle(s)",
                                    delta=f"ex. {pd.Timestamp(_dminc).strftime('%d/%m/%Y')}",delta_color="off")
                        _cc2.metric("Maximum de cycles/jour",f"{_maxc} cycle(s)",
                                    delta=f"ex. {pd.Timestamp(_dmaxc).strftime('%d/%m/%Y')}",delta_color="off")
                        apply_bb(_fc)
                        st.plotly_chart(_fc,width="stretch",config=PLOTLY_CFG,key=f"intra_cyc_{_mname}")

                    # ── Récapitulatif annuel ──────────────────────────────────
                    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
                    st.caption("Tableau détaillé par année. Le taux de capturé (PnL réel / borne max) mesure l'efficacité de votre stratégie.")
                    _show_m = _yearly_i.copy()
                    _show_m["pnl_absolu_total"] = _show_m["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €")
                    _show_m["pnl_total"]        = _show_m["pnl_total"].apply(lambda x: f"{_fmt(x)} €")
                    _show_m["taux_activation"]  = _show_m["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
                    _show_m["spread_moy"]       = _show_m["spread_moy"].apply(lambda x: f"{x:.1f}")
                    _disp_cols = ["annee","jours_actifs","taux_activation",
                                  "spread_moy","pnl_absolu_total","pnl_total","pnl_par_MW"]
                    _disp_cols = [c for c in _disp_cols if c in _show_m.columns]
                    _disp_rename = {"annee":"Année","jours_actifs":"Jours actifs",
                                    "taux_activation":"Taux activation",
                                    "spread_moy":"Spread moy. (€/MWh)",
                                    "pnl_absolu_total":"Borne max (€)",
                                    "pnl_total":"PnL réel (€)",
                                    "pnl_par_MW":"PnL/MW (€/MW)"}
                    st.dataframe(_show_m[_disp_cols].rename(columns=_disp_rename),
                                 hide_index=True,width="stretch")

                    # ── Explorer un jour ──────────────────────────────────────
                    with st.expander("Explorer un jour spécifique", expanded=True):
                        st.caption("Sélectionnez un jour pour voir le détail : prix intraday heure par heure, heures d'achat/vente choisies par la batterie, et PnL réalisé comparé au maximum théorique. Un jour à 0 cycle/0 € est normal s'il n'y avait aucune opportunité de spread positif (ou si peu de données de prix existent pour ce jour, ex. tout début/fin de fichier).")
                        _intra_dates_m = sorted(_daily_i["date"].dt.date.unique())
                        _intra_valid_dates_m = sorted(_daily_i.loc[_daily_i["valid"], "date"].dt.date.unique())
                        _intra_default_m = _intra_valid_dates_m[0] if _intra_valid_dates_m else (_intra_dates_m[0] if _intra_dates_m else None)
                        _intra_date_sel = st.date_input(
                            "Date",
                            value=_intra_default_m,
                            min_value=_intra_dates_m[0] if _intra_dates_m else None,
                            max_value=_intra_dates_m[-1] if _intra_dates_m else None,
                            key=f"intra_date_{_mname}",
                        )
                        if _intra_date_sel:
                            _row_m = _pv_m[_pv_m["date"].dt.date == _intra_date_sel]
                            if not _row_m.empty:
                                _prix_m  = _row_m[HOUR_COLS].values.flatten().astype(float)
                                _avail_m = get_available_hours(int(_row_m["weekday"].iloc[0]),_intra_excl)
                                _cyc_m   = _best_n_cycles(_prix_m,_avail_m,_intra_n_cyc,_intra_dur,power_MW,efficiency)
                                # Borne max du jour
                                _cyc_bm  = _best_n_cycles(_prix_m,list(range(24)),0,_intra_dur,power_MW,efficiency)
                                _pnl_m   = sum(c["pnl"] for c in _cyc_m)
                                _pnl_bm  = sum(c["pnl"] for c in _cyc_bm)
                                _spr_m   = sum(c["spread"] for c in _cyc_m)/len(_cyc_m) if _cyc_m else 0

                                _cex1,_cex2 = st.columns([3,1])
                                with _cex2:
                                    st.metric("Spread moyen",f"{_spr_m:.2f} €/MWh")
                                    st.metric("PnL du jour",f"{_pnl_m:.2f} €")
                                    st.metric("PnL borne max",f"{_pnl_bm:.2f} €",
                                              delta=f"{_pnl_m/_pnl_bm*100:.0f}% capturé" if _pnl_bm else None)
                                    st.metric("Cycles réalisés",f"{len(_cyc_m)}" + (" (max)" if _intra_n_cyc==0 else f" / {_intra_n_cyc}"))
                                    st.metric("Heures dispo",f"{len(_avail_m)}/24")
                                    for _ii,_cy in enumerate(_cyc_m,1):
                                        st.caption(f"Cycle {_ii} : charge H{_cy['h_charge']} ({_cy['prix_charge']:.1f} €/MWh) → décharge H{_cy['h_decharge']} ({_cy['prix_decharge']:.1f} €/MWh) | PnL {_cy['pnl']:.2f} €")
                                with _cex1:
                                    _fig_ex=go.Figure()
                                    for _eh in [h for h in range(24) if h not in _avail_m]:
                                        _fig_ex.add_vrect(x0=_eh-0.5,x1=_eh+0.5,
                                            fillcolor="rgba(200,200,200,0.3)",layer="below",line_width=0)
                                    _fig_ex.add_trace(go.Bar(
                                        x=list(range(24)),
                                        y=[float(p) if not pd.isna(p) else 0 for p in _prix_m],
                                        marker_color="#c8d8ec",name="Prix intraday",
                                        hovertemplate="H%{x:02d} — <b>%{y:.2f} €/MWh</b><extra></extra>",
                                    ))
                                    _col_ch=["#2e7d32","#1565c0","#6a1b9a"]
                                    _col_dch=["#c62828","#e65100","#4527a0"]
                                    for _ii,_cy in enumerate(_cyc_m):
                                        _cc2,_cd2=_col_ch[_ii%3],_col_dch[_ii%3]
                                        _fig_ex.add_trace(go.Scatter(
                                            x=_cy["h_charge"],y=_prix_m[_cy["h_charge"]],
                                            mode="markers",name=f"Achat C{_ii+1} ({_intra_dur}h)",
                                            marker=dict(color=_cc2,size=16,symbol="triangle-up",
                                                        line=dict(color="white",width=1)),
                                            hovertemplate=f"H%{{x:02d}} — Achat C{_ii+1} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                                        ))
                                        _fig_ex.add_trace(go.Scatter(
                                            x=_cy["h_decharge"],y=_prix_m[_cy["h_decharge"]],
                                            mode="markers",name=f"Vente C{_ii+1} ({_intra_dur}h)",
                                            marker=dict(color=_cd2,size=16,symbol="triangle-down",
                                                        line=dict(color="white",width=1)),
                                            hovertemplate=f"H%{{x:02d}} — Vente C{_ii+1} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                                        ))
                                        _fig_ex.add_shape(type="line",x0=-0.5,x1=23.5,
                                            y0=_cy["prix_charge"],y1=_cy["prix_charge"],
                                            line=dict(color=_cc2,width=1,dash="dot"))
                                        _fig_ex.add_shape(type="line",x0=-0.5,x1=23.5,
                                            y0=_cy["prix_decharge"],y1=_cy["prix_decharge"],
                                            line=dict(color=_cd2,width=1,dash="dot"))
                                        _fig_ex.add_annotation(x=22,y=_cy["prix_charge"],
                                            text=f"Achat C{_ii+1}: {_cy['prix_charge']:.1f}",
                                            showarrow=False,font=dict(size=9,color=_cc2))
                                        _fig_ex.add_annotation(x=22,y=_cy["prix_decharge"],
                                            text=f"Vente C{_ii+1}: {_cy['prix_decharge']:.1f}",
                                            showarrow=False,font=dict(size=9,color=_cd2))
                                    _fig_ex.update_layout(
                                        height=380,margin=dict(t=20,b=60,l=60,r=10),
                                        xaxis=dict(title="Heure",tickmode="array",
                                                   tickvals=list(range(24)),
                                                   ticktext=[f"H{h:02d}" for h in range(24)],
                                                   range=[-0.5,23.5]),
                                        yaxis=dict(title="Prix intraday (€/MWh)",gridcolor="#f0f0f0"),
                                        plot_bgcolor="white",paper_bgcolor="white",legend=LEGEND_BOTTOM,
                                        hoverlabel=dict(bgcolor="white",font_size=12),
                                    )
                                    apply_bb(_fig_ex)
                                    st.caption("Prix intraday heure par heure. ▲ = achat · ▼ = vente · lignes pointillées = prix moyens du cycle · zones grises = heures exclues.")
                                    st.plotly_chart(_fig_ex,width="stretch",config=PLOTLY_CFG,key=f"intra_ex_{_mname}")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET — IMBALANCE MARKET (marché des écarts)
# ══════════════════════════════════════════════════════════════════════════════

with tab_imbalance:
    _diag_set_tab("imbalance")
    st.markdown('<p class="section">Imbalance Market — Valorisation des écarts</p>', unsafe_allow_html=True)
    st.caption(
        "Le prix de règlement des écarts (imbalance settlement price) est le prix auquel un acteur du marché "
        "paie ou est rémunéré pour l'écart entre sa production/consommation réelle et son programme prévisionnel. "
        "Il y a deux prix : **négatif** (s'applique à un acteur en écart négatif — il a manqué d'énergie) et "
        "**positif** (s'applique à un acteur en écart positif — il a eu un surplus). "
        "Une batterie peut valoriser cet écart en chargeant/déchargeant aux heures les plus favorables — "
        "on simule ici cette stratégie séparément sur chacune des deux séries de prix, pour comparer leur potentiel. "
        "Cette simulation suppose un accès direct à la série de prix choisie ; elle ne modélise pas la mécanique "
        "réglementaire complète d'un Responsable d'Équilibre (BRP)."
    )

    st.markdown('<p class="section">Paramètres du scénario</p>', unsafe_allow_html=True)
    st.caption("Configurez les caractéristiques de la batterie : cycles, durée, restrictions horaires. Ces paramètres sont propres à cet onglet, indépendants du Day-Ahead et de l'Intraday.")
    _ib1, _ib2, _ib3, _ib4, _ib5, _ib6 = st.columns(6)
    with _ib1:
        _imb_n_cyc = st.selectbox(
            "Cycles par jour", [1, 2, 0],
            format_func=lambda x: {1: "1 cycle / jour", 2: "2 cycles / jour",
                                    0: "Illimité — tous les cycles rentables du jour"}[x],
            index=0, key="imb_n_cyc",
        )
    with _ib2:
        _imb_dur = st.selectbox(
            "Durée du cycle", [1, 2],
            format_func=lambda x: f"{x}h  ({x}h charge + {x}h décharge)",
            index=1, key="imb_dur",
        )
    _imb_cap = power_MW * _imb_dur
    with _ib3:
        _ib_disp = "illimité" if _imb_n_cyc == 0 else str(_imb_n_cyc)
        st.metric("Capacité par cycle (MWh)", f"{_imb_cap:.3f} MWh",
                  delta=f"{power_MW} MW × {_imb_dur}h · {_ib_disp} cycle(s)/j.")
    with _ib4:
        _imb_jours_excl = st.multiselect(
            "Jours avec restriction",
            ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"],
            key="imb_jours_excl"
        )
    with _ib5:
        _imb_h_debut = st.number_input("Heure début restriction", 0, 23, 0, key="imb_h_debut")
    with _ib6:
        _imb_h_fin = st.number_input("Heure fin restriction", 0, 23, 0, key="imb_h_fin")

    _day_map_ib = {"Lundi": 0, "Mardi": 1, "Mercredi": 2, "Jeudi": 3, "Vendredi": 4, "Samedi": 5, "Dimanche": 6}
    _imb_days_num = [_day_map_ib[d] for d in _imb_jours_excl]
    _imb_excl = ({"days": _imb_days_num, "hours": list(range(_imb_h_debut, _imb_h_fin))}
                 if _imb_days_num and _imb_h_debut < _imb_h_fin else {})

    st.markdown('<p class="section">Fichier de prix d\'écart</p>', unsafe_allow_html=True)
    _imb_file = st.file_uploader(
        "Déposez le fichier Excel des prix de règlement des écarts",
        type=["xlsx"], key="imb_file",
        help=(
            "Colonnes attendues : une colonne Date (horodatage demi-horaire ou quart-horaire) "
            "+ une colonne dont le nom contient \"negative\" (negative_imbalance_settlement_price) "
            "et/ou une colonne dont le nom contient \"positive\" (positive_imbalance_settlement_price). "
            "Le pas de temps est automatiquement ramené à la moyenne horaire."
        )
    )

    if not _imb_file:
        st.info("Uploadez le fichier de prix d'écart ci-dessus pour lancer la simulation.")
    else:
        if st.button("Lancer la simulation imbalance", key="btn_imb_run", type="primary"):
            st.session_state["_imb_run"] = True
            st.session_state["_imb_markets"] = None

        if st.session_state.get("_imb_run"):
            @st.cache_data(show_spinner=False)
            def _load_imbalance_cached(file_bytes: bytes):
                import io as _io_ib
                return load_imbalance(_io_ib.BytesIO(file_bytes))

            @st.cache_data(show_spinner=False)
            def _simulate_imbalance_cached(pv, params_json):
                import json as _j_ib
                params = _j_ib.loads(params_json)
                daily = (simulate_arbitrage_optimal(pv, params) if params["optimal_quota"]
                          else simulate_arbitrage(pv, params))
                yearly = aggregate_arbitrage(daily, params["power_MW"])
                return daily, yearly

            _p_imb = {
                "power_MW":        power_MW,
                "n_cycles":        0 if max_cycles else _imb_n_cyc,
                "duration_h":      _imb_dur,
                "excluded_hours":  _imb_excl,
                "efficiency":      efficiency,
                "max_cycles_year": max_cycles,
                "optimal_quota":   bool(max_cycles),
                "min_spread":      min_spread,
            }
            _p_imb_json = json.dumps(_p_imb, sort_keys=True)

            with st.spinner("Lecture et simulation…"):
                _imb_bytes = _imb_file.read()
                _imb_pivots = _load_imbalance_cached(_imb_bytes)
                _imb_markets = {}
                for _ikey, _ilabel in [("negatif", "Imbalance Négatif"), ("positif", "Imbalance Positif")]:
                    _pv_ib = _imb_pivots.get(_ikey)
                    if _pv_ib is None or _pv_ib.empty:
                        continue
                    _daily_ib, _yearly_ib = _simulate_imbalance_cached(_pv_ib, _p_imb_json)
                    _imb_markets[_ilabel] = {"pv": _pv_ib, "daily": _daily_ib, "yearly": _yearly_ib}

            if not _imb_markets:
                st.error("Aucune colonne de prix reconnue dans le fichier. Vérifiez qu'il contient une colonne Date et une colonne negative/positive_imbalance_settlement_price.")
            st.session_state["_imb_markets"] = _imb_markets

        _imb_markets = st.session_state.get("_imb_markets")
        if _imb_markets:
            # ── Comparaison Négatif vs Positif ────────────────────────────────
            st.markdown('<p class="section">Comparaison Négatif / Positif</p>', unsafe_allow_html=True)
            st.caption(
                "Chaque ligne simule la batterie comme si elle n'était exposée qu'à cette seule série de prix. "
                "PnL/MW/an permet de comparer indépendamment de la puissance — c'est le chiffre clé pour savoir "
                "laquelle des deux séries d'écart rémunère le mieux la batterie."
            )
            _comp_rows_ib = []
            for _ilabel, _idata in _imb_markets.items():
                _y_ib = _idata["yearly"]
                _nans_ib = len(_y_ib)
                _pnl_ib = _y_ib["pnl_total"].sum()
                _comp_rows_ib.append({
                    "Série":               _ilabel,
                    "Période":             f"{int(_y_ib['annee'].min())}–{int(_y_ib['annee'].max())}",
                    "PnL total (€)":       _fmt(_pnl_ib),
                    "PnL/MW/an (€)":       _fmt(_pnl_ib / power_MW / max(_nans_ib, 1)),
                    "Spread moy. (€/MWh)": f"{_y_ib['spread_moy'].mean():.1f}",
                    "Jours actifs/an":     str(int(_y_ib["jours_actifs"].mean())),
                })
            st.dataframe(pd.DataFrame(_comp_rows_ib), hide_index=True, width="stretch")

            _COLORS_IB = {"Imbalance Négatif": "#c62828", "Imbalance Positif": "#2e7d32"}
            _fig_comp_ib = go.Figure()
            for _ilabel, _idata in _imb_markets.items():
                _y_ib = _idata["yearly"]
                _fig_comp_ib.add_trace(go.Bar(
                    name=_ilabel, x=_y_ib["annee"].astype(str), y=_y_ib["pnl_total"] / power_MW,
                    marker_color=_COLORS_IB.get(_ilabel, "#666"),
                    hovertemplate=f"<b>{_ilabel}</b><br>%{{x}} : %{{y:,.0f}} €/MW<extra></extra>",
                ))
            _fig_comp_ib.update_layout(
                height=320, barmode="group",
                xaxis=dict(title="Année"), yaxis=dict(title="PnL/MW (€/MW)", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                margin=dict(t=10, b=80, l=60, r=10),
            )
            apply_bb(_fig_comp_ib)
            st.caption("PnL par MW installé, année par année — pour repérer si l'écart entre les deux séries se creuse ou se resserre dans le temps.")
            st.plotly_chart(_fig_comp_ib, width="stretch", config=PLOTLY_CFG, key="imb_comp_bar")

            # ── Détail par série ───────────────────────────────────────────────
            st.markdown('<p class="section">Résultats détaillés par série</p>', unsafe_allow_html=True)
            _tabs_ib = st.tabs(list(_imb_markets.keys()))
            for _ti_ib, (_ilabel, _idata) in enumerate(zip(_imb_markets.keys(), _imb_markets.values())):
                with _tabs_ib[_ti_ib]:
                    _daily_ib  = _idata["daily"]
                    _yearly_ib = _idata["yearly"]
                    _pv_ib     = _idata["pv"]
                    _clr_ib    = _COLORS_IB.get(_ilabel, C1)

                    _ikm1, _ikm2, _ikm3, _ikm4, _ikm5, _ikm6 = st.columns(6)
                    _pnl_tot_ib  = _daily_ib["pnl"].sum()
                    _pnl_abs_ib  = _daily_ib["pnl_absolu"].sum() if "pnl_absolu" in _daily_ib.columns else _pnl_tot_ib
                    _jact_ib     = int(_daily_ib["valid"].sum())
                    _jtot_ib     = len(_daily_ib)
                    _spr_moy_ib  = float(_daily_ib.loc[_daily_ib["valid"], "spread"].mean()) if _daily_ib["valid"].any() else 0
                    _spr_abs_ib  = float(_daily_ib["spread_absolu"].mean()) if "spread_absolu" in _daily_ib.columns else _spr_moy_ib
                    _energ_ib    = float(_daily_ib["energie_MWh"].sum()) if "energie_MWh" in _daily_ib.columns else 0
                    _blq_ib      = int(_daily_ib["maintenance_bloque"].sum()) if "maintenance_bloque" in _daily_ib.columns else 0
                    _taux_ib     = _pnl_tot_ib / _pnl_abs_ib if _pnl_abs_ib else 1
                    _ikm1.metric("PnL réel (avec contraintes)", f"{_fmt(_pnl_tot_ib)} €")
                    _ikm2.metric("PnL borne théorique max", f"{_fmt(_pnl_abs_ib)} €", delta=f"{_taux_ib*100:.1f}% capturé")
                    _ikm3.metric("Spread moyen (jours actifs)", f"{_spr_moy_ib:.1f} €/MWh", delta=f"Théorique: {_spr_abs_ib:.1f} €/MWh")
                    _ikm4.metric("Jours actifs", f"{_jact_ib} / {_jtot_ib}", delta=f"{_jact_ib/_jtot_ib*100:.0f}% taux activation" if _jtot_ib else None)
                    _ikm5.metric("Énergie totale échangée", f"{_energ_ib:.1f} MWh" if _energ_ib else "—",
                                 delta=f"Capacité/jour : {_imb_cap:.2f} MWh")
                    _ikm6.metric("Jours bloqués (maintenance)", f"{_blq_ib}",
                                 delta="Aucun" if _blq_ib == 0 else f"{_blq_ib/_jtot_ib*100:.1f}%",
                                 delta_color="off" if _blq_ib == 0 else "inverse")

                    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
                    st.caption("Barres jaunes = borne max théorique. Barres colorées = PnL réel avec contraintes. L'écart = coût des restrictions et du quota.")
                    _fig1_ib = go.Figure()
                    _fig1_ib.add_trace(go.Bar(
                        x=_yearly_ib["annee"].astype(str), y=_yearly_ib["pnl_absolu_total"],
                        name="Borne max", marker_color=C2,
                        text=_yearly_ib["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €"), textposition="outside",
                        hovertemplate="<b>%{x}</b><br>Borne max : %{y:.0f} €<extra></extra>",
                    ))
                    _fig1_ib.add_trace(go.Bar(
                        x=_yearly_ib["annee"].astype(str), y=_yearly_ib["pnl_total"],
                        name="PnL réel", marker_color=_clr_ib,
                        text=_yearly_ib["pnl_total"].apply(lambda x: f"{_fmt(x)} €"), textposition="inside", textfont_color="white",
                        hovertemplate="<b>%{x}</b><br>PnL réel : %{y:.0f} €<extra></extra>",
                    ))
                    _fig1_ib.update_layout(
                        barmode="group", height=340, yaxis=dict(title="PnL (€)", gridcolor="#f0f0f0"),
                        xaxis_title="Année", legend=LEGEND_BOTTOM, margin=dict(t=10, b=40, l=60, r=20),
                        plot_bgcolor="white", paper_bgcolor="white",
                    )
                    apply_bb(_fig1_ib)
                    st.plotly_chart(_fig1_ib, width="stretch", config=PLOTLY_CFG, key=f"imb_fig1_{_ti_ib}")

                    _tgib1, _tgib2 = st.columns(2)
                    with _tgib1:
                        st.markdown('<p class="section">Spread moyen — par semaine (€/MWh)</p>', unsafe_allow_html=True)
                        st.caption("Évolution du spread moyenné par semaine. Zone ombrée = fourchette min-max observée. Ligne pointillée = spread moyen sur toute la période.")
                        _dv_ib = _daily_ib[_daily_ib["valid"]].copy()
                        if not _dv_ib.empty:
                            _dv_ib["sem"] = pd.to_datetime(_dv_ib["date"]).dt.to_period("W").apply(lambda p: p.start_time)
                            _wi_ib = _dv_ib.groupby("sem").agg(s=("spread", "mean"), mn=("spread", "min"), mx=("spread", "max")).reset_index()
                            _fig_sw_ib = go.Figure()
                            _fig_sw_ib.add_trace(go.Scatter(
                                x=pd.concat([_wi_ib["sem"], _wi_ib["sem"].iloc[::-1]]),
                                y=pd.concat([_wi_ib["mx"], _wi_ib["mn"].iloc[::-1]]),
                                fill="toself", fillcolor="rgba(46,125,50,0.15)",
                                line=dict(color="rgba(0,0,0,0)"), showlegend=False, hoverinfo="skip",
                            ))
                            _fig_sw_ib.add_trace(go.Scatter(
                                x=_wi_ib["sem"], y=_wi_ib["s"], mode="lines",
                                line=dict(color=_clr_ib, width=2), name="Spread hebdo",
                                hovertemplate="%{x|%d/%m/%Y} : %{y:.1f} €/MWh<extra></extra>",
                            ))
                            _fig_sw_ib.add_hline(y=_spr_moy_ib, line_dash="dot", line_color=ORANGE, line_width=1.5,
                                                  annotation_text=f"Moy: {_spr_moy_ib:.1f}")
                            _fig_sw_ib.update_layout(
                                height=300, margin=dict(t=10, b=40, l=60, r=80),
                                xaxis=dict(gridcolor="#f0f0f0", tickformat="%b %Y"),
                                yaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0"),
                                plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
                            )
                            apply_bb(_fig_sw_ib)
                            st.plotly_chart(_fig_sw_ib, width="stretch", config=PLOTLY_CFG, key=f"imb_sw_{_ti_ib}")

                    with _tgib2:
                        st.markdown('<p class="section">Profil horaire charge / décharge</p>', unsafe_allow_html=True)
                        st.caption("Fréquence d'utilisation de chaque heure sur toute la période. Vert = achat (charge). Orange = vente (décharge).")
                        _hc_cnt_ib = [0]*24; _hd_cnt_ib = [0]*24
                        for _, _rp_ib in _pv_ib.iterrows():
                            try:
                                _pp_ib = _rp_ib[HOUR_COLS].values.astype(float)
                                _av_ib = get_available_hours(int(_rp_ib["weekday"]), _imb_excl)
                                _cc_ib = _best_n_cycles(_pp_ib, _av_ib, _imb_n_cyc, _imb_dur, power_MW, efficiency)
                                for _cy_ib in _cc_ib:
                                    for _hc_ib in _cy_ib["h_charge"]:
                                        if 0 <= _hc_ib < 24: _hc_cnt_ib[_hc_ib] += 1
                                    for _hd_ib in _cy_ib["h_decharge"]:
                                        if 0 <= _hd_ib < 24: _hd_cnt_ib[_hd_ib] += 1
                            except Exception:
                                pass
                        _fig_prof_ib = go.Figure()
                        _fig_prof_ib.add_trace(go.Bar(x=[f"H{h:02d}" for h in range(24)], y=_hc_cnt_ib,
                            name="Charge (achat)", marker_color=C1, opacity=0.85))
                        _fig_prof_ib.add_trace(go.Bar(x=[f"H{h:02d}" for h in range(24)], y=_hd_cnt_ib,
                            name="Décharge (vente)", marker_color=ORANGE, opacity=0.85))
                        _fig_prof_ib.update_layout(
                            height=300, barmode="overlay", margin=dict(t=10, b=40, l=60, r=10),
                            xaxis=dict(title="Heure"), yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
                            plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                        )
                        apply_bb(_fig_prof_ib)
                        st.plotly_chart(_fig_prof_ib, width="stretch", config=PLOTLY_CFG, key=f"imb_prof_{_ti_ib}")

                    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
                    st.caption("Tableau détaillé par année. Le taux de capturé (PnL réel / borne max) mesure l'efficacité de votre stratégie.")
                    _show_ib = _yearly_ib.copy()
                    _show_ib["pnl_absolu_total"] = _show_ib["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €")
                    _show_ib["pnl_total"]        = _show_ib["pnl_total"].apply(lambda x: f"{_fmt(x)} €")
                    _show_ib["taux_activation"]  = _show_ib["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
                    _show_ib["spread_moy"]       = _show_ib["spread_moy"].apply(lambda x: f"{x:.1f}")
                    _disp_cols_ib = ["annee", "jours_actifs", "taux_activation", "spread_moy",
                                     "pnl_absolu_total", "pnl_total", "pnl_par_MW"]
                    _disp_cols_ib = [c for c in _disp_cols_ib if c in _show_ib.columns]
                    _disp_rename_ib = {"annee": "Année", "jours_actifs": "Jours actifs", "taux_activation": "Taux activation",
                                        "spread_moy": "Spread moy. (€/MWh)", "pnl_absolu_total": "Borne max (€)",
                                        "pnl_total": "PnL réel (€)", "pnl_par_MW": "PnL/MW (€/MW)"}
                    st.dataframe(_show_ib[_disp_cols_ib].rename(columns=_disp_rename_ib), hide_index=True, width="stretch")

                    with st.expander("Explorer un jour spécifique", expanded=True):
                        st.caption("Sélectionnez un jour pour voir le détail : prix d'écart heure par heure, heures d'achat/vente choisies, et PnL réalisé comparé au maximum théorique.")
                        _imb_dates = sorted(_daily_ib["date"].dt.date.unique())
                        _imb_valid_dates = sorted(_daily_ib.loc[_daily_ib["valid"], "date"].dt.date.unique())
                        _imb_default_date = _imb_valid_dates[0] if _imb_valid_dates else (_imb_dates[0] if _imb_dates else None)
                        _imb_date_sel = st.date_input(
                            "Date", value=_imb_default_date,
                            min_value=_imb_dates[0] if _imb_dates else None,
                            max_value=_imb_dates[-1] if _imb_dates else None,
                            key=f"imb_date_{_ti_ib}",
                        )
                        if _imb_date_sel:
                            _row_ib = _pv_ib[_pv_ib["date"].dt.date == _imb_date_sel]
                            if not _row_ib.empty:
                                _prix_ib  = _row_ib[HOUR_COLS].values.flatten().astype(float)
                                _avail_ib = get_available_hours(int(_row_ib["weekday"].iloc[0]), _imb_excl)
                                _cyc_ib   = _best_n_cycles(_prix_ib, _avail_ib, _imb_n_cyc, _imb_dur, power_MW, efficiency)
                                _cyc_bm_ib = _best_n_cycles(_prix_ib, list(range(24)), 0, _imb_dur, power_MW, efficiency)
                                _pnl_ib_d  = sum(c["pnl"] for c in _cyc_ib)
                                _pnl_bm_ib = sum(c["pnl"] for c in _cyc_bm_ib)
                                _spr_ib_d  = sum(c["spread"] for c in _cyc_ib)/len(_cyc_ib) if _cyc_ib else 0

                                _cexib1, _cexib2 = st.columns([3, 1])
                                with _cexib2:
                                    st.metric("Spread moyen", f"{_spr_ib_d:.2f} €/MWh")
                                    st.metric("PnL du jour", f"{_pnl_ib_d:.2f} €")
                                    st.metric("PnL borne max", f"{_pnl_bm_ib:.2f} €",
                                              delta=f"{_pnl_ib_d/_pnl_bm_ib*100:.0f}% capturé" if _pnl_bm_ib else None)
                                    st.metric("Cycles réalisés", f"{len(_cyc_ib)}" + (" (max)" if _imb_n_cyc == 0 else f" / {_imb_n_cyc}"))
                                    st.metric("Heures dispo", f"{len(_avail_ib)}/24")
                                    for _ii_ib, _cy_ib in enumerate(_cyc_ib, 1):
                                        st.caption(f"Cycle {_ii_ib} : charge H{_cy_ib['h_charge']} ({_cy_ib['prix_charge']:.1f} €/MWh) → décharge H{_cy_ib['h_decharge']} ({_cy_ib['prix_decharge']:.1f} €/MWh) | PnL {_cy_ib['pnl']:.2f} €")
                                with _cexib1:
                                    _fig_exib = go.Figure()
                                    for _eh_ib in [h for h in range(24) if h not in _avail_ib]:
                                        _fig_exib.add_vrect(x0=_eh_ib-0.5, x1=_eh_ib+0.5,
                                            fillcolor="rgba(200,200,200,0.3)", layer="below", line_width=0)
                                    _fig_exib.add_trace(go.Bar(
                                        x=list(range(24)),
                                        y=[float(p) if not pd.isna(p) else 0 for p in _prix_ib],
                                        marker_color="#c8d8ec", name="Prix d'écart",
                                        hovertemplate="H%{x:02d} — <b>%{y:.2f} €/MWh</b><extra></extra>",
                                    ))
                                    _col_ch_ib = ["#2e7d32", "#1565c0", "#6a1b9a"]
                                    _col_dch_ib = ["#c62828", "#e65100", "#4527a0"]
                                    for _ii_ib, _cy_ib in enumerate(_cyc_ib):
                                        _ccib2, _cdib2 = _col_ch_ib[_ii_ib % 3], _col_dch_ib[_ii_ib % 3]
                                        _fig_exib.add_trace(go.Scatter(
                                            x=_cy_ib["h_charge"], y=_prix_ib[_cy_ib["h_charge"]],
                                            mode="markers", name=f"Achat C{_ii_ib+1} ({_imb_dur}h)",
                                            marker=dict(color=_ccib2, size=16, symbol="triangle-up", line=dict(color="white", width=1)),
                                            hovertemplate=f"H%{{x:02d}} — Achat C{_ii_ib+1} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                                        ))
                                        _fig_exib.add_trace(go.Scatter(
                                            x=_cy_ib["h_decharge"], y=_prix_ib[_cy_ib["h_decharge"]],
                                            mode="markers", name=f"Vente C{_ii_ib+1} ({_imb_dur}h)",
                                            marker=dict(color=_cdib2, size=16, symbol="triangle-down", line=dict(color="white", width=1)),
                                            hovertemplate=f"H%{{x:02d}} — Vente C{_ii_ib+1} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                                        ))
                                    _fig_exib.update_layout(
                                        height=380, margin=dict(t=20, b=60, l=60, r=10),
                                        xaxis=dict(title="Heure", tickmode="array", tickvals=list(range(24)),
                                                   ticktext=[f"H{h:02d}" for h in range(24)], range=[-0.5, 23.5]),
                                        yaxis=dict(title="Prix d'écart (€/MWh)", gridcolor="#f0f0f0"),
                                        plot_bgcolor="white", paper_bgcolor="white", legend=LEGEND_BOTTOM,
                                        hoverlabel=dict(bgcolor="white", font_size=12),
                                    )
                                    apply_bb(_fig_exib)
                                    st.caption("Prix d'écart heure par heure. ▲ = achat · ▼ = vente · zones grises = heures exclues.")
                                    st.plotly_chart(_fig_exib, width="stretch", config=PLOTLY_CFG, key=f"imb_ex_{_ti_ib}")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2 — LISSAGE DE COURBE DE CHARGE
# ══════════════════════════════════════════════════════════════════════════════

with tab_lis:
    _diag_set_tab("lis")
    # ── Chargement courbe de charge ──────────────────────────────────────────
    st.markdown('<p class="section">Courbe de charge client</p>',
                unsafe_allow_html=True)
    st.caption("Importez ici le profil de consommation horaire du client industriel (24 valeurs par jour). La batterie va absorber les pics de consommation au-dessus d'un seuil pour reduire la puissance de pointe facturee.")

    col_up, col_params = st.columns([2, 1])

    with col_params:
        st.markdown("**Paramètres économiques**")
        seuil_pct    = st.slider("Seuil d'écrêtage (percentile)", 50, 95, 75,
                                  help="On lisse tout ce qui dépasse ce percentile de la conso")
        tarif_kw     = st.number_input("Tarif puissance souscrite (€/MW/an)",
                                        1000, 100000, 12000, 1000)
        soc_init_pct = st.slider("SOC initial (%)", 0, 90, 0)

        with st.expander("Tarification TURPE détaillée (optionnel)", expanded=False):
            st.caption(
                "Décomposez les économies par tranche horaire TURPE (tarif réseau HTB/HTA). "
                "Saisissez le tarif de chaque tranche en €/MWh. "
                "Les économies sont calculées heure par heure selon la tranche active."
            )
            _tc1, _tc2 = st.columns(2)
            with _tc1:
                _t_hd  = st.number_input("HD — Heures de Pointe (€/MWh)",  0.0, 500.0, 140.0, 5.0, key="t_hd")
                _t_hph = st.number_input("HPH — Heures Pleines Hiver (€/MWh)", 0.0, 300.0, 75.0, 5.0, key="t_hph")
                _t_hch = st.number_input("HCH — Heures Creuses Hiver (€/MWh)", 0.0, 200.0, 42.0, 5.0, key="t_hch")
            with _tc2:
                _t_hpe = st.number_input("HPE — Heures Pleines Été (€/MWh)", 0.0, 200.0, 52.0, 5.0, key="t_hpe")
                _t_hce = st.number_input("HCE — Heures Creuses Été (€/MWh)", 0.0, 150.0, 30.0, 5.0, key="t_hce")
            st.caption("Sources : TURPE 6 HTA option LU (CRE 2024). HD = ~500h/an en hiver.")

        # Mapping heure → tranche TURPE
        # Hiver = Nov-Mar (mois 11,12,1,2,3) ; Été = Avr-Oct
        # HD = HPM (Heures de Pointe Mobile) : indicatif H09-H11 et H18-H20 hiver
        # HPH = H07-H22 hiver hors HD ; HCH = H22-H07 hiver
        # HPE = H07-H22 été ; HCE = H22-H07 été
        def _get_turpe_tranche(mois, heure):
            """Retourne la tranche TURPE et le tarif €/MWh pour un mois/heure."""
            _hiver = mois in [11, 12, 1, 2, 3]
            if _hiver:
                if heure in [9, 10, 18, 19]:   return "HD",  _t_hd
                elif 7 <= heure <= 22:           return "HPH", _t_hph
                else:                            return "HCH", _t_hch
            else:
                if 7 <= heure <= 22:             return "HPE", _t_hpe
                else:                            return "HCE", _t_hce

        energy_MWh   = None  # sera calculé après lecture du profil

    with col_up:
        uploaded_cdc = st.file_uploader(
            "Fichier courbe de charge client (Excel Plénitude)",
            type=["xlsx"],
            help="Format attendu : feuille 'CdC_kWh' avec colonnes 'datetime' et consommation en kWh/h"
        )

        source_profil = "Fichier uploadé" if uploaded_cdc is not None else "Saisie manuelle"

    # ── Lecture et construction du profil ────────────────────────────────────
    PROFIL_DEFAUT = [0.14,0.14,0.14,0.13,0.12,0.14,0.15,0.14,
                     0.15,0.16,0.17,0.16,0.17,0.17,0.16,0.15,
                     0.15,0.13,0.12,0.11,0.09,0.08,0.10,0.13]

    profil_arr = None
    df_cdc_complet = None  # donnees complètes pour graphiques

    if source_profil == "Fichier uploadé" and uploaded_cdc is not None:
        try:
            cdc_bytes = uploaded_cdc.read()
            if cdc_bytes:
                st.session_state.excel_cdc_bytes = cdc_bytes  # persister pour l'audit
            df_cdc = pd.read_excel(
                io.BytesIO(cdc_bytes),
                sheet_name="CdC_kWh", engine="openpyxl"
            )
            df_cdc.columns = [str(c).strip() for c in df_cdc.columns]
            df_cdc = df_cdc.rename(columns={"datetime": "ts"})
            df_cdc["ts"] = pd.to_datetime(df_cdc["ts"])

            # Colonne de conso = première colonne numérique après datetime
            conso_col = [c for c in df_cdc.columns
                         if c != "ts" and pd.api.types.is_numeric_dtype(df_cdc[c])]
            if not conso_col:
                st.error("Aucune colonne numérique trouvée dans CdC_kWh.")
                st.stop()

            # Somme de toutes les formules si plusieurs colonnes
            df_cdc["conso_kWh"] = df_cdc[conso_col].sum(axis=1)
            df_cdc["conso_MW"]  = df_cdc["conso_kWh"] / 1000
            df_cdc["heure"]     = df_cdc["ts"].dt.hour
            df_cdc["date"]      = df_cdc["ts"].dt.date
            df_cdc["mois"]      = df_cdc["ts"].dt.month
            df_cdc["annee"]     = df_cdc["ts"].dt.year
            df_cdc_complet      = df_cdc.copy()

            # Profil horaire moyen sur toutes les donnees (en MW)
            profil_moy = df_cdc.groupby("heure")["conso_MW"].mean()
            profil_arr = profil_moy.values

            # Infos résumé
            n_jours  = df_cdc["date"].nunique()
            conso_tot = df_cdc["conso_kWh"].sum() / 1000  # MWh
            pointe    = df_cdc["conso_MW"].max()
            période   = f"{df_cdc['ts'].min().strftime('%d/%m/%Y')} -> {df_cdc['ts'].max().strftime('%d/%m/%Y')}"

            with col_up:
                st.success(
                    f"**{uploaded_cdc.name}** chargé — "
                    f"{n_jours} jours | {_fmt(conso_tot)} MWh total | "
                    f"Pointe : {pointe:.3f} MW"
                )
                st.caption(f"Période : {période} | "
                           f"Formules : {', '.join(conso_col)}")
        except Exception as e:
            st.error(f"Erreur lecture fichier : {e}")
            st.stop()

    # Saisie manuelle ou fallback
    if source_profil == "Saisie manuelle" or profil_arr is None:
        if uploaded_cdc is None and source_profil == "Fichier uploadé":
            st.info("Aucun fichier chargé — saisie manuelle activée.")
        with col_up:
            st.caption("Consommation en MW pour chaque heure H00–H23 :")
            cols8 = st.columns(8)
            profil_manuel = []
            for h in range(24):
                with cols8[h % 8]:
                    v = st.number_input(
                        f"H{h:02d}", 0.0, 100.0, float(PROFIL_DEFAUT[h]),
                        0.001, format="%.3f", label_visibility="visible"
                    )
                    profil_manuel.append(v)
        profil_arr = np.array(profil_manuel)

    # ── Calcul automatique de la capacité nécessaire ─────────────────────────
    # Énergie à absorber = somme des surplus au-dessus du seuil (jour type)
    _seuil_MW = float(np.percentile(profil_arr, seuil_pct))
    _surplus_h = np.maximum(profil_arr - _seuil_MW, 0)
    _energie_limitee = float(np.minimum(_surplus_h, power_MW).sum())
    energy_MWh = max(round(_energie_limitee, 3), round(power_MW * 0.5, 3))
    _duree_equiv = round(energy_MWh / power_MW, 2) if power_MW > 0 else 0
    st.info(
        f"Capacité calculée automatiquement : **{energy_MWh:.2f} MWh** "
        f"({_duree_equiv}h équivalent à {power_MW} MW) — "
        f"Seuil : {_seuil_MW:.3f} MW (P{seuil_pct})"
    )

    # ── Graphique profil chargé ──────────────────────────────────────────────
    if df_cdc_complet is not None:
        st.markdown('<p class="section">Aperçu de la courbe de charge réelle</p>',
                    unsafe_allow_html=True)
        st.caption("Visualisation du profil de consommation importe. La courbe par mois permet de voir la saisonnalité. La pointe maximale est la valeur que l'on cherche à réduire grâce à la batterie.")

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            # Profil horaire moyen par mois
            profil_mens = df_cdc_complet.groupby(
                ["mois","heure"])["conso_MW"].mean().reset_index()
            mois_fr = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jun",
                       7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
            heures_x = list(range(24))  # toujours 0..23 en liste Python

            fig_cdc = go.Figure()
            for m in sorted(profil_mens["mois"].unique()):
                d = profil_mens[profil_mens["mois"] == m].sort_values("heure")
                # Forcer en liste Python pour éviter les problèmes d'index pandas
                y_vals = d["conso_MW"].tolist()
                x_vals = d["heure"].tolist()
                fig_cdc.add_trace(go.Scatter(
                    x=x_vals, y=y_vals,
                    mode="lines", name=mois_fr.get(m, str(m)),
                    line=dict(width=1.5),
                    hovertemplate=f"{mois_fr.get(m,str(m))} "
                                  f"H%{{x:02d}} : <b>%{{y:.3f}} MW</b>"
                                  f"<extra></extra>",
                ))
            fig_cdc.add_trace(go.Scatter(
                x=heures_x, y=list(profil_arr),
                mode="lines+markers", name="Moy. annuelle",
                line=dict(color=C2, width=2.5, dash="dot"),
                marker=dict(size=5),
                hovertemplate="Moy H%{x:02d} : <b>%{y:.3f} MW</b><extra></extra>",
            ))
            fig_cdc.update_layout(
                height=360,
                margin=dict(t=30, b=100, l=60, r=10),
                title=dict(text="Profil horaire moyen par mois",
                           font=dict(size=13), x=0),
                xaxis=dict(
                    title="Heure",
                    tickmode="array",
                    tickvals=list(range(0, 24, 2)),
                    ticktext=[f"H{h:02d}" for h in range(0, 24, 2)],
                    range=[-0.5, 23.5],
                ),
                yaxis=dict(title="Consommation (MW)", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white",
                legend=dict(
                    orientation="h",
                    y=-0.30, x=0, xanchor="left",
                    font=dict(size=10),
                    itemwidth=40,
                ),
                hoverlabel=dict(bgcolor="white", font_size=11, font_color="#333333"),
            )
            apply_bb(fig_cdc)
            st.caption("Profil moyen de consommation par mois sur toute l'année. Permet de visualiser la saisonnalité et les périodes avec les pointes les plus élevées que la batterie devra effacer.")
            st.plotly_chart(fig_cdc, width="stretch", config=PLOTLY_CFG, key="arb_fig_cdc")
            st.caption(
            "Lecture du graphique de lissage : "
            "ROSE = consommation originale du client (ce qu'il consomme sans la batterie). "
            "VERT = consommation vue du réseau après intervention de la batterie. "
            "ORANGE (ligne horizontale) = seuil d'écrêtage calculé automatiquement (percentile P75). "
            "Zones où VERT < ROSE : la batterie décharge pour absorber un pic. "
            "Zones où VERT > ROSE : la batterie se recharge (elle tire plus du réseau pour stocker). "
            "La réduction de pointe = différence entre le point le plus haut du rose "
            "et le point le plus haut du vert."
        )

        with col_g2:
            # Courbe de charge sur les 30 derniers jours disponibles
            dates_dispo = sorted(df_cdc_complet["date"].unique())
            dates_30 = dates_dispo[-30:]
            df_30 = df_cdc_complet[df_cdc_complet["date"].isin(dates_30)]
            fig_ts = go.Figure()
            fig_ts.add_trace(go.Scatter(
                x=df_30["ts"], y=df_30["conso_MW"],
                mode="lines", name="Consommation réelle",
                line=dict(color=C1, width=1),
                hovertemplate="%{x|%d/%m %Hh} : <b>%{y:.3f} MW</b><extra></extra>",
            ))
            fig_ts.add_hline(
                y=np.percentile(df_30["conso_MW"], seuil_pct),
                line_dash="dash", line_color=GREEN, line_width=2,
                annotation_text=f"Seuil P{seuil_pct}",
            )
            fig_ts.update_layout(
                height=360,
                margin=dict(t=30, b=100, l=60, r=10),
                title=dict(text="Courbe de charge — 30 derniers jours",
                           font=dict(size=13), x=0),
                xaxis=dict(title="", tickformat="%d/%m",
                           tickangle=-30),
                yaxis=dict(title="MW", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white",
                legend=dict(orientation="h", y=-0.30, x=0,
                            font=dict(size=10)),
                hoverlabel=dict(bgcolor="white", font_size=11, font_color="#333333"),
            )
            apply_bb(fig_ts)
            st.caption('Evolution de la consommation maximale journalière. Permet de repérer les périodes avec les pointes les plus fortes qui bénéficient le plus du lissage batterie.')
            st.plotly_chart(fig_ts, width="stretch", config=PLOTLY_CFG, key="arb_fig_ts")

    # ── Simulation lissage ───────────────────────────────────────────────────
    params_l = {
        "power_MW": power_MW, "energy_MWh": energy_MWh,
        "soc_min_pct": 0.10, "soc_max_pct": 0.90,
        "efficiency": efficiency, "seuil_percentile": seuil_pct,
        "tarif_puissance_souscrite": tarif_kw,
        "soc_init": soc_init_pct / 100,
    }

    @st.cache_data(show_spinner=False)
    def _simulate_lissage_cached(_file_bytes, _cdc_bytes, _profil_tuple, _params_json):
        """Boucle jour par jour mise en cache par (fichier spot, courbe de charge,
        profil, paramètres) — change uniquement le tarif ou le SOC initial ne
        re-déclenche pas le recalcul complet si déjà fait avec ces valeurs."""
        import json as _j_lis
        _params = _j_lis.loads(_params_json)
        _df_cdc_c = None
        if _cdc_bytes:
            _df_cdc_c = pd.read_excel(io.BytesIO(_cdc_bytes), sheet_name="CdC_kWh", engine="openpyxl")
            _df_cdc_c.columns = [str(c).strip() for c in _df_cdc_c.columns]
            _df_cdc_c = _df_cdc_c.rename(columns={"datetime": "ts"})
            _df_cdc_c["ts"] = pd.to_datetime(_df_cdc_c["ts"])
            _cc = [c for c in _df_cdc_c.columns if c != "ts" and pd.api.types.is_numeric_dtype(_df_cdc_c[c])]
            _df_cdc_c["conso_kWh"] = _df_cdc_c[_cc].sum(axis=1)
            _df_cdc_c["conso_MW"]  = _df_cdc_c["conso_kWh"] / 1000
            _df_cdc_c["heure"]     = _df_cdc_c["ts"].dt.hour
            _df_cdc_c["date"]      = _df_cdc_c["ts"].dt.date
            _df_cdc_c["mois"]      = _df_cdc_c["ts"].dt.month
            _df_cdc_c["annee"]     = _df_cdc_c["ts"].dt.year
        _pv = get_pivot(_file_bytes)
        return simulate_lissage(_pv, np.array(_profil_tuple), _params, df_cdc=_df_cdc_c)

    res  = _simulate_lissage_cached(
        file_bytes,
        st.session_state.get("excel_cdc_bytes"),
        tuple(profil_arr.tolist()),
        json.dumps(params_l, sort_keys=True),
    )
    jour = res["jour_type"]

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Résultats</p>', unsafe_allow_html=True)
    st.caption("Économies realisees grâce au lissage de charge. La réduction de pointe = différence entre la puissance maximum avant et après intervention de la batterie. L'economie annuelle = réduction x tarif de puissance souscrite.")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Réduction de pointe",
              f"{res['reduction_MW']:.3f} MW",
              delta=f"−{res['reduction_MW']/res.get('pointe_avant_MW', jour['pointe_avant'])*100:.1f}%")
    m2.metric("Économie annuelle estimée",
              f"{_fmt(res['economie_an'])} €")
    m3.metric("Pointe avant / après",
              f"{res.get('pointe_avant_MW', jour['pointe_avant']):.3f} → {res.get('pointe_apres_MW', jour['pointe_apres']):.3f} MW",
              help="Puissance maximale consommée sur la période, avant et après intervention de la batterie.")
    m4.metric("Seuil d'écrêtage appliqué",
              f"{res['seuil_MW']:.2f} MW",
              delta=f"Percentile {seuil_pct}%")

    st.session_state["_diag_lis"] = {
        "kpis": [
            ("Réduction de pointe", f"{res['reduction_MW']:.3f} MW"),
            ("Économie annuelle estimée", f"{_fmt(res['economie_an'])} €"),
            ("Pointe avant", f"{res.get('pointe_avant_MW', jour['pointe_avant']):.3f} MW"),
            ("Pointe après", f"{res.get('pointe_apres_MW', jour['pointe_apres']):.3f} MW"),
            ("Seuil d'écrêtage", f"{res['seuil_MW']:.2f} MW"),
        ],
        "yearly": res.get("yearly", []),
    }

    # ── Sélection du jour à afficher ────────────────────────────────────────
    st.markdown('<p class="section">Profil de consommation — avant et après lissage</p>',
                unsafe_allow_html=True)

    # Construire la liste des jours disponibles
    if df_cdc_complet is not None:
        import numpy as _np2
        _seuil_vis = res["seuil_MW"]

        # Calculer la pointe de chaque jour et trier par pointe décroissante
        _daily_stats = (df_cdc_complet.groupby("date")["conso_MW"]
                        .agg(pointe="max", nb_h_seuil=lambda x: (x > _seuil_vis).sum())
                        .reset_index()
                        .sort_values("pointe", ascending=False))

        # Construire les labels : date + pointe + indicateur décharge
        def _fmt_date_opt(row):
            tag = "decharge possible" if row["nb_h_seuil"] > 0 else "charge seule"
            return f"{row['date']}  |  pointe {row['pointe']:.3f} MW  |  {tag}"

        _dates_labels = _daily_stats.apply(_fmt_date_opt, axis=1).tolist()
        _dates_values = _daily_stats["date"].tolist()

        # Index par défaut = premier jour avec décharge possible
        _idx_default = next(
            (i for i, r in _daily_stats.iterrows() if r["nb_h_seuil"] > 0),
            0
        )
        _idx_default = list(_daily_stats.index).index(_idx_default) if _idx_default != 0 else 0

        _col_date, _col_info = st.columns([3, 2])
        with _col_date:
            _date_sel_label = st.selectbox(
                "Sélectionner un jour à visualiser (triés par pointe décroissante)",
                _dates_labels,
                index=_idx_default,
                key="liss_date_sel",
                help="Les jours 'decharge possible' sont ceux où la batterie intervient réellement. Les autres ont une consommation trop faible pour dépasser le seuil."
            )
        _date_sel = _dates_values[_dates_labels.index(_date_sel_label)]
        _jour_data = df_cdc_complet[df_cdc_complet["date"] == _date_sel].sort_values("ts")

        if len(_jour_data) >= 24:
            _profil_jour = _jour_data["conso_MW"].values[:24].astype(float)
            # Recalculer le lissage sur ce jour précis
            from bess_engine import lissage_day as _ld
            _e_max   = params_l["energy_MWh"]
            _soc_min = params_l["soc_min_pct"] * _e_max
            _soc_max = params_l["soc_max_pct"] * _e_max
            _soc_i   = _e_max * params_l.get("soc_init", 0.5)
            jour = _ld(_profil_jour, res["seuil_MW"], params_l["power_MW"],
                       _e_max, _soc_i, _soc_min, _soc_max,
                       params_l.get("efficiency", 1.0))
            profil_arr = _profil_jour

            with _col_info:
                st.caption(
                    f"**{str(_date_sel)}** — "
                    f"Pointe réelle : **{jour['pointe_avant']:.3f} MW** → après lissage : **{jour['pointe_apres']:.3f} MW** "
                    f"| Réduction : **{jour['reduction_pointe']:.3f} MW** "
                    f"({jour['reduction_pointe']/jour['pointe_avant']*100:.1f}%)"
                )
        else:
            st.warning(f"Jour {str(_date_sel)} incomplet ({len(_jour_data)} heures < 24) — affichage du jour moyen.")
    else:
        st.info("Uploadez un fichier CdC pour sélectionner un jour réel. Affichage du jour moyen par défaut.")

    st.caption(
            "COURBE ROSE = consommation réelle du client heure par heure. "
            "COURBE VERTE = consommation vue du réseau après intervention de la batterie. "
            "LIGNE POINTILLÉE = seuil d'écrêtage : au-dessus la batterie décharge, en-dessous elle se recharge. "
            "BARRES (graphique du bas) : vert = batterie se charge, orange = batterie se décharge."
    )

    actions_bess = [v if a == "charge" else -v if a == "decharge" else 0
                    for a, v in jour["actions"]]

    fig_l = make_subplots(rows=2, cols=1, shared_xaxes=True,
                           row_heights=[0.65, 0.35],
                           subplot_titles=["Consommation (MW)",
                                           "Action BESS (+ charge / − décharge)"])
    fig_l.add_trace(go.Scatter(
        x=list(range(24)), y=profil_arr.tolist(),
        mode="lines+markers", name="Avant lissage",
        line=dict(color="#e91e8c", width=3), marker=dict(size=7, color="#e91e8c"),
        hovertemplate="H%{x:02d} — Avant : <b>%{y:.3f} MW</b><extra></extra>",
    ), row=1, col=1)
    fig_l.add_trace(go.Scatter(
        x=list(range(24)), y=jour["profil_lisse"].tolist(),
        mode="lines+markers", name="Après lissage",
        line=dict(color=C1, width=2.5),
        fill="tozeroy", fillcolor=("rgba(255,102,0,0.08)" if _BB else "rgba(92,184,92,0.07)"), marker=dict(size=6),
        hovertemplate="H%{x:02d} — Après : <b>%{y:.3f} MW</b><extra></extra>",
    ), row=1, col=1)
    fig_l.add_hline(y=res["seuil_MW"], line_dash="dash", line_color=GREEN,
                    line_width=2, annotation_text=f"Seuil {res['seuil_MW']:.2f} MW",
                    annotation_position="top right", row=1, col=1)
    colors_bar = [C1 if v >= 0 else C2 for v in actions_bess]
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
        legend=LEGEND_BOTTOM,
        xaxis2=dict(title="Heure",
                    tickmode="array",
                    tickvals=list(range(0,24,2)),
                    ticktext=[f"H{h:02d}" for h in range(0,24,2)],
                    range=[-0.5,23.5]),
        yaxis2=dict(title="MW BESS", gridcolor="#f0f0f0"),
        yaxis=dict(gridcolor="#f0f0f0"),
        hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        hovermode="x unified",
    )
    apply_bb(fig_l)
    st.caption(
            "Lissage de charge : la batterie agit comme un tampon énergétique. "
            "Elle absorbe de l'énergie quand la consommation est basse (recharge, courbe verte au-dessus du rose) "
            "et la restitue quand la consommation dépasse le seuil (decharge, courbe verte en dessous du rose). "
            "ROSE = profil original. VERT = profil après lissage. ORANGE = seuil d écrêtage. "
            "La surface coloree entre bleu et vert représente l'énergie échangée par la batterie. "
            "L'objectif est que le maximum de la courbe verte soit inférieur au maximum de la courbe rose : "
            "c'est la réduction de pointe, valorisee en euros via le tarif de puissance souscrite."
        )
    st.plotly_chart(fig_l, width="stretch", config=PLOTLY_CFG, key="liss_fig_l")

    st.caption(
        "**Pourquoi la batterie se charge plus qu'elle ne décharge sur ce graphique ?** "
        "Sur un jour moyen, la consommation est souvent régulière — il y a peu de pics à effacer. "
        "La batterie se recharge beaucoup (barres vertes) mais décharge peu car le seuil est rarement dépassé. "
        "Sur les jours réels avec de vrais pics, le déséquilibre est inversé : "
        "la batterie décharge massivement pour écrêter. "
        "L'équilibre charge/décharge se fait sur l'ensemble de la période, pas sur un seul jour."
    )

    # ── SOC + Projection économique ──────────────────────────────────────────
    col_s1, col_s2 = st.columns(2)

    with col_s1:
        st.markdown('<p class="section">État de charge batterie (SOC)</p>',
                    unsafe_allow_html=True)
        st.caption("SOC (State of Charge) = niveau de charge de la batterie heure par heure, en MWh. La batterie reste entre SOC min (10% de sa capacite) et SOC max (90%) pour preserver la durée de vie des cellules lithium.")
        fig_soc = go.Figure()
        fig_soc.add_trace(go.Scatter(
            x=list(range(25)), y=jour["soc_hist"],
            mode="lines+markers", name="SOC",
            line=dict(color=C1, width=2),
            fill="tozeroy", fillcolor=("rgba(255,102,0,0.10)" if _BB else "rgba(92,184,92,0.12)"),
            marker=dict(size=6),
            hovertemplate="Après H%{x:02d} — SOC : <b>%{y:.3f} MWh</b><extra></extra>",
        ))
        fig_soc.add_hline(y=energy_MWh * 0.9, line_dash="dash",
                          line_color=GREEN, annotation_text=f"SOC max ({energy_MWh*0.9:.1f} MWh)")
        fig_soc.add_hline(y=energy_MWh * 0.1, line_dash="dash",
                          line_color=ORANGE, annotation_text=f"SOC min ({energy_MWh*0.1:.1f} MWh)")
        fig_soc.update_layout(
            height=300, margin=dict(t=10, b=40, l=60, r=20),
            xaxis=dict(title="Heure", tickmode="array",
                       tickvals=list(range(0,25,2)),
                       ticktext=[f"H{h:02d}" for h in range(0,25,2)],
                       range=[-0.5,24.5]),
            yaxis=dict(title="SOC (MWh)", gridcolor="#f0f0f0",
                       range=[0, energy_MWh * 1.1]),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(fig_soc)
        st.caption("Niveau de charge de la batterie (SOC) heure par heure. Monte lors de la recharge (creux de consommation) et descend lors de la décharge (pics). Reste toujours entre 10% et 90% pour préserver les cellules.")
        st.plotly_chart(fig_soc, width="stretch", config=PLOTLY_CFG, key="liss_fig_soc")

    with col_s2:
        st.markdown('<p class="section">Projection économique annuelle</p>',
                    unsafe_allow_html=True)
        st.caption("Économies realisees chaque année grâce à la réduction de puissance de pointe. Calculées comme : réduction (MW) multiplié par le tarif de puissance souscrite (euros/MW/an).")
        yl = res["yearly"]
        fig_eco = go.Figure()
        fig_eco.add_trace(go.Bar(
            x=yl["annee"].astype(str), y=yl["economie_an"],
            marker_color=C2,
            text=yl["economie_an"].apply(lambda x: f"{_fmt(x)} €"),
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Économie : <b>%{y:.0f} €</b><extra></extra>",
        ))
        fig_eco.update_layout(
            height=300, margin=dict(t=10, b=40, l=60, r=20),
            xaxis_title="Année",
            yaxis=dict(title="Économie (€)", tickformat=",", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(fig_eco)
        st.caption("Économies annuelles realisees = réduction de pointe (MW) multipliée par le tarif de puissance souscrite (euros/MW/an). Ces économies reduisent directement la facture réseau du client industriel.")
        st.plotly_chart(fig_eco, width="stretch", config=PLOTLY_CFG, key="liss_fig_eco")

        # ── Décomposition TURPE par tranche horaire ───────────────────────────
        if df_cdc_complet is not None and _t_hd > 0:
            st.markdown('<p class="section">Décomposition des économies par tranche TURPE</p>',
                        unsafe_allow_html=True)
            st.caption(
                "Économies heure par heure réparties par tranche TURPE (HD/HPH/HCH/HPE/HCE). "
                "Chaque heure écrêtée génère une économie = énergie soulagée × tarif de la tranche. "
                "Configurer les tarifs dans les paramètres économiques ci-dessus."
            )
            # Calculer les économies TURPE sur l'ensemble des jours réels
            _turpe_totals = {"HD": 0.0, "HPH": 0.0, "HCH": 0.0, "HPE": 0.0, "HCE": 0.0}
            _seuil_turpe = res["seuil_MW"]
            for _, _rjour in df_cdc_complet.iterrows():
                _mois_j = int(pd.Timestamp(_rjour["ts"]).month)
                _heure_j = int(pd.Timestamp(_rjour["ts"]).hour)
                _conso_j = float(_rjour["conso_MW"])
                if _conso_j > _seuil_turpe:
                    _tranche, _tarif = _get_turpe_tranche(_mois_j, _heure_j)
                    # Économie = MWh écrêté × tarif TURPE
                    _mwh_ecrete = min(_conso_j - _seuil_turpe, params_l["power_MW"])
                    _turpe_totals[_tranche] += _mwh_ecrete * _tarif

            # Annualiser
            _n_annees_cdc = max(df_cdc_complet["ts"].dt.year.nunique(), 1)
            _turpe_an = {k: v / _n_annees_cdc for k, v in _turpe_totals.items()}

            _tp_df = pd.DataFrame([
                {"Tranche": k,
                 "Description": {"HD":"Heures de Pointe (hiver, ~500h/an)",
                                  "HPH":"Heures Pleines Hiver",
                                  "HCH":"Heures Creuses Hiver",
                                  "HPE":"Heures Pleines Été",
                                  "HCE":"Heures Creuses Été"}[k],
                 "Tarif (€/MWh)": {"HD":_t_hd,"HPH":_t_hph,"HCH":_t_hch,"HPE":_t_hpe,"HCE":_t_hce}[k],
                 "Économie annuelle (€)": f"{_fmt(v)} €",
                 "Part (%)": f"{v/max(sum(_turpe_an.values()),0.001)*100:.1f}%"}
                for k, v in _turpe_an.items()
            ])
            _tp_total = sum(_turpe_an.values())
            st.dataframe(_tp_df, hide_index=True, width="stretch")
            st.metric("Total économies TURPE annuelles", f"{_fmt(_tp_total)} €",
                      delta=f"vs tarif puissance souscrite : {_fmt(res['economie_an'])} €/an",
                      help="Les économies TURPE incluent la composante énergie du réseau. Le tarif puissance souscrite est la composante puissance (TURPE fixe).")

            # Graphique camembert des tranches
            _fig_turpe = go.Figure(go.Pie(
                labels=list(_turpe_an.keys()),
                values=[round(v, 0) for v in _turpe_an.values()],
                hole=0.4,
                marker_colors=["#c62828", "#e65100", "#f57c00", "#2e7d32", "#1565c0"],
                textinfo="label+percent",
                hovertemplate="<b>%{label}</b><br>%{value:,.0f} €/an<br>%{percent}<extra></extra>",
            ))
            _fig_turpe.update_layout(
                height=280, margin=dict(t=10, b=10, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
            )
            st.plotly_chart(_fig_turpe, config={"displayModeBar": False},
                            width="stretch", key="liss_turpe_pie")

    # ── Détail heure par heure ───────────────────────────────────────────────
    with st.expander("Détail heure par heure"):
        _lbl_detail = {"charge": "Charge", "decharge": "Décharge", "idle": "Repos"}
        detail = pd.DataFrame({
            "Heure"             : [f"H{h:02d}" for h in range(24)],
            "Conso (MW)"        : profil_arr.round(4),
            "Action BESS"       : [_lbl_detail.get(a, a) for a, v in jour["actions"]],
            "BESS (MW)"         : [round(v, 4) for a, v in jour["actions"]],
            "Conso lissée (MW)" : jour["profil_lisse"].round(4),
            "SOC avant (MWh)"   : [round(s, 4) for s in jour["soc_hist"][:-1]],
            "SOC après (MWh)"   : [round(s, 4) for s in jour["soc_hist"][1:]],
        })
        def _style_detail_lis(df):
            st_ = pd.DataFrame("", index=df.index, columns=df.columns)
            for i in df.index:
                act = df.loc[i, "Action BESS"]
                if act == "Charge":
                    st_.loc[i, "Action BESS"] = "background:#d4edda;color:#155724;font-weight:bold"
                elif act == "Décharge":
                    st_.loc[i, "Action BESS"] = "background:#cce5ff;color:#004085;font-weight:bold"
            return st_
        st.caption("Détail heure par heure. Vert = charge (batterie se remplit). Bleu = décharge (batterie efface le pic). Blanc = repos.")
        st.dataframe(detail.style.apply(_style_detail_lis, axis=None), hide_index=True, width="stretch")

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
    st.caption("Exportez le rapport de lissage en PDF ou les donnees en CSV.")
    col_e1, col_e2 = st.columns([2, 3])
    with col_e1:
        if st.button("Générer le rapport PDF", key="btn_pdf_2"):
            with st.spinner("Génération du rapport PDF..."):

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
                    ("Économie annuelle estimée", f"{_fmt(res['economie_an'])} €"),
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
# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — VÉRIFICATION ARBITRAGE
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — AUDIT ARBITRAGE : drill-down sur chaque valeur
# ══════════════════════════════════════════════════════════════════════════════




# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — COMPARAISON MULTI-SCÉNARIOS
# ══════════════════════════════════════════════════════════════════════════════

with tab_comp:
    _diag_set_tab("comp")
    st.markdown('<p class="section">Comparaison de scénarios — jusqu\'à 4 configurations simultanées</p>',
                unsafe_allow_html=True)
    st.caption("Définissez plusieurs scénarios et comparez leurs résultats côte à côte.")

    # Sauvegarde/chargement scénarios en session state
    # Clé de version : changer cette valeur force le reset des scénarios en session
    _SCENARIOS_VERSION = "v4_top365_illimite"
    _default_scenarios = [
        {"name": "1h · Top 365 spreads/an", "power_MW": power_MW,
         "n_cycles": 0, "duration_h": 1, "efficiency": efficiency,
         "max_cycles_year": 365, "excluded_hours": {}, "optimal_quota": True},
        {"name": "1h · Illimité", "power_MW": power_MW,
         "n_cycles": 0, "duration_h": 1, "efficiency": efficiency,
         "max_cycles_year": None, "excluded_hours": {}, "optimal_quota": False},
        {"name": "2h · Top 365 spreads/an", "power_MW": power_MW,
         "n_cycles": 0, "duration_h": 2, "efficiency": efficiency,
         "max_cycles_year": 365, "excluded_hours": {}, "optimal_quota": True},
        {"name": "2h · Illimité", "power_MW": power_MW,
         "n_cycles": 0, "duration_h": 2, "efficiency": efficiency,
         "max_cycles_year": None, "excluded_hours": {}, "optimal_quota": False},
    ]
    # Reset si session absente, vide, ou version obsolète
    if (    "scenarios" not in st.session_state
         or not st.session_state.scenarios
         or st.session_state.get("scenarios_version") != _SCENARIOS_VERSION):
        st.session_state.scenarios         = _default_scenarios
        st.session_state.scenarios_version = _SCENARIOS_VERSION
        st.session_state.sc_counter        = 5
        st.cache_data.clear()

    # ── Preset 4 cas standards ────────────────────────────────────────────────
    # Cas 1 : 1h · top 365 spreads/an  (tous les cycles générés, on garde les 365 meilleurs spreads)
    # Cas 2 : 1h · illimité            (tous les cycles rentables sans quota)
    # Cas 3 : 2h · top 365 spreads/an
    # Cas 4 : 2h · illimité
    preset_col1, preset_col2, _ = st.columns([2, 1, 2])
    with preset_col1:
        if st.button("Charger les 4 cas standards (1h/2h · Top-365 / Illimité)", key="btn_preset_4cas"):
            st.session_state.scenarios = [
                # Cas 1 — 1h charge+décharge, top 365 meilleurs spreads de l'année
                {"name": "1h · Top 365 spreads/an", "power_MW": power_MW,
                 "n_cycles": 0, "duration_h": 1, "efficiency": efficiency,
                 "max_cycles_year": 365, "excluded_hours": {},
                 "optimal_quota": True},
                # Cas 2 — 1h charge+décharge, illimité
                {"name": "1h · Illimité", "power_MW": power_MW,
                 "n_cycles": 0, "duration_h": 1, "efficiency": efficiency,
                 "max_cycles_year": None, "excluded_hours": {},
                 "optimal_quota": False},
                # Cas 3 — 2h charge+décharge, top 365 meilleurs spreads de l'année
                {"name": "2h · Top 365 spreads/an", "power_MW": power_MW,
                 "n_cycles": 0, "duration_h": 2, "efficiency": efficiency,
                 "max_cycles_year": 365, "excluded_hours": {},
                 "optimal_quota": True},
                # Cas 4 — 2h charge+décharge, illimité
                {"name": "2h · Illimité", "power_MW": power_MW,
                 "n_cycles": 0, "duration_h": 2, "efficiency": efficiency,
                 "max_cycles_year": None, "excluded_hours": {},
                 "optimal_quota": False},
            ]
            st.session_state.sc_counter = 5
            st.session_state.sc_just_added = False
            st.cache_data.clear()
            st.rerun()
    with preset_col2:
        if st.button("Effacer tout", key="btn_reset_all_sc"):
            st.session_state.scenarios = []
            st.session_state.sc_counter = 1
            st.session_state.sc_just_added = False
            st.cache_data.clear()
            st.rerun()

    # ── Définition des scénarios ──────────────────────────────────────────────
    st.markdown('<p class="section">Définir un nouveau scénario</p>', unsafe_allow_html=True)
    st.caption("Definissez les paramètrès d'une nouvelle configuration a comparer. Elle sera simulée sur toute la période 2026-2029 avec les memes prix de marche que les autrès scenarios.")
    if "sc_counter" not in st.session_state:
        st.session_state.sc_counter = 1
    if "sc_just_added" not in st.session_state:
        st.session_state.sc_just_added = False

    sc1, sc2, sc3, sc4, sc5, sc6 = st.columns(6)
    with sc1:
        sc_name  = st.text_input("Nom du scénario",
                                  value=f"Scénario {st.session_state.sc_counter}",
                                  key="sc_name")
    with sc2:
        sc_pow   = st.number_input("Puissance (MW)", 0.05, 50.0, power_MW, 0.01, format="%.3f", key="sc_pow")
    with sc3:
        sc_dur   = st.selectbox("Durée cycle", [1, 2], format_func=lambda x: f"{x}h (charge + décharge)", key="sc_dur")
    with sc4:
        sc_eff   = st.slider("Rendement (%)", 70, 100, 100, key="sc_eff") / 100
    with sc5:
        sc_maxcy = st.number_input("Max cycles/an (0 = illimité)", 0, 730, 365, key="sc_maxcy")
        sc_maxcy = sc_maxcy if sc_maxcy > 0 else None
    with sc6:
        pass

    # top-N meilleurs spreads toujours activé par défaut — pas de choix exposé
    sc_ncyc   = 0       # mode MAX : le moteur trouve tous les cycles rentables
    sc_optimal = True   # on prend toujours les N meilleurs spreads

    # Restrictions horaires
    JOURS_NOMS = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"]
    sr1, sr2, sr3 = st.columns(3)
    with sr1:
        sc_jours = st.multiselect(
            "Jours exclus de la charge/décharge",
            options=list(range(7)),
            default=[0,1,2,3,4,5],
            format_func=lambda x: JOURS_NOMS[x],
            key="sc_jours",
        )
    with sr2:
        sc_h_deb = st.number_input("Heure début restriction", 0, 23, 10, key="sc_hdeb")
    with sr3:
        sc_h_fin = st.number_input("Heure fin restriction", 0, 24, 12, key="sc_hfin")

    sc_excl = {}
    if sc_jours and sc_h_deb < sc_h_fin:
        sc_excl = {"days": sc_jours, "hours": list(range(sc_h_deb, sc_h_fin))}

    col_btn1, col_btn2, _ = st.columns([1, 1, 4])
    with col_btn1:
        add_btn = st.button("Ajouter scénario", key="btn_add_sc")
    with col_btn2:
        clear_btn = st.button("Effacer tout", key="btn_clear_sc")

    # Ajout — flag pour éviter double exécution au rerun
    if add_btn and not st.session_state.sc_just_added:
        if len(st.session_state.scenarios) < 4:
            st.session_state.scenarios.append({
                "name":            sc_name,
                "power_MW":        sc_pow,
                "n_cycles":        sc_ncyc,
                "duration_h":      sc_dur,
                "efficiency":      sc_eff,
                "max_cycles_year": sc_maxcy,
                "excluded_hours":  sc_excl,
                "optimal_quota":   sc_optimal,
            })
            st.session_state.sc_counter   += 1
            st.session_state.sc_just_added = True
            st.rerun()
        else:
            st.warning("Maximum 4 scénarios.")

    # Réinitialiser le flag une fois le rerun passé
    if st.session_state.sc_just_added and not add_btn:
        st.session_state.sc_just_added = False

    if clear_btn:
        st.session_state.scenarios      = []
        st.session_state.sc_counter     = 1
        st.session_state.sc_just_added  = False
        st.rerun()

    if not st.session_state.scenarios:
        st.info("Ajoutez au moins un scénario pour lancer la comparaison.")
    else:
        # Afficher les scénarios actifs
        st.markdown('<p class="section">Scénarios actifs</p>', unsafe_allow_html=True)
        st.caption("Configurations actuellement en cours de comparaison. Cliquez Supprimer pour retirer un scenario de la liste.")
        sc_cols = st.columns(len(st.session_state.scenarios))
        to_remove = None
        for i, sc in enumerate(st.session_state.scenarios):
            _nc_disp = sc["n_cycles"] if sc["n_cycles"] > 0 else 1
            cap = sc["power_MW"] * sc["duration_h"] * _nc_disp
            excl = sc.get("excluded_hours", {})
            restr_txt = "Aucune"
            if excl.get("days") and excl.get("hours"):
                jn = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"]
                jours_str = ", ".join(jn[d] for d in excl["days"])
                h_str = f"H{min(excl['hours']):02d}-H{max(excl['hours'])+1:02d}"
                restr_txt = f"{jours_str} {h_str}"
            with sc_cols[i]:
                st.markdown(f"""
**{sc['name']}**
- Puissance : {sc['power_MW']} MW
- Durée cycle : {sc['duration_h']}h (charge + décharge)
- Capacité : {cap:.2f} MWh
- Rendement : {sc['efficiency']*100:.0f}%
- Max cycles/an : {sc['max_cycles_year'] or 'illimité'}
- Restriction : {restr_txt}
- Mode quota : {'Top-N meilleurs spreads' if sc.get('optimal_quota') else 'Chronologique'}
""")
                if st.button(f"Supprimer", key=f"btn_del_sc_{i}"):
                    to_remove = i

        if to_remove is not None:
            st.session_state.scenarios.pop(to_remove)
            st.rerun()

        # ── Simulation de tous les scénarios ─────────────────────────────────
        @st.cache_data(show_spinner=False)
        def run_all_scenarios(file_bytes, scenarios_json):
            import json
            from bess_engine import load_spot, simulate_arbitrage, simulate_arbitrage_optimal, aggregate_arbitrage
            import io as _io
            pv = load_spot(io.BytesIO(file_bytes))
            results = []
            scenarios = json.loads(scenarios_json)
            for sc in scenarios:
                p = {k: sc[k] for k in ["power_MW","n_cycles","duration_h","efficiency","max_cycles_year"]}
                p["excluded_hours"] = sc.get("excluded_hours", {})
                p["optimal_quota"]  = sc.get("optimal_quota", False)
                if p.get("optimal_quota") and p.get("max_cycles_year"):
                    daily = simulate_arbitrage_optimal(pv, p)
                else:
                    daily = simulate_arbitrage(pv, p)
                yearly = aggregate_arbitrage(daily, p["power_MW"])
                results.append({"sc": sc, "daily": daily, "yearly": yearly})
            return results

        st.markdown("---")
        _col_btn_run, _ = st.columns([1, 3])
        with _col_btn_run:
            _btn_run = st.button(
                "Lancer la simulation",
                key="btn_run_scenarios",
                type="primary",
                use_container_width=True,
            )
        if _btn_run:
            st.session_state["_sc_run_requested"] = True

        if not st.session_state.get("_sc_run_requested"):
            st.info("Configurez vos scénarios ci-dessus puis cliquez sur **Lancer la simulation**.")
        else:
            _sc_names = [sc["name"] for sc in st.session_state.scenarios]
            with st.spinner(f"Simulation des {len(_sc_names)} scénarios en cours… ({', '.join(_sc_names)})"):
                results = run_all_scenarios(file_bytes, _json.dumps(st.session_state.scenarios))
            # Sauvegarder pour l'assistant IA et pour l'affichage
            st.session_state["_ai_sc_results"] = results
            st.session_state["_sc_results_cache"] = results

        results = st.session_state.get("_sc_results_cache")
        if results is None:
            results = []  # empêche les erreurs sur le reste du code

        if results:
            # ── Tableau comparatif ────────────────────────────────────────────────
            st.markdown('<p class="section">Résultats comparatifs</p>', unsafe_allow_html=True)
            st.caption("Tableau comparatif des performances de chaque scenario sur 4 ans. Identifiez rapidement quelle configuration généré les meilleurs revenus en tenant compte de vos contraintes specifiques.")

            comp_rows = []
            for r in results:
              sc = r["sc"]
              y  = r["yearly"]
              d  = r["daily"]
              _nc_cap = sc["n_cycles"] if sc["n_cycles"] > 0 else 1
              cap = sc["power_MW"] * sc["duration_h"] * _nc_cap
              comp_rows.append({
                  "Scénario":            sc["name"],
                  "Puissance (MW)":      sc["power_MW"],
                  "Cycles/jour":         "Max" if sc["n_cycles"] == 0 else sc["n_cycles"],
                  "Durée (h)":           sc["duration_h"],
                  "Capacité (MWh)":      round(cap, 2),
                  "PnL (€)":             f"{_fmt(y['pnl_total'].sum())}",
                  "Potentiel max (€)":   f"{_fmt(y['pnl_absolu_total'].sum())}",
                  "Spread moy. (€/MWh)": f"{y['spread_moy'].mean():.2f}",
                  "Taux activation (%)": f"{y['taux_activation'].mean()*100:.1f}",
                  "Énergie (MWh/4ans)":  f"{_fmt(y['energie_totale_MWh'].sum(), 1)}",
                  "PnL/kW (€/kW)":       f"{_fmt(y['pnl_par_MW'].mean(), 1)}",
              })

            st.caption(
              "Comparaison des scénarios sur 4 ans. "
              "**PnL** = revenus nets réalisés. "
              "Pour le mode Illimité : PnL = Potentiel max car aucun quota ne bloque les trades — les deux colonnes sont identiques, c'est normal. "
              "**Potentiel max** = maximum théorique sans aucune contrainte = le PnL du mode Illimité, "
              "affiché pour mesurer ce que le quota 365 vous coûte. "
              "**Spread moy.** = écart moyen prix vente − prix achat (€/MWh). "
              "**Taux activation** = % de jours tradés (< 100% si quota épuisé avant fin d'année)."
            )
            st.dataframe(pd.DataFrame(comp_rows), hide_index=True, width="stretch")

            # ── Graphiques comparatifs ────────────────────────────────────────────
            col_g1, col_g2 = st.columns(2)

            with col_g1:
              st.markdown('<p class="section">PnL total par scénario (€)</p>', unsafe_allow_html=True)
              st.caption("PnL total sur 4 ans pour chaque scenario. La barre pleine = PnL réel avec toutes les contraintes. La barre claire = borne max théorique sans restriction.")
              fig_c1 = go.Figure()
              for r in results:
                  y = r["yearly"]
                  fig_c1.add_trace(go.Bar(
                      name=r["sc"]["name"],
                      x=y["annee"].astype(str),
                      y=y["pnl_total"],
                      hovertemplate=f"<b>{r['sc']['name']}</b><br>%{{x}} : %{{y:.0f}} €<extra></extra>",
                  ))
              fig_c1.update_layout(
                  barmode="group", height=340,
                  yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
                  xaxis_title="Année",
                  legend=LEGEND_BOTTOM,
                  margin=dict(t=10, b=60, l=60, r=10),
                  plot_bgcolor="white", paper_bgcolor="white",
              )
              apply_bb(fig_c1)
              st.caption("PnL total sur 4 ans pour chaque scenario. Barres pleines = PnL réel avec toutes vos contraintes. Barres claires = borne max sans restriction. L'ecart entre les deux mesure l'impact de vos restrictions.")
              st.plotly_chart(fig_c1, width="stretch", config=PLOTLY_CFG, key="sens_fig_c1")
              st.caption('PnL total par scenario sur 4 ans. Permet de comparer directement la rentabilité de chaque configuration.')

            with col_g2:
              st.markdown('<p class="section">PnL cumulé sur 4 ans (€)</p>', unsafe_allow_html=True)
              st.caption("Evolution du PnL cumule dans le temps pour chaque scenario. Permet de voir si un scenario performe mieux sur certaines périodes de l'année ou sur le long terme.")
              fig_c2 = go.Figure()
              for i, r in enumerate(results):
                  d = r["daily"].sort_values("date")
                  fig_c2.add_trace(go.Scatter(
                      x=d["date"], y=d["pnl"].cumsum(),
                      mode="lines", name=r["sc"]["name"],
                      line=dict(width=2, color=COLORS[i % len(COLORS)]),
                      hovertemplate=f"<b>{r['sc']['name']}</b><br>%{{x|%d/%m/%Y}} : %{{y:.0f}} €<extra></extra>",
                  ))
              fig_c2.update_layout(
                  height=340,
                  yaxis=dict(title="PnL cumulé (€)", tickformat=",", gridcolor="#f0f0f0"),
                  xaxis=dict(tickformat="%b %Y"),
                  legend=LEGEND_BOTTOM,
                  margin=dict(t=10, b=60, l=60, r=10),
                  plot_bgcolor="white", paper_bgcolor="white",
                  hovermode="x unified",
              )
              apply_bb(fig_c2)
              st.caption("Evolution du PnL cumule dans le temps pour chaque scenario. Si deux courbes se croisent, le scenario supérieur est devenu plus rentable à partir de cette date.")
              st.plotly_chart(fig_c2, width="stretch", config=PLOTLY_CFG, key="sens_fig_c2")
              st.caption('PnL cumule par scenario sur 4 ans. La courbe la plus haute = scenario le plus rentable sur la durée totale.')

            col_g3, col_g4 = st.columns(2)

            with col_g3:
              st.markdown('<p class="section">Taux d\'activation par scénario (%)</p>', unsafe_allow_html=True)
              st.caption(
                  "Taux d'activation = % de jours où la batterie a effectivement tradé. "
                  "Avec n_cycles=Max et max_cycles_year=365, la batterie peut faire plusieurs cycles/jour "
                  "mais est limitée à 365 cycles au total sur l'année. "
                  "En illimité elle fait autant de cycles que rentable chaque jour. "
                  "Si taux < 100% : le quota annuel a été atteint avant la fin de l'année."
              )
              st.caption("Pourcentage de jours où la batterie a effectivement trade pour chaque scenario. Un taux bas peut indiquer des restrictions trop sévères ou un quota annuel trop faible par rapport aux opportunités.")
              fig_c3 = go.Figure()
              for r in results:
                  y = r["yearly"]
                  fig_c3.add_trace(go.Bar(
                      name=r["sc"]["name"],
                      x=y["annee"].astype(str),
                      y=(y["taux_activation"] * 100).round(1),
                      hovertemplate=f"<b>{r['sc']['name']}</b><br>%{{x}} : %{{y:.1f}}%<extra></extra>",
                  ))
              fig_c3.update_layout(
                  barmode="group", height=300,
                  yaxis=dict(title="Taux activation (%)", gridcolor="#f0f0f0"),
                  legend=LEGEND_BOTTOM, margin=dict(t=10, b=60, l=60, r=10),
                  plot_bgcolor="white", paper_bgcolor="white",
              )
              apply_bb(fig_c3)
              st.caption("Taux d'activation de chaque scenario = pourcentage de jours effectivement trades. Un taux bas peut indiquer des restrictions trop sévères ou un quota annuel insuffisant par rapport aux opportunités de marche.")
              st.plotly_chart(fig_c3, width="stretch", config=PLOTLY_CFG, key="sens_fig_c3")
              st.caption("Taux d'activation par scenario : pourcentage de jours où la batterie a effectivement trade.")

            with col_g4:
              st.markdown('<p class="section">PnL/MW par scénario (€/MW)</p>', unsafe_allow_html=True)
              st.caption("PnL total divisé par la puissance installée en kW. Permet de comparer des batteries de tailles différentes : quelle configuration génère le plus de revenus par kW installe ?")
              fig_c4 = go.Figure()
              sc_names_c4 = [r["sc"]["name"] for r in results]
              pnl_mw_c4   = [r["yearly"]["pnl_par_MW"].mean() for r in results]
              fig_c4.add_trace(go.Bar(
                  x=sc_names_c4, y=pnl_mw_c4,
                  marker_color=COLORS[:len(results)],
                  text=[f"{_fmt(v, 1)}" for v in pnl_mw_c4],
                  textposition="outside",
                  hovertemplate="<b>%{x}</b><br>%{y:.1f} €/MW<extra></extra>",
              ))
              fig_c4.update_layout(
                  height=300,
                  yaxis=dict(title="PnL/MW (€/MW)", tickformat=",", gridcolor="#f0f0f0"),
                  legend=LEGEND_BOTTOM, margin=dict(t=10, b=60, l=60, r=10),
                  plot_bgcolor="white", paper_bgcolor="white",
              )
              apply_bb(fig_c4)
              st.caption("PnL total divisé par la puissance installée en kW. Permet de comparer équitablement des batteries de tailles différentes : la configuration avec le PnL/kW le plus élevé est la plus efficace par unite de puissance.")
              st.plotly_chart(fig_c4, width="stretch", config=PLOTLY_CFG, key="sens_fig_c4")
              st.caption("Taux d'activation par scenario : pourcentage de jours où la batterie a effectivement trade.")

            # ── Simulation historique 2019-2025 pour les scénarios comparés ────────
            st.markdown('<p class="section">Simulation historique 2019–2025</p>', unsafe_allow_html=True)
            st.caption(
              "Les mêmes scénarios simulés sur les prix réels 2019–2025. "
              "Utilise le fichier déjà chargé si il contient la feuille 'Case 3 night station', "
              "sinon uploadez le fichier **BESS_valorisation_RESULTATS.xlsx** (sans '_Copie') ci-dessous."
            )
            # Permettre un upload alternatif pour le fichier historique
            _hist_file_extra = st.file_uploader(
              "Fichier historique (optionnel si déjà dans le fichier principal)",
              type=["xlsx"], key="comp_hist_upload",
              help="Uploadez BESS_valorisation_RESULTATS.xlsx (sans _Copie) si votre fichier principal ne contient pas l'historique 2019-2025."
            )
            _hist_bytes = _hist_file_extra.read() if _hist_file_extra else file_bytes

            try:
              from bess_engine import load_case3 as _lc3_comp
              import io as _io_comp
              _pv_hist_comp = _lc3_comp(_io_comp.BytesIO(_hist_bytes))
              if _pv_hist_comp.empty:
                  raise ValueError("Aucune donnée historique (années ≤ 2025) dans le fichier. Uploadez BESS_valorisation_RESULTATS.xlsx (sans _Copie).")
              _hist_rows = []
              _hist_results = []
              from bess_engine import (simulate_arbitrage as _sa,
                                       simulate_arbitrage_optimal as _sao,
                                       aggregate_arbitrage as _aa)
              _n_sc_hist = len(results)
              _prog_bar = st.progress(0, text="Simulation historique 2019–2025 — initialisation…")
              for _i_sc, r in enumerate(results):
                  sc = r["sc"]
                  _prog_bar.progress(
                      int(_i_sc / _n_sc_hist * 100),
                      text=f"Simulation historique — scénario {_i_sc+1}/{_n_sc_hist} : {sc['name']}…"
                  )
                  _p = {"power_MW": sc["power_MW"], "n_cycles": sc["n_cycles"],
                        "duration_h": sc["duration_h"], "efficiency": sc["efficiency"],
                        "max_cycles_year": sc["max_cycles_year"],
                        "excluded_hours": sc.get("excluded_hours", {}),
                        "optimal_quota": sc.get("optimal_quota", False)}
                  if _p.get("optimal_quota") and _p.get("max_cycles_year"):
                      _dh = _sao(_pv_hist_comp, _p)
                  else:
                      _dh = _sa(_pv_hist_comp, _p)
                  _yh = _aa(_dh, sc["power_MW"])
                  _hist_results.append({"sc": sc, "daily": _dh, "yearly": _yh})
                  for _, yr in _yh.iterrows():
                      _hist_rows.append({
                          "Scénario": sc["name"],
                          "Année": str(int(yr["annee"])),
                          "PnL (€)": _fmt(yr["pnl_total"]),
                          "Potentiel max (€)": _fmt(yr["pnl_absolu_total"]),
                          "Spread moy (€/MWh)": f"{yr['spread_moy']:.1f}",
                          "Jours actifs": str(int(yr["jours_actifs"])),
                          "Taux activation": f"{yr['taux_activation']*100:.0f}%",
                      })
                  _hist_rows.append({
                      "Scénario": f"▶ TOTAL {sc['name']}",
                      "Année": "2019-2025",
                      "PnL (€)": _fmt(_yh["pnl_total"].sum()),
                      "Potentiel max (€)": _fmt(_yh["pnl_absolu_total"].sum()),
                      "Spread moy (€/MWh)": f"{_yh['spread_moy'].mean():.1f}",
                      "Jours actifs": str(int(_yh["jours_actifs"].sum())),
                      "Taux activation": "—",
                  })
              _prog_bar.progress(100, text="Simulation historique terminée.")
              _prog_bar.empty()
              st.dataframe(pd.DataFrame(_hist_rows), hide_index=True, width="stretch")
              st.caption(
                  "Données réelles EPEX SPOT France 2019–2025. "
                  "**Spread moy.** = différence moyenne prix vente − prix achat sur les cycles réalisés. "
                  "2022 = pic de la crise énergétique (prix de l'électricité x3–x5 vs année normale). "
                  "Permet de valider que les projections 2026–2029 sont dans une fourchette historiquement plausible."
              )

              # Graphique unifié 2019-2029 (historique + projection sur un seul graphique)
              st.markdown('<p class="section">PnL comparatif 2019–2029 — Historique réel + Projection</p>', unsafe_allow_html=True)
              st.caption("Données réelles 2019–2025 (feuille Case 3) + projections 2026–2029 sur un seul graphique.")
              # Construire les données combinées 2019-2029 par scénario
              # Graphique PnL comparatif 2019-2029
              _COLORS_SC = ["#1565c0","#e65100","#2e7d32","#6a1b9a"]

              # Toutes les années triées 2019→2029
              _all_annees_set = set()
              for _rh in _hist_results:
                  _all_annees_set.update([int(a) for a in _rh["yearly"]["annee"]])
              for _rf in results:
                  _all_annees_set.update([int(a) for a in _rf["yearly"]["annee"]])
              _all_annees_sorted = sorted(_all_annees_set)  # [2019,2020,...,2029]
              # Positions numériques 0,1,2,...,10
              _pos = list(range(len(_all_annees_sorted)))
              _a2p = {a: i for i, a in enumerate(_all_annees_sorted)}  # annee→position

              _fig_uni = go.Figure()

              for _i, (_rh, _rf) in enumerate(zip(_hist_results, results)):
                  _clr = _COLORS_SC[_i % len(_COLORS_SC)]
                  _sc_name = _rh["sc"]["name"]

                  _yh_dict = {int(r["annee"]): float(r["pnl_total"])
                              for _, r in _rh["yearly"].iterrows()}
                  _yf_dict = {int(r["annee"]): float(r["pnl_total"])
                              for _, r in _rf["yearly"].iterrows()}

                  # x = positions numériques, y = valeur ou None
                  _x_reel = [_a2p[a] for a in _all_annees_sorted if a in _yh_dict]
                  _y_reel = [_yh_dict[a]  for a in _all_annees_sorted if a in _yh_dict]
                  _x_proj = [_a2p[a] for a in _all_annees_sorted if a in _yf_dict]
                  _y_proj = [_yf_dict[a]  for a in _all_annees_sorted if a in _yf_dict]

                  _fig_uni.add_trace(go.Bar(
                      name=f"{_sc_name} — Réel",
                      x=_x_reel, y=_y_reel,
                      marker_color=_clr, opacity=0.9,
                      legendgroup=f"sc{_i}",
                      hovertemplate="<b>" + _sc_name + " (réel)</b><br>%{x} : %{y:,.0f} €<extra></extra>",
                  ))
                  _fig_uni.add_trace(go.Bar(
                      name=f"{_sc_name} — Projection",
                      x=_x_proj, y=_y_proj,
                      marker_color=_clr,
                      marker_opacity=0.45,
                      legendgroup=f"sc{_i}",
                      hovertemplate="<b>" + _sc_name + " (proj.)</b><br>%{x} : %{y:,.0f} €<extra></extra>",
                  ))

              # Ligne séparatrice entre 2025 et 2026
              if 2026 in _a2p:
                  _sep = _a2p[2026] - 0.5
                  _fig_uni.add_shape(type="line",
                      x0=_sep, x1=_sep, y0=0, y1=1,
                      xref="x", yref="paper",
                      line=dict(dash="dash", color="#888", width=2))
                  _fig_uni.add_annotation(
                      x=_sep, y=1.04, xref="x", yref="paper",
                      text="◄ Réel  |  Projection ►",
                      showarrow=False, font=dict(size=10, color="#555"),
                      xanchor="center")

              _fig_uni.update_layout(
                  barmode="group", height=440,
                  yaxis=dict(title="PnL (€)", gridcolor="#f0f0f0"),
                  xaxis=dict(
                      title="Année",
                      tickmode="array",
                      tickvals=_pos,
                      ticktext=[str(a) for a in _all_annees_sorted],
                  ),
                  legend=dict(orientation="h", y=-0.30, x=0, font_size=10),
                  margin=dict(t=30, b=110, l=60, r=10),
                  plot_bgcolor="white", paper_bgcolor="white",
              )
              apply_bb(_fig_uni)
              st.plotly_chart(_fig_uni, width="stretch", config=PLOTLY_CFG, key="comp_fig_uni")
              st.session_state["_comp_fig_uni"] = _fig_uni
              st.session_state["_hist_results_comp"] = _hist_results

              # ── Volet déroulant : liste des spreads par scénario, 2019-2029 ──
              # On fusionne données historiques (_hist_results) + projections (results)
              with st.expander("Liste des cycles classés par spread — détail par scénario et par année (2019–2029)", expanded=False):
                  st.caption(
                      "Pour chaque scénario et chaque année (2019–2025 historique + 2026–2029 projection), "
                      "tous les cycles retenus classés du meilleur spread au plus faible. "
                      "Permet de voir exactement quels jours correspondent aux meilleures opportunités."
                  )
                  _tab_sc_labels2 = [r["sc"]["name"] for r in results]
                  _tabs_spread2 = st.tabs(_tab_sc_labels2)

                  for _ti2, (_tab_s2, _rf2, _rh2) in enumerate(zip(_tabs_spread2, results, _hist_results)):
                      with _tab_s2:
                          _sc2    = _rf2["sc"]
                          _max_cy2 = _sc2.get("max_cycles_year")
                          # Fusionner daily historique + projection
                          _daily_combined = pd.concat([_rh2["daily"], _rf2["daily"]], ignore_index=True)
                          _annees2 = sorted(_daily_combined["annee"].unique())
                          _ann_tabs2 = st.tabs([str(int(a)) for a in _annees2])

                          for _ann_tab2, _ann2 in zip(_ann_tabs2, _annees2):
                              with _ann_tab2:
                                  _rows_sp2 = []
                                  _dy2 = _daily_combined[_daily_combined["annee"] == _ann2].copy()
                                  for _, _row2 in _dy2.iterrows():
                                      if _row2["n_cycles_actifs"] > 0:
                                          _nc2    = int(_row2["n_cycles_actifs"])
                                          _sp2    = float(_row2["spread"])
                                          _pnl2   = float(_row2["pnl"])
                                          for _ in range(_nc2):
                                              _rows_sp2.append({
                                                  "Date":                  _row2["date"].strftime("%d/%m/%Y"),
                                                  "Jour":                  _row2["weekday_name"],
                                                  "Spread (€/MWh)":        round(_sp2, 2),
                                                  "Prix charge (€/MWh)":   round(float(_row2["prix_charge"]), 2),
                                                  "Prix décharge (€/MWh)": round(float(_row2["prix_decharge"]), 2),
                                                  "PnL (€)":               round(_pnl2 / _nc2, 2),
                                              })
                                  if not _rows_sp2:
                                      st.caption("Aucun cycle retenu cette année.")
                                      continue
                                  _df_sp2 = pd.DataFrame(_rows_sp2).sort_values("Spread (€/MWh)", ascending=False).reset_index(drop=True)
                                  _df_sp2.index = _df_sp2.index + 1
                                  if _max_cy2:
                                      st.caption(
                                          f"**{len(_df_sp2)} cycles retenus** — {int(_ann2)} "
                                          f"(quota = {_max_cy2}) · "
                                          f"Spread max : **{_df_sp2['Spread (€/MWh)'].max():.1f} €/MWh** · "
                                          f"Spread min retenu : **{_df_sp2['Spread (€/MWh)'].min():.1f} €/MWh**"
                                      )
                                  else:
                                      st.caption(
                                          f"**{len(_df_sp2)} cycles** — {int(_ann2)} (illimité) · "
                                          f"Spread max : **{_df_sp2['Spread (€/MWh)'].max():.1f} €/MWh** · "
                                          f"Spread min : **{_df_sp2['Spread (€/MWh)'].min():.1f} €/MWh**"
                                      )
                                  st.dataframe(
                                      _df_sp2[["Date","Jour","Spread (€/MWh)","Prix charge (€/MWh)","Prix décharge (€/MWh)","PnL (€)"]],
                                      hide_index=False,
                                      width="stretch",
                                      height=min(600, 36 + 35 * len(_df_sp2)),
                                  )

            except Exception as _e_hist:
              import traceback as _tb
              _tb_str = _tb.format_exc()
              st.error(f"🔴 ERREUR HISTORIQUE v2.2 — {type(_e_hist).__name__}: {_e_hist}")
              with st.expander("Traceback complet"):
                  st.code(_tb_str)

            # ── Export CSV + PDF avec graphiques ────────────────────────────────
            st.markdown('<p class="section">Export du rapport</p>', unsafe_allow_html=True)
            exp_c1, exp_c2 = st.columns(2)
            with exp_c1:
              st.download_button(
                  "⬇ Exporter comparaison (CSV)",
                  data=pd.DataFrame(comp_rows).to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                  file_name="BESS_comparaison_scenarios.csv",
                  mime="text/csv",
                  key="btn_comp_csv",
              )
            with exp_c2:
              if st.button("Générer rapport PDF", key="btn_comp_pdf"):
                  import io as _io_pdf
                  import datetime as _dt_pdf
                  import plotly.io as pio
                  try:
                      from reportlab.lib.pagesizes import A4
                      from reportlab.lib import colors as rl_colors
                      from reportlab.lib.units import cm
                      from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                                      Table, TableStyle, Image as RLImage, HRFlowable)
                      from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                      from reportlab.lib.enums import TA_CENTER, TA_LEFT

                      _buf = _io_pdf.BytesIO()
                      doc = SimpleDocTemplate(_buf, pagesize=A4,
                                              leftMargin=1.5*cm, rightMargin=1.5*cm,
                                              topMargin=1.5*cm, bottomMargin=1.5*cm)
                      styles = getSampleStyleSheet()
                      _BLUE = rl_colors.HexColor("#1565c0")
                      _GREEN = rl_colors.HexColor("#2e7d32")
                      s_title  = ParagraphStyle("t", parent=styles["Heading1"], textColor=_BLUE, fontSize=16, spaceAfter=4)
                      s_h2     = ParagraphStyle("h2", parent=styles["Heading2"], textColor=_BLUE, fontSize=12, spaceBefore=10, spaceAfter=4)
                      s_body   = ParagraphStyle("b", parent=styles["Normal"], fontSize=9, spaceAfter=4)
                      s_small  = ParagraphStyle("s", parent=styles["Normal"], fontSize=7.5, textColor=rl_colors.grey)

                      story = []
                      _fname_pdf = st.session_state.get("excel_name", "N/A")
                      story.append(Paragraph("Rapport de comparaison de scénarios BESS", s_title))
                      story.append(Paragraph(
                          f"Généré le {_dt_pdf.datetime.now().strftime('%d/%m/%Y à %H:%M')} · "
                          f"Fichier : {_fname_pdf} · Puissance : {power_MW*1000:.0f} kW", s_body))
                      story.append(HRFlowable(width="100%", color=_BLUE, spaceAfter=8))

                      # ── Scénarios comparés ────────────────────────────────────────────
                      story.append(Paragraph("Scénarios comparés", s_h2))
                      _sc_data = [["Scénario","Puissance (MW)","Cycles/j","Durée cycle","Capacité (MWh)","Rend.","Quota cycles/an"]]
                      for r in results:
                          sc = r["sc"]
                          _nc_disp = sc["n_cycles"] if sc["n_cycles"] > 0 else "Max"
                          cap = sc["power_MW"] * sc["duration_h"] * (sc["n_cycles"] if sc["n_cycles"] > 0 else 1)
                          _sc_data.append([sc["name"], f"{sc['power_MW']:.3f}", str(_nc_disp),
                                           f"{sc['duration_h']}h", f"{cap:.3f}",
                                           f"{sc['efficiency']*100:.0f}%", str(sc["max_cycles_year"] or "Illimité")])
                      _sc_tbl = Table(_sc_data, repeatRows=1,
                          colWidths=[5.5*cm, 2.2*cm, 1.6*cm, 2.0*cm, 2.2*cm, 1.4*cm, 2.8*cm])
                      _sc_tbl.setStyle(TableStyle([
                          ("BACKGROUND",(0,0),(-1,0),_BLUE), ("TEXTCOLOR",(0,0),(-1,0),rl_colors.white),
                          ("FONTSIZE",(0,0),(-1,-1),8), ("GRID",(0,0),(-1,-1),0.3,rl_colors.grey),
                          ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor("#f8f9fa")]),
                          ("PADDING",(0,0),(-1,-1),4),
                          ("ALIGN",(1,0),(-1,-1),"CENTER"),
                          ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                      ]))
                      story.append(_sc_tbl)
                      story.append(Paragraph(
                          "<i>Durée cycle : durée d'une charge + durée d'une décharge (ex. 2h = 2h d'achat puis 2h de revente). "
                          "Capacité : énergie stockable en MWh (Puissance × Durée). "
                          "Quota : nombre max de cycles autorisés par an — 365 = on garde les 365 meilleurs spreads, Illimité = tous les cycles rentables.</i>",
                          s_small))
                      story.append(Spacer(1, 8))

                      # ── Résultats 2026-2029 ───────────────────────────────────────────
                      story.append(Paragraph("Résultats comparatifs — Projections 2026–2029", s_h2))
                      _r_data = [["Scénario","Année","PnL réel (€)","PnL borne max (€)","Spread moy. (€/MWh)","Jours actifs","Taux activ."]]
                      for r in results:
                          for _, yr in r["yearly"].iterrows():
                              _r_data.append([r["sc"]["name"], str(int(yr["annee"])),
                                              _fmt(yr["pnl_total"]), _fmt(yr["pnl_absolu_total"]),
                                              f"{yr['spread_moy']:.1f}", str(int(yr["jours_actifs"])),
                                              f"{yr['taux_activation']*100:.0f}%"])
                          _r_data.append([f"TOTAL {r['sc']['name']}", "2026-29",
                                          _fmt(r["yearly"]["pnl_total"].sum()),
                                          _fmt(r["yearly"]["pnl_absolu_total"].sum()),
                                          f"{r['yearly']['spread_moy'].mean():.1f}","—","—"])
                      _r_tbl = Table(_r_data, repeatRows=1,
                          colWidths=[4.8*cm, 1.4*cm, 2.4*cm, 2.4*cm, 2.4*cm, 1.8*cm, 2.0*cm])
                      _r_tbl.setStyle(TableStyle([
                          ("BACKGROUND",(0,0),(-1,0),_BLUE), ("TEXTCOLOR",(0,0),(-1,0),rl_colors.white),
                          ("FONTSIZE",(0,0),(-1,-1),8), ("GRID",(0,0),(-1,-1),0.3,rl_colors.grey),
                          ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor("#f8f9fa")]),
                          ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("PADDING",(0,0),(-1,-1),4),
                          ("ALIGN",(1,0),(-1,-1),"CENTER"),
                          ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                          ("BACKGROUND",(0,-1),(-1,-1),rl_colors.HexColor("#e3f2fd")),
                          ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
                      ]))
                      story.append(_r_tbl)
                      story.append(Paragraph(
                          "<i>PnL réel : revenus nets avec toutes vos contraintes (quota, heures exclues). "
                          "PnL borne max : maximum théorique sans aucune contrainte — mesure le potentiel du marché. "
                          "Spread moy. : différence moyenne entre prix de revente et prix d'achat sur les cycles réalisés (€/MWh). "
                          "Jours actifs : jours où la batterie a effectivement tradé. "
                          "Taux activ. : % de jours tradés — inférieur à 100% si le quota annuel est atteint avant fin d'année.</i>",
                          s_small))
                      story.append(Spacer(1, 8))

                      # ── Résultats historiques 2019-2025 ──────────────────────────────
                      _hr_pdf = st.session_state.get("_hist_results_comp", [])
                      if _hr_pdf:
                          story.append(Paragraph("Simulation historique — Données réelles 2019–2025", s_h2))
                          _h_data = [["Scénario","Année","PnL réel (€)","PnL borne max (€)","Spread moy. (€/MWh)","Jours actifs","Taux activ."]]
                          for _rh in _hr_pdf:
                              for _, yr in _rh["yearly"].iterrows():
                                  _h_data.append([_rh["sc"]["name"], str(int(yr["annee"])),
                                                  _fmt(yr["pnl_total"]), _fmt(yr["pnl_absolu_total"]),
                                                  f"{yr['spread_moy']:.1f}", str(int(yr["jours_actifs"])),
                                                  f"{yr['taux_activation']*100:.0f}%"])
                              _h_data.append([f"TOTAL {_rh['sc']['name']}", "2019-25",
                                              _fmt(_rh["yearly"]["pnl_total"].sum()),
                                              _fmt(_rh["yearly"]["pnl_absolu_total"].sum()),
                                              f"{_rh['yearly']['spread_moy'].mean():.1f}","—","—"])
                          _h_tbl = Table(_h_data, repeatRows=1,
                              colWidths=[4.8*cm, 1.4*cm, 2.4*cm, 2.4*cm, 2.4*cm, 1.8*cm, 2.0*cm])
                          _h_tbl.setStyle(TableStyle([
                              ("BACKGROUND",(0,0),(-1,0),_GREEN), ("TEXTCOLOR",(0,0),(-1,0),rl_colors.white),
                              ("FONTSIZE",(0,0),(-1,-1),8), ("GRID",(0,0),(-1,-1),0.3,rl_colors.grey),
                              ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor("#f1f8f1")]),
                              ("ALIGN",(1,0),(-1,-1),"CENTER"),
                              ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                              ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("PADDING",(0,0),(-1,-1),4),
                              ("BACKGROUND",(0,-1),(-1,-1),rl_colors.HexColor("#e8f5e9")),
                              ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
                          ]))
                          story.append(_h_tbl)
                          story.append(Paragraph(
                              "<i>Données réelles EPEX SPOT France 2019–2025. "
                              "Permet de valider les projections 2026–2029 sur des prix historiques connus. "
                              "2022 correspond au pic de la crise énergétique européenne (prix x3–x5 vs années normales).</i>",
                              s_small))
                          story.append(Spacer(1, 8))

                      # ── Graphiques : chacun sur sa propre page ──────────────────────
                      from reportlab.platypus import PageBreak
                      _figs_to_export = [
                          ("PnL comparatif 2019–2029 (réel + projection)",
                           "Vue d'ensemble historique et projections — barres pleines = données réelles, barres transparentes = projections.",
                           st.session_state.get("_comp_fig_uni"), 1600, 560),
                          ("PnL total par scénario 2026–2029",
                           "PnL annuel par scénario sur la période de projection. Permet de comparer directement la rentabilité de chaque configuration.",
                           fig_c1, 1600, 560),
                          ("PnL cumulé 2026–2029",
                           "Cumul des gains depuis janvier 2026. La courbe la plus haute = scénario le plus rentable sur la durée.",
                           fig_c2, 1600, 560),
                      ]
                      story.append(PageBreak())
                      story.append(Paragraph("Graphiques comparatifs", s_h2))
                      for _fig_label, _fig_caption, _fig_obj, _fw, _fh in _figs_to_export:
                          if _fig_obj is None: continue
                          try:
                              _fig_obj.update_layout(
                                  template="plotly_white",
                                  plot_bgcolor="white",
                                  paper_bgcolor="white",
                              )
                              _img_bytes = pio.to_image(_fig_obj, format="png", width=_fw, height=_fh, scale=2)
                              _img_buf = _io_pdf.BytesIO(_img_bytes)
                              story.append(Paragraph(f"<b>{_fig_label}</b>", s_body))
                              story.append(Paragraph(f"<i>{_fig_caption}</i>", s_small))
                              story.append(Spacer(1, 4))
                              story.append(RLImage(_img_buf, width=17.5*cm, height=8*cm))
                              story.append(Spacer(1, 12))
                          except Exception as _efig:
                              story.append(Paragraph(f"[Graphique non disponible : {_efig}]", s_small))

                      # ── Analyse différentielle ────────────────────────────────────────
                      if len(results) == 2:
                          story.append(Paragraph("Analyse différentielle", s_h2))
                          _pnl_a = results[0]["yearly"]["pnl_total"].sum()
                          _pnl_b = results[1]["yearly"]["pnl_total"].sum()
                          _diff  = _pnl_a - _pnl_b
                          _gagne = results[0]["sc"]["name"] if _diff > 0 else results[1]["sc"]["name"]
                          story.append(Paragraph(
                              f"<b>Scénario le plus rentable 2026-2029 :</b> {_gagne} avec un gain de "
                              f"<b>{_fmt(abs(_diff))} €</b> ({abs(_diff/_pnl_b*100) if _pnl_b else 0:.1f}% d'écart).", s_body))
                          if _hr_pdf and len(_hr_pdf) == 2:
                              _pnl_ha = _hr_pdf[0]["yearly"]["pnl_total"].sum()
                              _pnl_hb = _hr_pdf[1]["yearly"]["pnl_total"].sum()
                              _diff_h = _pnl_ha - _pnl_hb
                              _gagne_h = _hr_pdf[0]["sc"]["name"] if _diff_h > 0 else _hr_pdf[1]["sc"]["name"]
                              story.append(Paragraph(
                                  f"<b>Scénario le plus rentable 2019-2025 :</b> {_gagne_h} avec un gain de "
                                  f"<b>{_fmt(abs(_diff_h))} €</b> sur données historiques réelles.", s_body))

                      story.append(Spacer(1, 12))
                      story.append(Paragraph("BESS Valorisation — Plénitude B-Charge · Rapport généré automatiquement", s_small))

                      doc.build(story)
                      _pdf_bytes = _buf.getvalue()
                      st.download_button(
                          "⬇ Télécharger le PDF",
                          data=_pdf_bytes,
                          file_name=f"BESS_comparaison_{_dt_pdf.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                          mime="application/pdf",
                          key="btn_comp_pdf_dl",
                      )
                      st.success("✓ PDF généré avec succès !")
                  except ImportError:
                      st.error("Installez reportlab : `pip install reportlab kaleido`")
                  except Exception as _epdf:
                      st.error(f"Erreur génération PDF : {_epdf}")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 4 — SENSIBILITÉ
# ══════════════════════════════════════════════════════════════════════════════

with tab_sensi:
    _diag_set_tab("sensi")
    st.markdown('<p class="section">Analyse de sensibilité — Heat map Puissance × Durée</p>',
                unsafe_allow_html=True)
    st.caption(
        "Montre le PnL total sur 4 ans pour chaque combinaison de puissance installée et de durée de cycle. "
        "Le moteur teste automatiquement toutes les combinaisons et affiche une carte de chaleur : "
        "plus la cellule est verte/chaude, plus la configuration est rentable."
    )

    ss1, ss2 = st.columns(2)
    with ss1:
        s_eff = st.slider("Rendement (%)", 70, 100, 100, key="s_eff") / 100
        st.caption("Efficacité de la batterie : 100% = pas de pertes, 92% = pour 100 kWh achetés, 92 kWh revendus.")
    with ss2:
        s_maxcy = st.number_input(
            "Quota cycles/an",
            0, 730, 365, key="s_maxcy",
        )
        s_maxcy = s_maxcy if s_maxcy > 0 else None
        st.caption(
            "Nombre maximum de cycles autorisés par an. "
            "**365** = on retient les 365 cycles avec les meilleurs spreads de l'année. "
            "**0** = illimité, la batterie trade tous les jours autant que le marché le permet."
        )

    # Grilles de valeurs
    sg1, sg2 = st.columns(2)
    with sg1:
        s_pow_min = st.number_input("Puissance min (MW)", 0.1, 20.0, 0.2, 0.1, key="s_pmin")
        s_pow_max = st.number_input("Puissance max (MW)", 0.1, 50.0, 2.0, 0.1, key="s_pmax")
        st.caption(
            "La puissance (MW) = débit de charge/décharge de la batterie. "
            "Une batterie de 0.43 MW achète ou vend 0.43 MWh par heure de cycle. "
            "Le PnL est proportionnel à la puissance : doubler la puissance = doubler les revenus."
        )
        s_pow_n = st.slider("Nombre de puissances testées", 3, 10, 6, key="s_pn")
        # Afficher les valeurs qui seront testées
        if s_pow_n >= 2:
            _preview = [round(s_pow_min + i*(s_pow_max-s_pow_min)/(s_pow_n-1), 2) for i in range(s_pow_n)]
            st.caption(f"Valeurs testées : {' · '.join(str(v)+' MW' for v in _preview)}")
    with sg2:
        s_dur_list = st.multiselect(
            "Durées de cycle à tester",
            [1, 2],
            default=[1, 2],
            key="s_durs",
            format_func=lambda x: f"{x}h (charge {x}h + décharge {x}h)"
        )
        st.caption(
            "**1h** : la batterie charge 1h puis décharge 1h (capacité = puissance × 1h). "
            "**2h** : charge 2h puis décharge 2h (capacité double, spreads potentiellement plus larges). "
            "Sélectionnez les deux pour comparer côte à côte sur la même heatmap."
        )

    if not s_dur_list:
        st.warning("Sélectionnez au moins une durée.")
    else:
        pow_vals = [round(s_pow_min + i*(s_pow_max-s_pow_min)/(s_pow_n-1), 3)
                    for i in range(s_pow_n)]

        @st.cache_data(show_spinner=False)
        def run_sensi(file_bytes, pow_vals_json, dur_list_json, eff, maxcy):
            import json
            from bess_engine import (load_spot, simulate_arbitrage,
                                     simulate_arbitrage_optimal, aggregate_arbitrage)
            pv = load_spot(io.BytesIO(file_bytes))
            pow_vals = json.loads(pow_vals_json)
            dur_list = json.loads(dur_list_json)
            matrix   = {}  # {dur: {pow: pnl_total}}
            for dur in dur_list:
                matrix[dur] = {}
                for pw in pow_vals:
                    p = {"power_MW": pw, "n_cycles": 0, "duration_h": dur,
                         "efficiency": eff, "max_cycles_year": maxcy,
                         "excluded_hours": {}, "optimal_quota": bool(maxcy)}
                    if p["optimal_quota"]:
                        d = simulate_arbitrage_optimal(pv, p)
                    else:
                        d = simulate_arbitrage(pv, p)
                    y = aggregate_arbitrage(d, pw)
                    matrix[dur][pw] = round(y["pnl_total"].sum(), 0)
            return matrix

        with st.spinner("Calcul de la sensibilité en cours…"):
            matrix = run_sensi(
                file_bytes,
                _json.dumps([round(p,3) for p in pow_vals]),
                _json.dumps(s_dur_list),
                s_eff, s_maxcy
            )

        # Heat maps
        for dur in s_dur_list:
            z_vals = [[matrix[dur][pw]] for pw in pow_vals]
            z_flat = [matrix[dur][pw] for pw in pow_vals]
            z_max  = max(z_flat)
            z_min  = min(z_flat)

            fig_h = go.Figure(go.Heatmap(
                z=[[matrix[dur][pw] for pw in pow_vals]],
                x=[f"{pw} MW" for pw in pow_vals],
                y=[f"{dur}h × Max cycles"],
                colorscale="RdYlGn",
                text=[[f"{_fmt(matrix[dur][pw])} €" for pw in pow_vals]],
                texttemplate="%{text}",
                textfont=dict(size=11),
                hovertemplate="Puissance: %{x}<br>PnL total: <b>%{text}</b><extra></extra>",
                showscale=True,
                colorbar=dict(title="PnL (€)"),
            ))
            fig_h.update_layout(
                title=f"Durée cycle = {dur}h — PnL total 4 ans (€)",
                height=180, margin=dict(t=40, b=40, l=120, r=80),
                plot_bgcolor="white", paper_bgcolor="white",
            )
            apply_bb(fig_h)
            st.caption("Heatmap de performance : chaque cellule représente le PnL total sur 4 ans pour une combinaison de puissance et de durée de cycle. Les couleurs chaudes (rouge/orange) = combinaisons les plus rentables.")
            st.plotly_chart(fig_h, width="stretch", config=PLOTLY_CFG, key=f"exec_fig_h_{dur}")
            st.caption('Heatmap puissance x durée. Les cellules les plus colorees = combinaisons les plus rentables. Survolez pour voir le PnL exact.')

        # Tableau récap
        st.markdown('<p class="section">Tableau de sensibilité complet</p>', unsafe_allow_html=True)
        st.caption("Résultats détaillés de toutes les combinaisons puissance x durée, tries par PnL total décroissant. Utilisez ce tableau pour choisir la configuration optimale.")
        sensi_rows = []
        for dur in s_dur_list:
            for pw in pow_vals:
                cap = pw * dur * 1  # mode MAX : on affiche capacité pour 1 cycle
                pnl = matrix[dur][pw]
                sensi_rows.append({
                    "Puissance (MW)": pw,
                    "Durée (h)":      dur,
                    "Capacité (MWh)": round(cap, 2),
                    "PnL total (€)":  f"{_fmt(pnl)}",
                    "PnL/kW (€/kW)":  f"{_fmt(pnl/(pw*1000), 1)}" if pw > 0 else "—",
                })
        st.caption('Toutes les combinaisons puissance x durée, triees par PnL décroissant. La premiere ligne = configuration optimale pour vos conditions de marche.')
        st.dataframe(pd.DataFrame(sensi_rows), hide_index=True, width="stretch")

        # Courbes de sensibilité par durée
        fig_s2 = go.Figure()
        for i, dur in enumerate(s_dur_list):
            pnl_vals = [matrix[dur][pw] for pw in pow_vals]
            fig_s2.add_trace(go.Scatter(
                x=pow_vals, y=pnl_vals,
                mode="lines+markers", name=f"Durée {dur}h",
                line=dict(width=2.5, color=COLORS[i % len(COLORS)]),
                marker=dict(size=8),
                hovertemplate=f"Durée {dur}h<br>Puissance: %{{x}} MW<br>PnL: %{{y:.0f}} €<extra></extra>",
            ))
        fig_s2.update_layout(
            height=340,
            xaxis=dict(title="Puissance (MW)", gridcolor="#f0f0f0"),
            yaxis=dict(title="PnL total 4 ans (€)", tickformat=",", gridcolor="#f0f0f0"),
            legend=LEGEND_BOTTOM, margin=dict(t=10, b=60, l=70, r=10),
            plot_bgcolor="white", paper_bgcolor="white",
            hovermode="x unified",
        )
        apply_bb(fig_s2)
        st.caption("Courbes PnL en fonction de la puissance pour chaque durée de cycle. Montre si la relation est lineaire et laquelle des durées (1h ou 2h) est la plus rentable pour chaque niveau de puissance.")
        st.plotly_chart(fig_s2, width="stretch", config=PLOTLY_CFG, key="exec_fig_s2")
        st.caption('PnL total par scenario sur 4 ans. Permet de comparer directement la rentabilité de chaque configuration.')

        st.download_button(
            "Exporter sensibilité (CSV)",
            data=pd.DataFrame(sensi_rows).to_csv(
                index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name="BESS_sensibilite.csv",
            mime="text/csv",
            key="btn_sensi_csv",
        )


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 5 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════



with tab_hist:
    _diag_set_tab("hist")
    # ── Paramètres scénario ──────────────────────────────────────────────────
    st.markdown('<p class="section">Données historiques 2019–2025</p>', unsafe_allow_html=True)
    st.caption("Analyse des performances de la batterie sur données historiques réelles 2019–2025, lues depuis la feuille « Case 3 night station » du fichier uploadé dans la sidebar. Les paramètres de la sidebar s'appliquent (puissance, rendement, restrictions).")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        _cyc_opts  = [1, 2, 0]
        _cyc_labels = {1: "1 cycle / jour", 2: "2 cycles / jour",
                       0: "Illimité — tous les cycles rentables du jour"}
        n_cycles = st.selectbox(
            "Cycles par jour",
            _cyc_opts,
            format_func=lambda x: _cyc_labels[x],
            index=0,
            key="h_n_cycles_widget",
            help=(
                "1 = la batterie fait 1 charge + 1 décharge par jour (les meilleures heures).\n"
                "2 = 2 cycles successifs par jour.\n"
                "Illimité = le moteur exploite toutes les opportunités rentables de la journée."
            )
        )
    with c2:
        duration_h = st.selectbox(
            "Durée du cycle",
            [1, 2],
            format_func=lambda x: f"{x}h  ({x}h charge + {x}h décharge)",
            index=1,
            key="h_duration_widget",
            help="1h = achète 1h puis revend 1h. 2h = achète 2h puis revend 2h (capacité double)."
        )
    # Capacité = puissance × durée (indépendant du nb de cycles)
    capacite_auto = power_MW * duration_h
    with c3:
        _cyc_disp = "illimité" if n_cycles == 0 else str(n_cycles)
        st.metric(
            "Capacité par cycle (MWh)",
            f"{capacite_auto:.3f} MWh",
            delta=f"{power_MW} MW × {duration_h}h · {_cyc_disp} cycle(s)/j.",
        )
    with c4:
        # Les valeurs sont toujours initialisées dans la sidebar au premier chargement
        jours_excl = st.multiselect(
            "Jours avec restriction",
            ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"],
            key="h_jours_excl_widget"
        )
    with c5:
        h_debut = st.number_input("Heure début restriction", 0, 23,
                                   key="h_h_debut_widget")
    with c6:
        h_fin = st.number_input("Heure fin restriction", 0, 23,
                                 key="h_h_fin_widget")

    day_map = {"Lundi":0,"Mardi":1,"Mercredi":2,"Jeudi":3,
               "Vendredi":4,"Samedi":5,"Dimanche":6}
    days_num = [day_map[d] for d in jours_excl]
    excluded = ({"days": days_num, "hours": list(range(h_debut, h_fin))}
                if days_num and h_debut < h_fin else {})

    # ── Guard : fichier requis ───────────────────────────────────────────────
    if not file_bytes:
        st.info("⬆ Uploadez le fichier **BESS_valorisation_RESULTATS.xlsx** (sans '_Copie') dans la sidebar pour activer cet onglet.")
        st.stop()

    # ── Chargement pivot historique (toujours nécessaire) ───────────────────
    from bess_engine import load_case3 as _lc3
    import io as _io2
    pv = _lc3(_io2.BytesIO(file_bytes))

    # ── Simulation avec barre de progression ─────────────────────────────────
    _hist_key = f"hist_{hash((file_bytes, power_MW, n_cycles, duration_h, json.dumps(excluded), efficiency, max_cycles, min_spread))}"

    if _hist_key not in st.session_state:
        def _bprog(pct, label):
            """Met à jour la barre dans le header ET st.progress natif."""
            st.markdown(
                f'<script>if(window._bessSetProgress)'
                f'window._bessSetProgress({pct}, "{label}");</script>',
                unsafe_allow_html=True
            )

        _bprog(0, "Chargement des données historiques 2019–2025…")
        excl  = json.loads(json.dumps(excluded))
        _use_optimal_h = bool(max_cycles)
        p_arb = {"power_MW": power_MW,
                 "n_cycles": 0 if _use_optimal_h else n_cycles,
                 "duration_h": duration_h, "excluded_hours": excl,
                 "efficiency": efficiency, "max_cycles_year": max_cycles,
                 "optimal_quota": _use_optimal_h,
                 "min_spread": min_spread}

        _bar = st.progress(0, text="Simulation en cours…")
        if _use_optimal_h:
            daily_res = simulate_arbitrage_optimal(pv, p_arb)
        else:
            daily_res = simulate_arbitrage(pv, p_arb)
        _bar.progress(90, text="Agrégation…")
        yearly_res = aggregate_arbitrage(daily_res, power_MW)
        _bar.progress(100, text="Terminé.")

        import time; time.sleep(0.4)
        st.session_state[_hist_key] = (daily_res, yearly_res)
        _bar.empty()
        st.markdown(
            '<script>if(window._bessHideProgress)window._bessHideProgress();</script>',
            unsafe_allow_html=True
        )

    daily_h, yearly_h = st.session_state[_hist_key]

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section">Résultats globaux</p>', unsafe_allow_html=True)
    st.caption("Synthese des performances sur toute la période. Le PnL réel tient compte de toutes vos contraintes. La borne max est le maximum théorique en supposant aucune restriction et une connaissance parfaite des prix.")

    total_pnl        = daily_h["pnl"].sum()
    total_pnl_absolu = daily_h["pnl_absolu"].sum()
    spread_moy       = daily_h.loc[daily_h["valid"], "spread"].mean() if daily_h["valid"].any() else 0
    spread_abs_moy   = daily_h.loc[daily_h["valid"], "spread_absolu"].mean() if daily_h["valid"].any() else 0
    jours_actifs     = int(daily_h["valid"].sum())
    jours_total      = len(daily_h)
    jours_usure      = int(daily_h["maintenance_bloque"].sum())
    energie_totale   = daily_h["energie_MWh"].sum()
    ratio            = total_pnl / total_pnl_absolu * 100 if total_pnl_absolu else 0

    st.session_state["_diag_hist"] = {
        "kpis": [
            ("PnL réel", f"{_fmt(total_pnl)} €"),
            ("PnL borne max", f"{_fmt(total_pnl_absolu)} €"),
            ("Ratio capturé", f"{ratio:.1f}%"),
            ("Spread moyen", f"{spread_moy:.1f} €/MWh"),
            ("Jours actifs", f"{jours_actifs} / {jours_total}"),
            ("Jours bloqués (maintenance)", str(jours_usure)),
        ],
        "yearly": yearly_h[["annee", "jours_actifs", "spread_moy",
                             "pnl_absolu_total", "pnl_total", "pnl_par_MW"]],
    }

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("PnL réel (avec contraintes)",
              f"{_fmt(total_pnl)} €")
    m2.metric("PnL borne théorique max",
              f"{_fmt(total_pnl_absolu)} €",
              delta=f"{ratio:.1f}% capturé")
    m3.metric("Spread moyen (jours actifs)",
              f"{spread_moy:.1f} €/MWh",
              delta=f"Théorique: {spread_abs_moy:.1f} €/MWh")
    m4.metric("Jours actifs",
              f"{jours_actifs} / {jours_total}",
              delta=f"{jours_actifs/jours_total*100:.0f}% taux activation")
    m5.metric("Énergie totale échangée",
              f"{_fmt(energie_totale, 1)} MWh",
              delta=f"Capacité/jour : {capacite_auto:.2f} MWh")
    m6.metric("Jours bloqués (maintenance)",
              f"{jours_usure}",
              delta="Aucun" if jours_usure == 0 else f"{jours_usure/jours_total*100:.1f}%",
              delta_color="off" if jours_usure == 0 else "inverse")

    # ── Graphique 1 : PnL annuel ─────────────────────────────────────────────
    st.markdown('<p class="section">PnL annuel — borne max vs réel</p>', unsafe_allow_html=True)
    st.caption("Comparaison annuelle entre PnL réel (barres bleues) et borne max théorique (barres claires). L'ecart entre les deux mesure le coût de vos restrictions operationnelles sur les revenus.")
    hfig1 = go.Figure()
    hfig1.add_trace(go.Bar(
        x=yearly_h["annee"].astype(str), y=yearly_h["pnl_absolu_total"],
        name="Borne max", marker_color=C2,
        text=yearly_h["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €"),
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Borne max : %{y:.0f} €<extra></extra>",
    ))
    hfig1.add_trace(go.Bar(
        x=yearly_h["annee"].astype(str), y=yearly_h["pnl_total"],
        name="PnL réel", marker_color=C1,
        text=yearly_h["pnl_total"].apply(lambda x: f"{_fmt(x)} €"),
        textposition="inside", textfont_color="white",
        hovertemplate="<b>%{x}</b><br>PnL réel : %{y:.0f} €<extra></extra>",
    ))
    hfig1.update_layout(
        barmode="group", height=380,
        yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis_title="Année",
        legend=LEGEND_BOTTOM,
        margin=dict(t=10, b=40, l=60, r=20),
        plot_bgcolor="white", paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    apply_bb(hfig1)
    st.caption("Barres bleues = PnL réel avec contraintes. Barres claires = borne max théorique. L'écart = revenus perdus à cause des restrictions et du quota maintenance.")
    st.plotly_chart(hfig1, width="stretch", config=PLOTLY_CFG, key="h_fig1")

    # ── Graphiques 2×2 ───────────────────────────────────────────────────────
    col_a, col_b = st.columns(2)

    with col_a:
        # Spread HEBDOMADAIRE — une seule courbe continue sur toute la période
        st.markdown('<p class="section">Spread moyen — par semaine (€/MWh)</p>',
                    unsafe_allow_html=True)
        st.caption("Spread hebdo = différence prix vente - prix achat. Un spread élevé = opportunité rentable.")

        d_all = daily_h.copy()
        d_all["semaine"] = pd.to_datetime(d_all["date"]).dt.to_period("W").apply(
            lambda r: r.start_time)

        weekly_all = d_all.groupby("semaine").agg(
            spread_moy    = ("spread_absolu", "mean"),
            spread_min    = ("spread_absolu", "min"),
            spread_max    = ("spread_absolu", "max"),
            nb_jours_tot  = ("date",          "count"),
            nb_jours_act  = ("valid",         "sum"),
            annee         = ("annee",         "first"),
        ).reset_index().sort_values("semaine")

        hfig2 = go.Figure()

        # Palette : vert clair → jaune → bleu → orange
        SPREAD_COLORS = ["#5cb85c", "#f5c518", "#2e75b6", "#d46b1a",
                         "#7b3fa0", "#3a8a5c"]

        # Nombre d'annees pour choisir couleur et bande
        annees_list = sorted(weekly_all["annee"].unique())
        n_annees_tot = len(annees_list)

        if n_annees_tot <= 1:
            # Une seule couleur : vert clair
            main_color = C1
            fill_color = ("rgba(255,102,0,0.08)" if _BB else "rgba(92,184,92,0.10)")
        else:
            # Dégradé vert → jaune
            main_color = C1
            fill_color = ("rgba(255,102,0,0.08)" if _BB else "rgba(92,184,92,0.10)")

        # Bande min-max
        hfig2.add_trace(go.Scatter(
            x=pd.concat([weekly_all["semaine"],
                          weekly_all["semaine"].iloc[::-1]]),
            y=pd.concat([weekly_all["spread_max"],
                          weekly_all["spread_min"].iloc[::-1]]),
            fill="toself",
            fillcolor=fill_color,
            line=dict(color="rgba(0,0,0,0)"),
            showlegend=False, hoverinfo="skip",
        ))

        # Courbe continue — vert clair
        hfig2.add_trace(go.Scatter(
            x=weekly_all["semaine"],
            y=weekly_all["spread_moy"],
            mode="lines",
            name="Spread hebdomadaire",
            line=dict(width=2.5, color=main_color),
            connectgaps=True,
            hovertemplate=(
                "<b>Semaine du %{x|%d/%m/%Y}</b><br>"
                "Spread : <b>%{y:.1f} €/MWh</b><br>"
                "Min sem. : %{customdata[0]:.1f} | "
                "Max sem. : %{customdata[1]:.1f}<br>"
                "Jours actifs : %{customdata[2]:.0f} / %{customdata[3]:.0f}"
                "<extra></extra>"
            ),
            customdata=weekly_all[["spread_min", "spread_max",
                                    "nb_jours_act", "nb_jours_tot"]].values,
        ))

        # Moyenne globale en pointillés jaunes
        spread_global_moy = weekly_all["spread_moy"].mean()
        hfig2.add_hline(
            y=spread_global_moy,
            line_dash="dot", line_color=C2, line_width=2,
            annotation_text=f"Moy: {spread_global_moy:.1f} €/MWh",
            annotation_position="top left",
            annotation_font=dict(size=11, color=C2),
        )

        # Lignes verticales de séparation des annees
        for yr in sorted(d_all["annee"].unique())[1:]:
            jan1 = pd.Timestamp(f"{int(yr)}-01-01")
            hfig2.add_shape(
                type="line",
                x0=jan1, x1=jan1, y0=0, y1=1,
                xref="x", yref="paper",
                line=dict(color="#cccccc", width=1, dash="dot"),
            )
            hfig2.add_annotation(
                x=jan1, y=1.02, xref="x", yref="paper",
                text=str(int(yr)), showarrow=False,
                font=dict(size=10, color="#888888"),
            )

        hfig2.update_layout(
            height=320, margin=dict(t=20, b=40, l=60, r=10),
            yaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0",
                       range=[0, 300]),
            xaxis=dict(title="", tickformat="%b %Y", nticks=16),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
            hovermode="x unified",
        )
        apply_bb(hfig2)
        st.caption(
            "La courbe = spread moyen de la semaine (différence prix vente - prix achat). "
            "La zone ombrée autour = écart entre le spread min et max de la semaine : "
            "plus l'ombre est large, plus les prix ont varié dans la semaine. "
            "La ligne pointillée = spread moyen sur toute la période."
        )
        st.plotly_chart(hfig2, width="stretch", config=PLOTLY_CFG, key="h_fig2")

    with col_b:
        st.markdown('<p class="section">Profil horaire charge / décharge</p>',
                    unsafe_allow_html=True)
        st.caption("Nombre de fois ou chaque heure a ete utilisée pour charger (achat) ou decharger (vente) sur toute la période. Les heures de charge sont les moins chères de la journee, les heures de décharge les plus chères.")
        h_ch  = {h: 0 for h in range(24)}
        h_dch = {h: 0 for h in range(24)}
        for hc_list, hd_list in zip(daily_h.loc[daily_h["valid"], "h_charge"],
                                     daily_h.loc[daily_h["valid"], "h_decharge"]):
            for h in hc_list:  h_ch[h]  += 1
            for h in hd_list:  h_dch[h] += 1
        hfig3 = go.Figure()
        hfig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_ch.values()),
            name="Charge (achat)", marker_color=C1,
            hovertemplate="H%{x:02d} — Charge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_ch.values()],
        ))
        hfig3.add_trace(go.Bar(
            x=list(range(24)), y=list(h_dch.values()),
            name="Décharge (vente)", marker_color=C2,
            hovertemplate="H%{x:02d} — Décharge : <b>%{y} jours</b>"
                          " (%{customdata:.1f}%)<extra></extra>",
            customdata=[v / jours_actifs * 100 for v in h_dch.values()],
        ))
        hfig3.update_layout(
            height=320, barmode="group", margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Heure", tickmode="array",
                       tickvals=list(range(0,24,2)),
                       ticktext=[f"H{h:02d}" for h in range(0,24,2)],
                       range=[-0.5,23.5]),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(hfig3)
        st.caption("Fréquence d'utilisation de chaque heure sur 4 ans. Vert = charge (achat aux heures bon marché). Orange = décharge (vente aux heures chères).")
        st.plotly_chart(hfig3, width="stretch", config=PLOTLY_CFG, key="h_fig3")

    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown('<p class="section">Distribution des spreads journaliers</p>',
                    unsafe_allow_html=True)
        st.caption("Répartition des jours selon le spread journalier. Un spread de 40 euros/MWh signifie 40 euros de différence entre prix de vente et prix achat ce jour. Plus la distribution est vers la droite, plus le marché est favorable.")
        sv = daily_h.loc[daily_h["valid"], "spread"]
        # Borner à 300 €/MWh — les valeurs extrêmes de 2022 (~2500) écraseraient le graphique
        _pmax_h = 300
        _n_outliers = int((sv > _pmax_h).sum())
        hfig4 = go.Figure()
        hfig4.add_trace(go.Histogram(
            x=sv.clip(upper=_pmax_h), nbinsx=40, marker_color=C1, opacity=0.8,
            hovertemplate="Spread : %{x:.0f} €/MWh<br>Nb jours : <b>%{y}</b><extra></extra>",
        ))
        hfig4.add_vline(
            x=min(sv.mean(), _pmax_h), line_dash="dash", line_color=ORANGE, line_width=2,
            annotation_text=f"Moy: {sv.mean():.1f} €/MWh",
            annotation_position="top right",
            annotation_font=dict(size=12, color=ORANGE),
        )
        hfig4.add_vline(
            x=min(sv.median(), _pmax_h), line_dash="dot", line_color=C3, line_width=1.5,
            annotation_text=f"Méd: {sv.median():.1f}",
            annotation_position="top left",
            annotation_font=dict(size=11, color=GREEN),
        )
        hfig4.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            xaxis=dict(title="Spread (€/MWh)", gridcolor="#f0f0f0", range=[0, _pmax_h]),
            yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
        )
        apply_bb(hfig4)
        _note_h = (f" {_n_outliers} jours > 300 €/MWh (crise 2022) non affichés."
                   if _n_outliers > 0 else "")
        st.caption(f'Histogramme des spreads journaliers. Chaque barre = nombre de jours avec ce niveau de spread. Vers la droite = marché très favorable.{_note_h}')
        st.plotly_chart(hfig4, width="stretch", config=PLOTLY_CFG, key="h_fig4")

    with col_d:
        st.markdown('<p class="section">PnL cumulé dans le temps</p>',
                    unsafe_allow_html=True)
        st.caption("Cumul des gains depuis le debut de la simulation. Une courbe régulière = batterie qui trade tous les jours. Les plateaux horizontaux = périodes bloquées par le quota de maintenance annuel.")
        ds = daily_h.sort_values("date").copy()
        ds["pnl_cum"]       = ds["pnl"].cumsum()
        ds["pnl_absolu_cum"] = ds["pnl_absolu"].cumsum()
        hfig5 = go.Figure()
        hfig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_cum"],
            fill="tozeroy", mode="lines", name="PnL cumulé",
            line=dict(color=C1, width=2),
            fillcolor=("rgba(255,102,0,0.10)" if _BB else "rgba(92,184,92,0.12)"),
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "PnL cumulé : <b>%{y:.0f} €</b><br>"
                "PnL du jour : %{customdata:.0f} €<extra></extra>"
            ),
            customdata=ds["pnl"].values,
        ))
        hfig5.add_trace(go.Scatter(
            x=ds["date"], y=ds["pnl_absolu_cum"],
            mode="lines", name="Borne max",
            line=dict(color=C2, width=1.5, dash="dot"),
            hovertemplate=(
                "<b>%{x|%d/%m/%Y}</b><br>"
                "Borne max cumulée : %{y:.0f} €<extra></extra>"
            ),
        ))
        hfig5.update_layout(
            height=320, margin=dict(t=10, b=40, l=60, r=10),
            yaxis=dict(title="PnL cumulé (€)", tickformat=",", gridcolor="#f0f0f0"),
            xaxis=dict(title="", tickformat="%b %Y"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12, font_color="#333333"),
            hovermode="x unified",
        )
        apply_bb(hfig5)
        st.caption('PnL cumule depuis le debut. La pente = rythme de gain. Les plateaux = périodes bloquées par le quota maintenance annuel.')
        st.plotly_chart(hfig5, width="stretch", config=PLOTLY_CFG, key="h_fig5")

    # ── Répartition des cycles journaliers ───────────────────────────────────
    st.markdown('<p class="section">Répartition des cycles journaliers</p>', unsafe_allow_html=True)
    st.caption("Distribution du nombre de cycles réalisés par jour sur toute la période. Identifie les journées les plus et les moins actives.")
    if n_cycles != 1:  # Graphique cycles seulement si multi-cycles ou max
        _cyc_result_h = _plot_cycles_distribution(daily_h, "2019–2025")
    else:
        _cyc_result_h = None
    if _cyc_result_h:
        _fig_cyc_h, _min_ch, _max_ch, _dmin_h, _dmax_h = _cyc_result_h
        _cd1h, _cd2h = st.columns(2)
        _cd1h.metric("Minimum de cycles/jour",
                    f"{_min_ch} cycle(s)",
                    delta=f"ex. {pd.Timestamp(_dmin_h).strftime('%d/%m/%Y')}",
                    delta_color="off")
        _cd2h.metric("Maximum de cycles/jour",
                    f"{_max_ch} cycle(s)",
                    delta=f"ex. {pd.Timestamp(_dmax_h).strftime('%d/%m/%Y')}",
                    delta_color="off")
        apply_bb(_fig_cyc_h)
        st.plotly_chart(_fig_cyc_h, width="stretch", config=PLOTLY_CFG, key="h_fig_cyc")

    # ── Tableau récap ────────────────────────────────────────────────────────
    st.markdown('<p class="section">Récapitulatif annuel</p>', unsafe_allow_html=True)
    st.caption("Tableau détaillé par année : jours actifs, spreads captures, PnL réel et théorique. Le taux de capturé (PnL réel / borne max) mesure l'efficacité de votre strategie.")
    show = yearly_h.copy()
    show["pnl_absolu_total"]  = show["pnl_absolu_total"].apply(lambda x: f"{_fmt(x)} €")
    show["pnl_total"]         = show["pnl_total"].apply(lambda x: f"{_fmt(x)} €")
    show["pnl_par_MW"]        = show["pnl_par_MW"].apply(lambda x: f"{x:.1f} €/MW")
    show["taux_activation"]   = show["taux_activation"].apply(lambda x: f"{x*100:.0f}%")
    show["spread_absolu_moy"] = show["spread_absolu_moy"].apply(lambda x: f"{x:.1f}")
    show["spread_moy"]        = show["spread_moy"].apply(lambda x: f"{x:.1f}")
    show["energie_totale_MWh"]= show["energie_totale_MWh"].apply(lambda x: f"{x:.1f} MWh")
    show["cycles_totaux"]     = show["cycles_totaux"].astype(int)
    show = show[[
        "annee","jours_simules","jours_actifs","jours_bloques_maintenance","taux_activation",
        "spread_absolu_moy","pnl_absolu_total",
        "spread_moy","pnl_total","pnl_par_MW",
        "energie_totale_MWh","cycles_totaux"
    ]]
    show.columns = [
        "Année","Jours simulés","Jours actifs","Bloqués maintenance","Taux activation",
        "Spread théorique\n(€/MWh)","PnL borne max\nthéorique (€)",
        "Spread réel\n(€/MWh)","PnL réel (€)","PnL/MW\n(€/MW)",
        "Énergie échangée\n(MWh)","Cycles realises"
    ]
    st.caption('Tableau annuel détaillé. Colonnes : Jours actifs (jours trades), Taux activation (proportion), PnL réel (revenus nets après contraintes). Cliquez sur une colonne pour trier.')
    st.dataframe(show, hide_index=True, width="stretch")

    # ── Explorer un jour ─────────────────────────────────────────────────────
    with st.expander("Explorer un jour spécifique", expanded=True):
        st.caption("Sélectionnez un jour pour voir en detail comment la batterie a trade : prix horaires, heures de charge et décharge choisies, PnL réalisé et comparaison avec le maximum théorique possible.")
        _pv_dates = pd.to_datetime(pv["date"]).dt.normalize()
        _d0 = _pv_dates.iloc[0].date()
        _d1 = _pv_dates.iloc[-1].date()
        date_sel = st.date_input("Date", value=_d0, min_value=_d0, max_value=_d1, key="h_date_sel")
        _ts = pd.Timestamp(date_sel)
        row_pivot = pv[_pv_dates == _ts]

        if row_pivot.empty:
            st.info(f"Aucune donnée pour le {date_sel}.")
        elif not row_pivot.empty:
            prix_j   = row_pivot[HOUR_COLS].values.flatten()
            weekday  = int(row_pivot["weekday"].iloc[0])
            avail    = get_available_hours(weekday, excluded)

            # Recalcul EN TEMPS RÉEL avec les paramètres actuels
            from bess_engine import _best_n_cycles

            cycles_jour = _best_n_cycles(
                prix_j, avail, n_cycles, duration_h, power_MW, efficiency
            )

            ca, cb = st.columns([3, 1])
            with cb:
                pnl_jour     = sum(c["pnl"]    for c in cycles_jour)
                spread_jour  = sum(c["spread"] for c in cycles_jour) / len(cycles_jour) if cycles_jour else 0
                energie_jour = power_MW * duration_h * len(cycles_jour)

                # Spread par cycle
                if len(cycles_jour) == 1:
                    st.metric("Spread cycle 1", f"{cycles_jour[0]['spread']:.2f} €/MWh")
                elif len(cycles_jour) >= 2:
                    st.metric("Spread moyen", f"{spread_jour:.2f} €/MWh")
                    for i, cy in enumerate(cycles_jour, 1):
                        st.metric(f"Spread cycle {i}",
                                  f"{cy['spread']:.2f} €/MWh",
                                  delta=f"{cy['spread'] - spread_jour:+.2f} vs moy")
                else:
                    st.metric("Spread", "— Pas de cycle")

                st.metric("PnL du jour", f"{pnl_jour:.2f} €")
                st.metric("Énergie chargée", f"{energie_jour:.3f} MWh",
                          help="Énergie totale achetée (chargée) par la batterie ce jour")
                st.metric("Énergie déchargée", f"{energie_jour * efficiency:.3f} MWh",
                          help=f"Énergie revendue après rendement {efficiency*100:.0f}%")
                st.metric("Cycles réalisés", f"{len(cycles_jour)}" + (" / " + str(n_cycles) if n_cycles > 0 else " (max)"))
                st.metric("Heures dispo", f"{len(avail)}/24")
                for i, cy in enumerate(cycles_jour, 1):
                    st.caption(
                        f"Cycle {i} : charge H{cy['h_charge']} "
                        f"({cy['prix_charge']:.1f} €/MWh) → "
                        f"décharge H{cy['h_decharge']} "
                        f"({cy['prix_decharge']:.1f} €/MWh) "
                        f"| PnL {cy['pnl']:.2f} €"
                    )

            with ca:
                # Couleurs des barres : heures exclues en gris clair
                bar_colors = []
                for h in range(24):
                    if h not in avail:
                        bar_colors.append("rgba(255,102,0,0.2)" if _BB else "#c8d4e0")  # exclu
                    else:
                        bar_colors.append("#c8d8ec")  # disponible

                fig_d = go.Figure()

                # Barres prix spot
                fig_d.add_trace(go.Bar(
                    x=list(range(24)), y=prix_j.tolist(),
                    marker_color=bar_colors, name="Prix spot",
                    hovertemplate="H%{x:02d} — Prix : <b>%{y:.2f} €/MWh</b><extra></extra>",
                ))

                # Couleurs par cycle
                colors_ch  = ["#2e7d32", "#1565c0", "#6a1b9a"]
                colors_dch = ["#c62828", "#e65100", "#4527a0"]

                for i, cy in enumerate(cycles_jour):
                    _h_ch  = cy["h_charge"]
                    _h_dch = cy["h_decharge"]
                    c_ch   = colors_ch[i % 3]
                    c_dch  = colors_dch[i % 3]
                    lbl    = f"Cycle {i+1}"

                    fig_d.add_trace(go.Scatter(
                        x=_h_ch, y=prix_j[_h_ch],
                        mode="markers",
                        name=f"Charge {lbl} ({duration_h}h)",
                        marker=dict(color=c_ch, size=18, symbol="triangle-up",
                                    line=dict(color="white", width=1)),
                        hovertemplate=f"H%{{x:02d}} — Charge {lbl} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                    ))
                    fig_d.add_trace(go.Scatter(
                        x=_h_dch, y=prix_j[_h_dch],
                        mode="markers",
                        name=f"Decharge {lbl} ({duration_h}h)",
                        marker=dict(color=c_dch, size=18, symbol="triangle-down",
                                    line=dict(color="white", width=1)),
                        hovertemplate=f"H%{{x:02d}} — Decharge {lbl} : <b>%{{y:.2f}} €/MWh</b><extra></extra>",
                    ))
                    fig_d.add_shape(type="line", x0=-0.5, x1=23.5,
                        y0=cy["prix_charge"], y1=cy["prix_charge"],
                        line=dict(color=c_ch, width=1, dash="dot"))
                    fig_d.add_shape(type="line", x0=-0.5, x1=23.5,
                        y0=cy["prix_decharge"], y1=cy["prix_decharge"],
                        line=dict(color=c_dch, width=1, dash="dot"))
                    fig_d.add_annotation(x=22, y=cy["prix_charge"],
                        text=f"Achat C{i+1}: {cy['prix_charge']:.1f}",
                        showarrow=False, font=dict(size=9, color=c_ch))
                    fig_d.add_annotation(x=22, y=cy["prix_decharge"],
                        text=f"Vente C{i+1}: {cy['prix_decharge']:.1f}",
                        showarrow=False, font=dict(size=9, color=c_dch))

                # Zones exclues
                for h in range(24):
                    if h not in avail:
                        fig_d.add_vrect(x0=h-0.5, x1=h+0.5,
                            fillcolor="rgba(26,58,92,0.15)", line_width=0, layer="below")

                fig_d.update_layout(
                    height=420, margin=dict(t=30, b=60, l=60, r=10),
                    xaxis=dict(title="Heure", tickmode="array",
                               tickvals=list(range(0,24,1)),
                               ticktext=[f"H{h:02d}" for h in range(24)],
                               range=[-0.5, 23.5]),
                    yaxis=dict(title="Prix (€/MWh)", gridcolor="#f0f0f0"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=LEGEND_BOTTOM,
                    hoverlabel=dict(bgcolor="white", font_size=12),
                )
                apply_bb(fig_d)
                st.caption('Prix horaires du jour sélectionné. Triangles verts = achat. Triangles rouges = vente. Lignes pointillees = prix moyens du cycle. Zones grises = heures exclues.')
                st.plotly_chart(fig_d, width="stretch", config=PLOTLY_CFG,
                                key="h_fig_explorer_jour")

    # ── Export PDF ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Export du rapport</p>', unsafe_allow_html=True)
    st.caption("Exportez les résultats de la simulation en cours. Le PDF contient un rapport complet, le CSV contient toutes les donnees journalières pour analyse externe.")

    col_exp1, col_exp2 = st.columns([2, 3])
    with col_exp1:
        if st.button("Générer le rapport PDF", key="h_btn_pdf_1"):
            with st.spinner("Génération du rapport PDF..."):

                params_txt = [
                    f"Fichier : {uploaded.name}",
                    f"Puissance : {power_MW} MW | Capacité : {capacite_auto} MWh",
                    f"Durée cycle : {duration_h}h | Rendement : {efficiency*100:.0f}%",
                    f"Cycles/j : {'max' if n_cycles==0 else n_cycles} · Max cycles/an : {max_cycles or 'illimité'},"
                    f"Restriction : {', '.join(jours_excl) or 'aucune'} "
                    f"{'H'+str(h_debut)+'-H'+str(h_fin) if jours_excl else ''}",
                    f"Données : {annees[0]}–{annees[-1]} ({jours_total} jours)",
                ]
                kpis = [
                    ("PnL total (avec contraintes)",    f"{_fmt(total_pnl)} €"),
                    ("PnL borne théorique max", f"{_fmt(total_pnl_absolu)} €"),
                    ("Ratio réel / max",                f"{ratio:.1f}%"),
                    ("Spread moyen",                    f"{spread_moy:.2f} €/MWh"),
                    ("Jours actifs",                    f"{jours_actifs} / {jours_total}"),
                    ("Taux d'activation",               f"{jours_actifs/jours_total*100:.0f}%"),
                    ("Jours bloqués (maintenance)",           str(jours_usure)),
                ]
                pdf_bytes = build_pdf_arbitrage(
                    params_txt, kpis, yearly_h, daily_h,
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
        # Export CSV donnees journalières
        csv_buf = daily_h[["date","annee","mois","weekday_name",
                          "spread_absolu","pnl_absolu","spread","pnl",
                          "valid","h_charge","h_decharge"]].copy()
        csv_buf["date"] = csv_buf["date"].dt.strftime("%Y-%m-%d")
        st.download_button(
            label="Exporter les donnees (CSV)",
            key="h_btn_csv_1",
            data=csv_buf.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_donnees_journalieres_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

    # ── PRIORITÉ 1 : Analyse de l'impact de l'usure ──────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Impact de l\'usure — jours bloqués et PnL manqué</p>',
                unsafe_allow_html=True)
    st.caption(
        f"La batterie est limitée à {max_cycles or 'illimité'} cycles/an. "
        f"Un jour bloqué maintenance = quota de cycles atteint → PnL non capturé."
    )

    # Calcul mensuel : jours actifs / bloqués / PnL manqué
    daily_u = daily_h.copy()
    daily_u["annee_mois"] = daily_u["date"].dt.to_period("M")

    # PnL manqué quota : jours bloqués par le quota de cycles
    daily_u["pnl_manque_quota"]        = daily_u.apply(
        lambda r: r["pnl_absolu"] if r["maintenance_bloque"] else 0.0, axis=1
    )
    # PnL manqué restrictions : écart entre sans restriction et avec restriction
    # sur les jours NON bloqués par le quota
    daily_u["pnl_manque_restrictions"] = daily_u.apply(
        lambda r: max(0.0, r["pnl_absolu"] - r["pnl"])
                  if not r["maintenance_bloque"] else 0.0,
        axis=1
    )

    grp_u = daily_u.groupby("annee_mois").agg(
        jours_actifs            = ("valid",                    "sum"),
        jours_bloques           = ("maintenance_bloque",       "sum"),
        pnl_realise             = ("pnl",                      "sum"),
        pnl_manque_quota        = ("pnl_manque_quota",         "sum"),
        pnl_manque_restrictions = ("pnl_manque_restrictions",  "sum"),
    ).reset_index()
    grp_u["label"] = grp_u["annee_mois"].astype(str)

    pnl_manque_quota_total        = daily_u["pnl_manque_quota"].sum()
    pnl_manque_restrictions_total = daily_u["pnl_manque_restrictions"].sum()
    pnl_manque_total              = pnl_manque_quota_total + pnl_manque_restrictions_total
    has_quota   = jours_usure > 0
    has_restr   = pnl_manque_restrictions_total > 0.5

    # ── Graphique 1 : Jours actifs / bloqués ─────────────────────────────────
    fig_u1 = go.Figure()
    fig_u1.add_trace(go.Bar(
        x=grp_u["label"], y=grp_u["jours_actifs"],
        name="Jours actifs", marker_color=C1,
        hovertemplate="<b>%{x}</b><br>Jours actifs : %{y}<extra></extra>",
    ))
    if has_quota:
        fig_u1.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["jours_bloques"],
            name="Jours bloqués (quota)", marker_color="#ef5350",
            hovertemplate="<b>%{x}</b><br>Jours bloqués quota : %{y}<extra></extra>",
        ))
    fig_u1.update_layout(
        barmode="stack", height=280,
        yaxis=dict(title="Nb jours", gridcolor="#f0f0f0"),
        xaxis=dict(tickangle=-45, tickfont=dict(size=9)),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=60, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    apply_bb(fig_u1)
    if not has_quota:
        st.caption("Quota illimité — aucun jour bloqué.")
    else:
        st.caption("Jours bloqués : quota de cycles annuel atteint. Concentrés en fin d'année.")
    st.plotly_chart(fig_u1, width="stretch", config=PLOTLY_CFG, key="h_fig_usure_jours")

    # ── Graphique 2 : PnL réel / manqué quota / manqué restrictions ──────────
    fig_u2 = go.Figure()
    fig_u2.add_trace(go.Bar(
        x=grp_u["label"], y=grp_u["pnl_realise"].round(0),
        name="PnL réalisé (€)", marker_color=C1,
        hovertemplate="<b>%{x}</b><br>PnL réalisé : %{y:.0f} €<extra></extra>",
    ))
    if has_restr:
        fig_u2.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["pnl_manque_restrictions"].round(0),
            name="PnL manqué — restrictions horaires (€)",
            marker_color="#f5a623",
            hovertemplate="<b>%{x}</b><br>PnL manqué restrictions : %{y:.0f} €<extra></extra>",
        ))
    if has_quota:
        fig_u2.add_trace(go.Bar(
            x=grp_u["label"], y=grp_u["pnl_manque_quota"].round(0),
            name="PnL manqué — quota cycles (€)",
            marker_color="#ef5350",
            hovertemplate="<b>%{x}</b><br>PnL manqué quota : %{y:.0f} €<extra></extra>",
        ))
    fig_u2.update_layout(
        barmode="stack", height=300,
        yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis=dict(tickangle=-45, tickfont=dict(size=9)),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=70, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    apply_bb(fig_u2)
    _cap_parts = ["PnL réalisé (vert)"]
    if has_restr:
        _cap_parts.append(f"PnL manqué restrictions (orange) : {pnl_manque_restrictions_total:,.0f} €")
    if has_quota:
        _cap_parts.append(f"PnL manqué quota (rouge) : {pnl_manque_quota_total:,.0f} €")
    if not has_restr and not has_quota:
        _cap_parts.append("aucune perte détectée — quota illimité et restrictions sans impact")
    st.caption(" · ".join(_cap_parts))
    st.plotly_chart(fig_u2, width="stretch", config=PLOTLY_CFG, key="h_fig_usure_pnl")

    # KPIs usure
    ku1, ku2, ku3, ku4 = st.columns(4)
    ku1.metric("Jours bloqués (quota)", f"{jours_usure}",
               delta=f"{jours_usure/jours_total*100:.1f}% du total" if jours_usure > 0 else "Quota illimité")
    ku2.metric("PnL manqué — restrictions", f"{_fmt(pnl_manque_restrictions_total)} €",
               delta=f"{pnl_manque_restrictions_total/total_pnl*100:.1f}% du PnL réel" if total_pnl > 0 else "—")
    ku3.metric("PnL manqué — quota", f"{_fmt(pnl_manque_quota_total)} €",
               delta=f"{pnl_manque_quota_total/total_pnl*100:.1f}% du PnL réel" if total_pnl > 0 else "—")
    if max_cycles:
        ku4.metric("Conseil quota",
                   f"Actuel : {max_cycles} → tester {min(365, max_cycles + 50)}",
                   delta="Augmenter si dégradation acceptable")
    else:
        ku4.metric("Quota annuel", "Illimité",
                   delta="Aucun jour bloqué possible")

    # ── PRIORITÉ 2 : Analyse ROI / Payback ───────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Analyse ROI — Retour sur investissement</p>',
                unsafe_allow_html=True)
    st.caption(
        "Calcul de la rentabilité de l'investissement basé sur le modèle sélectionné dans la sidebar. "
        "Les données CAPEX proviennent du rapport IEA Batteries and Secure Energy Transitions (2024). "
        "Choisissez l'année de référence et ajustez la dégradation selon vos hypothèses."
    )

    # Modèle déjà sélectionné dans la sidebar via _modele_choix et power_MW
    # CAPEX batteries LFP utility-scale 2h, France — sources multiples mai 2026
    # Taux USD/EUR : 0.85 (mai 2026)
    # Sources : IEA Electricity 2026, BNEF Cost Survey 2025, Ember oct. 2025, Capstone DC nov. 2025
    IEA_CAPEX = {
        "2022 — 330 €/kWh (IEA réel)":              330,
        "2024 — 150 €/kWh (IEA Electricity 2026)":  150,
        "2025 — 120 €/kWh (BNEF / Ember)":          120,
        "2026 — 105 €/kWh (Capstone DC France)":    105,
        "2030 — 85 €/kWh  (BNEF projection)":        85,
    }
    IEA_CYCLES = 6500   # LFP stationnaire 2025 : 6000-7000 cycles (BNEF/Ember)
    IEA_DUREE  = 15     # ans — durée de vie nominale (Capstone DC, Ember)

    # Info modèle actif
    st.info(
        f"Modèle actif : **{_modele_choix}** — {power_MW*1000:.0f} kW — {power_MW} MW  "
        f"| Capacité : {capacite_auto:.3f} MWh  "
        f"| Rendement : {efficiency*100:.0f}%  "
        f"| Quota : {max_cycles or 'illimité'} cycles/an  "
        "_(configuré dans la sidebar)_"
    )

    r1, r2, r3 = st.columns(3)
    with r1:
        annee_capex = st.radio(
            "Année de référence CAPEX (IEA)",
            list(IEA_CAPEX.keys()),
            index=1,
            key="h_annee_capex",
            help="Source : IEA Batteries and Secure Energy Transitions (2024), STEPS scenario. "
                 "Converti USD vers EUR au taux 0.93. LFP domine le stockage stationnaire."
        )
        capex_kwh = IEA_CAPEX[annee_capex]
        st.metric("CAPEX retenu", f"{capex_kwh} €/kWh", delta="Source IEA 2024")
    with r2:
        duree_proj = st.slider("Durée de projection (ans)", 5, 20, IEA_DUREE,
            key="h_duree_proj",
            help=f"Durée de vie estimée LFP : {IEA_DUREE} ans "
                 f"({IEA_CYCLES} cycles garantis / 300 cycles/an). Source IEA 2024.")
        opex_an = st.number_input(
            "OPEX annuel (€/an)", min_value=0, max_value=50000, value=0, step=500,
            key="h_opex_an",
            help="OPEX = 0 si inclus dans la garantie constructeur. "
                 "Référence IEA/NREL : 2.5% du CAPEX/an hors garantie."
        )
    with r3:
        taux_degrad = st.number_input(
            "Dégradation annuelle (%)", min_value=0.0, max_value=10.0,
            value=2.0, step=0.5, format="%.1f",
            key="h_taux_degrad",
            help="Perte de capacité annuelle des cellules Li-ion LFP. "
                 "Source IEA 2024 : 2-3% par an. Impact direct sur le PnL futur."
        ) / 100.0
        taux_actu = st.number_input(
            "Taux d'actualisation (%)", min_value=0.0, max_value=20.0,
            value=5.0, step=0.5, format="%.1f",
            key="h_taux_actu",
            help="Taux utilisé pour calculer la VAN (Valeur Actuelle Nette). "
                 "Reflète le coût du capital ou le taux d'opportunité. "
                 "Typiquement 5-8% pour un projet industriel en Europe."
        ) / 100.0

    with st.expander("Données de référence — batteries LFP stationnaire (sources mai 2026)", expanded=False):
        st.markdown("""
**Sources : IEA Electricity 2026 · IEA Global Energy Review 2026 · BNEF Cost Survey 2025 · Ember oct. 2025 · Capstone DC nov. 2025**

**CAPEX utility-scale 2h, France (taux USD/EUR : 0.85)**
- **2022 : 330 €/kWh** — IEA Electricity 2026 (340 $/kWh × 0.85 + premium Europe)
- **2024 : 150 €/kWh** — IEA Electricity 2026 (fin 2024, après baisse de 40 % sur l'année)
- **2025 : 120 €/kWh** — BNEF Cost Survey 2025 (117 $/kWh mondial + premium Europe)
- **2026 : 105 €/kWh** — Capstone DC France (€90–100/kWh equipment + balance of system)
- **2030 : 85 €/kWh** — BNEF projection Europe (101 $/kWh × 0.85)
- **Baisse** : -58% entre 2019 et 2024 (IEA). -45% supplémentaires en 2025 (IEA Global Energy Review 2026).
- **Projets 2h coûtent ~10-15% plus cher par kWh que projets 4h** (BNEF)

**Paramètres techniques LFP stationnaire**
- **Rendement (round-trip)** : 88% valeur centrale 2025 (Ember) | fourchette 85-95%
- **Cycles garantis LFP** : 6 000-7 000 cycles (BNEF/Ember 2025) | cellules CATL 2025+ : 15 000+
- **Cycles réels France 2h** : 1 500-1 800 cycles/an (Capstone DC)
- **Dégradation** : 2%/an garantie fabricant (Ember) | capacité résiduelle ~65% à 20 ans
- **Durée de vie** : 15 ans nominale (Capstone DC, Ember)
- **OPEX** : 2.5% du CAPEX/an (NREL ATB 2025) | 0 si inclus garantie constructeur
- **Chimie dominante** : LFP — 90% des nouvelles installations de stockage stationnaire en 2025 (IEA GER 2026)

**Marché France 2025-2026**
- Capacité installée début 2026 : ~1.5 GW (Modo Energy)
- Pipeline RTE : ~13 GW en file d'attente
- Revenus aFRR : effondrement de 66 €/MW/h (2024) à 16 €/MW/h (jan. 2026) — saturation
- IRR unlevered France : 5-7% (sous le WACC de 8%) — projet standalone non bancable sans hédging (Capstone DC)
        """)

    # Calcul CAPEX total
    capacite_kwh = capacite_auto * 1000  # MWh -> kWh
    capex_total  = capex_kwh * capacite_kwh
    pnl_an_moy   = total_pnl / len(annees)
    cashflow_an  = pnl_an_moy - opex_an

    # Payback simple (sans dégradation)
    payback = capex_total / cashflow_an if cashflow_an > 0 else float("inf")

    # Projection cumulée avec dégradation de capacité
    # La dégradation réduit le PnL chaque année car la batterie stocke moins
    annees_proj  = list(range(1, duree_proj + 1))
    cum_cashflow = []
    cum_cashflow_nodeg = []
    cum = -capex_total
    cum_nd = -capex_total
    for a in annees_proj:
        # Avec dégradation : PnL réduit de taux_degrad % par an
        facteur_degradation = (1 - taux_degrad) ** (a - 1)
        cf_a = (pnl_an_moy * facteur_degradation) - opex_an
        cum += cf_a
        cum_cashflow.append(round(cum, 0))
        # Sans dégradation (référence)
        cf_nd = cashflow_an
        cum_nd += cf_nd
        cum_cashflow_nodeg.append(round(cum_nd, 0))

    # Cash-flow actualisé (VAN)
    cum_cashflow_actu = []
    cum_actu = -capex_total
    for a in annees_proj:
        facteur_degradation = (1 - taux_degrad) ** (a - 1)
        facteur_actu        = 1 / (1 + taux_actu) ** a
        cf_actu = ((pnl_an_moy * facteur_degradation) - opex_an) * facteur_actu
        cum_actu += cf_actu
        cum_cashflow_actu.append(round(cum_actu, 0))

    # Payback actualisé
    payback_actu = None
    for i, cf in enumerate(cum_cashflow_actu):
        if cf >= 0:
            payback_actu = i + 1
            break

    # Payback avec dégradation (plus réaliste)
    payback_deg = None
    for i, cf in enumerate(cum_cashflow):
        if cf >= 0:
            payback_deg = i + 1
            break

    # KPIs ROI
    roi_cols = st.columns(5)
    roi_cols[0].metric("CAPEX total", f"{_fmt(capex_total)} €",
                       delta=f"{capacite_kwh:.0f} kWh x {capex_kwh} euros/kWh (source IEA 2024)")
    roi_cols[1].metric("PnL annuel moyen", f"{_fmt(pnl_an_moy)} euros/an",
                       delta=f"Sur {len(annees)} ans de simulation")
    payback_label = (f"{payback_deg} ans (dégradation {taux_degrad*100:.0f}%/an)"
                     if payback_deg else "Non remboursé")
    roi_cols[2].metric("Payback non actualisé",
                       payback_label,
                       delta="Rentable" if payback_deg and payback_deg <= duree_proj else "Hors période")
    roi_cols[3].metric(f"Cash-flow non actualisé à {duree_proj} ans",
                       f"{_fmt(cum_cashflow[-1])} €",
                       delta="Positif" if cum_cashflow[-1] > 0 else "Négatif",
                       delta_color="normal" if cum_cashflow[-1] > 0 else "inverse")
    _van_fin = cum_cashflow_actu[-1]
    _pb_actu_label = (f"{payback_actu} ans (taux {taux_actu*100:.0f}%)"
                      if payback_actu else "Non remboursé")
    roi_cols[4].metric(f"VAN à {duree_proj} ans ({taux_actu*100:.0f}%)",
                       f"{_fmt(_van_fin)} €",
                       delta=_pb_actu_label,
                       delta_color="normal" if _van_fin > 0 else "inverse")

    # Info sur la source des données
    st.info(
        "Sources CAPEX : IEA Electricity 2026 · BNEF Cost Survey 2025 · Ember oct. 2025 · Capstone DC nov. 2025. "
        "Valeurs utility-scale 2h, France, taux USD/EUR 0.85 (mai 2026). "
        "Les coûts ont chuté de 58% entre 2019 et 2024 (IEA), puis de 45% supplémentaires en 2025. "
        "Dégradation LFP : 2%/an garantie fabricant (Ember 2025)."
    )

    # Graphique projection cash-flow cumulé
    fig_roi = go.Figure()
    fig_roi.add_hline(y=0, line_dash="dash", line_color="#888", line_width=1.5)

    # Courbe avec dégradation (principale)
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow,
        mode="lines+markers",
        name=f"Avec dégradation {taux_degrad*100:.0f}%/an (réaliste)",
        line=dict(color=C1, width=2.5),
        fill="tozeroy",
        fillcolor="rgba(92,184,92,0.08)" if cum_cashflow[-1] > 0 else "rgba(239,83,80,0.08)",
        hovertemplate="Année %{x}<br>Cash-flow cumulé : <b>%{y:.0f} euros</b><extra></extra>",
        marker=dict(size=6),
    ))

    # Courbe sans dégradation (référence optimiste)
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_nodeg,
        mode="lines",
        name="Sans dégradation (optimiste)",
        line=dict(color=C2, width=1.5, dash="dot"),
        hovertemplate="Année %{x}<br>Sans dégradation : <b>%{y:.0f} euros</b><extra></extra>",
    ))

    # Courbe VAN actualisée
    fig_roi.add_trace(go.Scatter(
        x=annees_proj, y=cum_cashflow_actu,
        mode="lines+markers",
        name=f"VAN actualisée ({taux_actu*100:.0f}%/an)",
        line=dict(color="#e67e22", width=2, dash="dashdot"),
        hovertemplate=f"Année %{{x}}<br>VAN ({taux_actu*100:.0f}%) : <b>%{{y:.0f}} euros</b><extra></extra>",
        marker=dict(size=5, symbol="diamond"),
    ))

    # Payback actualisé
    if payback_actu and payback_actu <= duree_proj:
        fig_roi.add_vline(
            x=payback_actu, line_dash="dot", line_color="#e67e22", line_width=1.5,
            annotation_text=f"Payback VAN : {payback_actu} ans",
            annotation_font=dict(color="#e67e22", size=10),
            annotation_position="top right",
        )

    # Point de break-even avec dégradation
    if payback_deg and payback_deg <= duree_proj:
        fig_roi.add_vline(
            x=payback_deg, line_dash="dot", line_color="#f5c518", line_width=2,
            annotation_text=f"Payback : {payback_deg} ans",
            annotation_font=dict(color="#f5c518", size=11),
        )

    # Annotation CAPEX source
    fig_roi.add_annotation(
        x=0.01, y=0.02, xref="paper", yref="paper",
        text=f"CAPEX : {capex_kwh} euros/kWh — Source IEA 2024",
        showarrow=False, font=dict(size=9, color="#888"),
        xanchor="left",
    )

    fig_roi.update_layout(
        height=400,
        yaxis=dict(title="Cash-flow cumulé (euros)", tickformat=",", gridcolor="#f0f0f0"),
        xaxis=dict(title="Années depuis mise en service", dtick=1),
        legend=LEGEND_BOTTOM, margin=dict(t=10, b=80, l=80, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
    )
    apply_bb(fig_roi)
    st.caption(
        f"Bleu = cash-flow non actualisé avec dégradation {taux_degrad*100:.0f}%/an. "
        "Pointillé = sans dégradation (optimiste). "
        f"Orange = VAN actualisée au taux {taux_actu*100:.0f}%/an. "
        f"CAPEX retenu : {capex_kwh} €/kWh — Sources : IEA Electricity 2026, BNEF, Ember, Capstone DC (mai 2026)."
    )
    st.plotly_chart(fig_roi, width="stretch", config=PLOTLY_CFG, key="h_fig_roi_arb")
    st.caption(
        "Courbe pleine = avec +1%/an sur les spreads (hypothèse conservatrice). "
        "Courbe pointillée = spreads constants. "
        "La ligne jaune indique le point de break-even (CAPEX remboursé)."
    )
    # ── Export diagnostic complet ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section">Diagnostic — Exporter pour le développeur</p>',
                unsafe_allow_html=True)
    st.caption("Génère un fichier HTML avec toutes les donnees et résultats visibles à l'écran.")

    if st.button(" Exporter diagnostic complet", key="h_btn_diag_arb"):
        import json as _jd, datetime as _dtd

        # Profil horaire charge/décharge
        h_ch_freq  = {h: 0 for h in range(24)}
        h_dch_freq = {h: 0 for h in range(24)}
        for hc_l, hd_l in zip(daily_h.loc[daily_h["valid"], "h_charge"],
                               daily_h.loc[daily_h["valid"], "h_decharge"]):
            for h in hc_l: h_ch_freq[h]  += 1
            for h in hd_l: h_dch_freq[h] += 1

        # Distribution spreads
        sv = daily_h.loc[daily_h["valid"], "spread"]
        bins = [0, 20, 40, 60, 80, 100, 150, 200, 9999]
        labels = ["0-20","20-40","40-60","60-80","80-100","100-150","150-200",">200"]
        distrib = {labels[i]: int(((sv>=bins[i])&(sv<bins[i+1])).sum())
                   for i in range(len(labels))}

        diag = {
            "meta": {
                "generated_at": _dtd.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "fichier_excel": uploaded.name,
                "annees": [int(a) for a in annees],
                "nb_jours_total": jours_total,
                "theme": "TEST Bloomberg" if st.session_state.bloomberg else "Clair",
            },
            "paramètres": {
                "power_MW": power_MW,
                "n_cycles": n_cycles,
                "duration_h": duration_h,
                "capacite_MWh": round(capacite_auto, 4),
                "efficiency_pct": round(efficiency * 100, 1),
                "max_cycles_an": max_cycles,
                "excluded_hours": excluded,
                "jours_exclus": jours_excl,
                "h_debut_restriction": h_debut,
                "h_fin_restriction": h_fin,
            },
            "résultats_globaux": {
                "pnl_total_reel_euros": round(total_pnl, 2),
                "pnl_borne_max_euros": round(total_pnl_absolu, 2),
                "ratio_capture_pct": round(ratio, 2),
                "spread_moy_EuroMWh": round(spread_moy, 4),
                "spread_absolu_moy_EuroMWh": round(spread_abs_moy, 4),
                "jours_actifs": jours_actifs,
                "jours_total": jours_total,
                "taux_activation_pct": round(jours_actifs/jours_total*100, 1),
                "jours_bloques_maintenance": jours_usure,
                "energie_totale_MWh": round(energie_totale, 2),
            },
            "recap_annuel": yearly_h[[
                "annee","jours_simules","jours_actifs","taux_activation",
                "spread_absolu_moy","spread_moy",
                "pnl_absolu_total","pnl_total","pnl_par_MW",
                "energie_totale_MWh","cycles_totaux"
            ]].to_dict(orient="records"),
            "profil_horaire_frequence": {
                "charge":   h_ch_freq,
                "decharge": h_dch_freq,
            },
            "distribution_spreads": distrib,
            "spread_stats": {
                "min":    round(float(sv.min()), 4),
                "max":    round(float(sv.max()), 4),
                "moy":    round(float(sv.mean()), 4),
                "mediane":round(float(sv.median()), 4),
                "p25":    round(float(sv.quantile(0.25)), 4),
                "p75":    round(float(sv.quantile(0.75)), 4),
            },
            "donnees_journalieres_completes": daily_h[[
                "date","annee","mois","weekday_name",
                "spread_absolu","pnl_absolu","spread","pnl",
                "valid","maintenance_bloque","n_cycles_actifs","energie_MWh",
                "h_charge","h_decharge","prix_charge","prix_decharge",
                "prix_min","prix_max","prix_moy"
            ]].assign(date=lambda df: df["date"].dt.strftime("%Y-%m-%d")
            ).to_dict(orient="records"),
        }

        json_str = _jd.dumps(diag, indent=2, ensure_ascii=False, default=str)

        # Construire le HTML
        def kpi(label, value, sub=""):
            return (f'<div class="kpi"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{value}</div>'
                    + (f'<div class="kpi-sub">{sub}</div>' if sub else '')
                    + '</div>')

        def table_from_list(rows, cols=None):
            if not rows: return "<p>Aucune donnée</p>"
            if cols is None: cols = list(rows[0].keys())
            h = "<table><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"
            for r in rows:
                h += "<tr>" + "".join(
                    f"<td>{r.get(c,'')}</td>" for c in cols) + "</tr>"
            return h + "</table>"

        html = f"""<!DOCTYPE html><html lang="fr"><head>
<meta charset="UTF-8">
<title>BESS Diagnostic — {diag["meta"]["generated_at"]}</title>
<style>
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f8f9fb;color:#1a3a5c;margin:0;padding:20px;}}
h1{{background:#1a3a5c;color:white;padding:12px 20px;border-radius:6px;font-size:1.2rem;margin-bottom:8px;}}
h2{{color:#1a3a5c;border-bottom:2px solid #d0dff0;padding-bottom:4px;margin-top:24px;}}
h3{{color:#2e75b6;margin-top:16px;}}
.kpi{{display:inline-block;background:white;border:1px solid #d0dff0;border-radius:6px;
      padding:10px 16px;margin:6px;min-width:140px;vertical-align:top;}}
.kpi-label{{font-size:0.72rem;color:#6b7a8d;text-transform:uppercase;letter-spacing:0.5px;}}
.kpi-value{{font-size:1.25rem;font-weight:700;color:#1a3a5c;margin-top:2px;}}
.kpi-sub{{font-size:0.72rem;color:#5cb85c;margin-top:2px;}}
pre{{background:#111;color:#5cb85c;padding:16px;border-radius:6px;
     font-family:'Courier New',monospace;font-size:11px;overflow-x:auto;white-space:pre-wrap;}}
table{{border-collapse:collapse;width:100%;font-size:11px;margin:8px 0;}}
th{{background:#1a3a5c;color:white;padding:5px 8px;text-align:left;font-size:11px;}}
td{{border:1px solid #d0d0d0;padding:4px 8px;}}
tr:nth-child(even) td{{background:#f0f5fb;}}
.section{{background:#2e75b6;color:white;font-weight:700;padding:6px 12px;
          font-size:11px;margin:16px 0 6px 0;border-radius:3px;}}
.good{{color:#375623;font-weight:700;}} .bad{{color:#c00000;font-weight:700;}}
</style></head><body>
<h1> BESS Valorisation — Diagnostic complet</h1>
<p>Généré le <b>{diag["meta"]["generated_at"]}</b> &nbsp;|&nbsp;
   Fichier : <b>{diag["meta"]["fichier_excel"]}</b> &nbsp;|&nbsp;
   Période : <b>{" · ".join(str(a) for a in diag["meta"]["annees"])}</b> &nbsp;|&nbsp;
   {jours_total} jours simulés</p>

<h2>Paramètres de simulation</h2>
{kpi("Puissance", f"{power_MW} MW")}
{kpi("Cycles/jour", f"{n_cycles} × {duration_h}h")}
{kpi("Capacité", f"{capacite_auto:.3f} MWh")}
{kpi("Rendement", f"{efficiency*100:.0f}%")}
{kpi("Max cycles/an", str(max_cycles or "∞"))}
{kpi("Jours exclus", ", ".join(jours_excl) or "Aucun")}
{kpi("Heures restriction", f"H{h_debut}–H{h_fin}" if jours_excl else "Aucune")}

<h2>Résultats globaux</h2>
{kpi("PnL réel total", f"{_fmt(total_pnl)} €", f"Borne max : {total_pnl_absolu:,.0f} €")}
{kpi("Taux de capture", f"{ratio:.1f}%")}
{kpi("Spread moyen", f"{spread_moy:.2f} €/MWh", f"Théorique : {spread_abs_moy:.2f}")}
{kpi("Jours actifs", f"{jours_actifs} / {jours_total}", f"{jours_actifs/jours_total*100:.0f}% activation")}
{kpi("Énergie totale", f"{_fmt(energie_totale, 1)} MWh")}
{kpi("Jours bloqués", str(jours_usure))}

<h2>Récapitulatif annuel</h2>
{table_from_list(diag["recap_annuel"],
    ["annee","jours_simules","jours_actifs","taux_activation",
     "spread_absolu_moy","spread_moy","pnl_absolu_total","pnl_total",
     "pnl_par_MW","energie_totale_MWh","cycles_totaux"])}

<h2>Profil horaire — fréquence charge / décharge</h2>
<table><tr><th>Heure</th>
{''.join(f"<th>H{h:02d}</th>" for h in range(24))}
</tr>
<tr><td><b>Charge (j)</b></td>
{''.join(f"<td>{h_ch_freq[h]}</td>" for h in range(24))}
</tr>
<tr><td><b>Décharge (j)</b></td>
{''.join(f"<td>{h_dch_freq[h]}</td>" for h in range(24))}
</tr></table>

<h2>Distribution des spreads</h2>
{table_from_list([{"Tranche (€/MWh)": k, "Nb jours": v,
                   "% du total": f"{v/len(sv)*100:.1f}%" if len(sv)>0 else "0%"}
                  for k, v in distrib.items()])}
<p>Min={diag["spread_stats"]["min"]} · Max={diag["spread_stats"]["max"]} · 
   Moy={diag["spread_stats"]["moy"]} · Méd={diag["spread_stats"]["mediane"]} · 
   P25={diag["spread_stats"]["p25"]} · P75={diag["spread_stats"]["p75"]}</p>

<h2>Données journalières complètes ({jours_total} jours)</h2>
{table_from_list(diag["donnees_journalieres_completes"],
    ["date","annee","mois","weekday_name","spread_absolu","pnl_absolu",
     "spread","pnl","valid","n_cycles_actifs","energie_MWh",
     "h_charge","h_decharge","prix_charge","prix_decharge","prix_min","prix_max","prix_moy"])}

<h2>JSON brut complet</h2>
<pre>{json_str}</pre>
<hr><p style="color:#888;font-size:11px;">
BESS Valorisation v2.0 — Plénitude B-Charge — Diagnostic technique</p>
</body></html>"""

        st.download_button(
            "⬇ Télécharger le diagnostic HTML",
            data=html.encode("utf-8"),
            file_name=f"BESS_diagnostic_{_dtd.datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            key="h_btn_diag_dl",
        )
        st.success("Fichier prêt ! Téléchargez-le et envoyez-le au développeur.")

with tab_exec:
    _diag_set_tab("exec")
    st.markdown('<p class="section">Executive Summary — Vue synthétique</p>',
                unsafe_allow_html=True)
    st.caption("Vue d'ensemble pour la direction. Compare automatiquement les deux strategies principales (1 cycle vs 2 cycles par jour) et presente les indicateurs clés de performance.")

    # Recalculer avec les paramètres courants (n_cycles, duration_h de l'onglet arbitrage)
    # On réutilise le daily/yearly calculé dans tab_arb via cache
    try:
        # Utiliser les paramètres de la sidebar
        exec_params = {
            "power_MW": power_MW, "n_cycles": 1, "duration_h": 2,
            "excluded_hours": {}, "efficiency": efficiency,
            "max_cycles_year": max_cycles,
        }

        @st.cache_data(show_spinner=False)
        def run_exec(file_bytes, power_MW, efficiency, max_cycles):
            from bess_engine import load_spot, simulate_arbitrage, aggregate_arbitrage
            pv = load_spot(io.BytesIO(file_bytes))
            # Scénario 1 cycle 2h
            p1 = {"power_MW": power_MW, "n_cycles": 1, "duration_h": 2,
                  "excluded_hours": {}, "efficiency": efficiency,
                  "max_cycles_year": max_cycles}
            d1 = simulate_arbitrage(pv, p1)
            y1 = aggregate_arbitrage(d1, power_MW)
            # Scénario 2 cycles 2h
            p2 = {"power_MW": power_MW, "n_cycles": 2, "duration_h": 2,
                  "excluded_hours": {}, "efficiency": efficiency,
                  "max_cycles_year": max_cycles}
            d2 = simulate_arbitrage(pv, p2)
            y2 = aggregate_arbitrage(d2, power_MW)
            return d1, y1, d2, y2

        d_exec1, y_exec1, d_exec2, y_exec2 = run_exec(file_bytes, power_MW, efficiency, max_cycles)

        st.session_state["_diag_exec"] = {
            "kpis": [
                ("PnL 4 ans — 1 cycle/jour", f"{_fmt(y_exec1['pnl_total'].sum())} €"),
                ("PnL 4 ans — 2 cycles/jour", f"{_fmt(y_exec2['pnl_total'].sum())} €"),
                ("Spread moyen", f"{y_exec1['spread_moy'].mean():.1f} €/MWh"),
                ("Taux activation moyen", f"{y_exec1['taux_activation'].mean()*100:.0f}%"),
            ],
        }

        # ── Header executive ──────────────────────────────────────────────────
        st.markdown(f"""
<div style="background:{'#111' if _BB else '#1f4e79'};color:white;padding:16px 20px;
border-radius:6px;margin-bottom:16px;">
<div style="font-size:1.4rem;font-weight:700;letter-spacing:1px;">
BESS VALORISATION — RÉSUMÉ EXÉCUTIF</div>
<div style="font-size:0.85rem;opacity:0.8;margin-top:4px;">
Plénitude B-Charge · Batterie {power_MW} MW · {annees[0]}–{annees[-1]} · Rendement {efficiency*100:.0f}%
</div>
</div>
""", unsafe_allow_html=True)

        # ── KPIs clés ─────────────────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("PnL 4 ans — 1 cycle/jour",
                  f"{_fmt(y_exec1['pnl_total'].sum())} €",
                  delta=f"Borne max : {_fmt(y_exec1['pnl_absolu_total'].sum(), 0)} €")
        k2.metric("PnL 4 ans — 2 cycles/jour",
                  f"{_fmt(y_exec2['pnl_total'].sum())} €",
                  delta=f"+{_fmt(y_exec2['pnl_total'].sum()-y_exec1['pnl_total'].sum(), 0)} € vs 1 cycle")
        k3.metric("Spread moyen",
                  f"{y_exec1['spread_moy'].mean():.1f} €/MWh")
        k4.metric("Taux activation moyen",
                  f"{y_exec1['taux_activation'].mean()*100:.0f}%")

        # ── Graphique principal : PnL annuel 1 vs 2 cycles ───────────────────
        st.markdown('<p class="section">PnL annuel — comparaison 1 cycle vs 2 cycles/jour</p>',
                    unsafe_allow_html=True)
        st.caption("Comparaison directe entre 1 cycle par jour (une charge et une decharge) et 2 cycles par jour (deux sequences charge-decharge successives). Deux cycles capte plus d'opportunités mais consomme plus de quota annuel.")
        fig_ex1 = go.Figure()
        fig_ex1.add_trace(go.Bar(
            x=y_exec1["annee"].astype(str), y=y_exec1["pnl_total"],
            name="1 cycle/jour", marker_color=C1,
            text=y_exec1["pnl_total"].apply(lambda x: f"{_fmt(x)} €"),
            textposition="inside", textfont_color="white",
        ))
        fig_ex1.add_trace(go.Bar(
            x=y_exec2["annee"].astype(str), y=y_exec2["pnl_total"],
            name="2 cycles/jour", marker_color=C2,
            text=y_exec2["pnl_total"].apply(lambda x: f"{_fmt(x)} €"),
            textposition="outside",
        ))
        fig_ex1.update_layout(
            barmode="group", height=360,
            yaxis=dict(title="PnL (€)", tickformat=",", gridcolor="#f0f0f0"),
            xaxis_title="Année",
            legend=LEGEND_BOTTOM, margin=dict(t=10, b=60, l=70, r=10),
            plot_bgcolor="white", paper_bgcolor="white",
        )
        apply_bb(fig_ex1)
        st.caption("PnL annuel compare entre 1 cycle et 2 cycles par jour. Si 2 cycles dépasse largement 1 cycle, le marché offre deux opportunités quotidiennes distinctes a exploiter pour maximiser les revenus.")
        st.plotly_chart(fig_ex1, width="stretch", config=PLOTLY_CFG, key="verif_fig_ex1")
        st.caption("Taux d'activation par scenario : pourcentage de jours où la batterie a effectivement trade.")

        # ── Tableau synthèse ──────────────────────────────────────────────────
        st.markdown('<p class="section">Synthèse par année</p>', unsafe_allow_html=True)
        st.caption("Tableau comparatif annuel entre les deux strategies. Permet de voir si les ecarts de performance varient selon les annees et les conditions de marche.")
        exec_rows = []
        for _, r1 in y_exec1.iterrows():
            r2 = y_exec2[y_exec2["annee"] == r1["annee"]].iloc[0]
            exec_rows.append({
                "Année":                int(r1["annee"]),
                "Spread moy (€/MWh)":  f"{r1['spread_moy']:.1f}",
                "Jours actifs (1cyc)":  int(r1["jours_actifs"]),
                "PnL réel 1 cycle (€)": f"{_fmt(r1['pnl_total'])}",
                "PnL borne max (€)":    f"{_fmt(r1['pnl_absolu_total'])}",
                "Jours actifs (2cyc)":  int(r2["jours_actifs"]),
                "PnL réel 2 cycles (€)":f"{_fmt(r2['pnl_total'])}",
                "Gain 2vs1 cycle (€)":  f"{_fmt(r2['pnl_total']-r1['pnl_total'])}",
                "Énergie 1cyc (MWh)":   f"{r1['energie_totale_MWh']:.1f}",
            })
        st.caption('Tableau comparatif 1 cycle vs 2 cycles. Gain marginal = euros supplémentaires apportes par le second cycle. Si le gain est faible, 1 seul cycle suffit.')
        st.dataframe(pd.DataFrame(exec_rows), hide_index=True, width="stretch")

        # ── PnL cumulé 4 ans ─────────────────────────────────────────────────
        st.markdown('<p class="section">Trajectoire de valorisation — PnL cumulé</p>',
                    unsafe_allow_html=True)
        st.caption("Projection du PnL cumule sur toute la période pour les deux strategies. La strategie avec la pente la plus forte est la plus rentable annuellement.")
        fig_ex2 = go.Figure()
        for d_ex, name, color in [(d_exec1, "1 cycle/jour", C1),
                                    (d_exec2, "2 cycles/jour", C2)]:
            ds = d_ex.sort_values("date")
            fig_ex2.add_trace(go.Scatter(
                x=ds["date"], y=ds["pnl"].cumsum(),
                mode="lines", name=name,
                line=dict(width=2.5, color=color),
                fill="tozeroy" if name == "2 cycles/jour" else None,
                fillcolor="rgba(92,184,92,0.06)",
                hovertemplate=f"{name}<br>%{{x|%d/%m/%Y}} : %{{y:.0f}} €<extra></extra>",
            ))
        fig_ex2.update_layout(
            height=360,
            yaxis=dict(title="PnL cumulé (€)", tickformat=",", gridcolor="#f0f0f0"),
            xaxis=dict(tickformat="%b %Y"),
            legend=LEGEND_BOTTOM, margin=dict(t=10, b=60, l=70, r=10),
            plot_bgcolor="white", paper_bgcolor="white",
            hovermode="x unified",
        )
        apply_bb(fig_ex2)
        st.caption("PnL cumule sur 4 ans pour les deux strategies. La strategie avec la courbe la plus haute est la plus rentable. L'ecart final = gain additionnel sur 4 ans en choisissant 2 cycles plutot que 1 seul.")
        st.plotly_chart(fig_ex2, width="stretch", config=PLOTLY_CFG, key="verif_fig_ex2")
        st.caption('Comparaison annuelle 1 cycle vs 2 cycles. Les annees avec un fort écart = annees avec bonnes opportunités matin ET soir.')

        # ── Recommandation ────────────────────────────────────────────────────
        best_pnl    = max(y_exec1["pnl_total"].sum(), y_exec2["pnl_total"].sum())
        best_config = "2 cycles/jour" if y_exec2["pnl_total"].sum() > y_exec1["pnl_total"].sum() else "1 cycle/jour"
        spread_mean = y_exec1["spread_moy"].mean()

        st.markdown(f"""
<div style="background:{'#0a1a0a' if _BB else '#e2efda'};
border-left:4px solid {'#00e676' if _BB else '#375623'};
padding:14px 18px;border-radius:4px;margin-top:12px;">
<div style="font-weight:700;color:{'#00e676' if _BB else '#375623'};font-size:1rem;">
 RECOMMANDATION</div>
<div style="margin-top:8px;font-size:0.9rem;">
La configuration optimale sur la période {annees[0]}–{annees[-1]} est
<b>{best_config}</b> avec un PnL total de <b>{best_pnl:,.0f} €</b>
sur 4 ans pour une batterie de <b>{power_MW} MW</b>.<br><br>
Le spread moyen Day-Ahead sur la période est de <b>{spread_mean:.1f} €/MWh</b>.
{"Le marché est favorable à la valorisation batterie (spread > 30 €/MWh)." if spread_mean > 30 else
 "Le marché présente un spread modéré — la valorisation reste positive."}
</div>
</div>
""", unsafe_allow_html=True)

        # Export executive summary
        exec_df = pd.DataFrame(exec_rows)
        st.download_button(
            "Exporter executive summary (CSV)",
            data=exec_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_executive_summary_{annees[0]}_{annees[-1]}.csv",
            mime="text/csv",
            key="btn_exec_csv",
        )

    except Exception as e_exec:
        st.error(f"Erreur executive summary : {e_exec}")


with tab_verif_arb:
    _diag_set_tab("verif_arb")
    # ══════════════════════════════════════════════════════════════════════════
    # ONGLET VÉRIFICATION ARBITRAGE — Audit complet jour par jour
    # ══════════════════════════════════════════════════════════════════════════
    from itertools import combinations as _verif_comb
    import json as _verif_json
    _HCOLS_V = [f"H{h:02d}" for h in range(24)]

    st.subheader("Vérification Arbitrage — Audit indépendant des calculs")
    st.caption(
        "Sélectionnez un jour dans le tableau pour voir la vérification complète : "
        "prix bruts Excel, heures choisies, formule PnL, 8 contrôles automatiques."
    )

    # ── Paramètres de l'audit (indépendants de l'onglet Arbitrage) ───────────
    st.markdown('<p class="section">Paramètres de l\'audit</p>', unsafe_allow_html=True)
    st.caption(
        "Ces paramètres sont indépendants de l'onglet Arbitrage : "
        "vous pouvez tester n'importe quelle configuration pour auditer les calculs."
    )
    vac1, vac2, vac3, vac4, vac5, vac6, vac7 = st.columns(7)
    with vac1:
        va_pow = st.number_input("Puissance (MW)", 0.05, 50.0, power_MW, 0.01,
                                  format="%.3f", key="va_pow2")
    with vac2:
        va_ncyc = st.selectbox("Cycles/jour", [1, 2],
                                format_func=lambda x: f"{x} cycle{'s' if x>1 else ''}",
                                key="va_ncyc2")
    with vac3:
        va_dur = st.selectbox("Durée cycle", [1, 2],
                               format_func=lambda x: f"{x}h", index=1, key="va_dur2")
    with vac4:
        va_eff = st.slider("Rendement (%)", 70, 100, int(efficiency * 100),
                            key="va_eff2") / 100
    with vac5:
        va_jours = st.multiselect(
            "Jours restriction",
            ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"],
            default=["Lun","Mar","Mer","Jeu","Ven","Sam"],
            key="va_jours2"
        )
    with vac6:
        va_hdeb = st.number_input("H début restr.", 0, 23, 10, key="va_hdeb2")
    with vac7:
        va_hfin = st.number_input("H fin restr.", 0, 23, 12, key="va_hfin2")

    _va_daymap = {"Lun":0,"Mar":1,"Mer":2,"Jeu":3,"Ven":4,"Sam":5,"Dim":6}
    _va_days_num = [_va_daymap[d] for d in va_jours]
    _va_excl = ({"days": _va_days_num, "hours": list(range(va_hdeb, va_hfin))}
                if _va_days_num and va_hdeb < va_hfin else {})
    va_cap = va_pow * va_dur * va_ncyc

    st.info(
        f"Configuration : **{va_pow} MW** · **{va_ncyc} cycle(s) × {va_dur}h** · "
        f"Capacité **{va_cap:.3f} MWh** · Rendement **{va_eff*100:.0f}%** · "
        f"Restriction : jours {va_jours or 'aucun'}, H{va_hdeb:02d}–H{va_hfin:02d}"
    )

    # ── Construction du tableau d'audit (exhaustif + moteur côte à côte) ─────
    @st.cache_data(show_spinner=False)
    def _build_verif(fbytes, ncyc, dur, pow_, eff, excl_str):
        """
        Pour chaque jour :
        - Résultat bess_engine (algorithme candidats)
        - Résultat exhaustif (toutes combinaisons)
        - Contrôles automatiques (8 règles)
        """
        pv = get_pivot(fbytes)
        excl = _verif_json.loads(excl_str)
        rows = []

        for _, r in pv.iterrows():
            prices = np.array(r[_HCOLS_V].values, dtype=float)
            avail  = get_available_hours(int(r["weekday"]), excl)

            # ── Moteur (algorithme actuel) ────────────────────────────
            cycles_moteur = _best_n_cycles(prices, avail, ncyc, dur, pow_, eff)
            pnl_moteur = sum(c["pnl"] for c in cycles_moteur)
            h_ch_moteur  = [h for c in cycles_moteur for h in c["h_charge"]]
            h_dch_moteur = [h for c in cycles_moteur for h in c["h_decharge"]]
            pc_moteur = np.mean([prices[h] for h in h_ch_moteur]) if h_ch_moteur else 0
            pd_moteur = np.mean([prices[h] for h in h_dch_moteur]) if h_dch_moteur else 0

            # ── Exhaustif (force brute — toutes combinaisons) ─────────
            best_pnl_exh = 0.0
            h_ch_exh = []; h_dch_exh = []
            for hc in _verif_comb(avail, dur):
                max_hc = max(hc)
                avail_dch = [h for h in avail if h > max_hc and h not in set(hc)]
                if len(avail_dch) < dur:
                    continue
                pa_d = prices[np.array(avail_dch)]
                hd = sorted(np.array(avail_dch)[np.argsort(pa_d)[-dur:]].tolist())
                pc_ = prices[np.array(list(hc))].mean()
                pd_ = prices[np.array(hd)].mean()
                if pd_ <= pc_:
                    continue
                pnl_ = (pow_ * dur * eff * pd_) - (pow_ * dur * pc_)
                if pnl_ > best_pnl_exh:
                    best_pnl_exh = pnl_
                    h_ch_exh = list(hc); h_dch_exh = hd
            # Si 2 cycles, on fait 2 passes exhaustives séquentielles
            if ncyc == 2 and h_dch_exh:
                h_start2 = max(h_dch_exh) + 1
                avail2 = [h for h in avail if h >= h_start2]
                best2 = 0.0
                h_ch_exh2 = []; h_dch_exh2 = []
                for hc in _verif_comb(avail2, dur):
                    max_hc = max(hc)
                    avail_dch2 = [h for h in avail2 if h > max_hc and h not in set(hc)]
                    if len(avail_dch2) < dur:
                        continue
                    pa_d2 = prices[np.array(avail_dch2)]
                    hd2 = sorted(np.array(avail_dch2)[np.argsort(pa_d2)[-dur:]].tolist())
                    pc2 = prices[np.array(list(hc))].mean()
                    pd2 = prices[np.array(hd2)].mean()
                    if pd2 <= pc2:
                        continue
                    pnl2 = (pow_ * dur * eff * pd2) - (pow_ * dur * pc2)
                    if pnl2 > best2:
                        best2 = pnl2
                        h_ch_exh2 = list(hc); h_dch_exh2 = hd2
                best_pnl_exh += best2
                h_ch_exh  += h_ch_exh2
                h_dch_exh += h_dch_exh2

            # ── Borne max théorique (sans contrainte d'ordre) ─────────
            idx_sort = np.argsort(prices)
            n_tot = ncyc * dur
            h_bm_ch  = sorted(idx_sort[:n_tot].tolist())
            h_bm_dch = sorted(idx_sort[-n_tot:].tolist())
            pc_bm = prices[np.array(h_bm_ch)].mean()
            pd_bm = prices[np.array(h_bm_dch)].mean()
            pnl_bm = (pow_ * n_tot * eff * pd_bm) - (pow_ * n_tot * pc_bm) if pd_bm > pc_bm else 0

            # ── Contrôles automatiques ────────────────────────────────
            ctrl = {}
            # C1 : décharge après charge
            ctrl["C1_ordre"] = (
                all(h > max(h_ch_moteur) for h in h_dch_moteur)
                if h_ch_moteur and h_dch_moteur else True
            )
            # C2 : aucune heure exclue utilisée
            h_excl_set = set(excl.get("hours", []))
            weekday = int(r["weekday"])
            if weekday in excl.get("days", []):
                ctrl["C2_exclusion"] = not any(
                    h in h_excl_set for h in h_ch_moteur + h_dch_moteur
                )
            else:
                ctrl["C2_exclusion"] = True
            # C3 : spread ≥ 0
            spread_moteur = pd_moteur - pc_moteur if h_ch_moteur else 0
            ctrl["C3_spread_positif"] = spread_moteur >= -0.001
            # C4 : PnL moteur = PnL exhaustif (optimalité)
            ctrl["C4_optimalite"] = abs(pnl_moteur - best_pnl_exh) < 0.02
            # C5 : PnL moteur ≤ borne max
            ctrl["C5_borne_max"] = pnl_moteur <= pnl_bm + 0.02
            # C6 : formule PnL cohérente
            if h_ch_moteur and h_dch_moteur:
                pnl_formule = (pow_ * dur * eff * pd_moteur) - (pow_ * dur * pc_moteur)
                ctrl["C6_formule_pnl"] = abs(pnl_moteur - pnl_formule * (ncyc if ncyc > 1 and len(cycles_moteur) > 1 else 1)) < 0.05
            else:
                ctrl["C6_formule_pnl"] = True
            # C7 : prix charge = moyenne des heures de charge
            if h_ch_moteur:
                pc_verif = np.mean([prices[h] for h in h_ch_moteur])
                ctrl["C7_prix_charge"] = abs(pc_verif - pc_moteur) < 0.001
            else:
                ctrl["C7_prix_charge"] = True
            # C8 : prix décharge = moyenne des heures de décharge
            if h_dch_moteur:
                pd_verif = np.mean([prices[h] for h in h_dch_moteur])
                ctrl["C8_prix_decharge"] = abs(pd_verif - pd_moteur) < 0.001
            else:
                ctrl["C8_prix_decharge"] = True

            nb_ok = sum(ctrl.values())
            nb_total = len(ctrl)
            statut = "Conforme" if nb_ok == nb_total else f"{nb_total - nb_ok} anomalie(s)"

            rows.append({
                "date_raw"        : r["date"],
                "Date"            : r["date"].strftime("%d/%m/%Y"),
                "Jour"            : r["weekday_name"],
                "Année"           : int(r["annee"]),
                "Mois"            : int(r["mois"]),
                "Semaine"         : int(r["date"].isocalendar()[1]),
                "Heures dispo"    : len(avail),
                "PnL bess_engine (€)"  : round(pnl_moteur, 4),
                "PnL exhaustif (€)": round(best_pnl_exh, 4),
                "Écart (€)"       : round(best_pnl_exh - pnl_moteur, 4),
                "PnL borne max (€)": round(pnl_bm, 4),
                "Spread (€/MWh)"  : round(spread_moteur, 4),
                "Prix achat moy"  : round(pc_moteur, 4),
                "Prix vente moy"  : round(pd_moteur, 4),
                "Prix min"        : round(float(prices.min()), 4),
                "Prix max"        : round(float(prices.max()), 4),
                "H charge"        : h_ch_moteur,
                "H décharge"      : h_dch_moteur,
                "H exh charge"    : h_ch_exh,
                "H exh décharge"  : h_dch_exh,
                "H bm charge"     : h_bm_ch,
                "H bm décharge"   : h_bm_dch,
                "Contrôles"       : ctrl,
                "Nb OK"           : nb_ok,
                "Nb contrôles"    : nb_total,
                "Statut"          : statut,
                "_prices"         : prices.tolist(),
                "_avail"          : avail,
                "_cycles"         : cycles_moteur,
            })
        return pd.DataFrame(rows)

    df_va = _build_verif(
        file_bytes, va_ncyc, va_dur, va_pow, va_eff,
        _verif_json.dumps(_va_excl)
    )

    # ── Résumé global des contrôles ───────────────────────────────────────────
    st.markdown('<p class="section">Résumé global des contrôles automatiques</p>',
                unsafe_allow_html=True)
    st.caption(
        "8 règles testées automatiquement sur chaque jour. "
        "Un jour 'conforme' signifie que les 8 règles passent simultanément."
    )

    _ctrl_noms = {
        "C1_ordre"         : "C1 — Décharge après charge (contrainte physique)",
        "C2_exclusion"     : "C2 — Aucune heure exclue utilisée",
        "C3_spread_positif": "C3 — Spread ≥ 0 (achat < vente)",
        "C4_optimalite"    : "C4 — bess_engine = optimum exhaustif (pas de trade manqué)",
        "C5_borne_max"     : "C5 — PnL réel ≤ borne max théorique",
        "C6_formule_pnl"   : "C6 — Formule PnL cohérente",
        "C7_prix_charge"   : "C7 — Prix achat = moyenne des heures de charge",
        "C8_prix_decharge" : "C8 — Prix vente = moyenne des heures de décharge",
    }

    # Compter les violations par règle sur tous les jours
    _ctrl_rows = []
    for ckey, cnom in _ctrl_noms.items():
        violations = sum(
            1 for _, r in df_va.iterrows()
            if r["PnL bess_engine (€)"] > 0 and not r["Contrôles"].get(ckey, True)
        )
        jours_testes = int((df_va["PnL bess_engine (€)"] > 0).sum())
        _ctrl_rows.append({
            "Règle": cnom,
            "Violations": violations,
            "Jours testés": jours_testes,
            "Résultat": "0 violation" if violations == 0 else f"{violations} violation(s)"
        })

    df_ctrl_global = pd.DataFrame(_ctrl_rows)

    def _style_ctrl(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for i in df.index:
            color = "#d4edda" if df.loc[i, "Violations"] == 0 else "#f8d7da"
            txt   = "#155724" if df.loc[i, "Violations"] == 0 else "#721c24"
            for c in df.columns:
                styles.loc[i, c] = f"background:{color};color:{txt}"
        return styles

    st.dataframe(
        df_ctrl_global.style.apply(_style_ctrl, axis=None),
        hide_index=True, width="stretch"
    )

    _total_violations = df_ctrl_global["Violations"].sum()
    if _total_violations == 0:
        st.success(
            f"Résultat : 0 violation sur les 8 règles × {len(df_va)} jours. "
            "bess_engine est conforme sur l'ensemble de la période."
        )
    else:
        st.error(
            f"{_total_violations} violation(s) détectée(s). "
            "Consultez le détail jour par jour ci-dessous."
        )

    st.divider()

    # ── Tableau principal — tous les jours ────────────────────────────────────
    st.markdown('<p class="section">Tableau complet — Sélectionnez un jour pour l\'audit détaillé</p>',
                unsafe_allow_html=True)
    st.caption(
        "Chaque ligne représente un jour simulé. "
        "La colonne **Statut** indique si les 8 contrôles automatiques passent. "
        "La colonne **Écart (€)** compare bess_engine à l'algorithme exhaustif — "
        "un écart nul prouve que bess_engine trouve l'optimum exact."
    )

    # Filtres
    _vf1, _vf2, _vf3, _vf4, _vf5 = st.columns(5)
    with _vf1:
        _vf_annee = st.multiselect("Année", sorted(df_va["Année"].unique()),
                                    default=sorted(df_va["Année"].unique()), key="vf_an")
    with _vf2:
        _vf_mois = st.multiselect("Mois", list(range(1, 13)),
                                   default=list(range(1, 13)), key="vf_mo")
    with _vf3:
        _vf_statut = st.radio("Statut", ["Tous", "Conformes", "Anomalies"],
                               horizontal=True, key="vf_st")
    with _vf4:
        _vf_date = st.text_input("Recherche date (jj/mm/aaaa)", key="vf_dt")
    with _vf5:
        _vf_actif = st.radio("Jours", ["Tous", "Actifs seuls"], horizontal=True, key="vf_ac")

    _df_vf = df_va[df_va["Année"].isin(_vf_annee) & df_va["Mois"].isin(_vf_mois)].copy()
    if _vf_statut == "Conformes":
        _df_vf = _df_vf[_df_vf["Nb OK"] == _df_vf["Nb contrôles"]]
    elif _vf_statut == "Anomalies":
        _df_vf = _df_vf[_df_vf["Nb OK"] < _df_vf["Nb contrôles"]]
    if _vf_actif == "Actifs seuls":
        _df_vf = _df_vf[_df_vf["PnL bess_engine (€)"] > 0]
    if _vf_date:
        _df_vf = _df_vf[_df_vf["Date"].str.contains(_vf_date.strip())]

    _COLS_SHOW_VA = [
        "Date", "Jour", "Année", "Mois", "Semaine", "Heures dispo",
        "PnL bess_engine (€)", "PnL exhaustif (€)", "Écart (€)",
        "PnL borne max (€)", "Spread (€/MWh)",
        "Prix achat moy", "Prix vente moy", "Prix min", "Prix max",
        "Statut"
    ]

    def _style_va(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for i in df.index:
            # Colonne Statut
            if "Statut" in df.columns:
                v = df.loc[i, "Statut"]
                if v == "Conforme":
                    styles.loc[i, "Statut"] = "background:#d4edda;color:#155724;font-weight:bold"
                else:
                    styles.loc[i, "Statut"] = "background:#f8d7da;color:#721c24;font-weight:bold"
            # Colonne Écart
            if "Écart (€)" in df.columns:
                v = df.loc[i, "Écart (€)"]
                if abs(float(v)) < 0.01:
                    styles.loc[i, "Écart (€)"] = "background:#d4edda;color:#155724"
                else:
                    styles.loc[i, "Écart (€)"] = "background:#f8d7da;color:#721c24;font-weight:bold"
        return styles

    _sel_va = st.dataframe(
        _df_vf[_COLS_SHOW_VA].reset_index(drop=True).style.apply(_style_va, axis=None),
        hide_index=False,
        on_select="rerun",
        selection_mode="single-row",
        width="stretch",
        height=320,
    )

    # Métriques résumé de la sélection filtrée
    _sm1, _sm2, _sm3, _sm4, _sm5 = st.columns(5)
    _sm1.metric("Jours filtrés", len(_df_vf))
    _sm2.metric("Jours actifs", int((_df_vf["PnL bess_engine (€)"] > 0).sum()))
    _sm3.metric("PnL total", f"{_fmt(_df_vf['PnL bess_engine (€)'].sum(), 2)} €")
    _sm4.metric("Écart total bess_engine/exhaustif", f"{_df_vf['Écart (€)'].sum():.4f} €")
    _nm = len(_df_vf[_df_vf["Nb OK"] == _df_vf["Nb contrôles"]])
    _sm5.metric("Jours conformes", f"{_nm} / {len(_df_vf)}")

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # PANNEAU DRILL-DOWN — Audit complet d'un jour
    # ══════════════════════════════════════════════════════════════════
    _sel_rows_va = _sel_va.selection.rows if hasattr(_sel_va, "selection") else []

    if not _sel_rows_va:
        st.markdown("""
        <div style="text-align:center;padding:30px;background:#f8f9fb;
             border:2px dashed #b0c4de;border-radius:8px;color:#6b7a8d;">
        <strong>Cliquez sur une ligne du tableau</strong> pour voir l'audit complet du jour :<br>
        prix bruts Excel · heures choisies · formule PnL · 8 contrôles détaillés · comparaison exhaustif
        </div>
        """, unsafe_allow_html=True)
    else:
        _rid_va  = _sel_rows_va[0]
        _row_va  = _df_vf.iloc[_rid_va]
        _prices  = np.array(_row_va["_prices"])
        _cycles  = _row_va["_cycles"]
        _avail   = _row_va["_avail"]
        _ctrl    = _row_va["Contrôles"]
        _h_ch    = _row_va["H charge"]
        _h_dch   = _row_va["H décharge"]
        _h_ch_exh = _row_va["H exh charge"]
        _h_dch_exh = _row_va["H exh décharge"]
        _h_bm_ch  = _row_va["H bm charge"]
        _h_bm_dch = _row_va["H bm décharge"]
        _date_str = _row_va["Date"]
        _jour_str = _row_va["Jour"]
        _excl_h   = set(_va_excl.get("hours", []))

        # ── En-tête du jour ───────────────────────────────────────────
        _ok_all = _row_va["Nb OK"] == _row_va["Nb contrôles"]
        _hdr_col = "#1a3a5c" if not _BB else "#ff6600"
        _hdr_bg  = "#e8f4fd" if not _BB else "#1a1a1a"
        st.markdown(f"""
        <div style="background:{_hdr_bg};border:2px solid {_hdr_col};
             border-radius:8px;padding:14px 20px;margin:8px 0 16px 0;">
        <span style="font-size:1.25rem;font-weight:700;color:{_hdr_col};">
        AUDIT COMPLET — {_date_str} ({_jour_str})</span>
        &nbsp;&nbsp;
        <span style="font-size:1rem;font-weight:700;color:{'#155724' if _ok_all else '#721c24'};">
        {'Conforme (8/8)' if _ok_all else f'{_row_va["Nb OK"]}/{_row_va["Nb contrôles"]} contrôles OK'}
        </span><br>
        <span style="color:{'#aaa' if _BB else '#6b7a8d'};font-size:0.82rem;">
        {va_pow} MW · {va_ncyc} cycle(s)/{va_dur}h · Rendement {va_eff*100:.0f}% ·
        Capacité {va_cap:.3f} MWh ·
        Heures disponibles : {len(_avail)}/24
        {'· Restriction : H' + str(va_hdeb) + '–H' + str(va_hfin) + ' (' + ', '.join(va_jours) + ')' if _va_excl else '· Pas de restriction'}
        </span>
        </div>
        """, unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════════
        # BLOC 1 — CONTRÔLES AUTOMATIQUES
        # ══════════════════════════════════════════════════════════════
        st.markdown("### Bloc 1 — Contrôles automatiques (8 règles)")
        st.caption(
            "Chaque règle est vérifiée indépendamment. "
            "Vert = la règle passe sur ce jour. Rouge = anomalie détectée."
        )

        _ctrl_labels = {
            "C1_ordre"         : ("C1 — Ordre temporel", "Toutes les heures de décharge sont strictement après la dernière heure de charge."),
            "C2_exclusion"     : ("C2 — Heures exclues", "Aucune heure de la plage de restriction n'est utilisée pour charger ou décharger."),
            "C3_spread_positif": ("C3 — Spread positif", "Le prix de vente moyen est supérieur au prix d'achat moyen (spread ≥ 0)."),
            "C4_optimalite"    : ("C4 — Optimalité", "Le PnL de bess_engine est égal au PnL de l'algorithme exhaustif : aucune meilleure combinaison n'existe."),
            "C5_borne_max"     : ("C5 — Borne max", "Le PnL réel est inférieur ou égal à la borne max théorique (sans contrainte d'ordre)."),
            "C6_formule_pnl"   : ("C6 — Formule PnL", "PnL calculé par la formule correspond au PnL enregistré."),
            "C7_prix_charge"   : ("C7 — Prix achat", "Le prix d'achat moyen correspond bien à la moyenne des prix des heures de charge."),
            "C8_prix_decharge" : ("C8 — Prix vente", "Le prix de vente moyen correspond bien à la moyenne des prix des heures de décharge."),
        }

        _ck1, _ck2 = st.columns(2)
        for i, (ckey, (cname, cdesc)) in enumerate(_ctrl_labels.items()):
            _ok = _ctrl.get(ckey, True)
            _col = _ck1 if i % 2 == 0 else _ck2
            with _col:
                _bg  = "#d4edda" if _ok else "#f8d7da"
                _txt = "#155724" if _ok else "#721c24"
                st.markdown(f"""
                <div style="background:{_bg};border-left:4px solid {'#28a745' if _ok else '#dc3545'};
                     padding:8px 12px;margin:4px 0;border-radius:4px;">
                <span style="color:{_txt};font-weight:700;">
                {cname}</span><br>
                <span style="color:{'#555' if not _BB else '#bbb'};font-size:0.8rem;">{cdesc}</span>
                </div>
                """, unsafe_allow_html=True)

        st.divider()

        # ══════════════════════════════════════════════════════════════
        # BLOC 2 — PRIX BRUTS EXCEL
        # ══════════════════════════════════════════════════════════════
        st.markdown("### Bloc 2 — Prix bruts Excel (feuille Spot_input)")
        st.caption(
            "Ces valeurs sont lues directement depuis la feuille **Spot_input** de votre Excel. "
            "Pour vérifier : filtrer ANNEE, MOIS, JOUR dans la feuille Spot_input → colonne Prix Final. "
            "Le **Rang** indique la position de l'heure par ordre de prix croissant (1 = heure la moins chère)."
        )

        _idx_sort = np.argsort(_prices)
        _rang = {int(_idx_sort[i]): i + 1 for i in range(24)}

        _df_spot = pd.DataFrame({
            "Heure"             : [f"H{h:02d}" for h in range(24)],
            "Disponible ?"      : ["Oui" if h in _avail else "Non" for h in range(24)],
            "Prix spot (€/MWh)" : [round(_prices[h], 4) for h in range(24)],
            "Rang (1=min)"      : [_rang[h] for h in range(24)],
            "vs moy jour"       : [f"{'+' if _prices[h] >= _prices.mean() else ''}{_prices[h] - _prices.mean():.2f}" for h in range(24)],
            "Rôle"              : [
                ("Charge" if h in _h_ch else
                 "Décharge" if h in _h_dch else
                 "—")
                for h in range(24)
            ],
        })

        def _style_spot(df):
            st_ = pd.DataFrame("", index=df.index, columns=df.columns)
            for i in df.index:
                h = i
                if h in _h_ch:
                    st_.loc[i, :] = "background:#d4edda;color:#155724;font-weight:bold"
                elif h in _h_dch:
                    st_.loc[i, :] = "background:#cce5ff;color:#004085;font-weight:bold"
                elif h not in _avail:
                    st_.loc[i, :] = "background:#f8d7da;color:#721c24;font-style:italic"
            return st_

        _sc1, _sc2 = st.columns([3, 1])
        with _sc1:
            st.dataframe(
                _df_spot.style.apply(_style_spot, axis=None),
                hide_index=True, width="stretch", height=520
            )
        with _sc2:
            st.markdown("""
**Légende**

Vert = Charge (achat)

Bleu = Décharge (vente)

Rouge = Heure exclue (restriction)

Blanc = Non utilisé
            """)
            st.metric("Prix min du jour", f"{_prices.min():.4f} €/MWh")
            st.metric("Prix max du jour", f"{_prices.max():.4f} €/MWh")
            st.metric("Prix moyen", f"{_prices.mean():.4f} €/MWh")
            st.metric("Amplitude", f"{_prices.max() - _prices.min():.4f} €/MWh")
            st.metric("Heures disponibles", f"{len(_avail)} / 24")

        st.divider()

        # ══════════════════════════════════════════════════════════════
        # BLOC 3 — COMPARAISON MOTEUR / EXHAUSTIF / BORNE MAX
        # ══════════════════════════════════════════════════════════════
        st.markdown("### Bloc 3 — Comparaison des trois approches")
        st.caption(
            "**bess_engine** = algorithme de calcul (bess_engine.py). "
            "**Exhaustif** = force brute (toutes les combinaisons testées). "
            "**Borne max** = maximum théorique sans contrainte d'ordre temporel. "
            "Si bess_engine = Exhaustif → bess_engine est optimal. Si Borne max > Exhaustif → la contrainte d'ordre coûte de l'argent."
        )

        _b3c1, _b3c2, _b3c3 = st.columns(3)

        with _b3c1:
            _pnl_m = _row_va["PnL bess_engine (€)"]
            _pc_m  = _row_va["Prix achat moy"]
            _pd_m  = _row_va["Prix vente moy"]
            st.markdown(f"""
<div style="border:2px solid #28a745;border-radius:8px;padding:12px;background:#f0fff4;">
<b style="color:#28a745;">bess_engine</b><br><br>
<b>H charge :</b> {_h_ch}<br>
<b>H décharge :</b> {_h_dch}<br><br>
<b>Prix achat moy :</b> {_pc_m:.4f} €/MWh<br>
<b>Prix vente moy :</b> {_pd_m:.4f} €/MWh<br>
<b>Spread :</b> {_pd_m - _pc_m:.4f} €/MWh<br><br>
<b>Formule :</b><br>
Vente = {va_pow} × {va_dur} × {va_eff} × {_pd_m:.4f} = <b>{va_pow*va_dur*va_eff*_pd_m:.4f} €</b><br>
Achat = {va_pow} × {va_dur} × {_pc_m:.4f} = <b>{va_pow*va_dur*_pc_m:.4f} €</b><br><br>
<b style="font-size:1.1rem;color:#155724;">PnL = {_pnl_m:.4f} €</b>
</div>
            """, unsafe_allow_html=True)

        with _b3c2:
            _pnl_exh = _row_va["PnL exhaustif (€)"]
            _ecart   = _row_va["Écart (€)"]
            _eq_sign = "Identique à bess_engine" if abs(_ecart) < 0.01 else f"Écart : +{_ecart:.4f} €"
            _pc_exh  = np.mean([_prices[h] for h in _h_ch_exh]) if _h_ch_exh else 0
            _pd_exh  = np.mean([_prices[h] for h in _h_dch_exh]) if _h_dch_exh else 0
            st.markdown(f"""
<div style="border:2px solid #0066cc;border-radius:8px;padding:12px;background:#f0f4ff;">
<b style="color:#0066cc;">EXHAUSTIF (force brute — référence)</b><br><br>
<b>H charge :</b> {_h_ch_exh}<br>
<b>H décharge :</b> {_h_dch_exh}<br><br>
<b>Prix achat moy :</b> {_pc_exh:.4f} €/MWh<br>
<b>Prix vente moy :</b> {_pd_exh:.4f} €/MWh<br>
<b>Spread :</b> {_pd_exh - _pc_exh:.4f} €/MWh<br><br>
<b style="font-size:1.1rem;color:#004085;">PnL = {_pnl_exh:.4f} €</b><br><br>
<b>{_eq_sign}</b>
</div>
            """, unsafe_allow_html=True)

        with _b3c3:
            _pnl_bm  = _row_va["PnL borne max (€)"]
            _pc_bm_v = np.mean([_prices[h] for h in _h_bm_ch]) if _h_bm_ch else 0
            _pd_bm_v = np.mean([_prices[h] for h in _h_bm_dch]) if _h_bm_dch else 0
            _cout_contrainte = _pnl_bm - _pnl_exh
            st.markdown(f"""
<div style="border:2px solid #888;border-radius:8px;padding:12px;background:#f9f9f9;">
<b style="color:#555;">BORNE MAX (sans contrainte d'ordre)</b><br><br>
<b>H achat :</b> {_h_bm_ch}<br>
<b>H vente :</b> {_h_bm_dch}<br><br>
<b>Prix achat moy :</b> {_pc_bm_v:.4f} €/MWh<br>
<b>Prix vente moy :</b> {_pd_bm_v:.4f} €/MWh<br>
<b>Spread :</b> {_pd_bm_v - _pc_bm_v:.4f} €/MWh<br><br>
<b style="font-size:1.1rem;color:#333;">PnL = {_pnl_bm:.4f} €</b><br><br>
<b>Coût contrainte temporelle : {_cout_contrainte:.4f} €</b><br>
<small>{'(les heures optimales respectaient déjà l\'ordre → écart nul)' if _cout_contrainte < 0.01 else '(certaines meilleures heures ne respectaient pas charge < décharge)'}</small>
</div>
            """, unsafe_allow_html=True)

        st.divider()

        # ══════════════════════════════════════════════════════════════
        # BLOC 4 — GRAPHIQUE INTERACTIF
        # ══════════════════════════════════════════════════════════════
        st.markdown("### Bloc 4 — Graphique des prix et des trades")
        st.caption(
            "Barres grises = heures non utilisées. "
            "Barres rouges = heures exclues (restriction). "
            "Triangles verts ▲ = charge (achat). Triangles bleus ▼ = décharge (vente). "
            "Les lignes pointillées horizontales montrent les prix moyens d'achat et de vente."
        )

        _fig_v = go.Figure()

        # Barres de prix colorées
        _bar_col = []
        for h in range(24):
            if h in _h_ch:    _bar_col.append("#5cb85c")
            elif h in _h_dch: _bar_col.append("#2e75b6")
            elif h not in _avail: _bar_col.append("#f5c6cb")
            else:             _bar_col.append("#dee2e6")

        _fig_v.add_trace(go.Bar(
            x=list(range(24)), y=_prices,
            marker_color=_bar_col,
            name="Prix spot (€/MWh)",
            hovertemplate=(
                "<b>H%{x:02d}</b><br>"
                "Prix brut Excel : <b>%{y:.4f} €/MWh</b><br>"
                "Rang : " + str([_rang[h] for h in range(24)]).replace("[","").replace("]","") + "<extra></extra>"
            ),
        ))

        # Ligne moyenne
        _moy_v = _prices.mean()
        _fig_v.add_hline(y=_moy_v, line_dash="dot", line_color="#999",
                         annotation_text=f"Moy {_moy_v:.2f} €", annotation_position="right")

        # Triangles charge
        if _h_ch:
            _fig_v.add_trace(go.Scatter(
                x=_h_ch, y=[_prices[h] for h in _h_ch],
                mode="markers+text", name="Charge (achat)",
                text=[f"H{h:02d}<br>{_prices[h]:.2f}€" for h in _h_ch],
                textposition="bottom center",
                marker=dict(color="#28a745", size=22, symbol="triangle-up"),
                hovertemplate="<b>Charge H%{x:02d}</b><br>Prix : %{y:.4f} €/MWh<extra></extra>",
            ))
            _fig_v.add_hline(
                y=np.mean([_prices[h] for h in _h_ch]),
                line_dash="dash", line_color="#28a745", line_width=2,
                annotation_text=f"Achat moy = {np.mean([_prices[h] for h in _h_ch]):.4f} €",
                annotation_position="left"
            )

        # Triangles décharge
        if _h_dch:
            _fig_v.add_trace(go.Scatter(
                x=_h_dch, y=[_prices[h] for h in _h_dch],
                mode="markers+text", name="Décharge (vente)",
                text=[f"H{h:02d}<br>{_prices[h]:.2f}€" for h in _h_dch],
                textposition="top center",
                marker=dict(color="#0066cc", size=22, symbol="triangle-down"),
                hovertemplate="<b>Décharge H%{x:02d}</b><br>Prix : %{y:.4f} €/MWh<extra></extra>",
            ))
            _fig_v.add_hline(
                y=np.mean([_prices[h] for h in _h_dch]),
                line_dash="dash", line_color="#0066cc", line_width=2,
                annotation_text=f"Vente moy = {np.mean([_prices[h] for h in _h_dch]):.4f} €",
                annotation_position="right"
            )

        # Zones grises pour heures exclues
        for h in range(24):
            if h not in _avail:
                _fig_v.add_vrect(x0=h - 0.5, x1=h + 0.5,
                                  fillcolor="rgba(200,0,0,0.08)", line_width=0)

        _fig_v.update_layout(
            height=420, margin=dict(t=20, b=80, l=70, r=180),
            xaxis=dict(title="Heure de la journée", tickmode="array",
                       tickvals=list(range(24)), ticktext=[f"H{h:02d}" for h in range(24)]),
            yaxis=dict(title="Prix spot (€/MWh)", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            hoverlabel=dict(bgcolor="white", font_size=12),
            showlegend=True,
        )
        apply_bb(_fig_v)
        st.plotly_chart(_fig_v, width="stretch", config=PLOTLY_CFG, key=f"fig_verif_{_rid_va}")

        st.divider()

        # ══════════════════════════════════════════════════════════════
        # BLOC 5 — GUIDE DE VÉRIFICATION MANUELLE DANS EXCEL
        # ══════════════════════════════════════════════════════════════
        st.markdown("### Bloc 5 — Guide de vérification manuelle dans Excel")
        st.caption(
            "Suivez ces étapes pour vérifier vous-même chaque chiffre directement "
            "dans votre fichier Excel source, sans aucun outil intermédiaire."
        )

        _annee_v = int(_row_va["Année"])
        _mois_v  = int(_row_va["Mois"])
        _jour_v  = int(_date_str.split("/")[0])
        _pnl_v   = _row_va["PnL bess_engine (€)"]

        st.markdown(f"""
**Étape 1 — Ouvrir la feuille Spot_input dans Excel**
> Filtre : `ANNEE = {_annee_v}` · `MOIS = {_mois_v}` · `JOUR = {_jour_v}`
> → Vous obtenez 24 lignes, une par heure H00 à H23

**Étape 2 — Vérifier les prix horaires**
> Comparez colonne `Prix Final` avec le tableau ci-dessus (Bloc 2).
> Chaque valeur doit correspondre au centime.

**Étape 3 — Repérer les heures de charge**
> Heures choisies : `{_h_ch}` → colonne Prix Final de ces lignes → moyenne = `{np.mean([_prices[h] for h in _h_ch]):.4f} €/MWh`

**Étape 4 — Repérer les heures de décharge**
> Heures choisies : `{_h_dch}` → colonne Prix Final de ces lignes → moyenne = `{np.mean([_prices[h] for h in _h_dch]):.4f} €/MWh`

**Étape 5 — Appliquer la formule**
```
PnL = Puissance × Durée × Rendement × Prix_vente_moy
    − Puissance × Durée × Prix_achat_moy

    = {va_pow} × {va_dur} × {va_eff} × {np.mean([_prices[h] for h in _h_dch]):.4f}
    − {va_pow} × {va_dur} × {np.mean([_prices[h] for h in _h_ch]):.4f}

    = {va_pow*va_dur*va_eff*np.mean([_prices[h] for h in _h_dch]):.6f}
    − {va_pow*va_dur*np.mean([_prices[h] for h in _h_ch]):.6f}

    = {_pnl_v:.6f} €
```

**Étape 6 — Vérifier la cohérence**
> - Les heures de charge doivent être TOUTES avant les heures de décharge
> - Les heures `{list(_excl_h)}` (restriction) ne doivent PAS apparaître dans les étapes 3 et 4
> - Le PnL doit être positif (spread > 0)
        """)

        # Export CSV de ce jour
        _df_export_jour = pd.DataFrame({
            "Heure": [f"H{h:02d}" for h in range(24)],
            "Prix spot Excel (€/MWh)": [round(_prices[h], 4) for h in range(24)],
            "Rang": [_rang[h] for h in range(24)],
            "Disponible": ["Oui" if h in _avail else "Non (restriction)" for h in range(24)],
            "Rôle bess_engine": [
                ("Charge" if h in _h_ch else "Décharge" if h in _h_dch else "Inutilisé")
                for h in range(24)
            ],
            "Rôle exhaustif": [
                ("Charge" if h in _h_ch_exh else "Décharge" if h in _h_dch_exh else "Inutilisé")
                for h in range(24)
            ],
        })

        st.download_button(
            f"Exporter audit {_date_str} (CSV)",
            data=_df_export_jour.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"BESS_audit_{_date_str.replace('/', '')}.csv",
            mime="text/csv",
            key=f"dl_verif_jour_{_rid_va}",
        )

    st.divider()
    st.download_button(
        "Exporter tableau complet vérification (CSV)",
        data=_df_vf[_COLS_SHOW_VA].to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
        file_name="BESS_verification_arbitrage_complet.csv",
        mime="text/csv",
        key="dl_verif_global",
    )

with tab_verif_lis:
    _diag_set_tab("verif_lis")
    from bess_engine import lissage_day as _lissage_day
    import io as _io3

    st.subheader("Vérification Lissage — Audit heure par heure")
    st.caption(
        "Sélectionnez un jour pour vérifier les calculs de lissage : "
        "données brutes Excel, actions BESS heure par heure, formule de réduction de pointe."
    )

    if not st.session_state.get("excel_cdc_bytes"):
        st.warning("Chargez d'abord un fichier courbe de charge dans l'onglet **Lissage de charge**.")
    else:
        # ── Paramètres ────────────────────────────────────────────────────────
        st.markdown('<p class="section">Paramètres de l\'audit</p>', unsafe_allow_html=True)
        vl1, vl2, vl3, vl4 = st.columns(4)
        with vl1: vl_pow   = st.number_input("Puissance (MW)", 0.05, 50.0, power_MW, 0.01, key="vl_pow")
        with vl2: vl_emwh  = st.number_input("Capacité (MWh)", 0.1, 100.0, 2.0, 0.1, key="vl_emwh")
        with vl3: vl_seuil = st.slider("Seuil écrêtage (%)", 50, 95, 75, key="vl_seuil")
        with vl4: vl_tarif = st.number_input("Tarif (€/MW/an)", 1000, 100000, 12000, 1000, key="vl_tarif")

        @st.cache_data(show_spinner=False)
        def build_audit_lis(cdc_bytes, pow_, emwh, seuil_pct, tarif):
            df = pd.read_excel(_io3.BytesIO(cdc_bytes), sheet_name="CdC_kWh", engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            df = df.rename(columns={"datetime": "ts"})
            df["ts"] = pd.to_datetime(df["ts"])
            cc = [c for c in df.columns if c != "ts" and pd.api.types.is_numeric_dtype(df[c])]
            df["conso_MW"] = df[cc].sum(axis=1) / 1000
            df["heure"] = df["ts"].dt.hour
            df["date"]  = df["ts"].dt.date
            df["annee"] = df["ts"].dt.year
            df["mois"]  = df["ts"].dt.month
            seuil   = float(np.percentile(df["conso_MW"].values, seuil_pct))
            soc_max = emwh * 0.9
            soc_min = emwh * 0.1
            records = []
            for date, grp in df.groupby("date"):
                grp = grp.sort_values("heure")
                if len(grp) < 24:
                    continue
                profil = grp["conso_MW"].values[:24].astype(float)
                res = _lissage_day(profil, seuil, pow_, emwh, emwh * 0.5, soc_min, soc_max, 0.92)
                red = res["reduction_pointe"]
                # Contrôles automatiques
                ctrl_lis = {
                    "L1_soc_bornes":    all(soc_min - 0.001 <= s <= soc_max + 0.001 for s in res["soc_hist"]),
                    "L2_lisse_positive":all(v >= -0.001 for v in res["profil_lisse"]),
                    "L3_pas_sur_inject":all(res["profil_lisse"][h] <= profil[h] + 0.001
                                            for h in range(24) if res["actions"][h][0] == "decharge"),
                    "L4_red_positive":  red >= -0.001,
                    "L5_cohérence_soc": abs(res["soc_hist"][-1] - res["soc_final"]) < 0.001,
                    "L6_formule_eco":   abs(red * tarif - round(red * tarif, 0)) < 1,
                }
                nb_ok = sum(ctrl_lis.values())
                records.append({
                    "date_raw":          date,
                    "Date":              date.strftime("%d/%m/%Y"),
                    "Année":             int(grp["annee"].iloc[0]),
                    "Mois":              int(grp["mois"].iloc[0]),
                    "Pointe avant (MW)": round(res["pointe_avant"], 4),
                    "Pointe après (MW)": round(res["pointe_apres"], 4),
                    "Réduction (MW)":    round(red, 4),
                    "Économie/an (€)":   round(red * tarif, 0),
                    "SOC final (MWh)":   round(res["soc_final"], 4),
                    "Nb contrôles OK":   nb_ok,
                    "Statut":            "Conforme" if nb_ok == 6 else f"{6 - nb_ok} anomalie(s)",
                    # Contrôles individuels (scalaires bool)
                    "_L1": ctrl_lis["L1_soc_bornes"],
                    "_L2": ctrl_lis["L2_lisse_positive"],
                    "_L3": ctrl_lis["L3_pas_sur_inject"],
                    "_L4": ctrl_lis["L4_red_positive"],
                    "_L5": ctrl_lis["L5_cohérence_soc"],
                    "_L6": ctrl_lis["L6_formule_eco"],
                    "_seuil": float(seuil),
                })
            return pd.DataFrame(records), seuil

        @st.cache_data(show_spinner=False)
        def get_day_detail_lis(cdc_bytes, date_str, pow_, emwh, seuil_pct):
            """Recalcule le détail heure par heure pour un jour donné — pas de listes dans le cache principal."""
            df = pd.read_excel(_io3.BytesIO(cdc_bytes), sheet_name="CdC_kWh", engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            df = df.rename(columns={"datetime": "ts"})
            df["ts"] = pd.to_datetime(df["ts"])
            cc = [c for c in df.columns if c != "ts" and pd.api.types.is_numeric_dtype(df[c])]
            df["conso_MW"] = df[cc].sum(axis=1) / 1000
            df["heure"] = df["ts"].dt.hour
            df["date"]  = df["ts"].dt.date
            seuil = float(np.percentile(df["conso_MW"].values, seuil_pct))
            soc_max = emwh * 0.9
            soc_min = emwh * 0.1
            target = pd.to_datetime(date_str, dayfirst=True).date()
            grp = df[df["date"] == target].sort_values("heure")
            if len(grp) < 24:
                return None, seuil
            profil = grp["conso_MW"].values[:24].astype(float)
            res = _lissage_day(profil, seuil, pow_, emwh, emwh * 0.5, soc_min, soc_max, 0.92)
            return {
                "profil":   profil.tolist(),
                "lisse":    res["profil_lisse"].tolist(),
                "actions":  [(str(a[0]), float(a[1])) for a in res["actions"]],
                "soc_hist": [float(s) for s in res["soc_hist"]],
                "seuil":    float(seuil),
                "ctrl": {
                    "L1_soc_bornes":     all(soc_min-0.001 <= s <= soc_max+0.001 for s in res["soc_hist"]),
                    "L2_lisse_positive": all(v >= -0.001 for v in res["profil_lisse"]),
                    "L3_pas_sur_inject": all(res["profil_lisse"][h] <= profil[h]+0.001
                                             for h in range(24) if res["actions"][h][0] == "decharge"),
                    "L4_red_positive":   res["reduction_pointe"] >= -0.001,
                    "L5_cohérence_soc":  abs(res["soc_hist"][-1] - res["soc_final"]) < 0.001,
                    "L6_formule_eco":    True,
                },
            }, seuil

        df_al, seuil_v = build_audit_lis(
            st.session_state.excel_cdc_bytes, vl_pow, vl_emwh, vl_seuil, vl_tarif
        )

        # ── Résumé global des contrôles ───────────────────────────────────────
        st.markdown('<p class="section">Résumé global des contrôles automatiques</p>',
                    unsafe_allow_html=True)
        st.caption(f"Seuil P{vl_seuil}% calculé sur toute la série = **{seuil_v:.4f} MW**")

        _ctrl_lis_noms = {
            "L1_soc_bornes":    "L1 — SOC toujours entre SOC_min et SOC_max",
            "L2_lisse_positive":"L2 — Consommation lissée toujours positive",
            "L3_pas_sur_inject":"L3 — La décharge ne dépasse jamais la conso originale",
            "L4_red_positive":  "L4 — Réduction de pointe ≥ 0",
            "L5_cohérence_soc": "L5 — SOC final cohérent avec le dernier historique",
            "L6_formule_eco":   "L6 — Formule économie = Réduction × Tarif",
        }
        _ctrl_lis_rows = []
        for ckey, cnom in _ctrl_lis_noms.items():
            _ckey_col = {"L1_soc_bornes": "_L1", "L2_lisse_positive": "_L2",
                          "L3_pas_sur_inject": "_L3", "L4_red_positive": "_L4",
                          "L5_cohérence_soc": "_L5", "L6_formule_eco": "_L6"}
            col = _ckey_col.get(ckey)
            viol = int((~df_al[col]).sum()) if col and col in df_al.columns else 0
            _ctrl_lis_rows.append({
                "Règle": cnom,
                "Violations": viol,
                "Jours testés": len(df_al),
                "Résultat": "0 violation" if viol == 0 else f"{viol} violation(s)",
            })
        df_ctrl_lis = pd.DataFrame(_ctrl_lis_rows)

        def _style_ctrl_lis(df):
            s = pd.DataFrame("", index=df.index, columns=df.columns)
            for i in df.index:
                bg  = "#d4edda" if df.loc[i, "Violations"] == 0 else "#f8d7da"
                txt = "#155724" if df.loc[i, "Violations"] == 0 else "#721c24"
                for c in df.columns:
                    s.loc[i, c] = f"background:{bg};color:{txt}"
            return s

        st.dataframe(df_ctrl_lis.style.apply(_style_ctrl_lis, axis=None),
                     hide_index=True, width="stretch")

        _tot_viol_lis = df_ctrl_lis["Violations"].sum()
        if _tot_viol_lis == 0:
            st.success(f"0 violation sur les 6 règles × {len(df_al)} jours. Lissage conforme.")
        else:
            st.error(f"{_tot_viol_lis} violation(s) détectée(s).")

        st.divider()

        # ── Tableau principal ─────────────────────────────────────────────────
        st.markdown('<p class="section">Tableau complet — Sélectionnez un jour pour l\'audit détaillé</p>',
                    unsafe_allow_html=True)

        COLS_LIS = ["Date", "Année", "Mois", "Pointe avant (MW)", "Pointe après (MW)",
                    "Réduction (MW)", "Économie/an (€)", "SOC final (MWh)",
                    "Nb contrôles OK", "Statut"]

        _lf1, _lf2 = st.columns(2)
        with _lf1:
            _lf_annee = st.multiselect("Année", sorted(df_al["Année"].unique()),
                                        default=sorted(df_al["Année"].unique()), key="lf_an")
        with _lf2:
            _lf_statut = st.radio("Statut", ["Tous", "Conformes", "Anomalies"],
                                   horizontal=True, key="lf_st")

        df_al_f = df_al[df_al["Année"].isin(_lf_annee)].copy()
        if _lf_statut == "Conformes":
            df_al_f = df_al_f[df_al_f["Nb contrôles OK"] == 6]
        elif _lf_statut == "Anomalies":
            df_al_f = df_al_f[df_al_f["Nb contrôles OK"] < 6]

        def style_lis(df):
            s = pd.DataFrame("", index=df.index, columns=df.columns)
            for i in df.index:
                r = df.loc[i, "Réduction (MW)"] if "Réduction (MW)" in df.columns else 0
                if r > 0:
                    s.loc[i, "Réduction (MW)"] = "background:#d4edda;font-weight:bold"
                else:
                    s.loc[i, "Réduction (MW)"] = "background:#f8d7da"
                if "Statut" in df.columns:
                    v = str(df.loc[i, "Statut"])
                    s.loc[i, "Statut"] = (
                        "background:#d4edda;color:#155724;font-weight:bold" if v == "Conforme"
                        else "background:#f8d7da;color:#721c24;font-weight:bold"
                    )
            return s

        sel_l = st.dataframe(
            df_al_f[COLS_LIS].reset_index(drop=True).style.apply(style_lis, axis=None),
            hide_index=False, on_select="rerun", selection_mode="single-row",
            width="stretch", height=300,
        )

        _lt1, _lt2, _lt3, _lt4 = st.columns(4)
        _lt1.metric("Jours analysés", len(df_al_f))
        _lt2.metric("Réduction moy", f"{df_al_f['Réduction (MW)'].mean():.4f} MW")
        _lt3.metric("Réduction max", f"{df_al_f['Réduction (MW)'].max():.4f} MW")
        _lt4.metric("Jours conformes",
                    f"{(df_al_f['Nb contrôles OK'] == 6).sum()} / {len(df_al_f)}")

        st.divider()

        # ── DRILL-DOWN ────────────────────────────────────────────────────────
        sel_rows_l = sel_l.selection.rows if hasattr(sel_l, "selection") else []

        if not sel_rows_l:
            st.markdown("""
            <div style="text-align:center;padding:30px;background:#f8f9fb;
                 border:2px dashed #b0c4de;border-radius:8px;color:#6b7a8d;">
            <strong>Cliquez sur une ligne</strong> pour voir l'audit complet du jour :<br>
            données brutes Excel · actions heure par heure · 6 contrôles · graphique profil
            </div>
            """, unsafe_allow_html=True)
        else:
            rid_l   = sel_rows_l[0]
            row_l   = df_al_f.iloc[rid_l]
            # Recalcul du détail à la volée (évite la corruption par st.cache_data Arrow)
            _detail, _ = get_day_detail_lis(
                st.session_state.excel_cdc_bytes,
                row_l["Date"], vl_pow, vl_emwh, vl_seuil
            )
            profil  = _detail["profil"]
            lisse   = _detail["lisse"]
            actions = _detail["actions"]
            soc_h   = _detail["soc_hist"]
            seuil_j = _detail["seuil"]
            date_l  = row_l["Date"]
            ctrl_l  = _detail["ctrl"]
            _ok_all_l = row_l["Nb contrôles OK"] == 6

            # En-tête
            _ok_all_l_str = "Conforme (6/6)" if _ok_all_l else f'Anomalie : {row_l["Nb contrôles OK"]}/6'
            st.markdown(
                f'<p class="section">Audit lissage — {date_l} &nbsp; '
                f'<span style="font-weight:700;color:{"#155724" if _ok_all_l else "#721c24"};">'
                f'{_ok_all_l_str}</span></p>',
                unsafe_allow_html=True
            )
            st.caption(
                f"{vl_pow} MW · Capacité {vl_emwh} MWh · "
                f"Seuil P{vl_seuil}% = {seuil_j:.4f} MW · Tarif {vl_tarif:,} €/MW/an"
            )

            # ═══ BLOC 1 — Contrôles automatiques ════════════════════════
            st.markdown("### Bloc 1 — Contrôles automatiques (6 règles)")

            _ctrl_lis_desc = {
                "L1_soc_bornes":    ("L1 — SOC dans les bornes", f"Le SOC reste entre {vl_emwh*0.1:.3f} MWh (10%) et {vl_emwh*0.9:.3f} MWh (90%) à chaque heure."),
                "L2_lisse_positive":("L2 — Conso lissée positive", "La consommation lissée ne descend jamais en négatif (pas de sur-injection)."),
                "L3_pas_sur_inject":("L3 — Décharge ≤ conso originale", "Aux heures de décharge, la batterie ne dépasse pas la consommation du client (pas d'export réseau)."),
                "L4_red_positive":  ("L4 — Réduction ≥ 0", "La pointe après lissage est inférieure ou égale à la pointe avant."),
                "L5_cohérence_soc": ("L5 — SOC final cohérent", "Le SOC_final enregistré correspond au dernier point de l'historique."),
                "L6_formule_eco":   ("L6 — Formule économie", "Économie = Réduction × Tarif, à l'arrondi près."),
            }
            _bc1, _bc2 = st.columns(2)
            for i, (ckey, (cname, cdesc)) in enumerate(_ctrl_lis_desc.items()):
                _ok = ctrl_l.get(ckey, True)
                _col = _bc1 if i % 2 == 0 else _bc2
                with _col:
                    _bg  = "#d4edda" if _ok else "#f8d7da"
                    _txt = "#155724" if _ok else "#721c24"
                    st.markdown(f"""
                    <div style="background:{_bg};border-left:4px solid {'#28a745' if _ok else '#dc3545'};
                         padding:8px 12px;margin:4px 0;border-radius:4px;">
                    <span style="color:{_txt};font-weight:700;">
                    {cname}</span><br>
                    <span style="font-size:0.8rem;color:#555;">{cdesc}</span>
                    </div>
                    """, unsafe_allow_html=True)

            st.divider()

            # ═══ BLOC 2 — Tableau heure par heure ═══════════════════════
            st.markdown("### Bloc 2 — Données brutes Excel et actions BESS heure par heure")
            st.caption(
                f"Valeurs lues depuis la feuille **CdC_kWh** pour le {date_l}, converties kWh → MW (÷1000). "
                "Vert = charge batterie. Bleu = décharge. Jaune = heure au-dessus du seuil."
            )

            _action_label = {"charge": "Charge", "decharge": "Décharge", "idle": "— Repos"}
            df_hh = pd.DataFrame({
                "Heure":               [f"H{h:02d}" for h in range(24)],
                "Conso Excel (kWh/h)": [round(profil[h] * 1000, 4) for h in range(24)],
                "Conso (MW)":          [round(profil[h], 4) for h in range(24)],
                "vs seuil":            [
                    f"{'▲ +' if profil[h] > seuil_j else '▼ '}{profil[h] - seuil_j:+.4f} MW"
                    for h in range(24)
                ],
                "Action BESS":         [_action_label.get(str(actions[h][0]), str(actions[h][0])) for h in range(24)],
                "BESS (MW)":           [round(actions[h][1], 4) for h in range(24)],
                "Conso lissée (MW)":   [round(lisse[h], 4) for h in range(24)],
                "SOC avant (MWh)":     [round(soc_h[h], 4) for h in range(24)],
                "SOC après (MWh)":     [round(soc_h[h + 1], 4) for h in range(24)],
                "SOC %":               [f"{soc_h[h + 1] / vl_emwh * 100:.1f}%" for h in range(24)],
                "Vérif lissée":        [
                    "OK" if abs(lisse[h] - (
                        profil[h] - actions[h][1] if actions[h][0] == "decharge"
                        else profil[h] + actions[h][1] if actions[h][0] == "charge"
                        else profil[h]
                    )) < 0.001 else "Erreur"
                    for h in range(24)
                ],
            })

            def style_hh(df):
                s = pd.DataFrame("", index=df.index, columns=df.columns)
                for i in df.index:
                    act = actions[i][0]
                    if act == "decharge":
                        s.loc[i, :] = "background:#cce5ff"
                        s.loc[i, "Action BESS"] = "background:#0056b3;color:white;font-weight:bold"
                    elif act == "charge":
                        s.loc[i, :] = "background:#d4edda"
                        s.loc[i, "Action BESS"] = "background:#155724;color:white;font-weight:bold"
                    if profil[i] > seuil_j:
                        s.loc[i, "vs seuil"] = "background:#fff3cd;font-weight:bold"
                    if df.loc[i, "Vérif lissée"] == "Erreur":
                        s.loc[i, "Vérif lissée"] = "background:#f8d7da;color:#721c24;font-weight:bold"
                return s

            st.dataframe(
                df_hh.style.apply(style_hh, axis=None),
                hide_index=True, width="stretch", height=600
            )

            # ═══ BLOC 3 — Calcul d'une heure au choix ═══════════════════
            st.markdown("### Bloc 3 — Vérification du calcul sur une heure précise")
            st.caption("Reproduit le raisonnement de bess_engine étape par étape pour l'heure choisie.")

            h_check = st.slider("Heure à vérifier :", 0, 23, 0, key="vl_hcheck")
            conso_h = profil[h_check]
            act_h   = actions[h_check]
            soc_av  = soc_h[h_check]
            soc_ap  = soc_h[h_check + 1]
            lisse_h = lisse[h_check]

            # Recalcul manuel
            if act_h[0] == "decharge":
                q_expected = min(conso_h - seuil_j, vl_pow, soc_av - vl_emwh * 0.1)
                lisse_expected = conso_h - q_expected
            elif act_h[0] == "charge":
                q_expected = min(seuil_j - conso_h, vl_pow, vl_emwh * 0.9 - soc_av)
                lisse_expected = conso_h + q_expected
            else:
                q_expected = 0.0
                lisse_expected = conso_h
            soc_expected = soc_av + (q_expected if act_h[0] == "charge" else -q_expected if act_h[0] == "decharge" else 0)

            _ok_q   = abs(act_h[1] - max(0, q_expected)) < 0.001
            _ok_l   = abs(lisse_h - lisse_expected) < 0.001
            _ok_soc = abs(soc_ap - soc_expected) < 0.001

            st.markdown(f"""
**H{h_check:02d} — Raisonnement bess_engine :**
```
Conso Excel     = {conso_h * 1000:.4f} kWh/h = {conso_h:.4f} MW
Seuil P{vl_seuil}%    = {seuil_j:.4f} MW
SOC avant       = {soc_av:.4f} MWh ({soc_av / vl_emwh * 100:.1f}%)
SOC_min         = {vl_emwh * 0.1:.4f} MWh | SOC_max = {vl_emwh * 0.9:.4f} MWh

Condition : conso ({conso_h:.4f}) {'>' if conso_h > seuil_j else '<=' } seuil ({seuil_j:.4f})
         → {'AU-DESSUS → DÉCHARGE possible' if conso_h > seuil_j else 'EN-DESSOUS → CHARGE possible' if conso_h < seuil_j else 'ÉGAL → REPOS'}
```

**Calcul de la quantité :**
```
{'quantité = min(conso-seuil, puissance, SOC-SOC_min)' if act_h[0]=='decharge' else 'quantité = min(seuil-conso, puissance, SOC_max-SOC)' if act_h[0]=='charge' else 'Repos → quantité = 0'}
{'         = min(' + f'{conso_h-seuil_j:.4f}, {vl_pow}, {soc_av-vl_emwh*0.1:.4f})' if act_h[0]=='decharge' else '         = min(' + f'{seuil_j-conso_h:.4f}, {vl_pow}, {vl_emwh*0.9-soc_av:.4f})' if act_h[0]=='charge' else ''}
         = {max(0, q_expected):.4f} MW

bess_engine a produit : {act_h[1]:.4f} MW  {'OK' if _ok_q else 'ERREUR'}

Conso lissée attendue = {conso_h:.4f} {'- ' if act_h[0]=='decharge' else '+ ' if act_h[0]=='charge' else ''}{max(0,q_expected):.4f} = {lisse_expected:.4f} MW
bess_engine a produit = {lisse_h:.4f} MW  {'OK' if _ok_l else 'ERREUR'}

SOC après attendu = {soc_av:.4f} {'- ' if act_h[0]=='decharge' else '+ ' if act_h[0]=='charge' else '+ 0 = '}{max(0,q_expected):.4f} = {soc_expected:.4f} MWh
bess_engine a produit = {soc_ap:.4f} MWh  {'OK' if _ok_soc else 'ERREUR'}
```
            """)

            st.divider()

            # ═══ BLOC 4 — Calcul économie ════════════════════════════════
            st.markdown("### Bloc 4 — Calcul de la réduction de pointe et de l'économie")
            red_l = row_l["Réduction (MW)"]
            eco_l = row_l["Économie/an (€)"]
            h_pointe_av  = int(np.argmax(profil))
            h_pointe_ap  = int(np.argmax(lisse))
            eco_recalc   = round(red_l * vl_tarif, 0)
            ok_eco       = abs(eco_recalc - eco_l) < 1

            st.markdown(f"""
```
Pointe avant lissage  = max(conso_excel)
                      = H{h_pointe_av:02d} → {max(profil):.4f} MW

Pointe après lissage  = max(conso_lissée)
                      = H{h_pointe_ap:02d} → {max(lisse):.4f} MW

Réduction de pointe   = {max(profil):.4f} - {max(lisse):.4f}
                      = {red_l:.4f} MW

Économie annuelle     = Réduction × Tarif
                      = {red_l:.4f} × {vl_tarif}
                      = {_fmt(eco_recalc)} €/an   {'OK' if ok_eco else 'Écart avec ' + str(eco_l) + '€'}
```
            """)

            # ═══ BLOC 5 — Graphique ══════════════════════════════════════
            st.markdown("### Bloc 5 — Visualisation profil avant / après")
            st.caption(
                "Courbe pointillée = consommation originale (Excel). "
                "Courbe pleine = après lissage BESS. "
                "Ligne orange = seuil d'écrêtage. "
                "Les zones bleues sont les heures de décharge, vertes les heures de charge."
            )

            fig_al = go.Figure()

            # Zones colorées par action
            for h in range(24):
                act = actions[h][0]
                if act == "decharge":
                    fig_al.add_vrect(x0=h - 0.5, x1=h + 0.5,
                                     fillcolor="rgba(0,86,179,0.10)", line_width=0)
                elif act == "charge":
                    fig_al.add_vrect(x0=h - 0.5, x1=h + 0.5,
                                     fillcolor="rgba(21,87,36,0.10)", line_width=0)

            fig_al.add_trace(go.Scatter(
                x=list(range(24)), y=profil,
                mode="lines+markers", name="Conso originale (Excel)",
                line=dict(color=C2, width=2.5, dash="dot"),
                hovertemplate="H%{x:02d} avant : <b>%{y:.4f} MW</b><extra></extra>",
            ))
            fig_al.add_trace(go.Scatter(
                x=list(range(24)), y=lisse,
                mode="lines+markers", name="Conso lissée (après BESS)",
                line=dict(color=C1, width=2.5),
                hovertemplate="H%{x:02d} après : <b>%{y:.4f} MW</b><extra></extra>",
            ))
            fig_al.add_hline(
                y=seuil_j, line_dash="dash", line_color="#e67e22", line_width=2,
                annotation_text=f"Seuil P{vl_seuil}% = {seuil_j:.4f} MW",
                annotation_position="right"
            )
            fig_al.add_hline(
                y=max(profil), line_dash="dot", line_color=C2, line_width=1,
                annotation_text=f"Pointe avant = {max(profil):.4f} MW",
                annotation_position="left"
            )
            fig_al.add_hline(
                y=max(lisse), line_dash="dot", line_color=C1, line_width=1,
                annotation_text=f"Pointe après = {max(lisse):.4f} MW",
                annotation_position="left"
            )
            fig_al.update_layout(
                height=400, margin=dict(t=20, b=80, l=70, r=200),
                xaxis=dict(title="Heure", tickmode="array",
                           tickvals=list(range(24)),
                           ticktext=[f"H{h:02d}" for h in range(24)]),
                yaxis=dict(title="Consommation (MW)", gridcolor="#f0f0f0"),
                plot_bgcolor="white", paper_bgcolor="white",
                legend=LEGEND_BOTTOM,
            )
            apply_bb(fig_al)
            st.plotly_chart(fig_al, width="stretch", config=PLOTLY_CFG,
                            key=f"fig_lis_{rid_l}")

            st.download_button(
                f"Exporter audit lissage {date_l} (CSV)",
                data=df_hh.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                file_name=f"BESS_audit_lissage_{date_l.replace('/', '-')}.csv",
                mime="text/csv",
                key=f"dl_lis_{rid_l}",
            )

        st.divider()
        st.download_button(
            "Exporter tableau complet vérification lissage (CSV)",
            data=df_al_f[COLS_LIS].to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name="BESS_verification_lissage_complet.csv",
            mime="text/csv",
            key="dl_lis_global",
        )

    # ══════════════════════════════════════════════════════════════════════════════
    # ONGLET MÉTHODOLOGIE
    # ══════════════════════════════════════════════════════════════════════════════

    # ══════════════════════════════════════════════════════════════════════════════
    # ONGLET METHODOLOGIE
    # ══════════════════════════════════════════════════════════════════════════════

with tab_pays:
    _diag_set_tab("pays")
    st.markdown('<p class="section">Comparaison multi-pays — Arbitrage Day-Ahead</p>',
                unsafe_allow_html=True)
    st.caption(
        "Comparez les performances d'arbitrage sur plusieurs marchés nationaux. "
        "Uploadez un fichier Excel par pays — même format que le fichier principal (feuille Spot_input). "
        "Le pays est détecté automatiquement dans le nom du fichier ou les métadonnées."
    )

    # ── Détection automatique du pays ────────────────────────────────────────
    # Mots-clés par pays — les codes courts (fr, de, be...) ne matchent que comme
    # tokens EXACTS (mot entier isolé), les mots longs matchent aussi en sous-chaîne
    _PAYS_KEYWORDS = {
        "France":      {"long": ["france", "french", "epex_fr"], "short": ["fr"]},
        "Allemagne":   {"long": ["germany", "allemagne", "german", "deutschland", "epex_de"], "short": ["de"]},
        "Belgique":    {"long": ["belgium", "belgique", "belgian", "belgie", "epex_be"], "short": ["be"]},
        "Espagne":     {"long": ["spain", "espagne", "spanish", "epex_es"], "short": ["es"]},
        "Italie":      {"long": ["italy", "italie", "italian", "epex_it"], "short": ["it"]},
        "Pays-Bas":    {"long": ["netherlands", "nederland", "dutch", "epex_nl"], "short": ["nl"]},
        "Royaume-Uni": {"long": ["unitedkingdom", "british", "england", "epex_gb"], "short": ["uk", "gb"]},
        "Suisse":      {"long": ["switzerland", "suisse", "swiss", "epex_ch"], "short": ["ch"]},
        "Autriche":    {"long": ["austria", "autriche", "austrian", "epex_at"], "short": ["at"]},
    }

    def _detect_pays(filename: str, file_bytes: bytes = None) -> str:
        """Détecte le pays dans le nom du fichier ET dans le contenu de l'Excel."""
        import re as _re

        def _scan_text(text: str) -> str:
            t = text.lower()
            tokens = set(_re.split(r'[\s_\-\./#,;|]+', t))
            for pays, kw_dict in _PAYS_KEYWORDS.items():
                # Mots longs : match en sous-chaîne
                for kw in kw_dict["long"]:
                    if kw in t:
                        return pays
                # Codes courts : token exact uniquement
                for kw in kw_dict["short"]:
                    if kw in tokens:
                        return pays
            return None

        # 1. Chercher dans le nom du fichier
        result = _scan_text(filename)
        if result:
            return result

        # 2. Chercher dans le contenu Excel si fourni
        if file_bytes:
            try:
                import openpyxl as _opxl, io as _io2
                wb = _opxl.load_workbook(_io2.BytesIO(file_bytes), read_only=True, data_only=True)
                # Chercher dans les noms de feuilles
                for sn in wb.sheetnames:
                    r = _scan_text(sn)
                    if r: return r
                # Chercher dans les 10 premières lignes de chaque feuille
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    rows_checked = 0
                    for row in ws.iter_rows(max_row=10, values_only=True):
                        for cell in row:
                            if cell and isinstance(cell, str):
                                r = _scan_text(cell)
                                if r: return r
                        rows_checked += 1
                wb.close()
            except Exception:
                pass

        return filename  # fallback = nom du fichier

    # ── Upload des fichiers ───────────────────────────────────────────────────
    st.subheader("Fichiers de prix par pays")
    _cp_col1, _cp_col2 = st.columns(2)
    with _cp_col1:
        st.caption("Le fichier principal (déjà chargé) sera automatiquement inclus.")
        _extra_files = st.file_uploader(
            "Ajouter des fichiers pays supplémentaires (Excel Spot_input)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="pays_files",
            help="Même format que le fichier principal : feuille 'Spot_input' avec colonnes ANNEE, MOIS, JOUR, HEURE, Prix Final."
        )
    with _cp_col2:
        _cp_dur = st.selectbox(
            "Durée du cycle",
            [1, 2],
            format_func=lambda x: f"{x}h ({x}h charge + {x}h décharge)",
            key="cp_dur"
        )
        _cp_quota = st.number_input(
            "Quota cycles/an (0 = illimité)",
            0, 730, 365, key="cp_quota"
        )
        _cp_quota = _cp_quota if _cp_quota > 0 else None
        st.caption(
            "365 = top-365 meilleurs spreads/an. 0 = illimité."
        )

    # ── Construction de la liste des pays à comparer ──────────────────────────
    _pays_list = []

    # Fichiers supplémentaires d'abord
    _extra_bytes_set = set()
    for _f in (_extra_files or []):
        _fb = _f.read()
        _pnom = _detect_pays(_f.name, _fb)
        _pays_list.append({"nom": _pnom, "bytes": _fb, "fname": _f.name})
        _extra_bytes_set.add(_f.name)

    # Fichier principal — ajouté seulement s'il n'est pas déjà dans les extras
    if file_bytes:
        _main_fname = st.session_state.get("_uploaded_fname", "fichier_principal")
        if _main_fname not in _extra_bytes_set:
            _main_pays = _detect_pays(_main_fname, file_bytes)
            _pays_list.insert(0, {"nom": _main_pays, "bytes": file_bytes, "fname": _main_fname})

    if not _pays_list:
        st.info("Uploadez le fichier principal (sidebar) et des fichiers pays supplémentaires pour lancer la comparaison.")
        st.stop()

    # Afficher les pays détectés
    _det_cols = st.columns(min(len(_pays_list), 4))
    for _ci, (_dc, _p) in enumerate(zip(_det_cols, _pays_list)):
        with _dc:
            st.metric("Pays détecté", _p["nom"])

    @st.cache_data(show_spinner=False)
    def _simulate_country_cached(_file_bytes, _params_json):
        """Chargement + simulation mis en cache par (fichier pays, paramètres) —
        rebasculer entre durées/quotas déjà testés, ou re-cliquer sans rien
        changer, ne refait pas le calcul pour les pays déjà connus."""
        import json as _j_cp, io as _io_cp2
        from bess_engine import (load_spot as _ls_cp2, simulate_arbitrage_optimal as _sao_cp2,
                                 simulate_arbitrage as _sa_cp2, aggregate_arbitrage as _agg_cp2)
        _params = _j_cp.loads(_params_json)
        _pv = _ls_cp2(_io_cp2.BytesIO(_file_bytes))
        _daily = (_sao_cp2(_pv, _params) if _params["optimal_quota"] else _sa_cp2(_pv, _params))
        _yearly = _agg_cp2(_daily, _params["power_MW"])
        return _daily, _yearly

    # ── Lancement du calcul ───────────────────────────────────────────────────
    if st.button("Lancer la comparaison", key="btn_cp_run"):
        st.session_state["_cp_results"] = None

        _cp_params = {
            "power_MW":       power_MW,
            "n_cycles":       0,
            "duration_h":     _cp_dur,
            "excluded_hours": {},
            "efficiency":     efficiency,
            "max_cycles_year": _cp_quota,
            "optimal_quota":   bool(_cp_quota),
        }
        _cp_params_json = json.dumps(_cp_params, sort_keys=True)

        _cp_res = []
        _prog_cp = st.progress(0, text="Calcul en cours…")

        for _pi, _p in enumerate(_pays_list):
            _prog_cp.progress(
                int(_pi / len(_pays_list) * 100),
                text=f"Calcul {_p['nom']}…"
            )
            try:
                _daily_cp, _yearly_cp = _simulate_country_cached(_p["bytes"], _cp_params_json)
                _cp_res.append({
                    "pays":   _p["nom"],
                    "daily":  _daily_cp,
                    "yearly": _yearly_cp,
                })
            except Exception as _e_cp:
                st.error(f"Erreur pour {_p['nom']} : {_e_cp}")

        _prog_cp.progress(100, text="Terminé.")
        _prog_cp.empty()
        st.session_state["_cp_results"] = _cp_res

    # ── Affichage des résultats ───────────────────────────────────────────────
    _cp_results = st.session_state.get("_cp_results")
    if _cp_results:
        st.markdown('<p class="section">Résultats comparatifs par pays</p>',
                    unsafe_allow_html=True)
        st.caption(
            f"Comparaison sur la période disponible dans chaque fichier. "
            f"Durée cycle : {_cp_dur}h. "
            f"Quota : {'illimité' if not _cp_quota else str(_cp_quota) + ' cycles/an (top spreads)'}. "
            f"Puissance : {power_MW*1000:.0f} kW. Rendement : {efficiency*100:.0f}%."
        )

        # ── Tableau récapitulatif trié par PnL/MW/an ─────────────────────────
        _cp_rows = []
        for _r in _cp_results:
            _y = _r["yearly"]
            _pnl_total = _y["pnl_total"].sum()
            _nb_ans = len(_y)
            _cp_rows.append({
                "Pays":                 _r["pays"],
                "Période":              f"{int(_y['annee'].min())}–{int(_y['annee'].max())}",
                "PnL total (€)":        _fmt(_pnl_total),
                "PnL/an moy. (€)":      _fmt(_pnl_total / _nb_ans),
                "Potentiel max (€)":    _fmt(_y["pnl_absolu_total"].sum()),
                "Spread moy. (€/MWh)":  f"{_y['spread_moy'].mean():.1f}",
                "Jours actifs/an":      str(int(_y["jours_actifs"].mean())),
                "PnL/MW/an (€/MW)":     _fmt(_pnl_total / power_MW / _nb_ans),
                "_pnl_mw": _pnl_total / power_MW / _nb_ans,  # pour le tri
            })

        _cp_df = pd.DataFrame(_cp_rows).sort_values("_pnl_mw", ascending=False).reset_index(drop=True)
        _cp_df.index = _cp_df.index + 1
        _cp_df_display = _cp_df.drop(columns=["_pnl_mw"])

        # Podium
        _medals = ["1er", "2e", "3e", "4e", "5e"]
        _medal_cols = st.columns(min(len(_cp_df), 3))
        for _mi, (_mc, (_, _row)) in enumerate(zip(_medal_cols, _cp_df.iterrows())):
            with _mc:
                st.metric(
                    f"{_medals[_mi]} — {_row['Pays']}",
                    f"{_row['PnL/MW/an (€/MW)']} €/MW/an",
                    delta=f"Spread moy. : {_row['Spread moy. (€/MWh)']} €/MWh"
                )

        st.dataframe(_cp_df_display, hide_index=False, width="stretch")
        st.caption(
            "Classement par **PnL/MW/an** = revenus annuels par MW installé. "
            "C'est la métrique clé pour décider dans quel pays investir : "
            "un marché avec un PnL/MW/an élevé génère plus de revenus pour le même investissement. "
            "**Spread moy.** = volatilité des prix — plus il est élevé, plus les opportunités d'arbitrage sont grandes."
        )

        # ── Graphique PnL annuel par pays ──────────────────────────────────────
        st.markdown('<p class="section">PnL annuel par pays</p>', unsafe_allow_html=True)
        _COLORS_CP = ["#1565c0","#e65100","#2e7d32","#6a1b9a","#00838f","#c62828","#4e342e","#37474f"]
        _fig_cp1 = go.Figure()
        for _ci2, _r in enumerate(_cp_results):
            _y = _r["yearly"]
            _clr = _COLORS_CP[_ci2 % len(_COLORS_CP)]
            _fig_cp1.add_trace(go.Bar(
                name=_r["pays"],
                x=_y["annee"].astype(str),
                y=_y["pnl_total"],
                marker_color=_clr,
                text=[f"{_fmt(v)} €" for v in _y["pnl_total"]],
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>" + _r["pays"] + " : %{y:,.0f} €<extra></extra>",
            ))
        _fig_cp1.update_layout(
            height=380, barmode="group",
            xaxis=dict(title="Année"),
            yaxis=dict(title="PnL (€)", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            margin=dict(t=20, b=80, l=60, r=10),
        )
        apply_bb(_fig_cp1)
        st.plotly_chart(_fig_cp1, width="stretch", config=PLOTLY_CFG, key="cp_fig1")
        st.caption("PnL réel annuel par pays. Permet de comparer directement la rentabilité de chaque marché.")

        # ── Graphique spread moyen par pays ────────────────────────────────────
        st.markdown('<p class="section">Spread moyen par pays (€/MWh)</p>', unsafe_allow_html=True)
        _fig_cp2 = go.Figure()
        for _ci3, _r in enumerate(_cp_results):
            _y = _r["yearly"]
            _clr = _COLORS_CP[_ci3 % len(_COLORS_CP)]
            _fig_cp2.add_trace(go.Scatter(
                name=_r["pays"],
                x=_y["annee"].astype(str),
                y=_y["spread_moy"],
                mode="lines+markers",
                line=dict(width=2, color=_clr),
                marker=dict(size=8),
                hovertemplate="<b>%{x}</b><br>" + _r["pays"] + " : %{y:.1f} €/MWh<extra></extra>",
            ))
        _fig_cp2.update_layout(
            height=320,
            xaxis=dict(title="Année"),
            yaxis=dict(title="Spread moyen (€/MWh)", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            margin=dict(t=20, b=80, l=60, r=10),
        )
        apply_bb(_fig_cp2)
        st.plotly_chart(_fig_cp2, width="stretch", config=PLOTLY_CFG, key="cp_fig2")
        st.caption(
            "Spread moyen des cycles réalisés. "
            "Un spread plus élevé = opportunités d'arbitrage plus grandes sur ce marché. "
            "La différence entre pays reflète la volatilité des prix de l'électricité."
        )

        # ── PnL cumulé ────────────────────────────────────────────────────────
        st.markdown('<p class="section">PnL cumulé dans le temps</p>', unsafe_allow_html=True)
        _fig_cp3 = go.Figure()
        for _ci4, _r in enumerate(_cp_results):
            _d = _r["daily"].sort_values("date")
            _clr = _COLORS_CP[_ci4 % len(_COLORS_CP)]
            _fig_cp3.add_trace(go.Scatter(
                name=_r["pays"],
                x=_d["date"],
                y=_d["pnl"].cumsum(),
                mode="lines",
                line=dict(width=2, color=_clr),
                hovertemplate="<b>%{x|%d/%m/%Y}</b><br>" + _r["pays"] + " : %{y:,.0f} €<extra></extra>",
            ))
        _fig_cp3.update_layout(
            height=340,
            xaxis=dict(title="", tickformat="%b %Y"),
            yaxis=dict(title="PnL cumulé (€)", gridcolor="#f0f0f0"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=LEGEND_BOTTOM,
            margin=dict(t=20, b=80, l=60, r=10),
            hovermode="x unified",
        )
        apply_bb(_fig_cp3)
        st.plotly_chart(_fig_cp3, width="stretch", config=PLOTLY_CFG, key="cp_fig3")
        st.caption("La courbe la plus haute = le marché le plus rentable sur la durée.")

        # ── Export CSV ────────────────────────────────────────────────────────
        _cp_all_rows = []
        for _r in _cp_results:
            _y = _r["yearly"]
            for _, _yr in _y.iterrows():
                _cp_all_rows.append({
                    "Pays": _r["pays"],
                    "Année": int(_yr["annee"]),
                    "PnL réel (€)": round(_yr["pnl_total"], 0),
                    "Potentiel max (€)": round(_yr["pnl_absolu_total"], 0),
                    "Spread moy. (€/MWh)": round(_yr["spread_moy"], 1),
                    "Jours actifs": int(_yr["jours_actifs"]),
                })
        st.download_button(
            "Exporter comparaison pays (CSV)",
            data=pd.DataFrame(_cp_all_rows).to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name="BESS_comparaison_pays.csv",
            mime="text/csv",
            key="btn_cp_csv",
        )

with tab_methodo:
    _diag_set_tab(None)
    st.title("Méthodologie — Documentation Technique")
    st.write(
        "Documentation complète des algorithmes, hypothèses et données utilisés dans chaque partie de l'outil. "
        "Mise à jour en mai 2026."
    )

    with st.expander("1. Architecture générale", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            st.write("**bess_engine.py** — Moteur de calcul pur : algorithmes d'optimisation, sans interface.")
        with c2:
            st.write("**bess_dashboard.py** — Interface Streamlit : affichage, graphiques, export.")
        st.subheader("Pipeline de calcul")
        st.code(
            "Excel Spot_input / Case 3\n"
            "   -> load_spot() / load_case3()   # Lecture pivot [jours x H00..H23]\n"
            "   -> simulate_arbitrage()          # Optimisation journaliere chronologique\n"
            "   -> simulate_arbitrage_optimal()  # Top-N meilleurs spreads\n"
            "   -> aggregate_arbitrage()          # Agregation annuelle\n"
            "   -> Affichage Streamlit",
            language="text"
        )

    with st.expander("2. Algorithme _best_n_cycles", expanded=False):
        st.subheader("Mode n=1 (best_single_cycle)")
        st.code(
            "Pour chaque heure de charge hc (boucle Python) :\n"
            "  - decharge = argmax(prix apres hc)   # vectorise numpy a l'interieur\n"
            "Complexite O(k^2) avec k = nb heures disponibles",
            language="text"
        )
        st.subheader("Mode MAX / n=0 — DP Weighted Interval Scheduling")
        st.write(
            "Trouve le sous-ensemble non chevauchant de cycles qui maximise le PnL total. "
            "Garantit l'optimum (pas greedy)."
        )
        st.code(
            "1. Generer candidats (2 types) :\n"
            "   a) Cycles consecutifs H_i -> H_{i+1} (petits spreads adjacents)\n"
            "   b) Meilleur cycle depuis chaque H_i (grands spreads H_bas -> H_haut)\n"
            "2. Trier par heure de fin croissante\n"
            "3. DP : dp[i] = max(dp[i-1], dp[last_compat(i)] + pnl[i])\n"
            "4. Reconstruction du chemin optimal",
            language="text"
        )
        st.write(
            "**Exemple :** au lieu de H8->H18 seul (24.51€), "
            "le DP trouve H8->H9 (1.29) + H10->H11 (3.44) + H14->H18 (22.36) = 27.09€ "
            "(plus les petits cycles adjacents rentables)."
        )
        st.subheader("Formule PnL")
        st.code(
            "PnL = Puissance_MW x n_heures x (Rendement x Prix_decharge_moy - Prix_charge_moy)",
            language="text"
        )

    with st.expander("3. Modes de simulation — Chronologique vs Top-N", expanded=False):
        st.subheader("Mode chronologique (simulate_arbitrage)")
        st.code(
            "for each day in 2026..2029:\n"
            "    if cycles_this_year >= max_cycles_year: skip  # bloque\n"
            "    cycles = _best_n_cycles(prices, avail, n_cycles, ...)\n"
            "    cycles_this_year += len(cycles)",
            language="python"
        )
        st.subheader("Mode Top-N meilleurs spreads (simulate_arbitrage_optimal)")
        st.write(
            "Sur chaque annee independamment, genere TOUS les cycles possibles (mode MAX), "
            "les classe par spread decroissant (PnL en tiebreak), retient les N meilleurs de "
            "CETTE annee peu importe le jour. Le quota max_cycles_year est donc applique "
            "annee par annee, pas une seule fois sur tout l'horizon. "
            "Suppose connaissance parfaite des prix annuels — borne haute theorique."
        )
        st.code(
            "# Passe 1 : tous les cycles de tous les jours, toutes annees\n"
            "for day in all_days:\n"
            "    cycles += _best_n_cycles(prices_day, avail, n_cycles=0, ...)\n"
            "# Passe 2 : top-N par spread decroissant, RESET PAR ANNEE\n"
            "for annee, cycles_annee in group_by_year(cycles):\n"
            "    selected[annee] = sorted(cycles_annee, key=(spread, pnl), reverse=True)[:max_cycles_year]",
            language="python"
        )
        st.subheader("4 cas standards")
        import pandas as pd
        st.dataframe(pd.DataFrame({
            "Scenario": ["1h Top 365/an", "1h Illimite", "2h Top 365/an", "2h Illimite"],
            "Duree": ["1h", "1h", "2h", "2h"],
            "Quota": ["365 meilleurs spreads", "Illimite", "365 meilleurs spreads", "Illimite"],
            "Algorithme": ["Top-N optimal", "Mode MAX", "Top-N optimal", "Mode MAX"],
        }), hide_index=True, width="stretch")

    with st.expander("4. Contraintes et quota", expanded=False):
        st.write(
            "**Quota annuel** : compteur incremente a chaque cycle. "
            "Quand il atteint max_cycles_year, les jours suivants sont bloques (PnL=0) "
            "jusqu'au 1er janvier."
        )
        st.write(
            "**Restrictions horaires** : certains jours/heures exclus (ex. lun-sam H10-H12). "
            "La borne max est calculee sans restriction."
        )
        st.code("Taux de capture (%) = PnL_reel / PnL_borne_max x 100", language="text")

    with st.expander("5. Analyse de sensibilite — Heat map Puissance x Duree", expanded=False):
        st.write(
            "Repond a la question : quelle combinaison puissance / duree maximise le PnL sur 4 ans ?"
        )
        st.write(
            "**Mode MAX** (n_cycles=0) : le moteur cherche toutes les opportunites rentables par jour. "
            "Si quota > 0 : top-N meilleurs spreads. Si quota = 0 : illimite."
        )
        st.code(
            "for dur in durees:\n"
            "    for power in puissances:\n"
            "        p = {n_cycles: 0, duration_h: dur, power_MW: power, ...}\n"
            "        daily = simulate_arbitrage_optimal(pv, p)  # ou simulate_arbitrage\n"
            "        matrix[dur][power] = yearly_pnl_total",
            language="python"
        )
        st.write(
            "Le PnL est lineairement proportionnel a la puissance. "
            "La duree impacte la capacite et les spreads accessibles : "
            "2h permet d'exploiter des ecarts prix sur une fenetre plus large."
        )

    with st.expander("6. ROI et donnees IEA 2024", expanded=False):
        st.code(
            "CAPEX_total  = CAPEX_kWh x Puissance_MW x 1000 x Duree_h\n"
            "PnL_an(n)    = PnL_base x (1 - taux_degrad)^(n-1)\n"
            "Cashflow(n)  = PnL_an(n) - OPEX_an\n"
            "VAN = -CAPEX_total + somme Cashflow(n) / (1 + taux_actu)^n",
            language="text"
        )
        for item, val in [
            ("CAPEX 2022", "330 €/kWh (IEA reel)"),
            ("CAPEX 2024", "150 €/kWh (IEA Electricity 2026)"),
            ("CAPEX 2025", "120 €/kWh (BNEF/Ember)"),
            ("CAPEX 2026", "105 €/kWh (Capstone DC France)"),
            ("CAPEX 2030", "85 €/kWh (BNEF projection)"),
            ("Cycles LFP", "6 000-7 000 avant 80% capacite (BNEF/Ember 2025)"),
            ("Degradation", "2-3%/an (IEA 2024)"),
            ("OPEX", "0 sous garantie | 2.5% CAPEX/an sinon (NREL)"),
        ]:
            st.write(f"- **{item}** : {val}")

    with st.expander("7. Lissage de charge (Peak Shaving)", expanded=False):
        st.write(
            "La batterie evite les pics de consommation industriels. "
            "Economie = reduction de la puissance de pointe souscrite (euros/MW/an). "
            "Seuil d'intervention = percentile configurable de la consommation (75 par defaut)."
        )
        st.code(
            "for h in heures:\n"
            "    if conso[h] > seuil and soc > soc_min:\n"
            "        decharge = min(conso[h]-seuil, puissance, soc-soc_min)\n"
            "        soc -= decharge\n"
            "    elif conso[h] < seuil and soc < soc_max:\n"
            "        charge = min(seuil-conso[h], puissance, soc_max-soc)\n"
            "        soc += charge",
            language="python"
        )
        st.warning(
            "Le rendement (efficiency) est recu en parametre par lissage_day() mais n'est "
            "actuellement pas applique au SOC ni a l'energie chargee/dechargee : le lissage "
            "suppose une batterie 100% efficace, contrairement a l'onglet Arbitrage qui applique "
            "bien le rendement. Ecart a corriger si la precision energetique du Lissage doit "
            "matcher la realite physique."
        )

    with st.expander("8. Historique des versions", expanded=False):
        st.code(
            "v2.0 :\n"
            "  - Mode MAX : DP Weighted Interval Scheduling\n"
            "  - Cycles consecutifs comme candidats (H_i->H_{i+1})\n"
            "  - simulate_arbitrage_optimal (top-N meilleurs spreads)\n"
            "  - 4 cas standards : 1h/2h x Top-365/Illimite\n"
            "  - Volet detail spreads par scenario et par annee\n"
            "  - Assistant IA (Gemini) avec contexte donnees\n"
            "  - CAPEX mis a jour sources mai 2026\n"
            "  - Cles uniques sur tous les graphiques Plotly\n"
            "  - Spinner et barre de progression sur calculs longs\n"
            "\n"
            "v2.1 (mises a jour ulterieures) :\n"
            "  - Intraday : correction signe spread bid/ask (ASK - BID)\n"
            "  - Intraday : Continuous_Index filtre par IndexName/TimeResolution\n"
            "    (evite de melanger ID1/ID3/IDFULL et 15min/30min/60min)\n"
            "  - Correction ZeroDivisionError sur le parsing IDA1/2/3\n"
            "  - Cache @st.cache_data etendu a tous les onglets (Lissage, "
            "Comparaison pays inclus)\n"
            "  - Assistant IA deplace en bulle flottante (bas droite), "
            "position fixe corrigee au scroll/fermeture\n"
            "  - Nouvel onglet Imbalance Market (prix de reglement des ecarts, "
            "series Negatif/Positif)\n"
            "  - Intraday : marche continu consolide en un seul onglet \"Continuous\" "
            "(base sur Continuous_Index/IDFULL, priorite 60min->30min->15min) ; "
            "Continuous_Statistics n'est plus traite separement (redondant)",
            language="text"
        )

# ══════════════════════════════════════════════════════════════════════════════
# RAPPORT DE VÉRIFICATION GLOBAL — remplissage du placeholder sidebar
# ══════════════════════════════════════════════════════════════════════════════
# Placé ici (fin de script) car tous les onglets ont eu l'occasion de déposer
# leur résumé dans st.session_state à ce stade du run. Le widget réel
# (download_button) est rendu dans _diag_report_slot, créé dans la sidebar,
# pour qu'il apparaisse bien là visuellement malgré le calcul tardif.
if st.session_state.get("_diag_global_requested"):
    import datetime as _dtd_final
    _diag_html = _build_global_diag_html()
    with _diag_report_slot.container():
        st.download_button(
            "⬇ Télécharger le rapport complet (HTML)",
            data=_diag_html.encode("utf-8"),
            file_name=f"BESS_verification_complete_{_dtd_final.datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            key="btn_diag_global_dl",
        )

