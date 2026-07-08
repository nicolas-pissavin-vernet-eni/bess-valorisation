"""
================================================================================
BESS VALORISATION - Marché Day-Ahead (DA)
================================================================================
Outil de valorisation d'un système de stockage par batteries (BESS)
par arbitrage sur le marché spot électrique.

Fonctionnalités :
  - Pivot automatique des prix spot en 24 colonnes (1 ligne = 1 jour)
  - Calcul du spread sans contrainte (Perfect Foresight = borne max théorique)
  - Calcul du spread avec contraintes (heures exclues, jours, usure)
  - Durée de charge paramétrable : 1h ou 2h
  - Contrainte charge AVANT décharge (réalisme physique)
  - Contrainte d'usure (nombre de cycles max par an)
  - Export Excel avec graphiques (PnL, profil journalier, distribution spreads)

Pour ajouter un scénario : ajouter un bloc dans SCENARIOS (section CONFIG).
================================================================================
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import warnings
import xlsxwriter
warnings.filterwarnings("ignore")


# ==============================================================================
# ██████╗ CONFIG — TOUT CE QUI PEUT CHANGER EST ICI
# ==============================================================================

INPUT_EXCEL  = r"C:\Users\nicolas.pissavin\Documents\Duplication projet batteries\BESS_valorisation_RESULTATS - Copie.xlsx"
INPUT_SHEET  = "Spot_input"
OUTPUT_EXCEL = r"C:\Users\nicolas.pissavin\Documents\Duplication projet batteries\BESS_valorisation_PYTHON_RESULTATS.xlsx"

# ── Paramètres physiques de la batterie ───────────────────────────────────────
BATTERY = {
    "energy_MWh"     : 2.0,    # Capacité totale (MWh)
    "soc_min_pct"    : 0.10,   # SOC minimum (10% de la capacité)
    "soc_max_pct"    : 0.90,   # SOC maximum (90% de la capacité)
    "efficiency"     : 0.92,   # Rendement aller-retour (ex: 0.92 = 92%)
    "max_cycles_year": 300,    # Contrainte usure : nb max de cycles par an
                               # (None = pas de contrainte)
}

# ── Scénarios ─────────────────────────────────────────────────────────────────
# Chaque scénario définit :
#   name           : libellé
#   power_MW       : puissance de charge/décharge en MW
#   duration_h     : durée d'un cycle de charge OU décharge → 1 ou 2
#                    1h = 1 heure achetée + 1 heure vendue
#                    2h = 2 heures achetées + 2 heures vendues
#   excluded_hours : {"days": [0..6], "hours": [0..23]}
#                    days   : jours concernés (0=Lun, 1=Mar, ..., 6=Dim)
#                    hours  : heures exclues ces jours-là
#                    {}     : aucune exclusion
#   description    : texte libre
SCENARIOS = [
    {
        "name"          : "Perfect Foresight 2h — sans contrainte",
        "power_MW"      : 0.43,
        "duration_h"    : 2,
        "excluded_hours": {},
        "description"   : "Borne max théorique — 430 KW, 2h, aucune restriction",
    },
    {
        "name"          : "Perfect Foresight 1h — sans contrainte",
        "power_MW"      : 0.43,
        "duration_h"    : 1,
        "excluded_hours": {},
        "description"   : "Borne max théorique — 430 KW, 1h, aucune restriction",
    },
    {
        "name"          : "Case 1 — 430KW, 2h, Lun-Sam 10h-12h exclus",
        "power_MW"      : 0.43,
        "duration_h"    : 2,
        "excluded_hours": {
            "days" : [0, 1, 2, 3, 4, 5],   # Lun à Sam
            "hours": [10, 11],
        },
        "description"   : "430 KW, 2h — Lun-Sam indisponible 10h-12h",
    },
    {
        "name"          : "Case 2 — 215KW, 2h, Mar-Jeu 10h-14h exclus",
        "power_MW"      : 0.215,
        "duration_h"    : 2,
        "excluded_hours": {
            "days" : [1, 2, 3],             # Mar, Mer, Jeu
            "hours": [10, 11, 12, 13],
        },
        "description"   : "215 KW, 2h — Mar-Jeu indisponible 10h-14h",
    },
    {
        "name"          : "Case 3 — 215KW, 2h, toujours disponible",
        "power_MW"      : 0.215,
        "duration_h"    : 2,
        "excluded_hours": {},
        "description"   : "215 KW, 2h — Disponible 24h/24, 7j/7",
    },
    {
        "name"          : "Case 3b — 215KW, 1h, toujours disponible",
        "power_MW"      : 0.215,
        "duration_h"    : 1,
        "excluded_hours": {},
        "description"   : "215 KW, 1h — Disponible 24h/24, 7j/7",
    },
    # ── Pour ajouter un cas : copier un bloc ci-dessus ────────────────────────
    # {
    #     "name"          : "Case X — description",
    #     "power_MW"      : 0.5,
    #     "duration_h"    : 1,       # 1 ou 2
    #     "excluded_hours": {
    #         "days" : [0,1,2,3,4],  # Lun-Ven
    #         "hours": list(range(8, 20)),  # 8h-20h exclus
    #     },
    #     "description"   : "500 KW, 1h — semaine 8h-20h exclus",
    # },
]


# ==============================================================================
# ÉTAPE 1 — LECTURE ET PIVOT DES DONNÉES SPOT
# ==============================================================================

def load_and_pivot(path: str, sheet: str) -> pd.DataFrame:
    """
    Lit les prix spot (format long : 1 ligne par heure)
    et les pivote en format large : 1 ligne par jour, colonnes H00..H23.

    Ajoute aussi : date, weekday (0=Lun..6=Dim), nom du jour.
    """
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    df = df.rename(columns={
        "HEURE"      : "heure",
        "JOUR"       : "jour",
        "MOIS"       : "mois",
        "ANNEE"      : "annee",
        "Prix Final" : "prix",
    })

    required = {"annee", "mois", "jour", "heure", "prix"}
    missing  = required - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans {sheet} : {missing}")

    df = df[["annee", "mois", "jour", "heure", "prix"]].dropna()
    for col in ["annee", "mois", "jour", "heure"]:
        df[col] = df[col].astype(int)
    df["prix"] = df["prix"].astype(float)

    # ── Pivot : 1 jour × 24 heures ───────────────────────────────────────────
    pivot = df.pivot_table(
        index=["annee", "mois", "jour"],
        columns="heure",
        values="prix",
        aggfunc="first",
    )
    pivot.columns = [f"H{h:02d}" for h in pivot.columns]
    pivot = pivot.reset_index()

    # ── Métadonnées temporelles ───────────────────────────────────────────────
    pivot["date"] = pd.to_datetime(
        pivot[["annee", "mois", "jour"]].rename(
            columns={"annee": "year", "mois": "month", "jour": "day"}
        )
    )
    pivot["weekday"]      = pivot["date"].dt.weekday   # 0=Lun, 6=Dim
    pivot["weekday_name"] = pivot["date"].dt.strftime("%A")

    return pivot.sort_values("date").reset_index(drop=True)


# ==============================================================================
# ÉTAPE 2 — CALCUL DE L'ARBITRAGE (cœur métier)
# ==============================================================================

HOUR_COLS = [f"H{h:02d}" for h in range(24)]

def get_available_hours(row: pd.Series, excluded: dict) -> list[int]:
    """
    Retourne la liste des indices d'heures disponibles pour ce jour,
    en appliquant les exclusions du scénario.
    """
    all_hours = list(range(24))
    if not excluded:
        return all_hours

    weekday       = int(row["weekday"])
    days_excl     = excluded.get("days",  [])
    hours_excl    = excluded.get("hours", [])

    if weekday in days_excl:
        return [h for h in all_hours if h not in hours_excl]
    return all_hours


def compute_arbitrage(prices_24h: np.ndarray,
                      avail_hours: list[int],
                      duration_h: int,
                      power_MW: float,
                      efficiency: float) -> dict:
    """
    Calcule le meilleur arbitrage possible pour un jour donné.

    Paramètres :
      prices_24h : array de 24 prix (index = heure)
      avail_hours: heures disponibles (après exclusions)
      duration_h : 1 ou 2 → nombre d'heures de charge ET de décharge
      power_MW   : puissance en MW
      efficiency : rendement (ex 0.92)

    Retourne :
      spread_brut  : prix_decharge - prix_charge (€/MWh), sans rendement
      spread_net   : spread après prise en compte du rendement
      pnl          : PnL en € (= spread_net × power_MW × duration_h)
      valid        : True si le trade est physiquement réalisable
      h_charge     : heures de charge retenues
      h_decharge   : heures de décharge retenues
      prix_charge  : prix moyen de charge
      prix_decharge: prix moyen de décharge
    """
    p = prices_24h
    n = duration_h  # nombre d'heures par phase

    if len(avail_hours) < 2 * n:
        return _empty_result()

    avail = np.array(avail_hours)
    prices_avail = p[avail]

    # n heures les moins chères (charge) et n heures les plus chères (décharge)
    idx_asc  = np.argsort(prices_avail)
    idx_desc = idx_asc[::-1]

    h_charge_idx   = sorted(avail[idx_asc[:n]])
    h_decharge_idx = sorted(avail[idx_desc[:n]])

    prix_charge    = p[h_charge_idx].mean()
    prix_decharge  = p[h_decharge_idx].mean()

    spread_brut = prix_decharge - prix_charge

    # Contrainte physique : toute heure de charge < toute heure de décharge
    valid = (max(h_charge_idx) < min(h_decharge_idx)) and (spread_brut > 0)

    if not valid:
        return _empty_result()

    # Rendement : on achète plus qu'on ne vend en énergie
    # PnL = énergie_vendue × prix_décharge - énergie_achetée × prix_charge
    #      = (power × duration × efficiency) × prix_décharge
    #        - (power × duration) × prix_charge
    energy_charged    = power_MW * n            # MWh achetés
    energy_discharged = power_MW * n * efficiency  # MWh vendus

    pnl = energy_discharged * prix_decharge - energy_charged * prix_charge
    spread_net = prix_decharge * efficiency - prix_charge  # spread "net rendement"

    return {
        "spread_brut"  : round(spread_brut, 4),
        "spread_net"   : round(spread_net,  4),
        "pnl"          : round(pnl, 4),
        "valid"        : True,
        "h_charge"     : h_charge_idx,
        "h_decharge"   : h_decharge_idx,
        "prix_charge"  : round(prix_charge, 4),
        "prix_decharge": round(prix_decharge, 4),
    }


def _empty_result() -> dict:
    return {
        "spread_brut": 0, "spread_net": 0, "pnl": 0, "valid": False,
        "h_charge": [], "h_decharge": [], "prix_charge": 0, "prix_decharge": 0,
    }


# ==============================================================================
# ÉTAPE 3 — SIMULATION D'UN SCÉNARIO (avec contrainte usure)
# ==============================================================================

def simulate_scenario(pivot: pd.DataFrame, scenario: dict) -> pd.DataFrame:
    """
    Simule le scénario sur toutes les journées disponibles.

    Pour chaque jour :
      1. Détermine les heures disponibles
      2. Calcule l'arbitrage sans contrainte (Perfect Foresight local)
      3. Applique la contrainte d'usure (cycles max/an)

    Retourne un DataFrame journalier avec tous les résultats.
    """
    power_MW   = scenario["power_MW"]
    duration_h = scenario["duration_h"]
    excluded   = scenario.get("excluded_hours", {})
    efficiency = BATTERY["efficiency"]
    max_cycles = BATTERY.get("max_cycles_year")  # peut être None

    records = []

    # Compteur de cycles par année (pour la contrainte usure)
    cycles_by_year: dict[int, int] = {}

    for _, row in pivot.iterrows():
        annee   = int(row["annee"])
        prices  = row[HOUR_COLS].values.astype(float)
        avail   = get_available_hours(row, excluded)

        # ── Calcul sans contrainte (borne max du jour) ────────────────────────
        res_libre = compute_arbitrage(prices, list(range(24)), duration_h,
                                      power_MW, efficiency)

        # ── Calcul avec contraintes horaires ─────────────────────────────────
        res_contraint = compute_arbitrage(prices, avail, duration_h,
                                          power_MW, efficiency)

        # ── Contrainte usure : max N cycles/an ───────────────────────────────
        cycles_done = cycles_by_year.get(annee, 0)
        if max_cycles is not None and res_contraint["valid"]:
            if cycles_done >= max_cycles:
                res_contraint = _empty_result()
                usure_bloque  = True
            else:
                cycles_by_year[annee] = cycles_done + 1
                usure_bloque = False
        else:
            usure_bloque = False
            if res_contraint["valid"]:
                cycles_by_year[annee] = cycles_done + 1

        records.append({
            # Identification
            "scenario"        : scenario["name"],
            "date"            : row["date"],
            "annee"           : annee,
            "mois"            : int(row["mois"]),
            "jour"            : int(row["jour"]),
            "weekday"         : int(row["weekday"]),
            "weekday_name"    : row["weekday_name"],
            "n_heures_dispo"  : len(avail),
            "duration_h"      : duration_h,

            # Sans contrainte (Perfect Foresight local)
            "spread_libre"    : res_libre["spread_brut"],
            "pnl_libre"       : res_libre["pnl"],

            # Avec contraintes horaires
            "spread_contraint": res_contraint["spread_brut"],
            "spread_net"      : res_contraint["spread_net"],
            "pnl_contraint"   : res_contraint["pnl"],
            "valid"           : res_contraint["valid"],
            "usure_bloque"    : usure_bloque,

            # Détail trade
            "h_charge"        : str(res_contraint["h_charge"]),
            "h_decharge"      : str(res_contraint["h_decharge"]),
            "prix_charge"     : res_contraint["prix_charge"],
            "prix_decharge"   : res_contraint["prix_decharge"],

            # Pour graphiques
            "prix_min_jour"   : float(np.nanmin(prices)),
            "prix_max_jour"   : float(np.nanmax(prices)),
            "prix_moy_jour"   : float(np.nanmean(prices)),
        })

    return pd.DataFrame(records)


# ==============================================================================
# ÉTAPE 4 — AGRÉGATION PAR ANNÉE
# ==============================================================================

def aggregate_yearly(daily: pd.DataFrame, scenario: dict) -> pd.DataFrame:
    """Agrège les résultats journaliers par année."""
    power_MW = scenario["power_MW"]

    g = daily.groupby("annee")
    out = pd.DataFrame({
        "annee"              : g["annee"].first(),
        "jours_simules"      : g["date"].count(),
        "jours_actifs"       : g["valid"].sum(),
        "jours_bloques_usure": g["usure_bloque"].sum(),
        "spread_libre_moy"   : g["spread_libre"].mean().round(2),
        "pnl_libre_total"    : g["pnl_libre"].sum().round(0),
        "spread_contraint_moy": g.apply(
            lambda x: x.loc[x["valid"], "spread_contraint"].mean()
            if x["valid"].any() else 0
        ).round(2),
        "pnl_contraint_total": g["pnl_contraint"].sum().round(0),
    }).reset_index(drop=True)

    # €/MW normalisé
    out["pnl_libre_par_MW"]     = (out["pnl_libre_total"]     / (power_MW * 1000)).round(2)
    out["pnl_contraint_par_MW"] = (out["pnl_contraint_total"] / (power_MW * 1000)).round(2)
    out["taux_activation"]      = (out["jours_actifs"] / out["jours_simules"]).round(3)
    out["scenario"]             = scenario["name"]
    out["power_MW"]             = power_MW
    out["duration_h"]           = scenario["duration_h"]
    out["description"]          = scenario["description"]

    return out


# ==============================================================================
# ÉTAPE 5 — EXPORT EXCEL AVEC GRAPHIQUES
# ==============================================================================

def export_excel(all_daily: list, all_yearly: list,
                 pivot: pd.DataFrame, scenarios: list, output_path: str):

    wb  = xlsxwriter.Workbook(output_path)

    # ── Formats ───────────────────────────────────────────────────────────────
    F = {
        "title"  : wb.add_format({"bold":True,"font_size":14,"font_color":"#1F4E79"}),
        "h1"     : wb.add_format({"bold":True,"bg_color":"#1F4E79","font_color":"white",
                                   "border":1,"align":"center","valign":"vcenter","text_wrap":True}),
        "h2"     : wb.add_format({"bold":True,"bg_color":"#2E75B6","font_color":"white",
                                   "border":1,"align":"center"}),
        "num0"   : wb.add_format({"num_format":"#,##0",    "border":1}),
        "num2"   : wb.add_format({"num_format":"#,##0.00", "border":1}),
        "pct"    : wb.add_format({"num_format":"0.0%",     "border":1}),
        "str"    : wb.add_format({"border":1}),
        "alt"    : wb.add_format({"bg_color":"#DEEAF1","border":1}),
        "alt2"   : wb.add_format({"bg_color":"#DEEAF1","num_format":"#,##0.00","border":1}),
        "alt0"   : wb.add_format({"bg_color":"#DEEAF1","num_format":"#,##0","border":1}),
        "green"  : wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","bold":True,"border":1}),
        "red"    : wb.add_format({"bg_color":"#FCE4D6","num_format":"#,##0","border":1}),
        "sub"    : wb.add_format({"bold":True,"bg_color":"#BDD7EE","border":1}),
        "date"   : wb.add_format({"num_format":"dd/mm/yyyy","border":1}),
    }

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 1 : RÉSUMÉ
    # ══════════════════════════════════════════════════════════════════════════
    ws = wb.add_worksheet("Résumé")
    ws.set_column("A:A", 42)
    ws.set_column("B:K", 18)
    ws.set_row(0, 36)

    ws.write("A1", "BESS Valorisation DA — Résumé par scénario", F["title"])
    ws.write("A2", f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
                   f"Rendement batterie : {BATTERY['efficiency']*100:.0f}%  |  "
                   f"Max cycles/an : {BATTERY['max_cycles_year']}", F["str"])

    headers = ["Scénario","Puissance\n(MW)","Durée\ncycle",
               "Spread libre\nmoy. (€/MWh)","PnL libre\ntotal (€)","PnL libre\n(€/MW)",
               "Spread contraint\nmoy. (€/MWh)","PnL contraint\ntotal (€)","PnL contraint\n(€/MW)",
               "Jours actifs","Taux activation","Description"]
    for c, h in enumerate(headers):
        ws.write(3, c, h, F["h1"])

    row = 4
    for i, (yearly, sc) in enumerate(zip(all_yearly, scenarios)):
        agg = {
            "spread_libre_moy"    : yearly["spread_libre_moy"].mean(),
            "pnl_libre_total"     : yearly["pnl_libre_total"].sum(),
            "pnl_libre_par_MW"    : yearly["pnl_libre_par_MW"].mean(),
            "spread_contraint_moy": yearly["spread_contraint_moy"].mean(),
            "pnl_contraint_total" : yearly["pnl_contraint_total"].sum(),
            "pnl_contraint_par_MW": yearly["pnl_contraint_par_MW"].mean(),
            "jours_actifs"        : yearly["jours_actifs"].sum(),
            "taux_activation"     : yearly["taux_activation"].mean(),
        }
        f0 = F["num0"] if i%2==0 else F["alt0"]
        f2 = F["num2"] if i%2==0 else F["alt2"]
        fs = F["str"]  if i%2==0 else F["alt"]

        ws.write(row, 0,  sc["name"],                          fs)
        ws.write(row, 1,  sc["power_MW"],                      f2)
        ws.write(row, 2,  f"{sc['duration_h']}h",              fs)
        ws.write(row, 3,  round(agg["spread_libre_moy"],    2), f2)
        ws.write(row, 4,  round(agg["pnl_libre_total"],     0), f0)
        ws.write(row, 5,  round(agg["pnl_libre_par_MW"],    2), f2)
        ws.write(row, 6,  round(agg["spread_contraint_moy"],2), f2)
        ws.write(row, 7,  round(agg["pnl_contraint_total"], 0), F["green"] if i%2==0 else F["green"])
        ws.write(row, 8,  round(agg["pnl_contraint_par_MW"],2), f2)
        ws.write(row, 9,  int(agg["jours_actifs"]),             f0)
        ws.write(row, 10, round(agg["taux_activation"],     3), F["pct"])
        ws.write(row, 11, sc["description"],                    fs)
        row += 1

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 2 : DONNÉES PIVOTÉES (24 colonnes)
    # ══════════════════════════════════════════════════════════════════════════
    ws2 = wb.add_worksheet("Prix_spot_pivot")
    ws2.set_column("A:A", 12)
    ws2.set_column("B:D", 8)
    ws2.set_column("E:AB", 10)

    pivot_out = pivot[["date","annee","mois","jour","weekday_name"] + HOUR_COLS].copy()
    pivot_out["date"] = pivot_out["date"].dt.strftime("%Y-%m-%d")

    headers_p = ["Date","Année","Mois","Jour","Jour semaine"] + HOUR_COLS
    for c, h in enumerate(headers_p):
        ws2.write(0, c, h, F["h2"])
    for r, row_data in enumerate(pivot_out.values.tolist(), start=1):
        for c, val in enumerate(row_data):
            ws2.write(r, c, val, F["str"])

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 3 : RÉSULTATS ANNUELS
    # ══════════════════════════════════════════════════════════════════════════
    ws3 = wb.add_worksheet("Résultats_annuels")
    ws3.set_column("A:A", 42)
    ws3.set_column("B:P", 18)

    all_yr_df = pd.concat(all_yearly, ignore_index=True)
    cols_yr = ["scenario","annee","power_MW","duration_h",
               "jours_simules","jours_actifs","jours_bloques_usure","taux_activation",
               "spread_libre_moy","pnl_libre_total","pnl_libre_par_MW",
               "spread_contraint_moy","pnl_contraint_total","pnl_contraint_par_MW",
               "description"]
    all_yr_df = all_yr_df[cols_yr]

    for c, h in enumerate(cols_yr):
        ws3.write(0, c, h, F["h1"])
    for r, row_data in enumerate(all_yr_df.values.tolist(), start=1):
        fmt = F["str"] if r%2==1 else F["alt"]
        for c, val in enumerate(row_data):
            if isinstance(val, float):
                ws3.write(r, c, round(val, 4), fmt)
            else:
                ws3.write(r, c, val, fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 4 : DÉTAIL JOURNALIER
    # ══════════════════════════════════════════════════════════════════════════
    ws4 = wb.add_worksheet("Détail_journalier")
    ws4.set_column("A:A", 42)
    ws4.set_column("B:B", 12)
    ws4.set_column("C:S", 14)

    all_daily_df = pd.concat(all_daily, ignore_index=True)
    cols_d = ["scenario","date","annee","mois","jour","weekday_name",
              "n_heures_dispo","duration_h",
              "spread_libre","pnl_libre",
              "spread_contraint","spread_net","pnl_contraint",
              "valid","usure_bloque",
              "h_charge","h_decharge","prix_charge","prix_decharge",
              "prix_min_jour","prix_max_jour","prix_moy_jour"]
    all_daily_df = all_daily_df[cols_d].copy()
    all_daily_df["date"] = pd.to_datetime(all_daily_df["date"]).dt.strftime("%Y-%m-%d")

    if len(all_daily_df) > 500_000:
        all_daily_df = all_daily_df.head(500_000)

    for c, h in enumerate(cols_d):
        ws4.write(0, c, h, F["h2"])
    for r, row_data in enumerate(all_daily_df.values.tolist(), start=1):
        for c, val in enumerate(row_data):
            if isinstance(val, float):
                ws4.write(r, c, round(val, 4), F["str"])
            elif isinstance(val, bool):
                ws4.write(r, c, str(val), F["str"])
            else:
                ws4.write(r, c, val, F["str"])

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 5 : GRAPHIQUES
    # ══════════════════════════════════════════════════════════════════════════
    ws5 = wb.add_worksheet("Graphiques")
    ws5.set_column("A:A", 2)

    ws5.write("B2", "BESS Valorisation — Analyses graphiques", F["title"])

    # ── Données pour graphique 1 : PnL annuel par scénario ───────────────────
    ws_gdata = wb.add_worksheet("_graph_data")  # feuille intermédiaire cachée

    # PnL contraint par scénario × année
    pnl_pivot = all_yr_df.pivot_table(
        index="annee", columns="scenario", values="pnl_contraint_total", aggfunc="sum"
    )
    annees_list   = pnl_pivot.index.tolist()
    sc_names_list = pnl_pivot.columns.tolist()

    ws_gdata.write(0, 0, "Annee")
    for c, sc in enumerate(sc_names_list, start=1):
        ws_gdata.write(0, c, sc)
    for r, annee in enumerate(annees_list, start=1):
        ws_gdata.write(r, 0, int(annee))
        for c, sc in enumerate(sc_names_list, start=1):
            val = pnl_pivot.loc[annee, sc] if sc in pnl_pivot.columns else 0
            ws_gdata.write(r, c, round(float(val), 0))

    n_annees = len(annees_list)
    n_sc     = len(sc_names_list)

    # ── Graphique 1 : PnL contraint par année et scénario ────────────────────
    chart1 = wb.add_chart({"type": "column"})
    chart1.set_title({"name": "PnL contraint par année et scénario (€)"})
    chart1.set_x_axis({"name": "Année"})
    chart1.set_y_axis({"name": "PnL (€)", "num_format": "#,##0"})
    chart1.set_style(10)
    chart1.set_size({"width": 700, "height": 380})

    colors = ["#1F4E79","#2E75B6","#ED7D31","#A9D18E","#FF0000","#7030A0","#00B0F0"]
    for c_idx in range(n_sc):
        chart1.add_series({
            "name"      : ["_graph_data", 0, c_idx + 1],
            "categories": ["_graph_data", 1, 0, n_annees, 0],
            "values"    : ["_graph_data", 1, c_idx+1, n_annees, c_idx+1],
            "fill"      : {"color": colors[c_idx % len(colors)]},
        })
    ws5.insert_chart("B4", chart1)

    # ── Données graphique 2 : spread moyen par scénario ──────────────────────
    row_g2 = n_annees + 3
    ws_gdata.write(row_g2, 0, "Scenario")
    ws_gdata.write(row_g2, 1, "Spread libre moy (€/MWh)")
    ws_gdata.write(row_g2, 2, "Spread contraint moy (€/MWh)")
    for i, (yearly, sc) in enumerate(zip(all_yearly, scenarios)):
        ws_gdata.write(row_g2 + 1 + i, 0, sc["name"])
        ws_gdata.write(row_g2 + 1 + i, 1, round(float(yearly["spread_libre_moy"].mean()), 2))
        ws_gdata.write(row_g2 + 1 + i, 2, round(float(yearly["spread_contraint_moy"].mean()), 2))

    chart2 = wb.add_chart({"type": "bar"})
    chart2.set_title({"name": "Spread moyen : sans vs avec contraintes (€/MWh)"})
    chart2.set_x_axis({"name": "Spread (€/MWh)"})
    chart2.set_style(10)
    chart2.set_size({"width": 700, "height": 380})
    chart2.add_series({
        "name"      : ["_graph_data", row_g2, 1],
        "categories": ["_graph_data", row_g2+1, 0, row_g2+n_sc, 0],
        "values"    : ["_graph_data", row_g2+1, 1, row_g2+n_sc, 1],
        "fill"      : {"color": "#2E75B6"},
    })
    chart2.add_series({
        "name"      : ["_graph_data", row_g2, 2],
        "categories": ["_graph_data", row_g2+1, 0, row_g2+n_sc, 0],
        "values"    : ["_graph_data", row_g2+1, 2, row_g2+n_sc, 2],
        "fill"      : {"color": "#ED7D31"},
    })
    ws5.insert_chart("B26", chart2)

    # ── Données graphique 3 : profil horaire moyen (heure de charge/décharge) ─
    # Pour le 1er scénario contraint (Case 1), calculer la fréquence de charge par heure
    first_contraint_daily = all_daily[2] if len(all_daily) > 2 else all_daily[0]
    h_freq = {h: 0 for h in range(24)}
    h_dech_freq = {h: 0 for h in range(24)}

    for _, row_d in first_contraint_daily[first_contraint_daily["valid"]].iterrows():
        try:
            hc = eval(row_d["h_charge"])    # liste d'heures
            hd = eval(row_d["h_decharge"])
            for h in hc: h_freq[h] += 1
            for h in hd: h_dech_freq[h] += 1
        except Exception:
            pass

    row_g3 = row_g2 + n_sc + 3
    ws_gdata.write(row_g3, 0, "Heure")
    ws_gdata.write(row_g3, 1, "Fréq. charge")
    ws_gdata.write(row_g3, 2, "Fréq. décharge")
    for h in range(24):
        ws_gdata.write(row_g3 + 1 + h, 0, h)
        ws_gdata.write(row_g3 + 1 + h, 1, h_freq[h])
        ws_gdata.write(row_g3 + 1 + h, 2, h_dech_freq[h])

    chart3 = wb.add_chart({"type": "column"})
    chart3.set_title({"name": f"Profil horaire — fréquence charge/décharge ({scenarios[2]['name'] if len(scenarios)>2 else scenarios[0]['name']})"})
    chart3.set_x_axis({"name": "Heure de la journée"})
    chart3.set_y_axis({"name": "Nombre de jours"})
    chart3.set_style(10)
    chart3.set_size({"width": 700, "height": 380})
    chart3.add_series({
        "name"      : ["_graph_data", row_g3, 1],
        "categories": ["_graph_data", row_g3+1, 0, row_g3+24, 0],
        "values"    : ["_graph_data", row_g3+1, 1, row_g3+24, 1],
        "fill"      : {"color": "#2E75B6"},
    })
    chart3.add_series({
        "name"      : ["_graph_data", row_g3, 2],
        "categories": ["_graph_data", row_g3+1, 0, row_g3+24, 0],
        "values"    : ["_graph_data", row_g3+1, 2, row_g3+24, 2],
        "fill"      : {"color": "#FF0000"},
    })
    ws5.insert_chart("B48", chart3)

    # ── Données graphique 4 : évolution mensuelle PnL (1er scénario contraint) ─
    daily_c1 = all_daily[2] if len(all_daily) > 2 else all_daily[0]
    monthly = daily_c1.groupby(["annee","mois"])["pnl_contraint"].sum().reset_index()
    monthly["label"] = monthly["annee"].astype(str) + "-" + monthly["mois"].astype(str).str.zfill(2)

    row_g4 = row_g3 + 27
    ws_gdata.write(row_g4, 0, "Mois")
    ws_gdata.write(row_g4, 1, "PnL mensuel (€)")
    for i, mrow in enumerate(monthly.itertuples(), start=1):
        ws_gdata.write(row_g4 + i, 0, mrow.label)
        ws_gdata.write(row_g4 + i, 1, round(float(mrow.pnl_contraint), 0))

    n_months = len(monthly)
    chart4 = wb.add_chart({"type": "line"})
    chart4.set_title({"name": f"PnL mensuel — {scenarios[2]['name'] if len(scenarios)>2 else scenarios[0]['name']}"})
    chart4.set_x_axis({"name": "Mois"})
    chart4.set_y_axis({"name": "PnL (€)", "num_format": "#,##0"})
    chart4.set_style(10)
    chart4.set_size({"width": 700, "height": 380})
    chart4.add_series({
        "name"      : ["_graph_data", row_g4, 1],
        "categories": ["_graph_data", row_g4+1, 0, row_g4+n_months, 0],
        "values"    : ["_graph_data", row_g4+1, 1, row_g4+n_months, 1],
        "line"      : {"color": "#1F4E79", "width": 2},
        "marker"    : {"type": "circle", "size": 4, "fill": {"color": "#1F4E79"}},
    })
    ws5.insert_chart("K4", chart4)

    # ── Masquer la feuille de données intermédiaires ──────────────────────────
    ws_gdata.hide()

    # ══════════════════════════════════════════════════════════════════════════
    # FEUILLE 6 : PARAMÈTRES
    # ══════════════════════════════════════════════════════════════════════════
    ws6 = wb.add_worksheet("Paramètres")
    ws6.set_column("A:A", 32)
    ws6.set_column("B:B", 20)
    ws6.set_column("C:C", 55)

    ws6.write("A1", "Paramètres de simulation", F["title"])

    ws6.write(2, 0, "Paramètre batterie", F["h1"])
    ws6.write(2, 1, "Valeur",             F["h1"])
    ws6.write(2, 2, "Description",        F["h1"])
    bat_params = [
        ("energy_MWh",      BATTERY["energy_MWh"],      "Capacité totale (MWh)"),
        ("soc_min",         BATTERY["soc_min_pct"],     "SOC minimum (fraction de la capacité)"),
        ("soc_max",         BATTERY["soc_max_pct"],     "SOC maximum (fraction de la capacité)"),
        ("efficiency",      BATTERY["efficiency"],      "Rendement aller-retour (ex: 0.92 = 92%)"),
        ("max_cycles/an",   BATTERY["max_cycles_year"], "Contrainte usure — None = pas de limite"),
    ]
    for i, (p, v, d) in enumerate(bat_params):
        fmt = F["str"] if i%2==0 else F["alt"]
        ws6.write(3+i, 0, p,       fmt)
        ws6.write(3+i, 1, str(v),  fmt)
        ws6.write(3+i, 2, d,       fmt)

    ws6.write(10, 0, "Scénario",         F["h1"])
    ws6.write(10, 1, "Puissance / Durée",F["h1"])
    ws6.write(10, 2, "Description",      F["h1"])
    for i, sc in enumerate(scenarios):
        fmt = F["str"] if i%2==0 else F["alt"]
        ws6.write(11+i, 0, sc["name"],                        fmt)
        ws6.write(11+i, 1, f"{sc['power_MW']} MW / {sc['duration_h']}h", fmt)
        ws6.write(11+i, 2, sc["description"],                 fmt)

    ws6.write(12+len(scenarios), 0,
              f"Fichier source : {INPUT_EXCEL}", F["str"])
    ws6.write(13+len(scenarios), 0,
              f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}", F["str"])

    wb.close()
    print(f"  ✅ Fichier Excel généré : {output_path}")


# ==============================================================================
# POINT D'ENTRÉE
# ==============================================================================

def main():
    print("=" * 68)
    print("  BESS VALORISATION — Marché Day-Ahead")
    print("=" * 68)

    # 1. Chargement et pivot
    print(f"\n📂 Lecture : {Path(INPUT_EXCEL).name}")
    pivot = load_and_pivot(INPUT_EXCEL, INPUT_SHEET)
    annees = sorted(pivot["annee"].unique())
    print(f"   → {len(pivot):,} jours | années : {annees}")
    print(f"   → Format pivoté : {len(pivot)} lignes × 24 colonnes H00-H23")

    # 2. Simulation
    all_daily  = []
    all_yearly = []

    for sc in SCENARIOS:
        print(f"\n⚡ {sc['name']}")
        daily  = simulate_scenario(pivot, sc)
        yearly = aggregate_yearly(daily, sc)
        all_daily.append(daily)
        all_yearly.append(yearly)

        for _, yr in yearly.iterrows():
            print(f"   {int(yr['annee'])} | "
                  f"Spread libre: {yr['spread_libre_moy']:5.1f} €/MWh | "
                  f"Spread contraint: {yr['spread_contraint_moy']:5.1f} €/MWh | "
                  f"PnL: {yr['pnl_contraint_total']:8,.0f} € | "
                  f"({yr['pnl_contraint_par_MW']:.1f} €/MW) | "
                  f"Actifs: {int(yr['jours_actifs'])}j "
                  f"({'⚠️ usure: '+str(int(yr['jours_bloques_usure']))+'j' if yr['jours_bloques_usure']>0 else ''})")

    # 3. Export
    print(f"\n📊 Export Excel + graphiques...")
    export_excel(all_daily, all_yearly, pivot, SCENARIOS, OUTPUT_EXCEL)

    print("\n" + "=" * 68)
    print("  Simulation terminée.")
    print("=" * 68)


if __name__ == "__main__":
    main()