"""
BESS Engine — Moteur de calcul pur (sans UI)
Deux modes :
  - arbitrage  : achat aux heures creuses, vente aux heures de pointe
  - lissage    : écrêtage des pics de consommation d'un client
"""

import pandas as pd
import numpy as np
from pathlib import Path


# ──────────────────────────────────────────────────────────────────────────────
# LECTURE & PIVOT
# ──────────────────────────────────────────────────────────────────────────────

HOUR_COLS = [f"H{h:02d}" for h in range(24)]


def load_spot(path: str) -> pd.DataFrame:
    """
    Lit le fichier Excel Spot_input et retourne un DataFrame pivoté :
    1 ligne = 1 jour, colonnes H00..H23 = prix spot horaires.
    """
    df = pd.read_excel(path, sheet_name="Spot_input", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={"HEURE":"heure","JOUR":"jour","MOIS":"mois",
                             "ANNEE":"annee","Prix Final":"prix"})
    df = df[["annee","mois","jour","heure","prix"]].dropna()
    for c in ["annee","mois","jour","heure"]:
        df[c] = df[c].astype(int)
    df["prix"] = df["prix"].astype(float)

    pivot = df.pivot_table(index=["annee","mois","jour"],
                           columns="heure", values="prix", aggfunc="first")
    pivot.columns = [f"H{h:02d}" for h in pivot.columns]
    pivot = pivot.reset_index()
    pivot["date"]         = pd.to_datetime(pivot[["annee","mois","jour"]]
                              .rename(columns={"annee":"year","mois":"month","jour":"day"}))
    pivot["weekday"]      = pivot["date"].dt.weekday
    pivot["weekday_name"] = pivot["date"].dt.strftime("%A")
    return pivot.sort_values("date").reset_index(drop=True)


# ──────────────────────────────────────────────────────────────────────────────
# MODE 1 — ARBITRAGE DA
# ──────────────────────────────────────────────────────────────────────────────

def get_available_hours(weekday: int, excluded: dict) -> list:
    if not excluded:
        return list(range(24))
    if weekday in excluded.get("days", []):
        return [h for h in range(24) if h not in excluded.get("hours", [])]
    return list(range(24))


def arbitrage_day(prices: np.ndarray, avail: list,
                  duration_h: int, power_MW: float, efficiency: float) -> dict:
    """Calcule le meilleur arbitrage pour 1 journée."""
    p = prices
    n = duration_h
    if len(avail) < 2 * n:
        return dict(spread_brut=0, spread_net=0, pnl=0, valid=False,
                    h_charge=[], h_decharge=[], prix_charge=0, prix_decharge=0)

    av = np.array(avail)
    pa = p[av]
    asc  = np.argsort(pa)
    desc = asc[::-1]

    hc = sorted(av[asc[:n]].tolist())
    hd = sorted(av[desc[:n]].tolist())

    pc = p[hc].mean()
    pd_ = p[hd].mean()
    spread = pd_ - pc

    valid = (max(hc) < min(hd)) and (spread > 0)
    if not valid:
        return dict(spread_brut=0, spread_net=0, pnl=0, valid=False,
                    h_charge=[], h_decharge=[], prix_charge=0, prix_decharge=0)

    pnl = (power_MW * n * efficiency * pd_) - (power_MW * n * pc)
    return dict(spread_brut=round(spread, 4),
                spread_net=round(pd_ * efficiency - pc, 4),
                pnl=round(pnl, 4), valid=True,
                h_charge=hc, h_decharge=hd,
                prix_charge=round(pc, 4), prix_decharge=round(pd_, 4))


def simulate_arbitrage(pivot: pd.DataFrame, params: dict) -> pd.DataFrame:
    """
    Simule l'arbitrage DA sur toutes les journées.
    params: power_MW, duration_h, excluded_hours, efficiency, max_cycles_year
    """
    power      = params["power_MW"]
    duration   = params["duration_h"]
    excluded   = params.get("excluded_hours", {})
    eff        = params.get("efficiency", 0.92)
    max_cy     = params.get("max_cycles_year", None)
    cycles_yr  = {}
    records    = []

    for _, row in pivot.iterrows():
        prices  = row[HOUR_COLS].values.astype(float)
        avail   = get_available_hours(int(row["weekday"]), excluded)
        annee   = int(row["annee"])

        # Borne max (sans contrainte horaire)
        res_libre = arbitrage_day(prices, list(range(24)), duration, power, eff)
        # Avec contraintes
        res       = arbitrage_day(prices, avail, duration, power, eff)

        # Contrainte usure
        usure_bloque = False
        cy = cycles_yr.get(annee, 0)
        if max_cy and res["valid"]:
            if cy >= max_cy:
                res = dict(spread_brut=0, spread_net=0, pnl=0, valid=False,
                           h_charge=[], h_decharge=[], prix_charge=0, prix_decharge=0)
                usure_bloque = True
            else:
                cycles_yr[annee] = cy + 1
        elif res["valid"]:
            cycles_yr[annee] = cy + 1

        records.append({
            "date"          : row["date"],
            "annee"         : annee,
            "mois"          : int(row["mois"]),
            "weekday"       : int(row["weekday"]),
            "weekday_name"  : row["weekday_name"],
            "n_heures_dispo": len(avail),
            "spread_libre"  : res_libre["spread_brut"],
            "pnl_libre"     : res_libre["pnl"],
            "spread"        : res["spread_brut"],
            "spread_net"    : res["spread_net"],
            "pnl"           : res["pnl"],
            "valid"         : res["valid"],
            "usure_bloque"  : usure_bloque,
            "h_charge"      : res["h_charge"],
            "h_decharge"    : res["h_decharge"],
            "prix_charge"   : res["prix_charge"],
            "prix_decharge" : res["prix_decharge"],
            "prix_min"      : float(np.nanmin(prices)),
            "prix_max"      : float(np.nanmax(prices)),
            "prix_moy"      : float(np.nanmean(prices)),
        })

    return pd.DataFrame(records)


# ──────────────────────────────────────────────────────────────────────────────
# MODE 2 — LISSAGE DE COURBE DE CHARGE
# ──────────────────────────────────────────────────────────────────────────────

def lissage_day(profil_MW: np.ndarray, seuil_MW: float,
                power_bess: float, energy_max: float,
                soc_init: float, soc_min: float, soc_max: float,
                efficiency: float) -> dict:
    """
    Lisse le profil de consommation d'un client sur 1 journée.

    Logique :
      - Si conso > seuil  → décharger la batterie (écrêter le pic)
      - Si conso < seuil  → charger la batterie (stocker pour plus tard)
      - Contraintes : SOC min/max, puissance max BESS

    Retourne le profil lissé, les actions BESS, le SOC heure par heure,
    et la réduction de pointe obtenue.
    """
    profil_lisse = profil_MW.copy().astype(float)
    actions      = []
    soc          = soc_init
    soc_hist     = [soc]
    energie_ch   = 0.0
    energie_dch  = 0.0

    for h in range(24):
        c = profil_MW[h]

        if c > seuil_MW and soc > soc_min:
            # Pic → décharger
            possible = min(c - seuil_MW, power_bess, soc - soc_min)
            soc -= possible
            profil_lisse[h] = c - possible
            actions.append(("decharge", round(possible, 4)))
            energie_dch += possible
        elif c < seuil_MW and soc < soc_max:
            # Creux → charger
            possible = min(seuil_MW - c, power_bess, soc_max - soc)
            soc += possible
            profil_lisse[h] = c + possible
            actions.append(("charge", round(possible, 4)))
            energie_ch += possible
        else:
            actions.append(("idle", 0))

        soc_hist.append(round(soc, 4))

    reduction_pointe = max(0, profil_MW.max() - profil_lisse.max())

    return {
        "profil_original" : profil_MW,
        "profil_lisse"    : profil_lisse,
        "actions"         : actions,
        "soc_hist"        : soc_hist,
        "soc_final"       : soc,
        "reduction_pointe": round(reduction_pointe, 4),
        "energie_chargee" : round(energie_ch, 4),
        "energie_dechargee": round(energie_dch, 4),
        "pointe_avant"    : round(profil_MW.max(), 4),
        "pointe_apres"    : round(profil_lisse.max(), 4),
    }


def simulate_lissage(pivot: pd.DataFrame, profil_client: np.ndarray,
                     params: dict) -> dict:
    """
    Simule le lissage sur toutes les journées.
    Le profil client (24 valeurs en MW) est supposé répétitif chaque jour
    (ou peut être saisonnalisé dans une version future).

    params: power_MW, energy_MWh, soc_min_pct, soc_max_pct,
            seuil_pct (percentile de la courbe pour définir le seuil),
            efficiency, tarif_puissance_souscrite (€/kW/an)
    """
    power      = params["power_MW"]
    e_max      = params["energy_MWh"]
    soc_min    = params["soc_min_pct"] * e_max
    soc_max    = params["soc_max_pct"] * e_max
    soc_init   = e_max * 0.5
    eff        = params.get("efficiency", 0.92)
    seuil_pct  = params.get("seuil_percentile", 75)
    tarif_kw   = params.get("tarif_puissance_souscrite", 12000)  # €/MW/an

    # Seuil = percentile de la courbe de charge client
    seuil = np.percentile(profil_client, seuil_pct)

    # Simuler 1 jour type (le profil est supposé constant)
    res_jour = lissage_day(profil_client, seuil, power, e_max,
                           soc_init, soc_min, soc_max, eff)

    reduction_mw  = res_jour["reduction_pointe"]
    # Économie annuelle = réduction de puissance souscrite
    economie_an   = reduction_mw * tarif_kw

    # Synthèse annuelle (même profil chaque jour = simplification)
    n_jours = len(pivot)
    n_annees = pivot["annee"].nunique()

    yearly = []
    for annee in sorted(pivot["annee"].unique()):
        n_j = len(pivot[pivot["annee"] == annee])
        yearly.append({
            "annee"           : int(annee),
            "jours"           : n_j,
            "reduction_pointe": reduction_mw,
            "pointe_avant_MW" : res_jour["pointe_avant"],
            "pointe_apres_MW" : res_jour["pointe_apres"],
            "economie_an"     : round(economie_an, 0),
            "seuil_MW"        : round(seuil, 4),
        })

    return {
        "jour_type"    : res_jour,
        "seuil_MW"     : seuil,
        "reduction_MW" : reduction_mw,
        "economie_an"  : economie_an,
        "yearly"       : pd.DataFrame(yearly),
        "params"       : params,
    }


# ──────────────────────────────────────────────────────────────────────────────
# AGRÉGATION ARBITRAGE
# ──────────────────────────────────────────────────────────────────────────────

def aggregate_arbitrage(daily: pd.DataFrame, power_MW: float) -> pd.DataFrame:
    g = daily.groupby("annee")
    out = pd.DataFrame({
        "annee"               : g["annee"].first(),
        "jours_simules"       : g["date"].count(),
        "jours_actifs"        : g["valid"].sum(),
        "jours_bloques_usure" : g["usure_bloque"].sum(),
        "taux_activation"     : (g["valid"].sum() / g["date"].count()).round(3),
        "spread_libre_moy"    : g["spread_libre"].mean().round(2),
        "pnl_libre_total"     : g["pnl_libre"].sum().round(0),
        "spread_moy"          : g.apply(lambda x: x.loc[x["valid"],"spread"].mean()
                                         if x["valid"].any() else 0).round(2),
        "pnl_total"           : g["pnl"].sum().round(0),
        "pnl_par_MW"          : (g["pnl"].sum() / (power_MW * 1000)).round(2),
    }).reset_index(drop=True)
    return out