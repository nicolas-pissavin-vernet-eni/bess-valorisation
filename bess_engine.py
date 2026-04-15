"""
BESS Engine — Moteur de calcul pur
Deux modes : arbitrage DA et lissage de charge.
"""

import pandas as pd
import numpy as np

HOUR_COLS = [f"H{h:02d}" for h in range(24)]


# ──────────────────────────────────────────────────────────────────────────────
# LECTURE & PIVOT
# ──────────────────────────────────────────────────────────────────────────────

def load_spot(source) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="Spot_input", engine="openpyxl")
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
    pivot["date"] = pd.to_datetime(
        pivot[["annee","mois","jour"]].rename(
            columns={"annee":"year","mois":"month","jour":"day"}))
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


def _best_cycle(prices: np.ndarray, hours: list, duration_h: int,
                power_MW: float, efficiency: float,
                forbidden_hours: set = None) -> dict:
    """
    Trouve le meilleur cycle charge+décharge sur un sous-ensemble d'heures.
    forbidden_hours : heures déjà utilisées par un cycle précédent.
    Retourne spread_brut, pnl, heures utilisées, ou None si invalide.
    """
    avail = [h for h in hours if forbidden_hours is None or h not in forbidden_hours]
    n = duration_h
    if len(avail) < 2 * n:
        return None

    av = np.array(avail)
    pa = prices[av]
    asc = np.argsort(pa)

    hc = sorted(av[asc[:n]].tolist())
    hd = sorted(av[asc[-n:]].tolist())

    pc  = prices[hc].mean()
    pd_ = prices[hd].mean()
    spread = pd_ - pc

    if not (max(hc) < min(hd) and spread > 0):
        return None

    pnl = (power_MW * n * efficiency * pd_) - (power_MW * n * pc)
    return dict(spread=round(spread,4), pnl=round(pnl,4),
                h_charge=hc, h_decharge=hd,
                prix_charge=round(pc,4), prix_decharge=round(pd_,4))


def _borne_max_day(prices: np.ndarray, n_cycles: int, duration_h: int,
                   power_MW: float, efficiency: float) -> float:
    """
    Borne max théorique : meilleur PnL SANS contrainte charge<décharge,
    SANS contrainte d'ordre. Simplement les n*duration_h heures les moins
    chères vs les n*duration_h heures les plus chères.
    """
    n = n_cycles * duration_h
    idx = np.argsort(prices)
    hc = idx[:n]
    hd = idx[-n:]
    pc  = prices[hc].mean()
    pd_ = prices[hd].mean()
    spread = pd_ - pc
    if spread <= 0:
        return 0.0
    return (power_MW * n * efficiency * pd_) - (power_MW * n * pc)


def simulate_arbitrage(pivot: pd.DataFrame, params: dict) -> pd.DataFrame:
    """
    params:
      power_MW         : puissance MW
      duration_h       : durée d'un demi-cycle (charge OU décharge) = 1 ou 2h
      n_cycles         : nombre de cycles par jour (1 ou 2)
      excluded_hours   : {"days":[...], "hours":[...]}
      efficiency       : rendement (ex 0.92)
      max_cycles_year  : contrainte usure (None = illimité)
    """
    power    = params["power_MW"]
    dur      = params["duration_h"]
    n_cyc    = params.get("n_cycles", 1)
    excluded = params.get("excluded_hours", {})
    eff      = params.get("efficiency", 0.92)
    max_cy   = params.get("max_cycles_year", None)

    # Capacité calculée automatiquement : puissance × durée × n_cycles
    capacite_MWh = power * dur * n_cyc

    cycles_yr = {}
    records   = []

    for _, row in pivot.iterrows():
        prices  = row[HOUR_COLS].values.astype(float)
        avail   = get_available_hours(int(row["weekday"]), excluded)
        annee   = int(row["annee"])

        # ── Borne max (SANS aucune contrainte) ───────────────────────────────
        pnl_absolu = _borne_max_day(prices, n_cyc, dur, power, eff)
        spread_absolu = (
            np.partition(prices, -n_cyc*dur)[-n_cyc*dur:].mean()
            - np.partition(prices, n_cyc*dur-1)[:n_cyc*dur].mean()
        ) if pnl_absolu > 0 else 0.0

        # ── Cycles avec contraintes ───────────────────────────────────────────
        cycles_result = []
        used_hours    = set()
        total_pnl     = 0.0
        spreads       = []

        for _ in range(n_cyc):
            cy = _best_cycle(prices, avail, dur, power, eff, used_hours)
            if cy is None:
                break
            cycles_result.append(cy)
            used_hours.update(cy["h_charge"] + cy["h_decharge"])
            total_pnl += cy["pnl"]
            spreads.append(cy["spread"])

        n_cy_actifs  = len(cycles_result)
        valid        = n_cy_actifs > 0
        spread_moy   = np.mean(spreads) if spreads else 0.0
        energie_MWh  = power * dur * n_cy_actifs * 2  # charge + décharge

        h_charge_all   = [h for cy in cycles_result for h in cy["h_charge"]]
        h_decharge_all = [h for cy in cycles_result for h in cy["h_decharge"]]
        prix_charge_moy   = np.mean([cy["prix_charge"]   for cy in cycles_result]) if cycles_result else 0
        prix_decharge_moy = np.mean([cy["prix_decharge"] for cy in cycles_result]) if cycles_result else 0

        # Contrainte usure
        usure_bloque = False
        cy_count = cycles_yr.get(annee, 0)
        if max_cy and valid:
            if cy_count + n_cy_actifs > max_cy:
                # Partiellement bloqué
                allowed = max_cy - cy_count
                if allowed <= 0:
                    total_pnl = 0; n_cy_actifs = 0; energie_MWh = 0
                    spread_moy = 0; valid = False; usure_bloque = True
                    h_charge_all = []; h_decharge_all = []
                else:
                    # Garder seulement les 'allowed' premiers cycles
                    total_pnl   = sum(c["pnl"] for c in cycles_result[:allowed])
                    spread_moy  = np.mean([c["spread"] for c in cycles_result[:allowed]])
                    energie_MWh = power * dur * allowed * 2
                    h_charge_all   = [h for c in cycles_result[:allowed] for h in c["h_charge"]]
                    h_decharge_all = [h for c in cycles_result[:allowed] for h in c["h_decharge"]]
                    n_cy_actifs = allowed
                    usure_bloque = True
                cycles_yr[annee] = cy_count + n_cy_actifs
            else:
                cycles_yr[annee] = cy_count + n_cy_actifs
        elif valid:
            cycles_yr[annee] = cy_count + n_cy_actifs

        records.append({
            "date"           : row["date"],
            "annee"          : annee,
            "mois"           : int(row["mois"]),
            "weekday"        : int(row["weekday"]),
            "weekday_name"   : row["weekday_name"],
            "n_heures_dispo" : len(avail),
            "duration_h"     : dur,
            "n_cycles"       : n_cyc,
            "capacite_MWh"   : round(capacite_MWh, 4),
            # Borne max sans aucune contrainte
            "spread_absolu"  : round(spread_absolu, 4),
            "pnl_absolu"     : round(pnl_absolu, 4),
            # Résultats réels avec contraintes
            "spread"         : round(spread_moy, 4),
            "pnl"            : round(total_pnl, 4),
            "valid"          : valid,
            "usure_bloque"   : usure_bloque,
            "n_cycles_actifs": n_cy_actifs,
            "energie_MWh"    : round(energie_MWh, 4),
            "h_charge"       : h_charge_all,
            "h_decharge"     : h_decharge_all,
            "prix_charge"    : round(prix_charge_moy, 4),
            "prix_decharge"  : round(prix_decharge_moy, 4),
            "prix_min"       : float(np.nanmin(prices)),
            "prix_max"       : float(np.nanmax(prices)),
            "prix_moy"       : float(np.nanmean(prices)),
        })

    return pd.DataFrame(records)


def aggregate_arbitrage(daily: pd.DataFrame, power_MW: float) -> pd.DataFrame:
    g = daily.groupby("annee")

    # Spread moyen : calculé UNIQUEMENT sur jours valides (base homogène)
    def spread_moy_valide(x, col):
        mask = x["valid"] & (x[col] > 0)
        return x.loc[mask, col].mean() if mask.any() else 0.0

    out = pd.DataFrame({
        "annee"              : g["annee"].first(),
        "jours_simules"      : g["date"].count(),
        "jours_actifs"       : g["valid"].sum(),
        "jours_bloques_usure": g["usure_bloque"].sum(),
        "taux_activation"    : (g["valid"].sum() / g["date"].count()).round(3),
        # Spreads sur base homogène (jours valides uniquement)
        "spread_absolu_moy"  : g.apply(lambda x: spread_moy_valide(x, "spread_absolu")).round(2),
        "spread_moy"         : g.apply(lambda x: spread_moy_valide(x, "spread")).round(2),
        # PnL
        "pnl_absolu_total"   : g["pnl_absolu"].sum().round(0),
        "pnl_total"          : g["pnl"].sum().round(0),
        "pnl_par_MW"         : (g["pnl"].sum() / (power_MW * 1000)).round(2),
        # Énergie et cycles
        "energie_totale_MWh" : g["energie_MWh"].sum().round(2),
        "cycles_totaux"      : g["n_cycles_actifs"].sum(),
    }).reset_index(drop=True)
    return out


# ──────────────────────────────────────────────────────────────────────────────
# MODE 2 — LISSAGE DE COURBE DE CHARGE
# ──────────────────────────────────────────────────────────────────────────────

def lissage_day(profil_MW, seuil_MW, power_bess, energy_max,
                soc_init, soc_min, soc_max, efficiency):
    profil_lisse = profil_MW.copy().astype(float)
    actions = []
    soc = soc_init
    soc_hist = [soc]
    for h in range(24):
        c = profil_MW[h]
        if c > seuil_MW and soc > soc_min:
            possible = min(c - seuil_MW, power_bess, soc - soc_min)
            soc -= possible
            profil_lisse[h] = c - possible
            actions.append(("decharge", round(possible, 4)))
        elif c < seuil_MW and soc < soc_max:
            possible = min(seuil_MW - c, power_bess, soc_max - soc)
            soc += possible
            profil_lisse[h] = c + possible
            actions.append(("charge", round(possible, 4)))
        else:
            actions.append(("idle", 0))
        soc_hist.append(round(soc, 4))
    reduction_pointe = max(0, profil_MW.max() - profil_lisse.max())
    return {
        "profil_original": profil_MW, "profil_lisse": profil_lisse,
        "actions": actions, "soc_hist": soc_hist, "soc_final": soc,
        "reduction_pointe": round(reduction_pointe, 4),
        "pointe_avant": round(profil_MW.max(), 4),
        "pointe_apres": round(profil_lisse.max(), 4),
    }


def simulate_lissage(pivot, profil_client, params):
    power  = params["power_MW"]
    e_max  = params["energy_MWh"]
    soc_min = params["soc_min_pct"] * e_max
    soc_max = params["soc_max_pct"] * e_max
    soc_init = e_max * params.get("soc_init", 0.5)
    eff    = params.get("efficiency", 0.92)
    seuil_pct = params.get("seuil_percentile", 75)
    tarif_kw  = params.get("tarif_puissance_souscrite", 12000)

    seuil     = np.percentile(profil_client, seuil_pct)
    res_jour  = lissage_day(profil_client, seuil, power, e_max,
                            soc_init, soc_min, soc_max, eff)
    economie_an = res_jour["reduction_pointe"] * tarif_kw

    yearly = [{"annee": int(a),
               "jours": len(pivot[pivot["annee"] == a]),
               "reduction_pointe": res_jour["reduction_pointe"],
               "pointe_avant_MW": res_jour["pointe_avant"],
               "pointe_apres_MW": res_jour["pointe_apres"],
               "economie_an": round(economie_an, 0),
               "seuil_MW": round(seuil, 4)}
              for a in sorted(pivot["annee"].unique())]

    return {"jour_type": res_jour, "seuil_MW": seuil,
            "reduction_MW": res_jour["reduction_pointe"],
            "economie_an": economie_an,
            "yearly": pd.DataFrame(yearly), "params": params}