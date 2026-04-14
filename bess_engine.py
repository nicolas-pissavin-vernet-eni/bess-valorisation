"""
BESS Engine — Moteur de calcul pur (sans UI)
Deux modes :
  - arbitrage  : achat aux heures creuses, vente aux heures de pointe
  - lissage    : écrêtage des pics de consommation d'un client
"""

import pandas as pd
import numpy as np

HOUR_COLS = [f"H{h:02d}" for h in range(24)]


def load_spot(source) -> pd.DataFrame:
    """
    Lit le fichier Excel Spot_input (chemin str ou BytesIO depuis upload Streamlit).
    Retourne un DataFrame pivoté : 1 ligne = 1 jour, colonnes H00..H23.
    """
    df = pd.read_excel(source, sheet_name="Spot_input", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={
        "HEURE": "heure", "JOUR": "jour", "MOIS": "mois",
        "ANNEE": "annee", "Prix Final": "prix"
    })
    df = df[["annee", "mois", "jour", "heure", "prix"]].dropna()
    for c in ["annee", "mois", "jour", "heure"]:
        df[c] = df[c].astype(int)
    df["prix"] = df["prix"].astype(float)

    pivot = df.pivot_table(
        index=["annee", "mois", "jour"],
        columns="heure", values="prix", aggfunc="first"
    )
    pivot.columns = [f"H{h:02d}" for h in pivot.columns]
    pivot = pivot.reset_index()
    pivot["date"] = pd.to_datetime(
        pivot[["annee", "mois", "jour"]].rename(
            columns={"annee": "year", "mois": "month", "jour": "day"}
        )
    )
    pivot["weekday"] = pivot["date"].dt.weekday
    pivot["weekday_name"] = pivot["date"].dt.strftime("%A")
    return pivot.sort_values("date").reset_index(drop=True)


def get_available_hours(weekday: int, excluded: dict) -> list:
    if not excluded:
        return list(range(24))
    if weekday in excluded.get("days", []):
        return [h for h in range(24) if h not in excluded.get("hours", [])]
    return list(range(24))


def arbitrage_day(prices, avail, duration_h, power_MW, efficiency):
    n = duration_h
    if len(avail) < 2 * n:
        return dict(spread_brut=0, spread_net=0, pnl=0, valid=False,
                    h_charge=[], h_decharge=[], prix_charge=0, prix_decharge=0)
    av = np.array(avail)
    asc = np.argsort(prices[av])
    hc = sorted(av[asc[:n]].tolist())
    hd = sorted(av[asc[-n:]].tolist())
    pc = prices[hc].mean()
    pd_ = prices[hd].mean()
    spread = pd_ - pc
    valid = (max(hc) < min(hd)) and (spread > 0)
    if not valid:
        return dict(spread_brut=0, spread_net=0, pnl=0, valid=False,
                    h_charge=[], h_decharge=[], prix_charge=0, prix_decharge=0)
    pnl = (power_MW * n * efficiency * pd_) - (power_MW * n * pc)
    return dict(spread_brut=round(spread, 4), spread_net=round(pd_ * efficiency - pc, 4),
                pnl=round(pnl, 4), valid=True, h_charge=hc, h_decharge=hd,
                prix_charge=round(pc, 4), prix_decharge=round(pd_, 4))


def simulate_arbitrage(pivot, params):
    power = params["power_MW"]
    duration = params["duration_h"]
    excluded = params.get("excluded_hours", {})
    eff = params.get("efficiency", 0.92)
    max_cy = params.get("max_cycles_year", None)
    cycles_yr = {}
    records = []

    for _, row in pivot.iterrows():
        prices = row[HOUR_COLS].values.astype(float)
        avail = get_available_hours(int(row["weekday"]), excluded)
        annee = int(row["annee"])
        res_libre = arbitrage_day(prices, list(range(24)), duration, power, eff)
        res = arbitrage_day(prices, avail, duration, power, eff)

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
            "date": row["date"], "annee": annee, "mois": int(row["mois"]),
            "weekday": int(row["weekday"]), "weekday_name": row["weekday_name"],
            "n_heures_dispo": len(avail), "duration_h": duration,
            "spread_libre": res_libre["spread_brut"], "pnl_libre": res_libre["pnl"],
            "spread": res["spread_brut"], "spread_net": res["spread_net"],
            "pnl": res["pnl"], "valid": res["valid"], "usure_bloque": usure_bloque,
            "h_charge": res["h_charge"], "h_decharge": res["h_decharge"],
            "prix_charge": res["prix_charge"], "prix_decharge": res["prix_decharge"],
            "prix_min": float(np.nanmin(prices)), "prix_max": float(np.nanmax(prices)),
            "prix_moy": float(np.nanmean(prices)),
        })
    return pd.DataFrame(records)


def aggregate_arbitrage(daily, power_MW):
    g = daily.groupby("annee")
    out = pd.DataFrame({
        "annee": g["annee"].first(),
        "jours_simules": g["date"].count(),
        "jours_actifs": g["valid"].sum(),
        "jours_bloques_usure": g["usure_bloque"].sum(),
        "taux_activation": (g["valid"].sum() / g["date"].count()).round(3),
        "spread_libre_moy": g["spread_libre"].mean().round(2),
        "pnl_libre_total": g["pnl_libre"].sum().round(0),
        "spread_moy": g.apply(
            lambda x: x.loc[x["valid"], "spread"].mean() if x["valid"].any() else 0
        ).round(2),
        "pnl_total": g["pnl"].sum().round(0),
        "pnl_par_MW": (g["pnl"].sum() / (power_MW * 1000)).round(2),
    }).reset_index(drop=True)
    return out


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
    power = params["power_MW"]
    e_max = params["energy_MWh"]
    soc_min = params["soc_min_pct"] * e_max
    soc_max = params["soc_max_pct"] * e_max
    soc_init = e_max * params.get("soc_init", 0.5)
    eff = params.get("efficiency", 0.92)
    seuil_pct = params.get("seuil_percentile", 75)
    tarif_kw = params.get("tarif_puissance_souscrite", 12000)

    seuil = np.percentile(profil_client, seuil_pct)
    res_jour = lissage_day(profil_client, seuil, power, e_max,
                           soc_init, soc_min, soc_max, eff)
    economie_an = res_jour["reduction_pointe"] * tarif_kw

    yearly = [{"annee": int(a), "jours": len(pivot[pivot["annee"] == a]),
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