"""
BESS Engine — Moteur de calcul pur
Deux modes : arbitrage DA et lissage de charge.

Logique de cycle (arbitrage) :
  - La batterie repart à vide à minuit (H00) chaque jour
  - 1 ou 2 cycles séquentiels dans la journée (H00–H23)
  - Un cycle = n heures de charge consécutives ou non,
    SUIVIES de n heures de décharge (toutes après la dernière heure de charge)
  - Le cycle 2 ne peut commencer (1ère heure de charge)
    qu'après la fin du cycle 1 (dernière heure de décharge)
  - Les heures ne doivent pas chevaucher les restrictions
"""

import pandas as pd
import numpy as np
from itertools import combinations as _comb

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


def load_imbalance(source) -> dict:
    """
    Charge les prix de règlement des écarts (imbalance settlement price).
    Source attendue : Excel avec une colonne Date (horodatage) + une colonne
    contenant "negative" et/ou une colonne contenant "positive" dans son nom
    (ex. negative_imbalance_settlement_price, positive_imbalance_settlement_price).
    Le pas de temps peut être demi-horaire ou quart-horaire (ou un mélange,
    ex. 30 min jusqu'en 2024 puis 15 min depuis 2025 sur les marchés qui ont
    migré) — tout est ramené à la moyenne horaire pour rejoindre le pivot
    standard H00..H23 utilisé par simulate_arbitrage / simulate_arbitrage_optimal.
    Retourne {"negatif": pivot, "positif": pivot} — l'une des deux clés peut
    être un DataFrame vide si la colonne correspondante est absente du fichier.
    """
    df = pd.read_excel(source, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    date_col = next((c for c in df.columns if c.strip().lower() == "date"), None)
    neg_col  = next((c for c in df.columns if "negative" in c.lower()), None)
    pos_col  = next((c for c in df.columns if "positive" in c.lower()), None)
    if date_col is None or (neg_col is None and pos_col is None):
        return {"negatif": pd.DataFrame(), "positif": pd.DataFrame()}

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])
    df["_h"] = df[date_col].dt.floor("h")

    def _build(col):
        if col is None:
            return pd.DataFrame()
        h = df.groupby("_h")[col].mean().reset_index()
        h["annee"] = h["_h"].dt.year
        h["mois"]  = h["_h"].dt.month
        h["jour"]  = h["_h"].dt.day
        h["heure"] = h["_h"].dt.hour
        pivot = h.pivot_table(index=["annee", "mois", "jour"], columns="heure",
                               values=col, aggfunc="first")
        pivot.columns = [f"H{int(c):02d}" for c in pivot.columns]
        pivot = pivot.reset_index()
        pivot["date"] = pd.to_datetime(
            pivot[["annee", "mois", "jour"]].rename(
                columns={"annee": "year", "mois": "month", "jour": "day"}))
        pivot["weekday"]      = pivot["date"].dt.weekday
        pivot["weekday_name"] = pivot["date"].dt.strftime("%A")
        return pivot.sort_values("date").reset_index(drop=True)

    return {"negatif": _build(neg_col), "positif": _build(pos_col)}


def load_case3(source) -> pd.DataFrame:
    """
    Charge les données historiques 2019-2025 depuis :
    - soit un fichier avec une feuille dont le nom contient 'case 3'
    - soit un fichier Excel dont la première feuille contient directement les données
    Structure attendue : colonnes Delivery, Année, 0..23 (prix H00-H23)
    """
    import openpyxl as _opxl
    import io as _io

    src_bytes = source if isinstance(source, (bytes, bytearray)) else None

    # Détecter la feuille à lire
    try:
        if src_bytes:
            wb = _opxl.load_workbook(_io.BytesIO(src_bytes), read_only=True)
        else:
            wb = _opxl.load_workbook(source, read_only=True)
        # Chercher une feuille avec "case 3" dans le nom
        sheet_name = next((s for s in wb.sheetnames if "case 3" in s.lower()), None)
        # Si pas trouvé, prendre la première feuille
        if sheet_name is None:
            sheet_name = wb.sheetnames[0]
        wb.close()
    except Exception:
        sheet_name = 0  # première feuille par défaut

    if src_bytes:
        df = pd.read_excel(_io.BytesIO(src_bytes), sheet_name=sheet_name, engine="openpyxl")
    else:
        df = pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl")

    # Normaliser la colonne de date
    # Structure Case 3 : colonne "Delivery" (datetime) + "Année" (int) + 0..23 (prix)
    if "Delivery" in df.columns:
        df["Delivery"] = pd.to_datetime(df["Delivery"])
        df["annee"] = df["Delivery"].dt.year.astype(int)
        df["mois"]  = df["Delivery"].dt.month.astype(int)
        df["jour"]  = df["Delivery"].dt.day.astype(int)
    elif "annee" in df.columns:
        pass  # déjà bon
    else:
        raise ValueError(f"Structure non reconnue. Colonnes : {df.columns.tolist()[:10]}")

    # Filtrer 2019-2025
    df = df[df["annee"] <= 2025].copy()

    # Colonnes heures : entiers 0..23
    hour_cols = {c for c in df.columns if isinstance(c, int) and 0 <= c <= 23}

    rows = []
    for _, row in df.iterrows():
        r = {"annee": int(row["annee"]), "mois": int(row["mois"]), "jour": int(row["jour"])}
        for h in range(24):
            r[f"H{h:02d}"] = float(row[h]) if h in hour_cols else np.nan
        rows.append(r)

    pivot = pd.DataFrame(rows)
    pivot["date"] = pd.to_datetime(
        pivot[["annee","mois","jour"]].rename(
            columns={"annee":"year","mois":"month","jour":"day"}))
    pivot["weekday"]      = pivot["date"].dt.weekday
    pivot["weekday_name"] = pivot["date"].dt.strftime("%A")
    return pivot.sort_values("date").reset_index(drop=True)


# ──────────────────────────────────────────────────────────────────────────────
# ARBITRAGE — FONCTIONS CŒUR
# ──────────────────────────────────────────────────────────────────────────────

def get_available_hours(weekday: int, excluded: dict) -> list:
    if not excluded:
        return list(range(24))
    if weekday in excluded.get("days", []):
        return [h for h in range(24) if h not in excluded.get("hours", [])]
    return list(range(24))


def _best_single_cycle(prices: np.ndarray,
                       avail_hours: list,
                       duration_h: int,
                       power_MW: float,
                       efficiency: float,
                       h_start: int = 0) -> dict | None:
    """
    Trouve le meilleur cycle charge+décharge sur les heures disponibles
    à partir de h_start.

    Règles :
    - n heures de charge parmi avail_hours >= h_start
    - n heures de décharge TOUTES après max(h_charge)
    - Maximise le PnL

    Implémentation vectorisée numpy :
    - n=1 : O(k²) avec k=len(hours), entièrement en numpy
    - n=2 : boucle sur C(k_cand,2) combinations avec k_cand ≤ 8
    """
    n = duration_h
    hours = [h for h in avail_hours if h >= h_start]
    if len(hours) < 2 * n:
        return None

    av = np.array(hours, dtype=np.int32)
    pa = prices[av]   # prix aux heures disponibles (indexés dans av)

    if n == 1:
        # ── Cas n=1 : vectorisé pur ──────────────────────────────────────────
        # Pour chaque heure de charge i, la meilleure décharge est max(pa[j]) avec j>i.
        # suffix_max[i] = max(pa[i], pa[i+1], ..., pa[-1])
        k = len(av)
        suffix_max_val = np.maximum.accumulate(pa[::-1])[::-1]  # max depuis i jusqu'à fin
        suffix_max_idx = np.zeros(k, dtype=np.int32)
        # Pour chaque i, trouver l'index j>i avec pa[j] = max(pa[i+1:])
        # On recalcule via argmax à droite
        best_pnl    = 0.0
        best_result = None
        for i in range(k - 1):
            pc = pa[i]
            # Meilleure décharge : max des prix après i
            right = pa[i + 1:]
            j_rel = int(np.argmax(right))
            pd_ = right[j_rel]
            spread = pd_ - pc
            if spread <= 0:
                continue
            pnl = power_MW * efficiency * pd_ - power_MW * pc
            if pnl > best_pnl:
                best_pnl = pnl
                hc = [int(av[i])]
                hd = [int(av[i + 1 + j_rel])]
                best_result = dict(
                    spread        = round(float(spread), 4),
                    pnl           = round(float(pnl), 4),
                    h_charge      = hc,
                    h_decharge    = hd,
                    prix_charge   = round(float(pc), 4),
                    prix_decharge = round(float(pd_), 4),
                    h_end         = hd[0],
                )
        return best_result

    # ── Cas n>=2 : vectorisé pour n=2, boucle pour n>2 ──────────────────────
    n_cand = min(2 * n + 6, len(av))
    idx_asc = np.argsort(pa)
    cand_av = av[idx_asc[:n_cand]]        # heures candidates (les moins chères)
    cand_pa = pa[idx_asc[:n_cand]]

    if n == 2:
        best_pnl    = 0.0
        best_result = None
        k = len(cand_av)
        for j in range(1, k):
            hj = int(cand_av[j])
            for i in range(j):
                hi = int(cand_av[i])
                # Les deux heures de charge sont hi et hj (hi < hj car cand_av trié par prix,
                # mais on veut max_hc = max(hi,hj) pour la contrainte de décharge)
                max_hc = max(hi, hj)
                pc = (prices[hi] + prices[hj]) / 2.0
                # Décharge : 2 heures APRÈS max(hi,hj), non utilisées pour la charge
                avail_dch = [h for h in hours if h > max_hc and h != hi and h != hj]
                if len(avail_dch) < 2:
                    continue
                pa_dch = prices[np.array(avail_dch)]
                top2   = np.argsort(pa_dch)[-2:]
                hd     = sorted([avail_dch[t] for t in top2])
                pd_    = prices[np.array(hd)].mean()
                spread = pd_ - pc
                if spread <= 0:
                    continue
                pnl = power_MW * 2 * efficiency * pd_ - power_MW * 2 * pc
                if pnl > best_pnl:
                    best_pnl    = pnl
                    best_result = dict(
                        spread        = round(float(spread), 4),
                        pnl           = round(float(pnl), 4),
                        h_charge      = sorted([hi, hj]),
                        h_decharge    = hd,
                        prix_charge   = round(float(pc), 4),
                        prix_decharge = round(float(pd_), 4),
                        h_end         = max(hd),
                    )
        return best_result

    # n > 2 : boucle générique sur combinations
    candidates = sorted(cand_av.tolist())
    best_pnl    = 0.0
    best_result = None

    for hc in _comb(candidates, n):
        hc     = list(hc)
        hc_set = set(hc)
        max_hc = max(hc)

        avail_dch = [h for h in hours if h > max_hc and h not in hc_set]
        if len(avail_dch) < n:
            continue

        pa_dch = prices[np.array(avail_dch)]
        hd     = sorted(np.array(avail_dch)[np.argsort(pa_dch)[-n:]].tolist())

        pc  = prices[np.array(hc)].mean()
        pd_ = prices[np.array(hd)].mean()
        spread = pd_ - pc
        if spread <= 0:
            continue

        pnl = (power_MW * n * efficiency * pd_) - (power_MW * n * pc)
        if pnl > best_pnl:
            best_pnl    = pnl
            best_result = dict(
                spread        = round(float(spread), 4),
                pnl           = round(float(pnl), 4),
                h_charge      = hc,
                h_decharge    = hd,
                prix_charge   = round(float(pc), 4),
                prix_decharge = round(float(pd_), 4),
                h_end         = max(hd),
            )

    return best_result


def _best_n_cycles(prices: np.ndarray, avail_hours: list, n_cycles: int,
                   duration_h: int, power_MW: float, efficiency: float) -> list:
    """
    Simule n_cycles cycles sequentiels sur la journee.
    - n_cycles=1 : meilleur cycle sur toute la journee
    - n_cycles=2 : teste toutes les coupures h* ; choisit la coupure maximisant PnL1+PnL2
    - n_cycles=0 : MAX — enchaîne autant de cycles rentables que les heures disponibles
                   permettent, séquentiellement, jusqu'à épuisement des créneaux
    """
    n = duration_h

    if n_cycles == 1:
        cy = _best_single_cycle(prices, avail_hours, n, power_MW, efficiency, 0)
        return [cy] if cy else []

    if n_cycles == 0:
        # Mode MAX — DP exact (Weighted Interval Scheduling).
        #
        # Pour n=1 : chaque cycle est (hc, hd) avec hc < hd.
        #   On génère 1 candidat par hc (best hd = argmax prix après hc) → O(24) candidats.
        # Pour n=2 : chaque cycle est ([hc1,hc2], [hd1,hd2]).
        #   On génère 1 candidat par hc_max (best hd en cherchant depuis hc_max+1) → O(24) candidats.
        # DP classique par intervalle [min(hc), max(hd)].

        A = list(avail_hours)
        if len(A) < 2 * n:
            return []

        # 1. Générer les candidats
        seen = {}

        if n == 1:
            # Pour n=1 : deux types de candidats
            # a) Tous les cycles consécutifs H_i → H_{i+1} (capturent les petits spreads)
            # b) Pour chaque heure de charge H_i, la meilleure décharge possible
            # Cela garantit qu'on ne rate ni les petits cycles adjacents
            # ni les grands cycles H_bas → H_haut
            for si in range(len(A) - 1):
                hc = A[si]
                # a) Cycle consécutif hc → hc+1 (si hc+1 est dans A)
                if A[si+1] == hc + 1 or si+1 < len(A):
                    hd_next = A[si+1]
                    spread = prices[hd_next] - prices[hc]
                    if spread > 0:
                        pnl = power_MW * efficiency * prices[hd_next] - power_MW * prices[hc]
                        key = (hc, hd_next)
                        if key not in seen or pnl > seen[key]['pnl']:
                            seen[key] = {
                                'spread': round(float(spread), 4),
                                'pnl': round(float(pnl), 4),
                                'h_charge': [hc], 'h_decharge': [hd_next],
                                'prix_charge': round(float(prices[hc]), 4),
                                'prix_decharge': round(float(prices[hd_next]), 4),
                                'h_end': hd_next,
                            }
                # b) Meilleur cycle depuis hc
                cy = _best_single_cycle(prices, A[si:], n, power_MW, efficiency, hc)
                if cy and cy['pnl'] > 0:
                    key = (min(cy['h_charge']), max(cy['h_decharge']))
                    if key not in seen or cy['pnl'] > seen[key]['pnl']:
                        seen[key] = cy
        else:
            for si in range(len(A) - 2*n + 1):
                cy = _best_single_cycle(prices, A[si:], n, power_MW, efficiency, A[si])
                if cy and cy['pnl'] > 0:
                    key = (min(cy['h_charge']), max(cy['h_decharge']))
                    if key not in seen or cy['pnl'] > seen[key]['pnl']:
                        seen[key] = cy
                window_end = A[:si + 2*n]
                cy2 = _best_single_cycle(prices, window_end, n, power_MW, efficiency, 0)
                if cy2 and cy2['pnl'] > 0:
                    key2 = (min(cy2['h_charge']), max(cy2['h_decharge']))
                    if key2 not in seen or cy2['pnl'] > seen[key2]['pnl']:
                        seen[key2] = cy2

        if not seen:
            return []

        # 2. DP par intervalle
        cands = sorted(seen.values(), key=lambda c: max(c['h_decharge']))
        N_c   = len(cands)
        ends   = [max(c['h_decharge']) for c in cands]
        starts = [min(c['h_charge'])   for c in cands]
        pnls   = [c['pnl']            for c in cands]

        def last_compat(i):
            lo, hi, res = 0, i-1, -1
            while lo <= hi:
                mid = (lo+hi)//2
                if ends[mid] < starts[i]: res = mid; lo = mid+1
                else: hi = mid-1
            return res

        dp = [0.0] * (N_c + 1)
        for i in range(N_c):
            p = last_compat(i)
            dp[i+1] = max(dp[i], (dp[p+1] if p >= 0 else 0.0) + pnls[i])

        # Reconstruction
        result, i = [], N_c - 1
        while i >= 0:
            p = last_compat(i)
            if abs(dp[i+1] - ((dp[p+1] if p >= 0 else 0.0) + pnls[i])) < 1e-9:
                result.append(cands[i]); i = p
            else:
                i -= 1
        return result[::-1]

    # n_cycles=2 : coupure optimale
    best_total  = None
    best_cycles = []

    for cut_h in range(n * 2 - 1, 24 - n * 2):
        avail1 = [h for h in avail_hours if h <= cut_h]
        avail2 = [h for h in avail_hours if h > cut_h]
        if len(avail1) < 2*n or len(avail2) < 2*n:
            continue
        cy1 = _best_single_cycle(prices, avail1, n, power_MW, efficiency, 0)
        cy2 = _best_single_cycle(prices, avail2, n, power_MW, efficiency, avail2[0])
        if cy1 and cy2:
            total = cy1['pnl'] + cy2['pnl']
            if best_total is None or total > best_total:
                best_total  = total
                best_cycles = [cy1, cy2]

    if not best_cycles:
        cy = _best_single_cycle(prices, avail_hours, n, power_MW, efficiency, 0)
        return [cy] if cy else []

    return best_cycles


def _borne_max_day(prices: np.ndarray, n_cycles: int, duration_h: int,
                   power_MW: float, efficiency: float) -> float:
    """Borne max théorique : n_cycles optimaux sur journée entière sans restriction."""
    cycles = _best_n_cycles(prices, list(range(24)), n_cycles, duration_h, power_MW, efficiency)
    return sum(c['pnl'] for c in cycles)


# ──────────────────────────────────────────────────────────────────────────────
# SIMULATION ARBITRAGE
# ──────────────────────────────────────────────────────────────────────────────

def simulate_arbitrage(pivot: pd.DataFrame, params: dict) -> pd.DataFrame:
    """
    Simule l'arbitrage day-ahead avec cycles séquentiels.

    Logique :
    - Batterie vide à H00 chaque jour
    - n_cycles cycles séquentiels dans la journée (H00–H23)
    - Cycle 2 commence après la fin (dernière heure décharge) du cycle 1
    - Chaque cycle : charge n heures, puis décharge n heures (toutes après max charge)
    - Restrictions horaires et quota annuel respectés
    """
    power    = params["power_MW"]
    dur      = params["duration_h"]
    n_cyc    = params.get("n_cycles", 1)
    excluded = params.get("excluded_hours", {})
    eff      = params.get("efficiency", 0.92)
    max_cy   = params.get("max_cycles_year", None)
    min_sp   = params.get("min_spread", None)  # spread minimum en €/MWh

    capacite_MWh = power * dur  # capacité nominale = puissance × durée 1 cycle

    cycles_yr = {}
    records   = []

    for _, row in pivot.iterrows():
        prices = row[HOUR_COLS].values.astype(float)
        avail  = get_available_hours(int(row["weekday"]), excluded)
        annee  = int(row["annee"])

        # ── Borne max (sans restriction) ─────────────────────────────────────
        cy_bm_list    = _best_n_cycles(prices, list(range(24)), n_cyc, dur, power, eff)
        pnl_absolu    = sum(c['pnl'] for c in cy_bm_list)
        spread_absolu = float(np.mean([c['spread'] for c in cy_bm_list])) if cy_bm_list else 0.0

        # ── Cycles avec contraintes ──────────────────────────────────────────
        cycles_result = _best_n_cycles(prices, avail, n_cyc, dur, power, eff)

        # Filtre spread minimum : on ne garde que les cycles rentables au-dessus du seuil
        if min_sp and cycles_result:
            cycles_result = [c for c in cycles_result if c["spread"] >= min_sp]

        n_cy_actifs = len(cycles_result)
        valid       = n_cy_actifs > 0
        total_pnl   = sum(c["pnl"]    for c in cycles_result)
        spread_moy  = float(np.mean([c["spread"] for c in cycles_result])) \
                      if cycles_result else 0.0
        energie_MWh = power * dur * n_cy_actifs * 2   # charge + décharge
        maintenance_bloque = False

        h_charge_all      = [h for c in cycles_result for h in c["h_charge"]]
        h_decharge_all    = [h for c in cycles_result for h in c["h_decharge"]]
        prix_charge_moy   = float(np.mean([c["prix_charge"]    for c in cycles_result])) \
                            if cycles_result else 0.0
        prix_decharge_moy = float(np.mean([c["prix_decharge"]  for c in cycles_result])) \
                            if cycles_result else 0.0

        # ── Quota maintenance ─────────────────────────────────────────────────
        if max_cy and valid:
            cy_count = cycles_yr.get(annee, 0)
            allowed  = min(n_cy_actifs, max_cy - cy_count)
            if allowed <= 0:
                valid             = False
                total_pnl         = 0.0
                spread_moy        = 0.0
                energie_MWh       = 0.0
                h_charge_all      = []
                h_decharge_all    = []
                prix_charge_moy   = 0.0
                prix_decharge_moy = 0.0
                n_cy_actifs       = 0
                maintenance_bloque = True
                cycles_yr[annee]  = cy_count
            elif allowed < n_cy_actifs:
                kept              = cycles_result[:allowed]
                total_pnl         = sum(c["pnl"]    for c in kept)
                spread_moy        = float(np.mean([c["spread"] for c in kept]))
                energie_MWh       = power * dur * allowed * 2
                h_charge_all      = [h for c in kept for h in c["h_charge"]]
                h_decharge_all    = [h for c in kept for h in c["h_decharge"]]
                prix_charge_moy   = float(np.mean([c["prix_charge"]   for c in kept]))
                prix_decharge_moy = float(np.mean([c["prix_decharge"] for c in kept]))
                n_cy_actifs       = allowed
                maintenance_bloque = True
                cycles_yr[annee]  = cy_count + allowed
            else:
                cycles_yr[annee]  = cy_count + n_cy_actifs
        elif valid:
            cycles_yr[annee] = cycles_yr.get(annee, 0) + n_cy_actifs

        records.append({
            "date"              : row["date"],
            "annee"             : annee,
            "mois"              : int(row["mois"]),
            "weekday"           : int(row["weekday"]),
            "weekday_name"      : row["weekday_name"],
            "n_heures_dispo"    : len(avail),
            "duration_h"        : dur,
            "n_cycles"          : n_cyc,
            "capacite_MWh"      : round(capacite_MWh, 4),
            "spread_absolu"     : round(spread_absolu, 4),
            "pnl_absolu"        : round(pnl_absolu, 4),
            "spread"            : round(spread_moy, 4),
            "pnl"               : round(total_pnl, 4),
            "valid"             : valid,
            "maintenance_bloque": maintenance_bloque,
            "n_cycles_actifs"   : n_cy_actifs,
            "energie_MWh"       : round(energie_MWh, 4),
            "h_charge"          : h_charge_all,
            "h_decharge"        : h_decharge_all,
            "prix_charge"       : round(prix_charge_moy, 4),
            "prix_decharge"     : round(prix_decharge_moy, 4),
            "prix_min"          : float(np.nanmin(prices)),
            "prix_max"          : float(np.nanmax(prices)),
            "prix_moy"          : float(np.nanmean(prices)),
        })

    return pd.DataFrame(records)


# ──────────────────────────────────────────────────────────────────────────────
# ARBITRAGE OPTIMAL (quota global annuel)
# ──────────────────────────────────────────────────────────────────────────────

def simulate_arbitrage_optimal(pivot: pd.DataFrame, params: dict) -> pd.DataFrame:
    """
    Variante optimale de simulate_arbitrage quand max_cycles_year est défini.

    Différence vs simulate_arbitrage (greedy) :
    - Greedy  : trade dans l'ordre chronologique jusqu'à épuiser le quota
    - Optimal : extrait TOUS les cycles candidats de toutes les journées,
                les trie par spread (puis PnL) décroissant, et sélectionne
                les max_cy MEILLEURS CYCLES — peu importe le jour d'origine
                et même si plusieurs cycles sélectionnés viennent du même jour.

    Exemple : quota=365, n_cycles=Max
    - Si un jour a 3 cycles très profitables, les 3 peuvent être retenus
    - Un jour avec un seul mauvais spread sera écarté même s'il est en janvier

    Paramètre d'activation : params["optimal_quota"] = True
    Sans ce paramètre ou max_cycles_year=0 → délègue à simulate_arbitrage standard.
    """
    max_cy   = params.get("max_cycles_year", None)
    min_sp   = params.get("min_spread", None)
    if not max_cy or not params.get("optimal_quota", False):
        return simulate_arbitrage(pivot, params)

    power    = params["power_MW"]
    dur      = params["duration_h"]
    n_cyc    = params.get("n_cycles", 1)
    excluded = params.get("excluded_hours", {})
    eff      = params.get("efficiency", 0.92)

    # ── Passe 1 : calculer TOUS les cycles candidats de toutes les journées ──
    # On force n_cycles=0 (mode MAX) pour énumérer tous les cycles rentables
    # de la journée, quelle que soit la config du scénario (1 cycle/j, 2, etc.).
    # La sélection des N meilleurs spreads se fait ensuite en Passe 2.
    day_potentials = []  # [{idx, annee, cycles: [...]}]
    all_cycles_flat = []  # [{idx_row, annee, cycle_idx, spread, pnl, cycle}]

    for idx_row, row in pivot.iterrows():
        prices = row[HOUR_COLS].values.astype(float)
        avail  = get_available_hours(int(row["weekday"]), excluded)
        annee  = int(row["annee"])

        # Mode MAX (0) pour avoir tous les cycles possibles du jour
        cycles = _best_n_cycles(prices, avail, 0, dur, power, eff)

        day_potentials.append({
            "idx"   : idx_row,
            "annee" : annee,
            "cycles": cycles,
        })

        # Aplatir chaque cycle individuel avec son identifiant
        for c_idx, cy in enumerate(cycles):
            # Filtre spread minimum si défini
            if cy["pnl"] > 0 and (not min_sp or cy["spread"] >= min_sp):
                all_cycles_flat.append({
                    "idx_row"  : idx_row,
                    "annee"    : annee,
                    "cycle_idx": c_idx,
                    "spread"   : cy["spread"],
                    "pnl"      : cy["pnl"],
                    "cycle"    : cy,
                })

    # ── Passe 2 : sélection des N meilleurs cycles par année ─────────────────
    # Trier tous les cycles par spread décroissant (puis pnl en tiebreak)
    # et prendre les max_cy meilleurs — peu importe le jour d'origine
    from collections import defaultdict
    by_year = defaultdict(list)
    for fc in all_cycles_flat:
        by_year[fc["annee"]].append(fc)

    # selected_cycles_map[idx_row] = liste ordonnée des cycle_idx retenus
    selected_cycles_map = defaultdict(list)
    for annee, flat_cycles in by_year.items():
        # Trier par spread décroissant, puis pnl décroissant en tiebreak
        flat_sorted = sorted(flat_cycles,
                             key=lambda x: (x["spread"], x["pnl"]),
                             reverse=True)
        quota_left = max_cy
        for fc in flat_sorted:
            if quota_left <= 0:
                break
            selected_cycles_map[fc["idx_row"]].append(fc["cycle_idx"])
            quota_left -= 1

    # Convertir en set pour lookup rapide (idx_row → set de cycle_idx retenus)
    selected_map = {
        idx_row: set(cycle_idxs)
        for idx_row, cycle_idxs in selected_cycles_map.items()
    }

    # ── Passe 3 : construire le DataFrame final ───────────────────────────────
    # Note : selected_map[idx_row] = set des cycle_idx retenus pour ce jour.
    # Un jour peut avoir 0, 1 ou plusieurs cycles retenus (si plusieurs spreads
    # de ce jour figurent dans le top-N global).
    records = []
    for dp in day_potentials:
        row        = pivot.loc[dp["idx"]]
        prices     = row[HOUR_COLS].values.astype(float)
        avail      = get_available_hours(int(row["weekday"]), excluded)
        annee      = dp["annee"]
        cycles_all = dp["cycles"]

        # Borne max (sans restriction ni quota)
        cy_bm      = _best_n_cycles(prices, list(range(24)), n_cyc, dur, power, eff)
        pnl_absolu = sum(c["pnl"] for c in cy_bm)
        spread_abs = float(np.mean([c["spread"] for c in cy_bm])) if cy_bm else 0.0

        retained_idxs = selected_map.get(dp["idx"], set())
        if retained_idxs:
            # Garder uniquement les cycles dont l'index figure dans retained_idxs
            kept  = [cy for i, cy in enumerate(cycles_all) if i in retained_idxs]
            valid = True
            # quota a écarté certains cycles de ce jour (pas tous retenus)
            maint = len(retained_idxs) < len(cycles_all)
        else:
            kept  = []
            valid = False
            maint = len(cycles_all) > 0  # potentiel existait mais quota global épuisé

        n_cy_actifs       = len(kept)
        total_pnl         = sum(c["pnl"]          for c in kept)
        spread_moy        = float(np.mean([c["spread"]        for c in kept])) if kept else 0.0
        energie_MWh       = power * dur * n_cy_actifs * 2
        h_charge_all      = [h for c in kept for h in c["h_charge"]]
        h_decharge_all    = [h for c in kept for h in c["h_decharge"]]
        prix_charge_moy   = float(np.mean([c["prix_charge"]   for c in kept])) if kept else 0.0
        prix_decharge_moy = float(np.mean([c["prix_decharge"] for c in kept])) if kept else 0.0

        records.append({
            "date"              : row["date"],
            "annee"             : annee,
            "mois"              : int(row["mois"]),
            "weekday"           : int(row["weekday"]),
            "weekday_name"      : row["weekday_name"],
            "n_heures_dispo"    : len(avail),
            "duration_h"        : dur,
            "n_cycles"          : n_cyc,
            "capacite_MWh"      : round(power * dur, 4),
            "spread_absolu"     : round(spread_abs, 4),
            "pnl_absolu"        : round(pnl_absolu, 4),
            "spread"            : round(spread_moy, 4),
            "pnl"               : round(total_pnl, 4),
            "valid"             : valid,
            "maintenance_bloque": maint,
            "n_cycles_actifs"   : n_cy_actifs,
            "energie_MWh"       : round(energie_MWh, 4),
            "h_charge"          : h_charge_all,
            "h_decharge"        : h_decharge_all,
            "prix_charge"       : round(prix_charge_moy, 4),
            "prix_decharge"     : round(prix_decharge_moy, 4),
            "prix_min"          : float(np.nanmin(prices)),
            "prix_max"          : float(np.nanmax(prices)),
            "prix_moy"          : float(np.nanmean(prices)),
        })

    return pd.DataFrame(records)


# ──────────────────────────────────────────────────────────────────────────────
# AGRÉGATION
# ──────────────────────────────────────────────────────────────────────────────

def aggregate_arbitrage(daily: pd.DataFrame, power_MW: float) -> pd.DataFrame:
    g = daily.groupby("annee")

    def spread_moy_valide(x, col):
        mask = x["valid"] & (x[col] > 0)
        return x.loc[mask, col].mean() if mask.any() else 0.0

    out = pd.DataFrame({
        "annee"                : g["annee"].first(),
        "jours_simules"        : g["date"].count(),
        "jours_actifs"         : g["valid"].sum(),
        "jours_bloques_maintenance": g["maintenance_bloque"].sum(),
        "taux_activation"      : (g["valid"].sum() / g["date"].count()).round(3),
        "spread_absolu_moy"    : g.apply(
            lambda x: spread_moy_valide(x[["valid","spread_absolu"]], "spread_absolu"),
            include_groups=False).round(2),
        "spread_moy"           : g.apply(
            lambda x: spread_moy_valide(x[["valid","spread"]], "spread"),
            include_groups=False).round(2),
        "pnl_absolu_total"     : g["pnl_absolu"].sum().round(0),
        "pnl_total"            : g["pnl"].sum().round(0),
        "pnl_par_MW"           : (g["pnl"].sum() / (power_MW * 1000)).round(2),
        "energie_totale_MWh"   : g["energie_MWh"].sum().round(2),
        "cycles_totaux"        : g["n_cycles_actifs"].sum(),
    }).reset_index(drop=True)
    return out


# ──────────────────────────────────────────────────────────────────────────────
# LISSAGE DE COURBE DE CHARGE
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


def simulate_lissage(pivot, profil_client, params, df_cdc=None):
    power     = params["power_MW"]
    e_max     = params["energy_MWh"]
    soc_min   = params["soc_min_pct"] * e_max
    soc_max   = params["soc_max_pct"] * e_max
    soc_init  = e_max * params.get("soc_init", 0.5)
    eff       = params.get("efficiency", 0.92)
    seuil_pct = params.get("seuil_percentile", 75)
    tarif_kw  = params.get("tarif_puissance_souscrite", 12000)

    if df_cdc is not None:
        seuil = np.percentile(df_cdc["conso_MW"].values, seuil_pct)
    else:
        seuil = np.percentile(profil_client, seuil_pct)

    res_jour = lissage_day(profil_client, seuil, power, e_max,
                           soc_init, soc_min, soc_max, eff)

    if df_cdc is not None:
        dates = sorted(df_cdc["date"].unique())
        records = []
        for date in dates:
            jour_data = df_cdc[df_cdc["date"] == date].sort_values("ts")
            if len(jour_data) < 24:
                continue
            profil_jour = jour_data["conso_MW"].values[:24].astype(float)
            annee = int(jour_data["annee"].iloc[0])
            res = lissage_day(profil_jour, seuil, power, e_max,
                              soc_init, soc_min, soc_max, eff)
            records.append({
                "date": date, "annee": annee,
                "reduction_pointe": res["reduction_pointe"],
                "pointe_avant": res["pointe_avant"],
                "pointe_apres": res["pointe_apres"],
            })
        df_daily = pd.DataFrame(records)
        yearly = []
        for annee, grp in df_daily.groupby("annee"):
            reduction_max = grp["reduction_pointe"].max()
            economie_an   = round(reduction_max * tarif_kw, 0)
            yearly.append({
                "annee": int(annee), "jours": len(grp),
                "reduction_pointe": round(reduction_max, 4),
                "pointe_avant_MW":  round(grp["pointe_avant"].max(), 4),
                "pointe_apres_MW":  round(grp["pointe_apres"].min(), 4),
                "economie_an":      economie_an,
                "seuil_MW":         round(seuil, 4),
            })
        reduction_globale = df_daily["reduction_pointe"].max()
        economie_globale  = reduction_globale * tarif_kw
    else:
        economie_an = res_jour["reduction_pointe"] * tarif_kw
        yearly = [{"annee": int(a),
                   "jours": len(pivot[pivot["annee"] == a]),
                   "reduction_pointe": res_jour["reduction_pointe"],
                   "pointe_avant_MW":  res_jour["pointe_avant"],
                   "pointe_apres_MW":  res_jour["pointe_apres"],
                   "economie_an":      round(economie_an, 0),
                   "seuil_MW":         round(seuil, 4)}
                  for a in sorted(pivot["annee"].unique())]
        reduction_globale = res_jour["reduction_pointe"]
        economie_globale  = economie_an

    return {
        "jour_type":    res_jour,
        "seuil_MW":     seuil,
        "reduction_MW": reduction_globale,
        "economie_an":  economie_globale,
        "yearly":       pd.DataFrame(yearly),
        "params":       params,
    }

# ══════════════════════════════════════════════════════════════════════════════
# INTRADAY DATA LOADER
# ══════════════════════════════════════════════════════════════════════════════

def load_intraday(source, max_trade_samples: int = 30, kinds: set | None = None,
                   index_names: set | None = None):
    """
    Charge données intraday depuis n'importe quel mélange de fichiers.
    Supports : ZIP EPEX (Statistics, Index, Trades annuels/journaliers), CSV IDA.
    max_trade_samples : nb max de fichiers Trades à lire pour le bid/ask —
                        les fichiers/ZIP Trades sont visités du plus récent
                        au plus ancien (date extraite du nom de fichier),
                        donc l'échantillon retenu est toujours le plus récent
                        disponible, même au milieu d'un gros ZIP multi-années.
    kinds : restreint l'extraction à un sous-ensemble de
            {"statistics", "index", "trades"}. None = tout (comportement par
            défaut). Permet, quand une seule archive contient plusieurs
            marchés (ex. export EPEX "EOD" complet), de faire un appel par
            marché sans re-décompresser inutilement les parties non désirées
            (en particulier les ZIP Trades annuels, qui peuvent peser
            plusieurs centaines de Mo).
    index_names : pour les fichiers Continuous_Index, quels IndexName extraire
            (ex. {"IDFULL"}, ou {"IDFULL","ID1","ID3"} pour comparer les 3
            références publiées dans le même fichier — ID1/ID3 = indice calculé
            1h/3h avant livraison, IDFULL = sur toute la session continue).
            None (défaut) = IDFULL uniquement, comportement historique, et la
            fonction renvoie un seul DataFrame. Si plusieurs noms sont demandés
            ET trouvés dans les données, renvoie un dict {nom: DataFrame} —
            un seul appel/parsing des fichiers, pas un par index.
    Attributs résultats :
      .report             — liste fichiers détectés
      .bid_ask_spread     — DataFrame {heure, wap_buy, wap_sell, spread} ou None
                            spread = WAP achats (ask) − WAP ventes (bid) ≥ 0 normalement.
      .trades_date_range  — (date_min, date_max) des trades effectivement
                            utilisés pour le bid/ask, ou None.
    """
    import zipfile as _zf, io as _io, re as _re
    from pathlib import Path as _Path
    import numpy as _np

    _idx_filter  = index_names or {'IDFULL'}
    rows_price   = []
    rows_trades  = []
    report       = []
    _trade_count = [0]  # compteur partagé pour limiter les Trades

    def _kind_of(fname_lower: str):
        """Type d'un fichier/ZIP déduit de son nom, ou None si ambigu."""
        if 'trade' in fname_lower:     return 'trades'
        if 'statistic' in fname_lower: return 'statistics'
        if 'ida' in fname_lower:       return 'ida'
        if 'index' in fname_lower:     return 'index'
        return None

    def _recency_key(name: str):
        """Clé de tri pour visiter les fichiers Trades du plus récent au plus
        ancien (date AAAAMMJJ ou année AAAA extraite du nom). Les fichiers
        non-Trades gardent un ordre stable après les Trades."""
        nl = name.lower()
        if 'trade' not in nl:
            return (1, 0, name)
        m = _re.search(r'(20\d{2})(\d{2})(\d{2})', name) or _re.search(r'(20\d{2})', name)
        val = int(m.group(0)) if m else 0
        return (0, -val, name)

    # ── Parsers ───────────────────────────────────────────────────────────────

    def _parse_stat_csv(data: bytes):
        txt = data.decode('utf-8', errors='replace')
        lines = txt.strip().split('\n')
        hi = next((i for i,l in enumerate(lines) if 'DeliveryStart' in l), None)
        if hi is None: return []
        hdr = [c.strip() for c in lines[hi].split(',')]
        pcol = next((c for c in ['WeightedAveragePrice','IndexPrice'] if c in hdr), None)
        if not pcol: return []
        # Statistics et Index publient le même créneau horaire à plusieurs
        # granularités (produits 15/30/60 min, + lignes "Base"/"Peak" pour
        # Index) — ce sont des agrégats IMBRIQUÉS des mêmes transactions
        # (4×15min = 2×30min = 1×60min), pas des observations indépendantes :
        # les mélanger ferait du double comptage et fausserait le prix.
        # Stratégie par heure : on garde en priorité la ligne 60 min (déjà
        # la moyenne pondérée exacte de l'heure) ; si elle est absente pour
        # une heure donnée (granularité 60 min non publiée sur cette période),
        # on reconstruit l'heure à partir des tranches 30 puis 15 min — pour
        # ne perdre aucune période plutôt que de filtrer strictement sur 60 min.
        # Pour Index, plusieurs indices (ID1/ID3/IDFULL) coexistent dans le
        # même CSV — on ne garde que ceux demandés via _idx_filter (IDFULL
        # par défaut), étiquetés pour pouvoir les séparer en aval.
        has_idxname = 'IndexName' in hdr
        has_timeres = 'TimeResolution' in hdr
        has_end     = 'DeliveryEnd' in hdr
        by_hour = {}  # (idx, date, heure) -> {60: [prix], 30: [prix], 15: [prix]}
        for line in lines[hi+1:]:
            if not line.strip(): continue
            vals = line.split(',')
            if len(vals) < len(hdr): continue
            row = dict(zip(hdr, vals))
            row_idx = row.get('IndexName', '').strip().upper() if has_idxname else None
            if has_idxname and row_idx not in _idx_filter:
                continue
            try:
                dt = pd.to_datetime(row['DeliveryStart'].strip())
                if has_timeres:
                    res_min = {'60min': 60, '30min': 30, '15min': 15}.get(
                        row.get('TimeResolution', '').strip().lower())
                    if res_min is None:
                        continue  # Base/Peak — agrégats non horaires, hors sujet
                elif has_end:
                    delta_min = (pd.to_datetime(row['DeliveryEnd'].strip()) - dt).total_seconds() / 60
                    res_min = int(delta_min) if delta_min in (15, 30, 60) else None
                    if res_min is None:
                        continue
                else:
                    res_min = 60  # pas d'info de résolution, on suppose horaire
                price = float(row[pcol].strip())
            except Exception:
                continue
            key = (row_idx, dt.date(), dt.hour)
            by_hour.setdefault(key, {}).setdefault(res_min, []).append(price)

        out = []
        for (row_idx, date, heure), by_res in by_hour.items():
            for res in (60, 30, 15):
                if res in by_res:
                    out.append({'date': date, 'heure': heure, 'idx': row_idx,
                                'prix': sum(by_res[res]) / len(by_res[res])})
                    break
        return out

    def _parse_trade_csv(data: bytes):
        out = []
        txt = data.decode('utf-8', errors='replace')
        lines = [l for l in txt.strip().split('\n') if not l.startswith('#')]
        if not lines: return out
        try:
            df_t = pd.read_csv(_io.StringIO('\n'.join(lines)))
        except Exception: return out
        if 'Product' not in df_t.columns: return out
        df_t = df_t[df_t['Product'].str.contains('Hour_Power', na=False)].copy()
        if 'DeliveryStart' not in df_t.columns: return out
        df_t['_dt'] = pd.to_datetime(df_t['DeliveryStart'], errors='coerce', utc=True)
        df_t = df_t.dropna(subset=['_dt'])
        for _, r in df_t.iterrows():
            try:
                out.append({'date': r['_dt'].date(), 'heure': int(r['_dt'].hour),
                             'side': str(r.get('Side','')).upper(),
                             'price': float(r['Price']),
                             'volume': float(r.get('Volume', 1)),
                             'produit': 'Quarter' if 'Quarter' in str(r['Product']) else 'Hour'})
            except Exception: pass
        return out

    def _parse_ida_csv(data: bytes):
        out = []
        txt = data.decode('utf-8', errors='replace')
        lines = txt.strip().split('\n')
        dh = next((i for i,l in enumerate(lines) if 'day,' in l and 'Hour' in l), None)
        if dh is None: return out
        cols = [c.strip() for c in lines[dh].split(',')]
        col_to_hour = {}
        for i, col in enumerate(cols[1:], 1):
            m = _re.match(r'Hour\s+(\d+)([AB]?)\s*Q', col, _re.IGNORECASE)
            if not m or m.group(2).upper() == 'B': continue
            col_to_hour[i] = int(float(m.group(1))) - 1
        for line in lines[dh+1:]:
            if not line.strip() or line.lower().startswith('nan'): continue
            vals = line.split(',')
            if len(vals) < 2: continue
            try:
                date = pd.to_datetime(vals[0].strip(), dayfirst=True).date()
            except Exception: continue
            hp = {}
            for i, h in col_to_hour.items():
                if i >= len(vals): continue
                try: v = float(vals[i].strip())
                except Exception: continue
                hp.setdefault(h, []).append(v)
            for h, pp in hp.items():
                if not pp: continue
                out.append({'date': date, 'heure': h, 'prix': sum(pp)/len(pp)})
        return out

    def _handle_file(name: str, raw: bytes, depth: int = 0, parent_path: str = ""):
        """Dispatche un fichier selon son type. parent_path trace le chemin ZIP→sous-ZIP."""
        if depth > 5: return
        fname = name.split('/')[-1].lower()
        full_path = f"{parent_path} → {name.split('/')[-1]}" if parent_path else name.split('/')[-1]

        # Filtre par marché demandé (kinds) — évite de décompresser des
        # branches entières (ex. ZIP Trades annuels de plusieurs centaines
        # de Mo) quand on ne veut que les Statistics ou que l'Index.
        # Nom ambigu (None) → on ne filtre pas, il faut l'ouvrir pour savoir.
        if kinds is not None:
            k = _kind_of(fname)
            if k is not None and k not in kinds:
                return

        # ZIP imbriqué → récursion
        if raw[:2] == b'PK':
            try:
                with _zf.ZipFile(_io.BytesIO(raw)) as inner:
                    members = [n for n in inner.namelist() if not n.endswith('/')]
                    # Visiter les Trades du plus récent au plus ancien : le
                    # quota max_trade_samples doit porter sur les données
                    # récentes, pas sur les premières années rencontrées
                    # dans un export EPEX historique 2007-2025.
                    members.sort(key=_recency_key)
                    for iname in members:
                        # Limiter les Trades pour le bid/ask
                        if 'trade' in iname.lower() and _trade_count[0] >= max_trade_samples:
                            continue
                        try:
                            iraw = inner.read(iname)
                            _handle_file(iname, iraw, depth+1, parent_path=full_path)
                        except Exception: pass
            except Exception: pass
            return

        # CSV Statistics / Index
        if fname.endswith('.csv') and ('statistic' in fname or 'index' in fname):
            r = _parse_stat_csv(raw)
            rows_price.extend(r)
            if r:
                idx_names = sorted({x['idx'] for x in r if x.get('idx')})
                report.append({'fichier': full_path, 'type_détecté': 'statistics/index',
                               'nb_jours': len({x['date'] for x in r}),
                               'index_names': idx_names, 'statut': 'OK'})
            return

        # CSV Trades
        if fname.endswith('.csv') and 'trade' in fname:
            if _trade_count[0] < max_trade_samples:
                r = _parse_trade_csv(raw)
                rows_trades.extend(r)
                _trade_count[0] += 1
                if r:
                    report.append({'fichier': full_path, 'type_détecté': 'trades',
                                   'nb_jours': len({x['date'] for x in r}), 'statut': 'OK'})
            return

        # CSV IDA
        if fname.endswith('.csv'):
            sample = raw[:500].decode('utf-8', errors='replace').lower()
            if 'hour' in sample and ('q1' in sample or 'ida' in fname):
                r = _parse_ida_csv(raw)
                rows_price.extend(r)
                if r:
                    report.append({'fichier': full_path, 'type_détecté': 'ida',
                                   'nb_jours': len({x['date'] for x in r}), 'statut': 'OK'})

    # ── Collecter sources ─────────────────────────────────────────────────────
    sources = source if isinstance(source, (list, tuple)) else [source]

    for src in sources:
        if hasattr(src, 'name') and hasattr(src, '_data'):
            fname, raw = src.name, src._data
        elif isinstance(src, (str, _Path)):
            fname = str(src)
            # Stream depuis fichier pour éviter de charger 500Mo en RAM
            with open(src, 'rb') as f:
                raw = f.read()
        elif isinstance(src, bytes):
            fname, raw = '', src
        elif hasattr(src, 'read'):
            fname = getattr(src, 'name', '')
            raw = src.read()
        else:
            continue
        _handle_file(fname, raw, depth=0)

    load_intraday.report = report

    # ── Bid/Ask spread depuis Trades ──────────────────────────────────────────
    # Side=BUY  : l'acheteur est l'agresseur, le trade s'exécute côté ASK.
    # Side=SELL : le vendeur est l'agresseur, le trade s'exécute côté BID.
    # Spread bid/ask = ASK − BID = WAP(BUY) − WAP(SELL)  (≥ 0 normalement).
    load_intraday.bid_ask_spread  = None
    load_intraday.bid_ask_spread_by_product = None
    load_intraday.trades_date_range = None
    if rows_trades:
        df_tr = pd.DataFrame(rows_trades)
        load_intraday.trades_date_range = (df_tr['date'].min(), df_tr['date'].max())

        def _wap_spread(group_cols):
            gb_buy  = df_tr[df_tr['side']=='BUY'].groupby(group_cols)
            gb_sell = df_tr[df_tr['side']=='SELL'].groupby(group_cols)
            wap_buy  = gb_buy.apply(lambda x: (x['price']*x['volume']).sum()/x['volume'].sum(),
                                     include_groups=False)
            wap_sell = gb_sell.apply(lambda x: (x['price']*x['volume']).sum()/x['volume'].sum(),
                                      include_groups=False)
            out = pd.concat([wap_buy.rename('wap_buy'), wap_sell.rename('wap_sell')], axis=1).dropna()
            out['spread'] = out['wap_buy'] - out['wap_sell']
            return out

        load_intraday.bid_ask_spread = _wap_spread('heure')
        if 'produit' in df_tr.columns:
            load_intraday.bid_ask_spread_by_product = _wap_spread(['produit', 'heure'])
        # Si pas de Statistics → utiliser WAP Trades comme prix
        if not rows_price:
            wap_all = df_tr.groupby(['date','heure']).apply(
                lambda x: (x['price']*x['volume']).sum()/x['volume'].sum(),
                include_groups=False
            ).reset_index()
            wap_all.columns = ['date','heure','prix']
            rows_price = wap_all.to_dict('records')

    if not rows_price:
        return pd.DataFrame()

    # ── Pivot ─────────────────────────────────────────────────────────────────
    def _build_pivot(sub_rows):
        df = pd.DataFrame(sub_rows)
        df = df.groupby(['date','heure'])['prix'].mean().reset_index()
        pivot = df.pivot(index='date', columns='heure', values='prix')
        for h in range(24):
            if h not in pivot.columns: pivot[h] = _np.nan
        pivot = pivot[[h for h in range(24)]]
        pivot.columns = HOUR_COLS
        pivot = pivot.reset_index()
        pivot['date']    = pd.to_datetime(pivot['date'])
        pivot['annee']   = pivot['date'].dt.year
        pivot['mois']    = pivot['date'].dt.month
        pivot['jour']    = pivot['date'].dt.day
        pivot['weekday'] = pivot['date'].dt.weekday
        _jours = ['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
        pivot['weekday_name'] = pivot['weekday'].apply(lambda x: _jours[int(x)])
        hcols = [c for c in HOUR_COLS if c in pivot.columns]
        pivot[hcols] = pivot[hcols].interpolate(axis=1, limit_direction='both')
        result = pivot.set_index(['annee','mois','jour'])[hcols+['weekday','weekday_name']].reset_index()
        result['date'] = pd.to_datetime(
            result['annee'].astype(str)+'-'+
            result['mois'].astype(str).str.zfill(2)+'-'+
            result['jour'].astype(str).str.zfill(2)
        )
        return result

    # Plusieurs IndexName demandés ET trouvés (ex. IDFULL+ID1+ID3) → un pivot
    # par index, sans re-décompresser/re-parser les fichiers une 2e/3e fois.
    _idx_present = {r.get('idx') for r in rows_price if r.get('idx') is not None}
    if index_names and len(_idx_present) > 1:
        return {name: _build_pivot([r for r in rows_price if r.get('idx') == name])
                for name in sorted(_idx_present)}
    return _build_pivot(rows_price)