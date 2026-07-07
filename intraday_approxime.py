import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================================
# PATHS
# ============================================================================

INPUT_EXCEL = r"C:\Users\nicolas.pissavin\Documents\Duplication projet batteries\BESS valorisation_V2.xlsx"
OUTPUT_EXCEL = r"C:\Users\nicolas.pissavin\Documents\Duplication projet batteries\BESS_valorisation_RESULTATS.xlsx"


# ============================================================================
# COPIE DU FICHIER (toujours au début)
# ============================================================================

shutil.copyfile(INPUT_EXCEL, OUTPUT_EXCEL)


# ============================================================================
# LECTURE DES DONNÉES SPOT
# ============================================================================

spot = pd.read_excel(INPUT_EXCEL, sheet_name="Spot_input", engine="openpyxl")

required = {"ANNEE", "MOIS", "JOUR", "HEURE", "Prix Final"}
if not required.issubset(spot.columns):
    raise ValueError("Colonnes manquantes dans Spot_input")


# ============================================================================
# OUVERTURE DU CLASSEUR
# ============================================================================

wb = load_workbook(OUTPUT_EXCEL)


# ============================================================================
# ✅ GARANTIR L’EXISTENCE DE BESS_Parametres
# ============================================================================

if "BESS_Parametres" not in wb.sheetnames:
    ws_p = wb.create_sheet("BESS_Parametres")
    ws_p.append(["Parametre", "Valeur", "Description"])
    ws_p.append(["Energy_MWh", 2, "Capacité batterie"])
    ws_p.append(["Power_MW", 1, "Puissance maximale"])
    ws_p.append(["Duration_h", 1, "Durée de fonctionnement"])
    ws_p.append(["SOC_min", 0.1, "SOC minimum"])
    ws_p.append(["SOC_max", 0.9, "SOC maximum"])
    ws_p.append(["Mode", "arbitrage", "arbitrage / lissage"])


# ============================================================================
# LECTURE PARAMÈTRES (DÉSORMAIS SÛRE)
# ============================================================================

params_ws = wb["BESS_Parametres"]
params = {}

for row in params_ws.iter_rows(min_row=2, values_only=True):
    params[row[0]] = row[1]

ENERGY_MAX = params["Energy_MWh"]
POWER_MW = params["Power_MW"]
SOC_MIN = params["SOC_min"] * ENERGY_MAX
SOC_MAX = params["SOC_max"] * ENERGY_MAX


# ============================================================================
# BESS_Simulation (industrialisation réelle)
# ============================================================================

soc = SOC_MAX
rows_sim = []

for (annee, mois, jour), day in spot.groupby(["ANNEE", "MOIS", "JOUR"]):

    day = day.drop_duplicates(subset="HEURE").sort_values("HEURE")
    if len(day) < 2:
        continue

    p_low = day["Prix Final"].quantile(0.25)
    p_high = day["Prix Final"].quantile(0.75)

    for _, r in day.iterrows():
        hour = int(r["HEURE"])
        price = r["Prix Final"]
        action = "idle"
        energy = 0
        value = 0

        if price <= p_low and soc < SOC_MAX:
            energy = min(POWER_MW, SOC_MAX - soc)
            soc += energy
            action = "charge"
            value = -energy * price

        elif price >= p_high and soc > SOC_MIN:
            energy = min(POWER_MW, soc - SOC_MIN)
            soc -= energy
            action = "discharge"
            value = energy * price
            energy = -energy

        rows_sim.append([
            annee, mois, jour, hour,
            action, energy, price, value, soc
        ])


# ============================================================================
# ÉCRITURE BESS_Simulation
# ============================================================================

if "BESS_Simulation" in wb.sheetnames:
    del wb["BESS_Simulation"]

ws = wb.create_sheet("BESS_Simulation")
ws.append([
    "ANNEE", "MOIS", "JOUR", "HEURE",
    "ACTION", "ENERGY_MWh",
    "PRICE_EUR_MWh", "VALUE_EUR", "SOC_after_MWh"
])

for r in rows_sim:
    ws.append(r)


# ============================================================================
# FORMATAGE SIMPLE
# ============================================================================

header_fill = PatternFill("solid", fgColor="DDDDDD")
charge_fill = PatternFill("solid", fgColor="CFE2F3")
discharge_fill = PatternFill("solid", fgColor="F4CCCC")

for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.fill = header_fill

for row in ws.iter_rows(min_row=2):
    if row[4].value == "charge":
        for c in row:
            c.fill = charge_fill
    elif row[4].value == "discharge":
        for c in row:
            c.fill = discharge_fill

for col in ws.columns:
    ws.column_dimensions[get_column_letter(col[0].column)].width = 18


# ============================================================================
# SAUVEGARDE
# ============================================================================

wb.save(OUTPUT_EXCEL)

print("✅ DA_Resultats conservée")
print("✅ BESS_Parametres garanti")
print("✅ BESS_Simulation ajoutée")
print("✅ Script stable (multi-run OK)")
