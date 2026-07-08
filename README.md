# BESS Valorisation - Dashboard

Outil de simulation de la rentabilité d'un système de stockage par batterie (BESS) sur les marchés électriques français : Day-Ahead, Intraday et Imbalance Market.

Développé dans le cadre d'un projet de fin d'études chez ENI Gas & Power.

## Onglets

| Onglet | Description |
|---|---|
| Arbitrage Day-Ahead | Simulation charge/décharge optimisée sur prix spot EPEX |
| Intraday | Analyse des marchés aux enchères IDA1/IDA3 et continu (IDFULL, ID1, ID3) |
| Imbalance Market | Analyse sur prix de règlement des écarts (négatif / positif) |
| Lissage de charge | Peak shaving sur courbe de charge industrielle |
| Comparaison scénarios | Comparaison de plusieurs configurations batterie |
| Sensibilité | Analyse de sensibilité sur les paramètres clés |
| Comparaison pays | Comparaison France / Allemagne sur données spot |
| Executive Summary | Synthèse des résultats pour présentation |

## Formats de fichiers attendus

**Données spot (requis pour tous les onglets)**
- Excel `.xlsx` avec une feuille `Spot_input`, colonnes : `ANNEE`, `MOIS`, `JOUR`, `HEURE`, `Prix Final`

**Intraday**
- ZIP EPEX `Continuous_Index-FR-*.zip` pour le marché continu (IDFULL / ID1 / ID3)
- CSV `pan-european_prices_france_IDA1_*.csv` et `IDA3_*.csv` pour les enchères
- IDA2 est rarement présent dans les exports EPEX (marché peu liquide) - l'outil fonctionne sans

**Imbalance Market**
- Excel `.xlsx` avec colonnes `Date`, `negative_imbalance_settlement_price`, `positive_imbalance_settlement_price`

**Lissage de charge**
- Excel `.xlsx` avec une feuille `CdC_kWh` (courbe de charge horaire en kWh)

**Comparaison pays**
- Un fichier Excel par pays, même format que le fichier spot principal

## Déploiement local

```bash
pip install -r requirements.txt
streamlit run bess_dashboard.py
```

## Assistant IA (optionnel)

L'outil intègre un assistant conversationnel basé sur l'API Gemini (Google). Il est optionnel et désactivé si aucune clé n'est fournie.

Pour l'activer :
1. Créer une clé sur [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. En local : créer `.streamlit/secrets.toml` avec `GEMINI_API_KEY = "AIza..."`
3. Sur Streamlit Cloud : ajouter la variable dans **Settings > Secrets**
