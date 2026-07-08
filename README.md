# BESS Valorisation — Dashboard

Outil d'analyse et de simulation de la rentabilité d'un système de stockage par batterie (BESS) sur les marchés électriques européens (Day-Ahead, Intraday, Imbalance).

## Utilisation

Déposez votre fichier Excel dans la sidebar (format `Spot_input` : colonnes ANNEE, MOIS, JOUR, HEURE, Prix Final) et parcourez les onglets :

- **Arbitrage Day-Ahead** — simulation cycle charge/décharge sur prix spot EPEX
- **Intraday** — comparaison IDA1/IDA2/IDA3 + Continuous (IDFULL, ID1, ID3) depuis les fichiers EPEX Continuous_Index FR
- **Imbalance Market** — prix de règlement des écarts (colonnes `negative_imbalance_settlement_price` / `positive_imbalance_settlement_price`)
- **Lissage de charge** — peak shaving industriel
- **Comparaison scénarios / Sensibilité / Comparaison pays** — analyses multi-configurations
- **Executive Summary** — synthèse pour présentation

## Clé API Gemini (assistant IA intégré)

L'assistant IA utilise l'API Gemini (Google). Pour l'activer :

1. Obtenez une clé gratuite sur [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. En local : créez `.streamlit/secrets.toml` avec `GEMINI_API_KEY = "AIza..."`
3. Sur Streamlit Cloud : ajoutez la variable dans **Settings → Secrets**

## Déploiement local

```bash
pip install -r requirements.txt
streamlit run bess_dashboard.py
```

## Formats de fichiers acceptés

| Onglet | Format attendu |
|---|---|
| Tous (données spot) | Excel `.xlsx` — feuille `Spot_input`, colonnes ANNEE/MOIS/JOUR/HEURE/Prix Final |
| Intraday | ZIP EPEX `Continuous_Index-FR-*.zip` ou `Intraday Continuous.zip` + CSV IDA1/2/3 |
| Imbalance Market | Excel `.xlsx` avec colonnes Date, negative/positive_imbalance_settlement_price |
| Lissage de charge | Excel `.xlsx` — feuille `CdC_kWh` (courbe de charge horaire) |
| Comparaison pays | Un Excel par pays, même format que le fichier principal |
