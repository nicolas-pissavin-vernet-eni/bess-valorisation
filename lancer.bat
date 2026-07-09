@echo off
echo Installation des dependances...
pip install -r requirements.txt
echo.
echo Lancement du dashboard...
streamlit run bess_dashboard.py
pause
