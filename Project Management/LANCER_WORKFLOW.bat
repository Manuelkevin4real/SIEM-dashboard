@echo off
:: ============================================================
::  CyberDashboard — Lanceur automatique complet
::  Double-cliquer pour démarrer tout le workflow
:: ============================================================

title CyberDashboard — Lanceur

echo.
echo  =====================================================
echo   CYBERSECURITY INCIDENT RESPONSE DASHBOARD
echo   Demarrage du workflow automatique...
echo  =====================================================
echo.

:: Vérification Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telecharge Python sur https://python.org
    pause
    exit /b 1
)

:: Installation des dépendances si nécessaire
echo [1/4] Verification des dependances Python...
pip install openpyxl pywinauto pygetwindow --break-system-packages -q
echo       OK.
echo.

:: Vérification du fichier Excel
if not exist "CyberDashboard_Data.xlsx" (
    echo [ERREUR] CyberDashboard_Data.xlsx introuvable dans ce dossier.
    echo Place ce .bat dans le meme dossier que les scripts Python et l'Excel.
    pause
    exit /b 1
)

echo [2/4] Lancement du simulateur de donnees (data_feeder.py)...
echo       Injection toutes les 60 minutes en arriere-plan.
start "CyberDashboard - Data Feeder" cmd /k "python data_feeder.py --rows 1 && echo. && echo Prochaine injection dans 60 min... && timeout /t 3600 /nobreak && python data_feeder.py --rows 1"
timeout /t 2 /nobreak >nul

echo [3/4] Lancement du moteur d'alertes (alert_engine.py)...
echo       Verification des seuils toutes les 60 minutes.
start "CyberDashboard - Alert Engine" cmd /k "python alert_engine.py && echo. && echo Prochaine verification dans 60 min... && timeout /t 3600 /nobreak && python alert_engine.py"
timeout /t 2 /nobreak >nul

echo [4/4] Lancement de l'actualisation Power BI (powerbi_autorefresh.py)...
echo       Actualisation toutes les 5 minutes.
echo       IMPORTANT : Power BI Desktop doit etre ouvert avec votre .pbix
echo.
start "CyberDashboard - PowerBI Refresh" cmd /k "python powerbi_autorefresh.py --interval 5"

echo.
echo  =====================================================
echo   Workflow demarre ! 3 fenetres ouvertes :
echo   - Data Feeder    : injection de donnees / 60 min
echo   - Alert Engine   : detection + mail    / 60 min
echo   - PowerBI Refresh: actualisation PBI   / 5 min
echo.
echo   Ferme les fenetres de commande pour tout arreter.
echo  =====================================================
echo.
pause
