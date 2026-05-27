# 🛡 CyberDashboard — Guide de déploiement complet

## Architecture du projet

```
CyberDashboard/
├── CyberDashboard_Data.xlsx        # Données + seuils + KPIs + journal
├── alert_engine.py                 # Moteur de détection et d'alertes mail
├── PowerBI_Configuration_Guide.txt # Guide étape par étape pour Power BI
└── README.md                       # Ce fichier
```

---

## 📊 Fichier Excel — Structure des onglets

| Onglet              | Rôle                                                           |
|---------------------|----------------------------------------------------------------|
| **Données**         | Historique 30j des indicateurs (mails, requêtes réseau, etc.) |
| **Configuration**   | Seuils d'alerte modifiables + paramètres SMTP                  |
| **KPI Dashboard**   | Tableau de bord avec formules Excel (vue rapide sans Power BI) |
| **Journal Alertes** | Historique automatique des alertes déclenchées                |
| **PowerBI — Guide** | Rappel des étapes de connexion Power BI                        |

### Indicateurs surveillés

| Indicateur              | Seuil AVERT. | Seuil CRITIQUE | Menace associée         |
|-------------------------|-------------|-----------------|-------------------------|
| Mails reçus             | 1 800       | 3 000           | Phishing campaign       |
| Requêtes réseau         | 12 000      | 25 000          | DDoS / scan massif      |
| Tentatives connexion    | 400         | 1 000           | Brute-force             |
| Alertes IDS             | 20          | 50              | Intrusion active        |
| Trafic entrant (MB)     | 300         | 600             | Exfiltration / DDoS     |
| Connexions refusées     | 200         | 500             | Scan agressif           |
| Processus suspects      | 5           | 10              | Malware actif           |

> ✏️ **Pour modifier un seuil** : ouvrir l'onglet **Configuration**, colonne B (Avert.) ou C (Critique).

---

## ⚙️ Installation du moteur d'alerte Python

### Prérequis

```bash
python --version   # Python 3.8+
pip install openpyxl
```

### Utilisation

```bash
# Analyse + envoi mail si seuil dépassé
python alert_engine.py

# Analyse sans envoi mail (test)
python alert_engine.py --dry-run

# Spécifier un autre fichier Excel
python alert_engine.py --file /chemin/vers/CyberDashboard_Data.xlsx
```

### Configuration du mail (Gmail)

1. Activer la **validation en deux étapes** sur votre compte Google
2. Générer un **mot de passe d'application** :
   → myaccount.google.com → Sécurité → Mots de passe des applications
3. Renseigner dans l'onglet **Configuration** du fichier Excel :
   - SMTP Serveur : `smtp.gmail.com`
   - SMTP Port    : `587`
   - Expéditeur   : `votre.email@gmail.com`
   - Mot de passe : le mot de passe d'application (16 caractères)
   - Destinataire : l'adresse de réception des alertes

### Configuration du mail (Office 365)

- SMTP Serveur : `smtp.office365.com`
- SMTP Port    : `587`
- Expéditeur   : `soc@votre-entreprise.fr`
- Mot de passe : mot de passe du compte (ou app password si MFA activé)

---

## 🔄 Automatisation (exécution toutes les heures)

### Windows — Planificateur de tâches

```powershell
# Ouvrir le Planificateur de tâches (taskschd.msc)
# Créer une tâche → Déclencheur : Toutes les heures
# Action :
#   Programme : C:\Python311\python.exe
#   Arguments : "C:\CyberDashboard\alert_engine.py"
#   Démarrer dans : C:\CyberDashboard\
```

### Linux / macOS — cron

```bash
crontab -e
# Ajouter la ligne suivante (exécution toutes les heures) :
0 * * * * cd /opt/cyberdashboard && python3 alert_engine.py >> /var/log/cyber_alerts.log 2>&1
```

---

## 📊 Connexion Power BI

Voir **PowerBI_Configuration_Guide.txt** pour le détail complet.

En résumé :
1. Power BI Desktop → Obtenir des données → Classeur Excel → `CyberDashboard_Data.xlsx`
2. Cocher les feuilles : **Données**, **Configuration**, **Journal Alertes**
3. Créer les visuels selon le guide (courbes, jauges, matrices, KPI cards)
4. Publier dans Power BI Service → configurer l'actualisation automatique

---

## 🔒 Sécurité

- Ne jamais committer le fichier Excel avec les identifiants SMTP renseignés
- Utiliser un compte dédié (ex. `soc-alerts@entreprise.fr`) pour l'envoi des mails
- Pour un déploiement en production, externaliser le mot de passe SMTP :
  ```bash
  export SMTP_PASSWORD="votre_mot_de_passe"
  ```
  Et modifier `alert_engine.py` ligne `"password"` :
  ```python
  "password": os.environ.get("SMTP_PASSWORD", str(ws["B17"].value or "")),
  ```

---

## 📬 Format du mail d'alerte

Le mail envoyé est en HTML responsive avec :
- **Bannière rouge/orange** selon le niveau de criticité
- **Tableau des indicateurs en alerte** (valeur, seuils, variation)
- **Interprétation automatique** (DDoS, phishing, brute-force, malware)
- **Tableau récapitulatif** de tous les indicateurs avec statut 🟢/🟡/🔴
- **Journal mis à jour** automatiquement dans l'Excel après chaque exécution

---

*Projet réalisé dans le cadre du module Cybersecurity Incident Response Dashboard — SME*
