"""
powerbi_autorefresh.py — Actualisation automatique de Power BI Desktop
=======================================================================
Force l'actualisation des données dans Power BI Desktop à intervalle
régulier, sans abonnement Pro ni Power BI Service.

Fonctionnement :
  1. Vérifie que Power BI Desktop est ouvert
  2. Met la fenêtre au premier plan
  3. Envoie le raccourci clavier Alt+F5 (Actualiser tout)
  4. Attend l'intervalle configuré, puis recommence

Usage :
    python powerbi_autorefresh.py                  # toutes les 5 minutes
    python powerbi_autorefresh.py --interval 2     # toutes les 2 minutes
    python powerbi_autorefresh.py --interval 1     # toutes les 60 secondes
    python powerbi_autorefresh.py --once           # une seule actualisation

Prérequis :
    pip install pywinauto pygetwindow

IMPORTANT : ce script est Windows uniquement.
Power BI Desktop doit être ouvert avec le fichier .pbix chargé.
"""

import argparse
import ctypes
import subprocess
import sys
import time
import datetime

# ── Vérification OS ──────────────────────────────────────────────────────────
if sys.platform != "win32":
    print("❌  Ce script est Windows uniquement (Power BI Desktop ne tourne que sur Windows).")
    sys.exit(1)

import winreg  # noqa: E402 — disponible uniquement sur Windows

try:
    import pygetwindow as gw
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pygetwindow", "--break-system-packages", "-q"])
    import pygetwindow as gw

try:
    from pywinauto import Application
    from pywinauto.keyboard import send_keys
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pywinauto", "--break-system-packages", "-q"])
    from pywinauto import Application
    from pywinauto.keyboard import send_keys

# ─────────────────────────────────────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────────────────────────────────────

POWERBI_WINDOW_TITLE = "Power BI Desktop"   # Titre partiel de la fenêtre
REFRESH_SHORTCUT     = "%{F5}"              # Alt+F5 = Actualiser tout dans Power BI
REFRESH_WAIT_SEC     = 8                    # Temps d'attente après envoi du raccourci
                                            # (augmenter si le fichier Excel est gros)

# ─────────────────────────────────────────────────────────────────────────────
# Utilitaires
# ─────────────────────────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"[{ts}]  {msg}", flush=True)


def find_powerbi_window():
    """Retourne la fenêtre Power BI Desktop si elle est ouverte, sinon None."""
    windows = gw.getWindowsWithTitle(POWERBI_WINDOW_TITLE)
    return windows[0] if windows else None


def bring_to_front(window) -> bool:
    """Met la fenêtre Power BI au premier plan."""
    try:
        if window.isMinimized:
            window.restore()
        window.activate()
        time.sleep(0.5)
        return True
    except Exception as e:
        log(f"⚠️   Impossible d'activer la fenêtre : {e}")
        return False


def trigger_refresh() -> bool:
    """
    Envoie Alt+F5 à Power BI Desktop pour déclencher l'actualisation.
    Retourne True si l'opération a réussi.
    """
    win = find_powerbi_window()
    if not win:
        log("❌  Power BI Desktop introuvable. Assure-toi qu'il est ouvert avec ton .pbix.")
        return False

    log(f"🔍  Fenêtre trouvée : '{win.title}'")

    if not bring_to_front(win):
        return False

    try:
        send_keys(REFRESH_SHORTCUT)
        log(f"⌨️   Raccourci Alt+F5 envoyé — actualisation en cours…")
        time.sleep(REFRESH_WAIT_SEC)
        log(f"✅  Actualisation terminée (attente de {REFRESH_WAIT_SEC}s).")
        return True
    except Exception as e:
        log(f"❌  Erreur lors de l'envoi du raccourci : {e}")
        return False


def check_powerbi_installed() -> bool:
    """Vérifie que Power BI Desktop est installé (via le registre Windows)."""
    keys_to_check = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    ]
    for root_key in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        for key_path in keys_to_check:
            try:
                key = winreg.OpenKey(root_key, key_path)
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        if "Power BI" in display_name:
                            return True
                    except (OSError, FileNotFoundError):
                        continue
            except (OSError, FileNotFoundError):
                continue
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Boucle principale
# ─────────────────────────────────────────────────────────────────────────────

def run_loop(interval_minutes: int):
    interval_sec = interval_minutes * 60
    log(f"🚀  Démarrage — actualisation toutes les {interval_minutes} minute(s).")
    log(f"    Appuie sur Ctrl+C pour arrêter.\n")

    refresh_count  = 0
    failure_count  = 0
    MAX_FAILURES   = 5  # Arrêt si Power BI reste introuvable trop longtemps

    while True:
        success = trigger_refresh()
        if success:
            refresh_count += 1
            failure_count  = 0
            log(f"    Total actualisations : {refresh_count}\n")
        else:
            failure_count += 1
            log(f"⚠️   Échec {failure_count}/{MAX_FAILURES}. "
                f"Nouvelle tentative dans {interval_minutes} min.\n")
            if failure_count >= MAX_FAILURES:
                log("❌  Trop d'échecs consécutifs. Vérifier que Power BI est ouvert.")
                log("    Arrêt du script.")
                sys.exit(1)

        log(f"💤  Prochaine actualisation dans {interval_minutes} minute(s)…")
        time.sleep(interval_sec)


# ─────────────────────────────────────────────────────────────────────────────
# Point d'entrée
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Actualisation automatique de Power BI Desktop"
    )
    parser.add_argument(
        "--interval", type=int, default=5,
        help="Intervalle entre chaque actualisation en minutes (défaut: 5)"
    )
    parser.add_argument(
        "--once", action="store_true",
        help="Actualiser une seule fois puis quitter"
    )
    args = parser.parse_args()

    # Vérification installation Power BI
    if not check_powerbi_installed():
        log("⚠️   Power BI Desktop ne semble pas installé sur ce poste.")
        log("    Téléchargement : https://powerbi.microsoft.com/fr-fr/desktop/")
        log("    Le script continue quand même…\n")

    # Vérification fenêtre ouverte
    win = find_powerbi_window()
    if not win:
        log("❌  Power BI Desktop n'est pas ouvert.")
        log("    1. Lance Power BI Desktop")
        log("    2. Ouvre ton fichier .pbix")
        log("    3. Relance ce script")
        sys.exit(1)

    if args.once:
        log("🔁  Mode --once : une seule actualisation.")
        success = trigger_refresh()
        sys.exit(0 if success else 1)

    if args.interval < 1:
        log("❌  L'intervalle minimum est 1 minute.")
        sys.exit(1)

    try:
        run_loop(args.interval)
    except KeyboardInterrupt:
        log("\n🛑  Arrêt demandé par l'utilisateur. Au revoir.")
        sys.exit(0)


if __name__ == "__main__":
    main()
