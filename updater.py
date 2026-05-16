"""
updater.py — Auto-updater silencieux via GitHub.

Au démarrage de l'exe, vérifie si une nouvelle version des scripts
est disponible sur GitHub et la télécharge dans _updates/.
Aucune exception n'est jamais levée : en cas d'échec, l'exe continue
avec la version précédente.
"""

import os
import sys
import json
import threading
import urllib.request

GITHUB_REPO    = "medboudia130278/Costing-webview"
GITHUB_RAW     = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main"
GITHUB_API     = f"https://api.github.com/repos/{GITHUB_REPO}/commits/main"

# Fichiers mis à jour automatiquement à chaque push
UPDATABLE_FILES = [
    "Preventive_corrective_Subcontract_planning_strip_MEP_day_Vref.py",
    "Preventive_corrective_Subcontract_Ovh_Renew_planning_MEP_Vref.py",
    "Shift_organization_with_technicians_24hours_7days_Vref.py",
    "Shift_organization_with_technicians_balanced_hours_restdays_Vref.py",
    "Linear_asset_with_day_premises_eng_droite_Vref.py",
    "APM_Linear_asset_with_day_premises_eng_droite_Vref.py",
    "Linear_asset_with_Ovh_Renewals_day_premises_eng_droite_Vref.py",
    "APM_Linear_asset_with_Ovh_Renewals_day_premises_eng_droite_Vref.py",
    "Assessment_Ops_per_shift_max_hours_balanced_charge_Vref.py",
    "team_implementation_vehicle_inspection_Vref.py",
    "build_benchmark_Synthesis_Recap_All_projects_Vref.py",
    "night_shift_assessment_different_ops_shifted_weekends_Vref.py",
    "web/index.html",
    "web/app.js",
    "web/style.css",
]


def _app_dir() -> str:
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))


def get_updates_dir() -> str:
    d = os.path.join(_app_dir(), "_updates")
    os.makedirs(os.path.join(d, "web"), exist_ok=True)
    return d


def _version_file() -> str:
    return os.path.join(_app_dir(), ".version")


def _get_local_sha() -> str | None:
    try:
        with open(_version_file(), encoding="utf-8") as f:
            return f.read().strip() or None
    except Exception:
        return None


def _save_sha(sha: str) -> None:
    try:
        with open(_version_file(), "w", encoding="utf-8") as f:
            f.write(sha)
    except Exception:
        pass


def _get_remote_sha(timeout: int = 5) -> str | None:
    try:
        req = urllib.request.Request(
            GITHUB_API,
            headers={"User-Agent": "MCCP-Updater/1.0"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode())["sha"]
    except Exception:
        return None


def _download_file(rel_path: str, dest_dir: str, timeout: int = 10) -> None:
    url  = f"{GITHUB_RAW}/{rel_path}"
    dest = os.path.join(dest_dir, *rel_path.split("/"))
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "MCCP-Updater/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        with open(dest, "wb") as f:
            f.write(r.read())


def _do_update() -> None:
    """Logique principale — appelée dans un thread avec timeout."""
    remote_sha = _get_remote_sha(timeout=5)
    if not remote_sha:
        return  # Pas de connexion

    if remote_sha == _get_local_sha():
        return  # Déjà à jour

    updates_dir = get_updates_dir()
    all_ok = True
    for rel_path in UPDATABLE_FILES:
        try:
            _download_file(rel_path, updates_dir, timeout=10)
        except Exception:
            all_ok = False

    if all_ok:
        _save_sha(remote_sha)


def run(timeout: int = 12) -> None:
    """
    Point d'entrée public.
    Lance la mise à jour dans un thread avec timeout global.
    Retourne toujours — l'exe démarre quoi qu'il arrive.
    """
    t = threading.Thread(target=_do_update, daemon=True)
    t.start()
    t.join(timeout=timeout)
