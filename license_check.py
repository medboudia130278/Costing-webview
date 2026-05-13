# license_check.py
import os, sys, json, base64, datetime, hashlib
from nacl.signing import VerifyKey  # pip install pynacl
from nacl.exceptions import BadSignatureError

PUBLIC_B64 = "TjPAccFbF/JpEKiT76wyLszP63Vkt2nRAEgox+2u3eo="  # ta clé publique

def _app_dir():
    return os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) \
           else os.path.dirname(os.path.abspath(__file__))

def _load_license(candidates):
    for p in candidates:
        if p and os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f), p
            except Exception:
                pass
    return None, None

def current_hwid():
    m = hashlib.sha256()
    m.update(os.environ.get("COMPUTERNAME","").encode())
    m.update(os.environ.get("SystemDrive","C:").encode())
    return m.hexdigest()[:24]

def verify_license(lic: dict) -> tuple[bool, str]:
    # Champs de base
    if not isinstance(lic, dict):
        return False, "Format de licence invalide."
    if "sig" not in lic:
        return False, "Licence invalide (signature manquante)."
    if "expires_at" not in lic:
        return False, "Licence invalide (champ 'expires_at' manquant)."

    # 1) Signature
    try:
        sig = base64.b64decode(lic["sig"])
    except Exception:
        return False, "Signature illisible (base64)."

    payload = dict(lic); payload.pop("sig", None)
    try:
        msg = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    except Exception:
        return False, "Payload illisible."

    try:
        vk = VerifyKey(base64.b64decode(PUBLIC_B64))
        vk.verify(msg, sig)
    except BadSignatureError:
        return False, "Signature de licence invalide."
    except Exception:
        return False, "Erreur de vérification de signature."

    # 2) Expiration
    try:
        exp = datetime.date.fromisoformat(payload["expires_at"])
        if datetime.date.today() > exp:
            return False, "Licence expirée."
    except Exception:
        return False, "Champ 'expires_at' invalide."

    # 3) HWID (si présent)
    if payload.get("hwid"):
        if payload["hwid"] != current_hwid():
            return False, "Licence non valide pour cette machine."

    return True, "OK"

def check_or_ask_license(*, show_ui: bool = True) -> tuple[bool, str]:
    # On cherche à côté de l'exe et en ProgramData
    candidates = [
        os.path.join(_app_dir(), "license.json"),
        os.path.join(os.environ.get("PROGRAMDATA", r"C:\ProgramData"),
                     "MaintenanceControlPanel", "license.json"),
    ]
    lic, path = _load_license(candidates)

    if not lic and show_ui:
        # Demander à l’utilisateur de sélectionner sa licence
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo("Activation", "Sélectionnez votre fichier de licence (license.json).")
        sel = filedialog.askopenfilename(title="Choisir la licence", filetypes=[("JSON","*.json")])
        if not sel:
            return False, "Activation annulée par l'utilisateur."
        with open(sel, "r", encoding="utf-8") as f:
            lic = json.load(f)
        # Essayer de copier en ProgramData (si droits), sinon à côté de l’exe
        try:
            target_dir = os.path.join(os.environ.get("PROGRAMDATA", r"C:\ProgramData"), "MaintenanceControlPanel")
            os.makedirs(target_dir, exist_ok=True)
            with open(os.path.join(target_dir, "license.json"), "w", encoding="utf-8") as g:
                json.dump(lic, g, separators=(",",":"), sort_keys=True)
        except Exception:
            try:
                with open(os.path.join(_app_dir(), "license.json"), "w", encoding="utf-8") as g:
                    json.dump(lic, g, separators=(",",":"), sort_keys=True)
            except Exception:
                pass

    if not lic:
        return False, "Licence introuvable."

    ok, msg = verify_license(lic)
    if not ok and show_ui:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Activation", f"Licence invalide : {msg}")
    return ok, msg

# Optionnel: utilitaire pour afficher l'HWID local
def print_hwid():
    print("HWID:", current_hwid())
