"""
main.py — PyWebView entry point for Maintenance Costing Control Panel
"""

import sys
import os
import threading
import queue
import time
import json
import traceback
from datetime import datetime

# ── DPI awareness (Windows, before any GUI) ─────────────────────────────────
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

# ── Resource path helper (dev vs. PyInstaller) ───────────────────────────────
def _resource_path(relative_path):
    # Vérifier d'abord si une version mise à jour existe dans _updates/
    updates_path = os.path.join(_app_dir(), "_updates", *relative_path.split("/"))
    if os.path.exists(updates_path):
        return updates_path
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

def _app_dir():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))

def _user_log_dir(appname="MaintenanceControlPanel"):
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, appname, "Logs")
    os.makedirs(path, exist_ok=True)
    return path

# ── License check ────────────────────────────────────────────────────────────
sys.path.insert(0, _app_dir())
from license_check import check_or_ask_license, verify_license

ok, msg = check_or_ask_license(show_ui=True)
if not ok:
    sys.exit(msg)

# ── Auto-updater (télécharge les scripts mis à jour depuis GitHub) ────────────
import updater
updater.run(timeout=12)
# Les scripts mis à jour dans _updates/ prennent la priorité sur les bundled
_updates_dir = os.path.join(_app_dir(), "_updates")
if _updates_dir not in sys.path:
    sys.path.insert(0, _updates_dir)

# ── Calculation imports ──────────────────────────────────────────────────────
import Preventive_corrective_Subcontract_planning_strip_MEP_day_Vref        as preventive_script
import Preventive_corrective_Subcontract_Ovh_Renew_planning_MEP_Vref        as ovh_renew_script
import Shift_organization_with_technicians_24hours_7days_Vref               as shift_24_7
import Shift_organization_with_technicians_balanced_hours_restdays_Vref     as shift_balanced_iterated
import Linear_asset_with_day_premises_eng_droite_Vref                       as linear_asset
import APM_Linear_asset_with_day_premises_eng_droite_Vref                   as APM_linear_asset
import Linear_asset_with_Ovh_Renewals_day_premises_eng_droite_Vref          as linear_asset_ovh_renew
import APM_Linear_asset_with_Ovh_Renewals_day_premises_eng_droite_Vref      as APM_linear_asset_ovh_renew
import Assessment_Ops_per_shift_max_hours_balanced_charge_Vref              as shift_assessment
import team_implementation_vehicle_inspection_Vref                          as team_implem_vehicles
import build_benchmark_Synthesis_Recap_All_projects_Vref                    as build_benchmark
import night_shift_assessment_different_ops_shifted_weekends_Vref           as night_shift

import webview

# ── Remplacer les dialogues tkinter par les équivalents PyWebView ────────────
# Les scripts de calcul utilisent tkinter.filedialog en interne.
# tkinter exige le thread principal — nos workers tournent en threads secondaires.
# On remplace globalement pour éviter "main thread is not in main loop".
import tkinter.filedialog as _tkfd

def _pywv_askopenfilename(**kwargs):
    w = webview.windows[0] if webview.windows else None
    if not w: return ""
    result = w.create_file_dialog(
        webview.OPEN_DIALOG, allow_multiple=False,
        file_types=("Excel files (*.xlsx;*.xls;*.xlsm)",)
    )
    return result[0] if result else ""

def _pywv_asksaveasfilename(**kwargs):
    w = webview.windows[0] if webview.windows else None
    if not w: return ""
    result = w.create_file_dialog(
        webview.SAVE_DIALOG,
        save_filename=kwargs.get("initialfile", "output.xlsx"),
        file_types=("Excel files (*.xlsx;*.xlsm)",)
    )
    return result[0] if result else ""

def _pywv_askdirectory(**kwargs):
    w = webview.windows[0] if webview.windows else None
    if not w: return ""
    result = w.create_file_dialog(webview.FOLDER_DIALOG)
    return result[0] if result else ""

_tkfd.askopenfilename   = _pywv_askopenfilename
_tkfd.asksaveasfilename = _pywv_asksaveasfilename
_tkfd.askdirectory      = _pywv_askdirectory

# Remplacer tkinter.messagebox par des boîtes Windows natives (ctypes)
# → thread-safe, rendu natif Windows, aucune dépendance tkinter
import tkinter.messagebox as _tkmb
import ctypes as _ctypes

_MB_OK          = 0x00
_MB_YESNO       = 0x04
_MB_OKCANCEL    = 0x01
_MB_ICONINFO    = 0x40
_MB_ICONWARN    = 0x30
_MB_ICONERROR   = 0x10
_MB_ICONQUESTION= 0x20
_IDYES, _IDOK   = 6, 1

def _win_msgbox(title, message, flags):
    return _ctypes.windll.user32.MessageBoxW(0, str(message), str(title), flags)

def _nat_askyesno(title="", message="", **kw):
    return _win_msgbox(title, message, _MB_YESNO | _MB_ICONQUESTION) == _IDYES

def _nat_askokcancel(title="", message="", **kw):
    return _win_msgbox(title, message, _MB_OKCANCEL | _MB_ICONQUESTION) == _IDOK

def _nat_showinfo(title="", message="", **kw):
    _win_msgbox(title, message, _MB_OK | _MB_ICONINFO)

def _nat_showerror(title="", message="", **kw):
    _win_msgbox(title, message, _MB_OK | _MB_ICONERROR)

def _nat_showwarning(title="", message="", **kw):
    _win_msgbox(title, message, _MB_OK | _MB_ICONWARN)

_tkmb.askyesno       = _nat_askyesno
_tkmb.askyesnocancel = _nat_askyesno
_tkmb.askokcancel    = _nat_askokcancel
_tkmb.showinfo       = _nat_showinfo
_tkmb.showerror      = _nat_showerror
_tkmb.showwarning    = _nat_showwarning

APP_TITLE = "Maintenance Costing Control Panel"
AUTHOR    = "Created by: Mohamed BOUDIA"

# ── Thread-safe stdout/stderr redirector ─────────────────────────────────────
class QueueRedirector:
    def __init__(self, q):
        self.q = q

    def write(self, msg):
        if msg:
            self.q.put(msg)

    def flush(self):
        pass

log_queue: queue.Queue = queue.Queue()

def _js_escape(text: str) -> str:
    """Escape a Python string so it is safe to embed inside JS backtick."""
    return (text
            .replace("\\", "\\\\")
            .replace("`", "\\`")
            .replace("$", "\\$")
            .replace("\r\n", "\\n")
            .replace("\r", "\\n")
            .replace("\n", "\\n"))

def _poll_log(window):
    """Background thread: drain queue and push lines to the JS log."""
    while True:
        time.sleep(0.1)
        lines = []
        try:
            while True:
                lines.append(log_queue.get_nowait())
        except queue.Empty:
            pass
        if lines and window:
            combined = "".join(lines)
            safe = _js_escape(combined)
            try:
                window.evaluate_js(f"appendLog(`{safe}`)")
            except Exception:
                pass

def _write_last_error_log(context: str, exc: Exception, tb_text: str):
    try:
        path = os.path.join(_user_log_dir(), "last_error.log")
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            "=" * 80,
            f"Timestamp : {stamp}",
            f"Context   : {context}",
            f"Error     : {type(exc).__name__}: {exc}",
            "-" * 80,
            tb_text.rstrip(),
            "",
        ]
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return path
    except Exception:
        return None

# ── API class exposed to JavaScript ─────────────────────────────────────────
class Api:
    """Methods callable from JS via window.pywebview.api.<method>(...)"""

    # ── internal run helper ──────────────────────────────────────────────────
    def _run_in_thread(self, card_id: str, context: str, func, *args):
        def worker():
            window = webview.windows[0] if webview.windows else None
            orig_out, orig_err = sys.stdout, sys.stderr

            # Redirecteur qui surveille les mots-clés d'erreur dans la sortie
            error_detected  = [False]
            captured_output = []
            _ERROR_KW = ("error", "erreur", "exception", "traceback", "[✗]")

            class MonitorRedirector:
                def __init__(self, q, flag, lines):
                    self.q = q
                    self.flag = flag
                    self.lines = lines
                def write(self, msg):
                    if msg:
                        self.q.put(msg)
                        self.lines.append(msg)
                        if any(kw in msg.lower() for kw in _ERROR_KW):
                            self.flag[0] = True
                def flush(self):
                    pass

            sys.stdout = MonitorRedirector(log_queue, error_detected, captured_output)
            sys.stderr = MonitorRedirector(log_queue, error_detected, captured_output)

            try:
                if window:
                    window.evaluate_js(f"setCardState('{card_id}', 'running')")
                log_queue.put(f"\n[→] Starting: {context}\n")
                func(*args)

                if error_detected[0]:
                    # Écrire last_error.log avec le contenu capturé
                    try:
                        err_path = os.path.join(_user_log_dir(), "last_error.log")
                        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        with open(err_path, "w", encoding="utf-8") as f:
                            f.write("=" * 80 + "\n")
                            f.write(f"Timestamp : {stamp}\n")
                            f.write(f"Context   : {context}\n")
                            f.write(f"Error     : Script logged an error (no exception raised)\n")
                            f.write("-" * 80 + "\n")
                            f.write("".join(captured_output))
                        log_queue.put(f"[i] Error log saved to: {err_path}\n")
                    except Exception:
                        pass
                    log_queue.put("\n[✗] Completed with errors — check log above.\n")
                    if window:
                        window.evaluate_js(f"setCardState('{card_id}', 'error')")
                else:
                    log_queue.put("\n[✓] Completed successfully.\n")
                    if window:
                        window.evaluate_js(f"setCardState('{card_id}', 'success')")

            except Exception as e:
                tb = traceback.format_exc()
                err_path = _write_last_error_log(context, e, tb)
                log_queue.put(f"\n[✗] {type(e).__name__}: {e}\n")
                log_queue.put(tb + "\n")
                if err_path:
                    log_queue.put(f"[i] Error log saved to: {err_path}\n\n")
                if window:
                    window.evaluate_js(f"setCardState('{card_id}', 'error')")
            finally:
                sys.stdout = orig_out
                sys.stderr = orig_err

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    # ── Section: Railway ────────────────────────────────────────────────────
    def run_railway_linear(self):
        self._run_in_thread(
            "card-railway-linear",
            "Railway Linear Assets (Prev/Correc)",
            linear_asset.main
        )

    def run_railway_ovh(self):
        self._run_in_thread(
            "card-railway-ovh",
            "Railway Linear Assets + Ovh/Renew",
            linear_asset_ovh_renew.main
        )

    # ── Section: APM ────────────────────────────────────────────────────────
    def run_apm_linear(self):
        self._run_in_thread(
            "card-apm-linear",
            "APM Linear Assets (Prev/Correc)",
            APM_linear_asset.main
        )

    def run_apm_ovh(self):
        self._run_in_thread(
            "card-apm-ovh",
            "APM Linear Assets + Ovh/Renew",
            APM_linear_asset_ovh_renew.main
        )

    # ── Section: Shifts / Planning ──────────────────────────────────────────
    def run_maintenance_planning(self):
        self._run_in_thread(
            "card-maintenance-planning",
            "Maintenance Planning",
            preventive_script.main
        )

    def run_maintenance_ovh(self):
        self._run_in_thread(
            "card-maintenance-ovh",
            "Maintenance Planning + Ovh/Renew",
            ovh_renew_script.main
        )

    def run_iterated_shift(self):
        self._run_in_thread(
            "card-iterated-shift",
            "Iterated Shift — Balanced Hours",
            shift_balanced_iterated.main
        )

    def run_shift_247(self, in_path: str, out_path: str):
        self._run_in_thread(
            "card-shift-247",
            "Shift 24/7",
            shift_24_7.main,
            in_path,
            out_path
        )

    # ── Section: Assessment ─────────────────────────────────────────────────
    def run_shift_assessment(self):
        self._run_in_thread(
            "card-shift-assessment",
            "Shift Assessment — Operators",
            shift_assessment.main
        )

    def run_night_shift(self):
        self._run_in_thread(
            "card-night-shift",
            "Night Shift Organization",
            night_shift.main
        )

    def run_team_vehicles(self):
        self._run_in_thread(
            "card-team-vehicles",
            "Team Implementation — Inspection Vehicles",
            team_implem_vehicles.main
        )

    # ── Section: Benchmark ──────────────────────────────────────────────────
    def run_benchmark(self):
        self._run_in_thread(
            "card-benchmark",
            "Benchmark — Projects",
            build_benchmark.main
        )

    # ── File dialogs ────────────────────────────────────────────────────────
    def open_file_dialog(self):
        """Open file picker; returns selected path string or empty string."""
        window = webview.windows[0] if webview.windows else None
        if window is None:
            return ""
        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel files (*.xlsx;*.xls;*.xlsm)", "All files (*.*)")
        )
        if result and len(result) > 0:
            return result[0]
        return ""

    def save_file_dialog(self):
        """Open save dialog; returns chosen path string or empty string."""
        window = webview.windows[0] if webview.windows else None
        if window is None:
            return ""
        result = window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="output.xlsx",
            file_types=("Excel Workbook (*.xlsx)", "Excel Macro-Enabled (*.xlsm)", "All files (*.*)")
        )
        if result and len(result) > 0:
            return result[0]
        return ""

    # ── Log actions ─────────────────────────────────────────────────────────
    def clear_log(self):
        window = webview.windows[0] if webview.windows else None
        if window:
            window.evaluate_js("clearLog()")
        return True

    def save_log(self, content: str):
        """Save log content to a text file chosen by user."""
        window = webview.windows[0] if webview.windows else None
        if window is None:
            return {"ok": False, "error": "No window"}
        result = window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="execution_log.txt",
            file_types=("Text files (*.txt)", "All files (*.*)")
        )
        if result and len(result) > 0:
            path = result[0]
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                return {"ok": True, "path": path}
            except Exception as e:
                return {"ok": False, "error": str(e)}
        return {"ok": False, "error": "Cancelled"}

    def open_last_error_log(self):
        path = os.path.join(_user_log_dir(), "last_error.log")
        if not os.path.exists(path):
            return {"ok": False, "error": "No error log found yet."}
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            return {"ok": True, "content": content}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── License info ─────────────────────────────────────────────────────────
    def get_license_info(self):
        candidates = [
            os.path.join(_app_dir(), "license.json"),
            os.path.join(
                os.environ.get("PROGRAMDATA", r"C:\ProgramData"),
                "MaintenanceControlPanel", "license.json"
            ),
        ]
        lic_path = next((p for p in candidates if os.path.exists(p)), None)
        if not lic_path:
            return {"ok": False, "error": "No license file found."}
        try:
            with open(lic_path, "r", encoding="utf-8") as f:
                lic = json.load(f)
            ok, msg = verify_license(lic)
            return {
                "ok": ok,
                "valid": ok,
                "status": msg,
                "name":     lic.get("name", "N/A"),
                "issued":   lic.get("issued_at", "N/A"),
                "expires":  lic.get("expires_at", "N/A"),
                "features": ", ".join(lic.get("features", [])) or "N/A",
                "hwid":     lic.get("hwid", "—"),
                "path":     lic_path,
            }
        except Exception as e:
            return {"ok": False, "error": str(e)}


# ── Application start ────────────────────────────────────────────────────────
def on_loaded():
    """Called once the WebView has finished loading index.html."""
    window = webview.windows[0]
    # Start log polling thread
    t = threading.Thread(target=_poll_log, args=(window,), daemon=True)
    t.start()
    # Send startup greeting
    log_queue.put(f"[✓] {APP_TITLE} — ready.\n")


def main():
    api    = Api()
    html   = _resource_path(os.path.join("web", "index.html"))
    window = webview.create_window(
        title    = "MCCP",
        url      = html,
        js_api   = api,
        width    = 1200,
        height   = 780,
        min_size = (1000, 660),
        resizable= True,
    )
    window.events.loaded += on_loaded
    webview.start(debug=False)


if __name__ == "__main__":
    main()
