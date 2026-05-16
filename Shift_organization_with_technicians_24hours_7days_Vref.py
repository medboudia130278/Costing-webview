# === Wrapped with main() by assistant ===
import os
import sys
import argparse

# Fix conflit Tcl/Tk entre plusieurs installations Python — doit s'exécuter avant import tkinter
if sys.platform == "win32":
    _py_dir = os.path.dirname(sys.executable)
    _tcl_path = os.path.join(_py_dir, "tcl", "tcl8.6")
    _tk_path  = os.path.join(_py_dir, "tcl", "tk8.6")
    if os.path.exists(_tcl_path):
        os.environ["TCL_LIBRARY"] = _tcl_path
    if os.path.exists(_tk_path):
        os.environ["TK_LIBRARY"] = _tk_path

import tkinter as tk
from tkinter import filedialog

# ============================== CORE PROCESS WRAPPER ==============================
def core_process(input_file=None, output_file=None, *, silent=False):
    import tkinter as _tk
    from tkinter import filedialog as _filedialog
    # Create a temporary hidden root only if we need dialogs
    temp_root = None
    needs_dialog = (input_file is None or output_file is None)
    if needs_dialog and getattr(_tk, "_default_root", None) is None:
        temp_root = _tk.Tk()
        temp_root.withdraw()

    # Filedialog wrappers: if path provided, use it; else fall back to dialogs
    def filedialog_askopen(**kwargs):
        if input_file is not None:
            return input_file
        return _filedialog.askopenfilename(**kwargs)

    def filedialog_asksave(**kwargs):
        if output_file is not None:
            return output_file
        return _filedialog.asksaveasfilename(**kwargs)

    try:
        import pandas as pd
        import math
        from datetime import timedelta
        import tkinter as tk
        from tkinter import filedialog
        
        # Créer une fenêtre Tkinter pour ouvrir le fichier d'entrée
        
        file_path = filedialog_askopen(
            title="Select the Excel file",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        
        if not file_path:
            print("No file selected. Exiting the program.")
            raise RuntimeError("User cancelled / exited")
        
        # Charger les données d'entrée
        tech_needed = pd.read_excel(file_path, sheet_name="Technicians Needed Per Day")
        _xl_gp = pd.ExcelFile(file_path)
        _pg_candidates = ['General Parameters', 'parametres_generaux', 'Parameters']
        _pg_sheet = next((s for s in _pg_candidates if s in _xl_gp.sheet_names), None)
        if _pg_sheet is None:
            raise ValueError(f"Sheet not found. Expected one of: {_pg_candidates}")
        general_params = pd.read_excel(_xl_gp, sheet_name=_pg_sheet)
        # Reconnaît 'Nom' ou 'Name', 'Valeur' ou 'Value'
        nom_col    = 'Nom'    if 'Nom'    in general_params.columns else 'Name'
        valeur_col = 'Valeur' if 'Valeur' in general_params.columns else 'Value'
        if nom_col != 'Nom' or valeur_col != 'Valeur':
            general_params = general_params.rename(columns={nom_col: 'Nom', valeur_col: 'Valeur'})

        # Récupérer les paramètres généraux
        max_hours_per_year = float(general_params.loc[general_params["Nom"] == "max_hours_per_year", "Valeur"].values[0])
        vacation_days = int(general_params.loc[general_params["Nom"] == "vacation_days", "Valeur"].values[0])
        consecutive_vacation_days = int(general_params.loc[general_params["Nom"] == "consecutive_vacation_days", "Valeur"].values[0])
        hours_per_shift = float(general_params.loc[general_params["Nom"] == "hours_per_shift", "Valeur"].values[0])
        technicians_per_supervisor = int(general_params.loc[general_params["Nom"] == "technicians_per_supervisor", "Valeur"].values[0])
        allowed_days_per_week = int(general_params.loc[general_params["Nom"] == "allowed_days_per_week", "Valeur"].values[0])
        
        # Initialiser les techniciens par sous-système (jour et nuit séparés)
        technicians = {}
        technician_stats = {}  # Initialiser technician_stats ici
        for subsystem in tech_needed["subsystem"].unique():
            day_needed_max = int(general_params.loc[
                (general_params["Nom"] == "max_day_technicians") &
                (general_params["Subsystem"] == subsystem),
                "Valeur"
            ].values[0])
            night_needed_max = int(general_params.loc[
                (general_params["Nom"] == "max_night_technicians") &
                (general_params["Subsystem"] == subsystem),
                "Valeur"
            ].values[0])
        
            technicians[subsystem] = {
                "day_techs": [
                    {
                        "name": f"Tech_{subsystem}_Day_{i + 1}",
                        "hours_worked": 0,
                        "max_hours_per_year": max_hours_per_year,
                        "last_shift": None,
                        "vacation_days": vacation_days,
                        "vacation_taken": 0,
                        "consecutive_vacation_days": consecutive_vacation_days,
                        "assigned_days_count": 0,  # Nouveau compteur pour le nombre de jours assignés
                        "work_streak": [] # Pour suivre les dates travaillés recentes
                    }
                    for i in range(day_needed_max)
                ],
                "night_techs": [
                    {
                        "name": f"Tech_{subsystem}_Night_{i + 1}",
                        "hours_worked": 0,
                        "max_hours_per_year": max_hours_per_year,
                        "last_shift": None,
                        "vacation_days": vacation_days,
                        "vacation_taken": 0,
                        "consecutive_vacation_days": consecutive_vacation_days,
                        "assigned_days_count": 0,
                        "work_streak": [] # Pour suivre les dates travaillés recentes
                    }
                    for i in range(night_needed_max)
                ]
            }
            
            # C'est la logique pour les superviseurs
            total_technicians = day_needed_max + night_needed_max
            if total_technicians < technicians_per_supervisor:
                num_supervisors = 1
            else:
                num_supervisors = (total_technicians + technicians_per_supervisor - 1) // technicians_per_supervisor  # Cela fait un arrondi à l'entier supérieur
        
            # Ajouter les superviseurs à la structure du sous-système
            technicians[subsystem]["supervisors"] = [
                {
                    "name": f"Supervisor_{subsystem}_{i + 1}",
                    "hours_worked": 0,
                    "max_hours_per_year": max_hours_per_year,
                    "last_shift": None,
                    "vacation_days": vacation_days,
                    "vacation_taken": 0,
                    "consecutive_vacation_days": consecutive_vacation_days
                }
                for i in range(num_supervisors)
            ]
        
        # Liste pour stocker les vacations
        vacation_schedule = []
        
        # Trier les dates dans l'ordre chronologique
        ordered_dates = tech_needed.sort_values(by="Date")
        
        # Initialiser la liste des shifts
        shifts = []  # Cette liste contiendra tous les shifts assignés (jour et nuit)
        # Variable pour suivre les technciens mobile pour chaque date
        assigned_techs_for_date = {}
        
        # Assignation des shifts jour/nuit pour chaque journée du planning
        for _, row in ordered_dates.iterrows():
            date, subsystem = row["Date"], row["subsystem"]
            day_needed = row["Day Technicians Needed"]
            night_needed = row["Night Technicians Needed"]
        
            # Initialiser la liste des techniciens déjà affectés pour la date courante
            if date not in assigned_techs_for_date:
                assigned_techs_for_date[date] = set()
        
            # Initialiser la liste des technciens deja affectés pour la date courante
            assigned_day = []
            assigned_night = []
        
            # Assignation des shifts de jour pour la journée courante
            available_day_techs = [
                tech for tech in technicians[subsystem]["day_techs"]
                if tech["hours_worked"] + hours_per_shift <= tech["max_hours_per_year"]
                and len([d for d in tech["work_streak"] if (date - d).days < 7]) < allowed_days_per_week
            ]
            # Trier les techniciens disponibles par nombre de jours assignés (priorité à ceux qui en ont le moins)
            available_day_techs.sort(key=lambda x: x["assigned_days_count"])
        
            for i in range(day_needed):
                if i < len(available_day_techs):
                    tech = available_day_techs[i]
                    assigned_day.append((date, subsystem, tech["name"], "Day"))
                    tech["last_shift"] = (date, "Day")
                    tech["hours_worked"] += hours_per_shift
                    tech["work_streak"].append(date)
                    assigned_techs_for_date[date].add(tech["name"])
                    # Ne pas incrémenter le compteur si la date est en juillet ou août
                    if date.month not in [7, 8]:
                        tech["assigned_days_count"] += 1
                else:
                    # Deuxième étape : Si on a besoin de plus de techniciens, on assigne les autres disponibles sans trier
                    other_day_techs = [
                        tech for tech in technicians[subsystem]["day_techs"]
                        if tech not in available_day_techs
                        and tech["hours_worked"] + hours_per_shift <= tech["max_hours_per_year"]
                        and len([d for d in tech["work_streak"] if (date - d).days < 7]) < allowed_days_per_week
                    ]
                    if other_day_techs:
                        tech = other_day_techs[0]
                        assigned_day.append((date, subsystem, tech["name"], "Day"))
                        tech["last_shift"] = (date, "Day")
                        tech["hours_worked"] += hours_per_shift
                        tech["work_streak"].append(date)
                        tech["assigned_days_count"] += 1
                        assigned_techs_for_date[date].add(tech["name"])
        
        
            # Assignation des shifts de nuit pour la journée courante
            available_night_techs = [
                tech for tech in technicians[subsystem]["night_techs"]
                if tech["hours_worked"] + hours_per_shift <= tech["max_hours_per_year"]
                and len([d for d in tech["work_streak"] if (date - d).days < 7]) < allowed_days_per_week
            ]
            # Trier les techniciens disponibles par nombre de jours assignés (priorité à ceux qui en ont le moins)
            available_night_techs.sort(key=lambda x: x["assigned_days_count"])
        
            for i in range(night_needed):
                if i < len(available_night_techs):
                    tech = available_night_techs[i]
                    assigned_night.append((date, subsystem, tech["name"], "Night"))
                    tech["last_shift"] = (date, "Night")
                    tech["hours_worked"] += hours_per_shift
                    tech["work_streak"].append(date)
                    assigned_techs_for_date[date].add(tech["name"])
                    # Ne pas incrémenter le compteur si la date est en juillet ou août
                    if date.month not in [7, 8]:
                        tech["assigned_days_count"] += 1
                else:
                    # Deuxième étape : Si on a besoin de plus de techniciens, on assigne les autres disponibles sans trier
                    other_night_techs = [
                        tech for tech in technicians[subsystem]["night_techs"]
                        if tech not in available_night_techs
                        and tech["hours_worked"] + hours_per_shift <= tech["max_hours_per_year"]
                        and len([d for d in tech["work_streak"] if (date - d).days < 7]) < allowed_days_per_week
                    ]
                    if other_night_techs:
                        tech = other_night_techs[0]
                        assigned_night.append((date, subsystem, tech["name"], "Night"))
                        tech["last_shift"] = (date, "Night")
                        tech["hours_worked"] += hours_per_shift
                        tech["work_streak"].append(date)
                        tech["assigned_days_count"] += 1
                        assigned_techs_for_date[date].add(tech["name"])
             
        
            # Ajouter les shifts assignés pour cette journée au planning
            shifts.extend(assigned_day + assigned_night)
        
            # Mettre à jour la liste de travail consécutif en retirant les dates qui sont au-delà de 7 jours
            for tech in technicians[subsystem]["day_techs"] + technicians[subsystem]["night_techs"]:
                tech["work_streak"] = [d for d in tech["work_streak"] if (date - d).days < 7]
        
        # Créer un DataFrame à partir des shifts assignés
        shifts_df = pd.DataFrame(shifts, columns=["Date", "Subsystem", "Technician", "Shift"])
        
        # Vérification des écarts après l'assignation des techniciens
        day_gap = day_needed - len(assigned_day)
        night_gap = night_needed - len(assigned_night)
        
        # Liste pour stocker les techniciens mobilisés
        mobilized_techs = []
        
        # Étape 2 : Mobilisation des techniciens disponibles d'autres sous-systèmes
        if day_gap > 0 or night_gap > 0:
            for other_subsystem, techs in technicians.items():
                if other_subsystem != subsystem:
                    # Mobilisation des techniciens de jour d'autres sous-systèmes
                    if day_gap > 0:
                        # Enregistrer la mobilisation
                        mobilized_techs.append({
                            "Technician": tech["name"],
                            "Mobilized From": other_subsystem,
                            "Mobilized To": subsystem,
                            "Date": date,
                            "Shift": "Day"
                        })
        
                    # Mobilisation des techniciens de nuit d'autres sous-systèmes
                    if night_gap > 0:
                        # Enregistrer la mobilisation
                        mobilized_techs.append({
                            "Technician": tech["name"],
                            "Mobilized From": other_subsystem,
                            "Mobilized To": subsystem,
                            "Date": date,
                            "Shift": "Night"
                            })
        
        # Ajouter les mobilisations au rapport de comparaison
        mobilization_report_df = pd.DataFrame(mobilized_techs)
        
        # Générer le rapport des heures travaillées
        hours_report = []
        for subsystem, techs in technicians.items():
            for shift_type, tech_list in techs.items():
                if shift_type == "supervisors":
                    continue
                for tech in tech_list:
                    hours_report.append({
                        "Technician": tech["name"],
                        "Subsystem": subsystem,
                        "Shift Type": "Day" if shift_type == "day_techs" else "Night",
                        "Hours Worked": tech["hours_worked"],
                        "Max Hours Per Year": tech["max_hours_per_year"],
                        "Remaining Hours": tech["max_hours_per_year"] - tech["hours_worked"]
                    })
        
        hours_report_df = pd.DataFrame(hours_report)
        
        # Identifier les jours libres pour chaque technicien
        all_dates = tech_needed["Date"].unique()  # Liste des dates uniques du calendrier
        technician_free_days = {}
        
        for _, row in shifts_df.iterrows():
            tech_name = row["Technician"]
            assigned_date = row["Date"]
        
            if tech_name not in technician_free_days:
                technician_free_days[tech_name] = {
                    "assigned_dates": set(),
                    "free_dates": set(all_dates)
                }
            technician_free_days[tech_name]["assigned_dates"].add(assigned_date)
        
        for tech_name, dates in technician_free_days.items():
            dates["free_dates"] = dates["free_dates"] - dates["assigned_dates"]
        
        free_days_report = []
        for subsystem, techs in technicians.items():
            for tech_name, dates in technician_free_days.items():
                if tech_name.startswith(f"Tech_{subsystem}"):
                    for free_date in sorted(dates["free_dates"]):
                        free_days_report.append({
                            "Technician": tech_name,
                            "Subsystem": subsystem,
                            "Free Date": free_date
                        })
        
        free_days_report_df = pd.DataFrame(free_days_report)
        
        # Identifier les jours de vacances et de repos
        vacation_schedule = []
        rest_schedule = []
        
        
        for tech_name, dates in technician_free_days.items():
            free_dates_sorted = sorted(dates["free_dates"])
            assigned_vacation_dates = set()  # Utiliser un set pour suivre les dates déjà assignées en vacances
            
            # Priorité pour les vacances en juillet et août
            july_august_dates = [date for date in free_dates_sorted if date.month in [7, 8]]
            other_dates = [date for date in free_dates_sorted if date.month not in [7, 8]]
        
            # Assigner les vacances consécutives
            consecutive_vacation_days_assigned = 0
            for date in july_august_dates:
                if consecutive_vacation_days_assigned < consecutive_vacation_days and date not in assigned_vacation_dates:
                    vacation_schedule.append({
                        "Technician": tech_name,
                        "Vacation Date": date
                    })
                    assigned_vacation_dates.add(date)
                    consecutive_vacation_days_assigned += 1
                
        
            # Assigner les jours de vacances restants
            total_vacation_days_assigned = consecutive_vacation_days_assigned
            for date in july_august_dates + other_dates:
                if total_vacation_days_assigned < vacation_days and date not in assigned_vacation_dates:
                    vacation_schedule.append({
                        "Technician": tech_name,
                        "Vacation Date": date
                    })
                    assigned_vacation_dates.add(date)
                    total_vacation_days_assigned += 1
                
        
            # Assigner les jours de repos
            recovery_rest_days_assigned = 0
            for date in free_dates_sorted:
                if recovery_rest_days_assigned < general_params.loc[general_params["Nom"] == "recovery_rest_days", "Valeur"].values[0] and date not in assigned_vacation_dates:
                    rest_schedule.append({
                        "Technician": tech_name,
                        "Rest Date": date
                    })
                    recovery_rest_days_assigned += 1
                
        
        # Créer des DataFrames pour les vacances et les jours de repos
        vacation_schedule_df = pd.DataFrame(vacation_schedule)
        rest_schedule_df = pd.DataFrame(rest_schedule)
        
        # Ajouter une colonne 'Type' pour différencier les vacations des jours de repos
        vacation_schedule_df["Type"] = "Vacation"
        rest_schedule_df["Type"] = "Rest"
        
        # Fusionner les deux DataFrames
        vacation_and_rest_schedule_df = pd.concat([vacation_schedule_df, rest_schedule_df], ignore_index=True)
        
        # Calculer le nombre total de jours dans le calendrier
        total_days_in_calendar = len(all_dates)
        
        # Calculer les jours assignés pour chaque technicien à partir de shifts_df
        assigned_days_count = shifts_df.groupby("Technician")["Date"].nunique()
        
        # Calculer les jours de vacances pris, les jours de repos pris et les jours disponibles
        
        for tech_name in technician_free_days.keys():
            assigned_days = assigned_days_count.get(tech_name, 0)
            vacation_days_taken = vacation_schedule_df[vacation_schedule_df["Technician"] == tech_name].shape[0]
            rest_days_taken = rest_schedule_df[rest_schedule_df["Technician"] == tech_name].shape[0]
            available_days = math.floor((max_hours_per_year/hours_per_shift) - assigned_days)
            max_available_days = len(all_dates) - assigned_days - vacation_days_taken - rest_days_taken
        
            technician_stats[tech_name] = {
                "assigned_days": assigned_days,
                "vacation_days_taken": vacation_days_taken,
                "rest_days_taken": rest_days_taken,
                "available_days": available_days,
                "max_available_days": max_available_days
            }
        
        # Ajouter les informations au DataFrame hours_report_df
        hours_report_enriched = []
        for _, row in hours_report_df.iterrows():
            tech_name = row["Technician"]
            stats = technician_stats[tech_name]
        
            row["Assigned Days"] = stats["assigned_days"]
            row["Vacation Days Taken"] = stats["vacation_days_taken"]
            row["Rest Days Taken"] = stats["rest_days_taken"]
            row["Available Days"] = stats["available_days"]
            row["Max Available Days"] = stats["max_available_days"]
        
            hours_report_enriched.append(row)
        
        # Créer un nouveau DataFrame enrichi
        hours_report_df = pd.DataFrame(hours_report_enriched)
        
        # Étape 1 : Générer un résumé du planning assigné
        assigned_summary = shifts_df.groupby(["Date", "Subsystem", "Shift"]).size().reset_index(name="Assigned Count")
        
        # Étape 2 : Fusionner avec les besoins pour comparaison
        # On prépare d'abord un DataFrame qui contient les besoins (tech_needed)
        tech_needed_summary = tech_needed.copy()
        tech_needed_summary.rename(columns={
            "Day Technicians Needed": "Day Needed",
            "Night Technicians Needed": "Night Needed"
        }, inplace=True)
        
        # Ajouter des colonnes pour le nombre de techniciens assignés de jour et de nuit
        tech_needed_summary["Day Assigned"] = 0
        tech_needed_summary["Night Assigned"] = 0
        
        # Étape 3 : Mettre à jour les colonnes "Day Assigned" et "Night Assigned" avec les valeurs assignées réelles
        for _, row in assigned_summary.iterrows():
            date, subsystem, shift, count = row["Date"], row["Subsystem"], row["Shift"], row["Assigned Count"]
            if shift == "Day":
                tech_needed_summary.loc[
                    (tech_needed_summary["Date"] == date) & (tech_needed_summary["subsystem"] == subsystem), "Day Assigned"
                ] = count
            elif shift == "Night":
                tech_needed_summary.loc[
                    (tech_needed_summary["Date"] == date) & (tech_needed_summary["subsystem"] == subsystem), "Night Assigned"
                ] = count
        
        # Étape 4 : Calculer le gap entre les besoins et l'assignation
        tech_needed_summary["Day Gap"] = tech_needed_summary["Day Needed"] - tech_needed_summary["Day Assigned"]
        tech_needed_summary["Night Gap"] = tech_needed_summary["Night Needed"] - tech_needed_summary["Night Assigned"]
        
        
        # Exporter le planning final vers un fichier Excel
        output_file = filedialog_asksave(
            title="Save Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if not output_file:
            print("No save location selected. Exiting the program.")
            raise RuntimeError("User cancelled / exited")
        
        try:
            with pd.ExcelWriter(output_file) as writer:
                shifts_df.to_excel(writer, sheet_name="final_schedule", index=False)
                free_days_report_df.to_excel(writer, sheet_name="free_days_report", index=False)
                vacation_and_rest_schedule_df.to_excel(writer, sheet_name="vacation_and_rest_schedule", index=False)
                hours_report_df.to_excel(writer, sheet_name="hours_report", index=False)
                tech_needed_summary.to_excel(writer, sheet_name="gap_analysis", index=False)
                if not mobilization_report_df.empty:
                    mobilization_report_df.to_excel(writer, sheet_name="mobilization_report", index=False)
            print(f"Planning successfully saved to {output_file}")
        except Exception as e:
            print(f"Error while saving the Excel file : {e}")
    finally:
        if temp_root is not None:
            try:
                temp_root.destroy()
            except Exception:
                pass

def main(input_file=None, output_file=None, *, silent=False):
    return core_process(input_file=input_file, output_file=output_file, silent=silent)

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Shift 24/7 organization (wrapped with main)")
    p.add_argument("--in", dest="input_file", help="Input Excel file")
    p.add_argument("--out", dest="output_file", help="Output Excel file")
    p.add_argument("--silent", action="store_true", help="Reduce popups (if any)")
    args = p.parse_args()
    main(args.input_file, args.output_file, silent=args.silent)
