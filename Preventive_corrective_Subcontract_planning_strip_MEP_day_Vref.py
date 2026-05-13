# === Wrapped with main() by assistant ===
import sys
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================== CORE PROCESS WRAPPER ==============================
def core_process(input_file=None, output_file=None, *, silent=False):
    # Create a temporary hidden root only if we need dialogs
    import tkinter as tk
    from tkinter import filedialog, messagebox
    temp_root = None
    if (input_file is None or output_file is None) and getattr(tk, "_default_root", None) is None:
        temp_root = tk.Tk()
        temp_root.withdraw()

    # Filedialog wrappers: if path provided, use it; else fall back to dialogs
    def filedialog_askopen(**kwargs):
        if input_file is not None:
            return input_file
        return filedialog.askopenfilename(**kwargs)

    def filedialog_asksave(**kwargs):
        if output_file is not None:
            return output_file
        return filedialog.asksaveasfilename(**kwargs)

    try:
        import pandas as pd
        import math
        import datetime
        import numpy as np # Importer la bibliothèque nécessaire pour travailler avec la loi de Poisson
        

        def read_sheet_strip_cells_or_empty(file_path: str, sheet_name: str, engine="openpyxl"):
            
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
            except ValueError:
                print(f"The sheet '{sheet_name}' does not exist or is invalid.")
                return pd.DataFrame()
            except Exception as e:
                print(f"Failed to read sheet '{sheet_name}' ({e}). An empty DataFrame will be used.")
                return pd.DataFrame()

            if df is None or df.shape[1] == 0:
                print(f"The sheet '{sheet_name}' does not contain any columns.")
                return pd.DataFrame()

            # strip des noms de colonnes (bords uniquement)
            df.columns = df.columns.map(lambda c: str(c).strip())

            # strip des cellules texte pour TOUTES les colonnes
            for col in df.columns:
                if df[col].dtype == "object":
                    df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)

            # "Visuellement vide" (que des NaN)
            if df.dropna(how="all").empty:
                print(f"The sheet '{sheet_name}' is empty or contains only NaN values.")
                return pd.DataFrame(columns=df.columns)

            return df

        def show_message(title, message):
            """Boîte de message compatible Windows / macOS / Linux."""
            try:
                messagebox.showinfo(title, message)
            except Exception:
                print(f"{title}: {message}")

        # Utiliser une boîte de dialogue pour demander à l'utilisateur de sélectionner un fichier Excel
        # Message interactif avant de choisir le fichier excel d'entrée
        if not silent:
            show_message(
                "Input File Selection",
                "Please select the Excel file that may include the merged linear assets file 'Project_name_Combined_Preventive_corrective_inputs_Vxx'.")
        print("Please select the input Excel file to run the Maintenance_planning script...")

        file_path = filedialog_askopen(
            title="Select the Excel file",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )

        # Vérifier si l'utilisateur a bien sélectionné un fichier
        if not file_path:
            print("No file selected. Exiting the program.")
            raise RuntimeError("User cancelled / exited")

        # Lire les données des différentes feuilles du fichier Excel
        try:
            _xl = pd.ExcelFile(file_path)
            _pg_candidates = ['parametres_generaux', 'General Parameters', 'Parameters']
            _pg_sheet = next((s for s in _pg_candidates if s in _xl.sheet_names), None)
            if _pg_sheet is None:
                raise ValueError(f"Sheet not found. Expected one of: {_pg_candidates}")
            general_parameters = pd.read_excel(_xl, sheet_name=_pg_sheet)
            maintenance_activities = pd.read_excel(_xl, sheet_name='maintenance_activities')
            equipments_per_premises = pd.read_excel(_xl, sheet_name='equipements_premises')
            premises_positions = pd.read_excel(_xl, sheet_name='premises_positions')
            print("Data successfully imported.")
        except Exception as e:
            print(f"Error while importing data : {e}")
            exit()

        # Nettoyer les noms des colonnes
        general_parameters.columns = general_parameters.columns.str.strip()
        for col in general_parameters.columns:
            if general_parameters[col].dtype == "object":
                general_parameters[col] = general_parameters[col].map(
                    lambda v: v.strip() if isinstance(v, str) else v
                )
        maintenance_activities.columns = maintenance_activities.columns.str.strip()
        for col in maintenance_activities.columns:
            if maintenance_activities[col].dtype == "object":
                maintenance_activities[col] = maintenance_activities[col].map(
                    lambda v: v.strip() if isinstance(v, str) else v
                )
        equipments_per_premises.columns = equipments_per_premises.columns.str.strip()
        for col in equipments_per_premises.columns:
            if equipments_per_premises[col].dtype == "object":
                equipments_per_premises[col] = equipments_per_premises[col].map(
                    lambda v: v.strip() if isinstance(v, str) else v
                )
        premises_positions.columns = premises_positions.columns.str.strip()
        for col in premises_positions.columns:
            if premises_positions[col].dtype == "object":
                premises_positions[col] = premises_positions[col].map(
                    lambda v: v.strip() if isinstance(v, str) else v
                )

        # Extraire les paramètres généraux en tant que dictionnaire
        # Reconnaît 'Nom' ou 'Name', 'Valeur' ou 'Value'
        nom_col    = 'Nom'    if 'Nom'    in general_parameters.columns else 'Name'
        valeur_col = 'Valeur' if 'Valeur' in general_parameters.columns else 'Value'
        params_dict = pd.Series(general_parameters[valeur_col].values, index=general_parameters[nom_col]).to_dict()

        # Paramètres importants récupérés depuis le fichier Excel
        weeks_per_year = int(params_dict['weeks_per_year'])
        max_night_hours = params_dict['night_work_hours']  # Plage de nuit maximale (par ex. 4 heures)
        max_day_hours = params_dict['day_work_hours']  # Plage de jour maximale
        average_speed_day = params_dict['average_speed_day_km_per_h']  # Vitesse moyenne de jour en km/h
        average_speed_night = params_dict['average_speed_night_km_per_h']  # Vitesse moyenne de nuit en km/h
        start_year = int(params_dict['planning_year'])
        contract_duration_years = int(params_dict['contract_duration_years'])
        days_per_week = int(params_dict['days_per_week'])
        max_hours_per_year_per_person = float(params_dict['max_hours_per_year_per_person']) # utilisé uniquement pour le corrective
        prev_unskilled_coef = float(params_dict['preventive_unskilled_coeficient']) # coefficient d'incompétence en préventive en génerale egal à 1
        corr_unskilled_coef = float(params_dict['corrective_unskilled_coeficient']) # coefficient d'incompétence en corrective dépendant des regions
        paliative_hours_coef = float(params_dict['paliative_hours_coeficient'])

        def get_optional_int(params, key):
            v = params.get(key, None)
            if v is None:
                return None
            # NaN
            if isinstance(v, float) and pd.isna(v):
                return None
            # Texte vide
            if isinstance(v, str):
                v = v.strip()
                if v == "":
                    return None
            # Excel peut renvoyer 10.0 -> ok
            return int(float(v))

        overhaul_start_offset = get_optional_int(params_dict, "correc_ovh_start_year")
        overhaul_end_offset   = get_optional_int(params_dict, "correc_ovh_end_year")

        overhaul_display_from_year = (start_year + overhaul_start_offset) if overhaul_start_offset is not None else None
        overhaul_display_to_year   = (start_year + overhaul_end_offset)   if overhaul_end_offset is not None else None

        # Lire les jours fériés à partir de la feuille 'days_off'
        try:
            days_off = pd.read_excel(file_path, sheet_name='days_off')
            days_off.columns = days_off.columns.str.strip()  # Nettoyer les noms des colonnes
            print("days off successfully imported.")
        except Exception as e:
            print(f"Error while importing days off : {e}")
            exit()

        # Lire les données de la feuille 'corrective_activities'
        corrective_activities = read_sheet_strip_cells_or_empty(file_path, 'corrective_activities')
        if corrective_activities.empty:
            print("No corrective activities to schedule.")
            corrective_activities = None

        try:
            # Lire la feuille subcontract_activities
            subcontract_activities = pd.read_excel(file_path, sheet_name='subcontract_activities')

            # Vérifier si la feuille est vide
            if subcontract_activities.empty:
                print("The 'subcontract_activities' sheet is empty. No subcontracting activities to schedule.")
                subcontract_activities = None  # Définir à None pour éviter d'exécuter la planification
            else:
                # Nettoyer les colonnes
                subcontract_activities.columns = subcontract_activities.columns.str.strip()
                for col in subcontract_activities.columns:
                    if subcontract_activities[col].dtype == "object":
                        subcontract_activities[col] = subcontract_activities[col].map(
                            lambda v: v.strip() if isinstance(v, str) else v
                        )
        except ValueError:
            # Si la feuille n'existe pas dans le fichier d'entrée
            print("The 'subcontract_activities' sheet is missing from the input file.")
            subcontract_activities = None  # Définir à None pour éviter d'exécuter la planification
        except Exception as e:
            print(f"Error while reading the 'subcontract_activities' sheet : {e}")
            subcontract_activities = None  # Définir à None pour éviter d'exécuter la planification

        # Traiter les données de soustraitance si elles existent
        if subcontract_activities is not None:
            # Filtrer pour garder uniquement les lignes avec strategy = 'Included'
            filtered_subcontracting = subcontract_activities[
                subcontract_activities['strategy'] == 'Included'
            ]

            # Sélectionner les colonnes pertinentes
            subcontracting_planning = filtered_subcontracting[[
                'subsystem', 'equipment', 'activity_description', 'type',
                'frequency', 'yearly_cost', 'currency', 'source'
            ]]

            # Vérifier si le DataFrame filtré est vide
            if subcontracting_planning.empty:
                print("No subcontracting activities with strategy 'Included' were found.")
            else:
                print("Subcontracting planning successfully generated.")
        else:
            print("No subcontracting data available for planning.")
            subcontracting_planning = pd.DataFrame()  # DataFrame vide pour éviter les erreurs plus tard

        # Lire la feuille DEQ_VMI_activities
        try:  
            deq_vmi_activities = pd.read_excel(file_path, sheet_name='DEQ_VMI_activities')

            # Vérifier si la feuille est vide
            if deq_vmi_activities.empty:
                print("The 'DEQ_VMI_activities' sheet is empty. No depot equipment activities to schedule.")
                deq_vmi_activities = None  # Définir à None pour éviter d'exécuter la planification
            else:
                # Nettoyer les colonnes
                deq_vmi_activities.columns = deq_vmi_activities.columns.str.strip()
                for col in deq_vmi_activities.columns:
                    if deq_vmi_activities[col].dtype == "object":
                        deq_vmi_activities[col] = deq_vmi_activities[col].map(
                            lambda v: v.strip() if isinstance(v, str) else v
                        )
        except ValueError:
            # Si la feuille n'existe pas dans le fichier d'entrée
            print("The 'DEQ_VMI_activities' is missing frome the input file.")
            deq_vmi_activities = None  # Définir à None pour éviter d'exécuter la planification
        except Exception as e:
            print(f"Error while reading the sheet 'DEQ_VMI_activities' : {e}")
            deq_vmi_activities = None  # Définir à None pour éviter d'exécuter la planification

        # Générer un calendrier pour une année donnée, en excluant les week-ends et les jours fériés
        def generate_calendar_for_year(year):
            calendar = []
            current_date = datetime.date(year, 1, 1)
            holidays =[]
            # Extraire les jours fériés s'ils existent et s'ils sont bien définis
            try:
                if 'days_off' in locals() and not days_off.empty:
                    holidays = [datetime.date(year, int(row['Month']), int(row['Day'])) for _, row in days_off.iterrows()]
            except Exception as e:
                print(f"Error while extracting days off : {e}")

            # Créer un calendrier sans week-ends et sans jours fériés
            while current_date.year == year:
                if current_date.weekday() < days_per_week and current_date not in holidays:  # Exclure les week-ends et les jours fériés
                    calendar.append(current_date)
                current_date += datetime.timedelta(days=1)
            return calendar

        # Générer le calendrier pour l'année de planification
        calendar_2025 = generate_calendar_for_year(start_year)
        print("Annual calendar successfully generated, excluding weekends and public holidays.")

        # Fonction pour diviser une activité qui dépasse la plage horaire (jour ou nuit) en plusieurs sessions
        def divide_activity(total_time, max_hours):
            sessions = []
            remaining_time = total_time
            while remaining_time > max_hours:
                sessions.append(max_hours)
                remaining_time -= max_hours
            if remaining_time > 0:
                sessions.append(remaining_time)
            return sessions

        # Fonction pour Vérifier les Périodes à Éviter (juillet, août et la dernière semaine de décembre)
        def is_date_in_avoid_period(date):
            if date.month in [7, 8]:  # Mois de juillet et août
                return True
            if date.month == 12 and date.day >= 25:  # Dernière semaine de décembre
                return True
            return False

        # Fonction pour calculer le temps de déplacement entre deux locaux techniques
        def calculate_travel_time(location1, location2, speed):
            position1 = premises_positions[premises_positions['Location'] == location1]['Position'].values[0]
            position2 = premises_positions[premises_positions['Location'] == location2]['Position'].values[0]
            distance = abs(position2 - position1)
            travel_time = distance / speed
            return travel_time

        # Fonction pour Trier les localisations par ordre croissant de distance par rapport à la localisation actuelle
        def sort_locations_by_proximity(current_location, locations, premises_positions):
            # Si aucune localisation de référence n'est spécifiée, retourner la liste originale
            if current_location is None:
                return locations

            # Vérifier si la localisation de référence est présente dans les données
            reference_position_data = premises_positions[premises_positions['Location'] == current_location]['Position'].values
            if len(reference_position_data) == 0:
                return locations

            # Extraire la position de référence
            reference_position = reference_position_data[0]

            # Calculer la distance entre la localisation actuelle et les autres localisations
            def get_distance(location):
                # Obtenir la position de la localisation actuelle
                position_data = premises_positions[premises_positions['Location'] == location]['Position'].values
                if len(position_data) == 0:
                    return float('inf')  # Si la position est inconnue, retourner une valeur infinie
                return abs(position_data[0] - reference_position)

            # Trier les localisations en fonction de la distance par rapport à `current_location`
            sorted_locations = sorted(locations, key=get_distance)
            return sorted_locations

        # Fonction pour calculer le nombre de composants à remplacer pendant la durée du contrat
        def calculate_replacements_poisson(mtbf, quantity, contract_duration_years, replacement_estimate=None):
            if replacement_estimate is not None:
                # Si remplacement basé sur estimation (10 %), répartir uniformément sur les années du contrat
                replacements_per_year = replacement_estimate / contract_duration_years
                return [replacements_per_year for _ in range(contract_duration_years)], [replacements_per_year for _ in range(contract_duration_years)]
            else:
                # Calcul basé sur MTBF
                expected_failures_per_unit = (contract_duration_years * 24 * 365) / mtbf
                total_replacements = expected_failures_per_unit * quantity
                
                # Calculer les remplacements par année en utilisant la loi de Poisson
                replacements_per_year_poisson = [np.random.poisson(total_replacements / contract_duration_years) for _ in range(contract_duration_years)]
                replacements_per_year_non_arrondi = [total_replacements / contract_duration_years for _ in range(contract_duration_years)]
                
                return replacements_per_year_poisson, replacements_per_year_non_arrondi

        # déclaration generic_year_counters au niveau global
        generic_year_counters = {}

        # Fonction pour planifier les activités sur le calendrier en utilisant des priorités et des fenêtres de flexibilité
        def assign_dates(activities_df, equipments_df, premises_list, max_hours, weeks_per_year, start_year, time_of_day):
            planning = []
            current_year = start_year
            calendar = generate_calendar_for_year(current_year)
            used_dates = {day: max_hours for day in calendar}  # Dictionnaire pour suivre les heures disponibles par date
            technicians_per_day = {}  # Dictionnaire pour suivre le nombre de techniciens nécessaires par jour

            # Créer une liste triée des activités par périodicité (priorité aux activités les plus fréquentes)
            activities_df = activities_df[activities_df['time_of_day'] == time_of_day].sort_values(by=['periodicity_weeks', 'duration_hours'], ascending=[True, True])
            print(f"{time_of_day} activities successfully sorted for subsystem {subsystem}.")

            # Initialiser la variable pour suivre la localisation précedente
            previous_location = None

            # Regrouper les équipements de tous les locaux techniques pour traiter les activités en fonction de leur fréquence
            for periodicity_weeks in sorted(activities_df['periodicity_weeks'].unique()):
                relevant_activities = activities_df[activities_df['periodicity_weeks'] == periodicity_weeks]

                # Trier les locaux techniques en fonction de la proximité par rapport à la localisation précédente
                if previous_location is None:
                    sorted_premises_list = premises_list
                else:
                    sorted_premises_list = sort_locations_by_proximity(previous_location, premises_list, premises_positions)

                for premises in sorted_premises_list:
                    premises_equipments = equipments_df[equipments_df['Location'] == premises]

                    for idx, row in premises_equipments.iterrows():
                        equipment = row['Equipment']
                        quantity = row['Quantity']

                        # Filtrer les activités pertinentes pour l'équipement actuel
                        equipment_activities = relevant_activities[relevant_activities['equipment'] == equipment]

                        for activity_idx, activity_row in equipment_activities.iterrows():
                            periodicity_weeks = activity_row['periodicity_weeks']
                            duration_hours = activity_row['duration_hours']
                            outage_time_before = activity_row['outage_time_before'] / 60  # Convertir en heures
                            outage_time_after = activity_row['outage_time_after'] / 60  # Convertir en heures
                            technicians_needed = activity_row['technicians_needed']
                            activity_name = activity_row['description']

                            # Calculer la durée totale de l'activité en fonction de la quantité d'équipements
                            total_activity_time = (duration_hours + outage_time_before + outage_time_after) * quantity * prev_unskilled_coef

                            # Déterminer la fenêtre de flexibilité pour l'activité en fonction de sa périodicité
                            if periodicity_weeks == 26:
                                flexibility_weeks = 1
                            elif periodicity_weeks == 52:
                                flexibility_weeks = 2
                            elif periodicity_weeks > 52:
                                flexibility_weeks = 3
                            else:
                                flexibility_weeks = 0  # Pas de flexibilité (cas improbable ici)

                            # Définir le nombre d'occurrences par année
                            if periodicity_weeks > 52:
                                # Pour les activités ayant une périodicité supérieure à 52 semaines, ne planifier que t0 sur l'année de départ
                                occurrences_per_year = 1
                            else:
                                # Pour les autres périodicités <= 52 semaines, calculer le nombre d'occurrences sur l'année
                                occurrences_per_year = math.floor(weeks_per_year / periodicity_weeks)

                            # Trouver une première occurrence (t0) qui respecte la contrainte d'éviter les périodes indésirables (seulement pour les activités > 13 semaines)
                            target_week = 0
                            max_attempts = 42  # Limiter le nombre maximal de tentatives pour éviter les boucles infinies
                            attempts = 0

                            while True:
                                # Convertir target_week en index du jour correspondant dans le calendrier
                                target_day_index = round(target_week / weeks_per_year * len(calendar))
                                target_day_index = min(target_day_index, len(calendar) - 1)
                                current_day = calendar[target_day_index]

                                # Vérifier si la première occurrence est dans une période à éviter pour les activités > 13 semaines
                                if periodicity_weeks > 13 and is_date_in_avoid_period(current_day):
                                    attempts += 1
                                    if attempts >= max_attempts:
                                        # Après un certain nombre de tentatives, autoriser la planification dans les périodes indésirables
                                        print(f"Maximum attempts reached for activity {activity_name}. Scheduling during restricted periods is now allowed.")
                                        break

                                    target_week += 1
                                    if target_week >= weeks_per_year:
                                        print(f"Unable to find a suitable date for activity {activity_name} outside restricted periods.")
                                        continue
                                    else:
                                        continue

                                # Vérifier les occurrences suivantes pour voir si elles tombent dans une période à éviter (seulement pour les activités > 13 semaines)
                                valid_t0 = True
                                if periodicity_weeks > 13:
                                    for occurrence in range(1, occurrences_per_year):
                                        future_week = target_week + occurrence * periodicity_weeks
                                        future_day_index = round(future_week / weeks_per_year * len(calendar))
                                        future_day_index = min(future_day_index, len(calendar) - 1)
                                        future_day = calendar[future_day_index]

                                        if is_date_in_avoid_period(future_day):
                                            valid_t0 = False
                                            break

                                # Si t0 et les occurrences suivantes sont valides, arrêter la recherche
                                if valid_t0:
                                    break
                                else:
                                    attempts += 1
                                    if attempts >= max_attempts:
                                        # Après un certain nombre de tentatives, autoriser la planification dans les périodes indésirables
                                        print(f"Maximum attempts reached for activity {activity_name}. Scheduling during restricted periods is now allowed.")
                                        break
                                    target_week += 1
                                    if target_week >= weeks_per_year:
                                        print(f"Unable to find a suitable date for activity {activity_name} outside restricted periods.")
                                        continue
                            
                            # Initialiser travel_time à 0
                            travel_time = 0

                            # Planifier les occurrences à partir de la première occurrence (t0)
                            for occurrence in range(occurrences_per_year):
                                target_week_adjusted = target_week + occurrence * periodicity_weeks

                                # Convertir target_week_adjusted en index du jour correspondant dans le calendrier
                                target_day_index = round(target_week_adjusted / weeks_per_year * len(calendar))
                                target_day_index = min(target_day_index, len(calendar) - 1)
                                current_day = calendar[target_day_index]

                                # Diviser l'activité si elle dépasse la plage horaire
                                sessions = divide_activity(total_activity_time, max_hours)

                                # Assigner chaque session sur des jours non consécutifs
                                for session in sessions:
                                    # Chercher le prochain jour disponible avec assez d'heures restantes
                                    while (
                                        used_dates.get(current_day, 0) < session or 
                                        (current_day in technicians_per_day and technicians_per_day[current_day] != technicians_needed)
                                    ):
                                        
                                        target_day_index += 1
                                        if target_day_index >= len(calendar):
                                            # Si nous atteignons la fin de l'année, générer une nouvelle année générique vierge
                                            generic_year_name = f"GEN{generic_year_counters[subsystem]}"
                                            generic_year_counters[subsystem] += 1

                                            # Réinitialiser le calendrier pour l'année générique vierge
                                            used_dates = {day: max_hours for day in calendar}
                                            technicians_per_day = {}
                                            # Revenir au début de l'année en utilisant un identifiant générique
                                            target_day_index = 0
                                            print(f"Switching to blank generic year : {generic_year_name}")

                                            # Utiliser le même calendrier pour l'année générique
                                            current_day = calendar[target_day_index]
                                        else:
                                            current_day = calendar[target_day_index]

                                    # Assigner la session
                                    used_dates[current_day] -= session
                                    technicians_per_day[current_day] = technicians_needed
                                    
                                    # Ajouter l'activité planifiée
                                    planning.append({
                                        'Location': premises,
                                        'Activity': activity_name,
                                        'Date': current_day,
                                        'Duration': round(session, 2),
                                        'Technicians Needed': technicians_needed,
                                        'Equipment': equipment,
                                        'Time of Day': time_of_day,
                                        'Travel Time': travel_time if previous_location != premises else 0,
                                        'Planning Year': start_year if generic_year_counters[subsystem] == 1 else f"GEN{generic_year_counters[subsystem] - 1}"
                                    })

                                    # Mettre à jour la localisation précedente
                                    previous_location = premises

                                    # Passer au prochain jour disponible (sans être consécutif)
                                    target_day_index += 1
                                    if target_day_index >= len(calendar):
                                        target_day_index = 0

                                    current_day = calendar[target_day_index]

            print(f"Date allocation completed for {time_of_day} activities in subsystem {subsystem}.")
            return planning

        # Obtenir la liste des sous-systèmes
        def uniques_or_empty(df, col):
            # Renvoie une Series ordonnée des valeurs uniques si df/col ok, sinon Series vide
            if isinstance(df, pd.DataFrame) and not df.empty and col in df.columns:
                return pd.Series(df[col].unique())
            return pd.Series(dtype=object)

        subsystems = pd.Index(pd.unique(pd.concat([
            pd.Series(maintenance_activities['subsystem'].unique()),
            uniques_or_empty(corrective_activities, 'subsystem'),
        ], ignore_index=True)))

        # Initialiser la liste pour accumuler les dataframes des plannings
        all_plannings = []
        all_technicians = []
        all_corrective_plannings = []

        # Boucle pour traiter chaque sous-système indépendamment
        for subsystem in subsystems:
            print(f"Treatment of subsystem : {subsystem}")
            # Initialiser le compteur d'année generiques pour le sous-système
            generic_year_counters[subsystem] = 1

            # Filtrer les données pour le sous-système en cours
            filtered_maintenance_activities = maintenance_activities[maintenance_activities['subsystem'] == subsystem]
            filtered_equipments_per_premises = equipments_per_premises[equipments_per_premises['subsystem'] == subsystem]
            filtered_premises_positions = premises_positions[premises_positions['subsystem'] == subsystem]

            # Ajouter la planification corrective pour le sous-système courant
            filtered_corrective_activities = (
                corrective_activities[corrective_activities['subsystem'] == subsystem]
                if corrective_activities is not None and 'subsystem' in corrective_activities.columns
                else pd.DataFrame()
            )

            # Créer la liste des locaux techniques à partir des équipements filtrés
            premises_list = filtered_equipments_per_premises['Location'].unique()

            # Générer les plannings pour les activités de jour et de nuit
            try:
                night_planning = assign_dates(filtered_maintenance_activities, filtered_equipments_per_premises, premises_list, max_night_hours, weeks_per_year, start_year, 'night')
                print(f"Night scheduling successfully completed for subsystem {subsystem}.")
            except Exception as e:
                print(f"An error occurred during night scheduling for subsystem {subsystem} : {e}")
                continue

            try:
                day_planning = assign_dates(filtered_maintenance_activities, filtered_equipments_per_premises, premises_list, max_day_hours, weeks_per_year, start_year, 'day')
                print(f"Day scheduling successfully completed for subsystem {subsystem}.")
            except Exception as e:
                print(f"An error occurred during day scheduling for subsystem {subsystem} : {e}")
                continue
            
            # Convertir les plannings en DataFrame
            night_planning_df = pd.DataFrame(night_planning)
            day_planning_df = pd.DataFrame(day_planning)

            # Ajouter la colonne 'subsystem' aux DataFrames de planning
            if not night_planning_df.empty:
                night_planning_df['subsystem'] = subsystem
                night_planning_df['Travel Time'] = night_planning_df['Travel Time'].astype(float)
            else:
                # Créer un DataFrame vide avec les colonnes nécessaires pour éviter des erreurs ultérieures
                night_planning_df = pd.DataFrame(columns=['Location', 'Activity', 'Date', 'Duration', 'Technicians Needed',
                                                    'Equipment', 'Time of Day', 'Travel Time', 'Planning Year', 'subsystem'])
            if not day_planning_df.empty:
                day_planning_df['subsystem'] = subsystem
                day_planning_df['Travel Time'] = day_planning_df['Travel Time'].astype(float)
            else:
                # Créer un DataFrame vide avec les colonnes nécessaires pour éviter des erreurs ultérieures
                day_planning_df = pd.DataFrame(columns=['Location', 'Activity', 'Date', 'Duration', 'Technicians Needed',
                                                    'Equipment', 'Time of Day', 'Travel Time', 'Planning Year', 'subsystem'])

            # Fonction pour calculer les distances et le temps total de déplacement entre différents locaux techniques
            def calculate_total_travel_time_for_group(group_df, speed):
                total_travel_time = 0
                previous_location = None

                # Trier les activités par location
                for index, row in group_df.iterrows():
                    current_location = row['Location']
                    if previous_location is not None and previous_location != current_location:
                        travel_time = calculate_travel_time(previous_location, current_location, speed)
                        total_travel_time += travel_time
                    previous_location = current_location

                return total_travel_time

            # Mettre à jour les plannings pour inclure les temps de déplacement

            for planning_df, time_of_day in [(night_planning_df, 'night'), (day_planning_df, 'day')]:
                # Grouper par date de planification et par année de planification
                grouped = planning_df.groupby(['Date', 'Planning Year' , 'subsystem'])

                for (date, planning_year, subsystem), group in grouped:
                    if len(group) > 1:  # Si plusieurs activités sont planifiées le même jour
                        speed = average_speed_night if time_of_day == 'night' else average_speed_day
                        total_travel_time = calculate_total_travel_time_for_group(group, speed)

                        # Répartir le temps de déplacement entre les activités
                        travel_time_per_activity = total_travel_time / len(group)

                        # Mettre à jour le DataFrame
                        for index in group.index:
                            planning_df.at[index, 'Travel Time'] += round(travel_time_per_activity, 2)

            # Filtrer les DataFrames non vides avant la concaténation
            planning_dfs = [night_planning_df, day_planning_df]
            non_empty_planning_dfs = [df for df in planning_dfs if not df.empty]

            # S'assurer qu'il y a des DataFrames à concaténer
            if non_empty_planning_dfs:
                combined_planning_df = pd.concat(non_empty_planning_dfs, ignore_index=True)
            else:
                # Créer un DataFrame vide avec les colonnes nécessaires si aucun planning n'est présent
                combined_planning_df = pd.DataFrame(columns=['Location', 'Activity', 'Date', 'Duration', 'Technicians Needed',
                                                        'Equipment', 'Time of Day', 'Travel Time', 'Planning Year', 'subsystem'])

            

            # "Year" = "Planning Year" (robuste même si DF vide)
            if 'Planning Year' in combined_planning_df.columns:
                combined_planning_df['Year'] = combined_planning_df['Planning Year'].astype(str)
            else:
                combined_planning_df['Year'] = pd.Series(index=combined_planning_df.index, dtype='object')

            # Calculer le nombre de techniciens nécessaires par jour de planification avant la transformation des dates
            technicians_max_per_day_day_shift = {}
            technicians_max_per_day_night_shift = {}

            # Calculer les techniciens de jour et de nuit avant transformation
            for planning_df, time_of_day in [(night_planning_df, 'night'), (day_planning_df, 'day')]:
                grouped = planning_df.groupby(['Date', 'Planning Year', 'subsystem'])

                for (date, planning_year, subsystem), group in grouped:
                    # Le nombre de techniciens pour la journée est le maximum parmi les activités
                    technicians_needed = group['Technicians Needed'].max()

                    # Stocker le nombre maximal de techniciens dans le bon dictionnaire (jour ou nuit)
                    if time_of_day == 'day':
                        if (date, planning_year, subsystem) not in technicians_max_per_day_day_shift:
                            technicians_max_per_day_day_shift[(date, planning_year, subsystem)] = technicians_needed
                        else:
                            technicians_max_per_day_day_shift[(date, planning_year, subsystem)] += technicians_needed
                    elif time_of_day == 'night':
                        if (date, planning_year, subsystem) not in technicians_max_per_day_night_shift:
                            technicians_max_per_day_night_shift[(date, planning_year, subsystem)] = technicians_needed
                        else:
                            technicians_max_per_day_night_shift[(date, planning_year, subsystem)] += technicians_needed

            # Transformer les dates des plannings pour les ramener à l'année 2025
            for planning_df in [night_planning_df, day_planning_df]:
                planning_df['Date'] = planning_df['Date'].apply(lambda d: d.replace(year=start_year))

            # Additionner les techniciens de jour et de nuit par journée après transformation des dates en 2025
            technicians_needed_per_day_2025 = {}

            for (date, planning_year, subsystem), technicians_needed in technicians_max_per_day_day_shift.items():
                # Transformer la date en 2025 tout en gardant le même jour et le même mois
                date_2025 = date.replace(year=start_year)

                # Ajouter les techniciens de jour pour chaque jour transformé en 2025
                if (date_2025, subsystem) not in technicians_needed_per_day_2025:
                    technicians_needed_per_day_2025[(date_2025, subsystem)] = {'day': technicians_needed, 'night': 0}
                else:
                    technicians_needed_per_day_2025[(date_2025, subsystem)]['day'] += technicians_needed

            for (date, planning_year, subsystem), technicians_needed in technicians_max_per_day_night_shift.items():
                # Transformer la date en 2025 tout en gardant le même jour et le même mois
                date_2025 = date.replace(year=start_year)

                # Ajouter les techniciens de nuit pour chaque jour transformé en 2025
                if (date_2025, subsystem) not in technicians_needed_per_day_2025:
                    technicians_needed_per_day_2025[(date_2025, subsystem)] = {'day': 0, 'night': technicians_needed}
                else:
                    technicians_needed_per_day_2025[(date_2025, subsystem)]['night'] += technicians_needed

            # Afficher le nombre de techniciens nécessaires par jour en 2025 (jour et nuit)
            for (date, subsystem), technicians in technicians_needed_per_day_2025.items():
                print(f"Date: {date}, subsystem: {subsystem}, Day Technicians Needed: {technicians['day']}, Night Technicians Needed: {technicians['night']}")

            # Convertir le dictionnaire des techniciens nécessaires en DataFrame
            technicians_needed_list = [
                {'Date': date, 'subsystem': subsystem, 'Day Technicians Needed': data['day'], 'Night Technicians Needed': data['night']}
                for (date, subsystem), data in technicians_needed_per_day_2025.items()
            ]
            technicians_needed_df = pd.DataFrame(technicians_needed_list)

            # Ajouter les Dataframes de jour et de nuit à la liste all plannings
            all_plannings.append(night_planning_df)
            all_plannings.append(day_planning_df)

            # Ajouter les informations des techniciens à la liste all technicians
            all_technicians.append(technicians_needed_df)

            # Planification des activités correctives pour le sous-système actuel
            corrective_planning = []
            if not filtered_corrective_activities.empty:
                for idx, row in filtered_corrective_activities.iterrows():
                    equipment = row['equipment']
                    element = row['element']
                    quantity_system_level = row['quantity_system_level']
                    mtbf = row['MTBF']
                    outage_time_before = row['lock_out_before'] # déja converti en heures sur fichier excel
                    outage_time_after = row['lock_in_after'] # déja converti en heures sur fichier excel
                    replacement_time = row['replacement_time']  # Temps de remplacement d'un composant en heures
                    staff_needed = row['staff_needed']  # Récupérer le nombre de techniciens nécessaires
                    preparation_time = row['preparation_time']
                    new_unit_price = row['new_unit_price']  # Récupérer le prix unitaire du remplacement
                    reparability = row['reparability']
                    reparable_unit_price = row['reparable_unit_price']
                    currency = row['currency']  # Récupérer la devise pour chaque élément
                    source = row['source']
                    material_type = row['material_type']
                    replacement_estimate = None

                    # Si MTBF est manquant ou nul, prendre 10 % de `quantity_system_level`
                    if pd.isna(mtbf) or mtbf == 0:
                        replacement_estimate = quantity_system_level * 0.1
                    
                    # Calculer les remplacements nécessaires
                    replacements_needed, replacements_non_arrondi = calculate_replacements_poisson(mtbf, quantity_system_level, contract_duration_years, replacement_estimate)

                    # Déterminer le coût unitaire à utiliser en fonction de la réparabilité
                    unit_price = new_unit_price if pd.isna(reparability) or reparability != "Y" else reparable_unit_price

                    # Déterminer le coût unitaire pour chaque scénario (avec et sans réparabilité)
                    unit_price_with_reparability = reparable_unit_price if pd.notna(reparability) and reparability == "Y" else 0
                    if pd.isna(unit_price_with_reparability) or reparability != "Y":
                        unit_price_without_reparability = new_unit_price
                    else:
                        unit_price_without_reparability = 0

                    # Planifier les remplacements sur la durée du contrat
                    for year, (replacements, replacements_non_arrondi_per_year) in enumerate(zip(replacements_needed, replacements_non_arrondi), start=start_year):
                        total_replacement_hours = replacements * corr_unskilled_coef * (replacement_time + outage_time_before + outage_time_after + preparation_time)
                        total_replacement_hours_estimated = replacements_non_arrondi_per_year * corr_unskilled_coef * (replacement_time + outage_time_before + outage_time_after + preparation_time)
                        
                        # Calculer le cout total estimé et en utilisant la loi de poisson

                        total_cost_with_reparability = replacements * unit_price_with_reparability
                        total_repair_cost_estimated = replacements_non_arrondi_per_year * unit_price_with_reparability
                        total_cost_without_reparability = replacements * unit_price_without_reparability
                        total_cost_estimated = replacements_non_arrondi_per_year * unit_price_without_reparability

                        # Calculer le nombre de techniciens corrective en se basant le max hours annuel de chaque techniciens
                        if not pd.isna(total_replacement_hours) and total_replacement_hours > 0:
                            headcounts = math.ceil((total_replacement_hours * staff_needed) / max_hours_per_year_per_person)
                        else:
                            headcounts = 0  # Pas besoin de techniciens si aucune heure n'est nécessaire

                        # Calculer également headcounts_estimated de manière similaire
                        if not pd.isna(total_replacement_hours_estimated) and total_replacement_hours_estimated > 0:
                            headcounts_estimated = math.ceil((total_replacement_hours_estimated * staff_needed) / max_hours_per_year_per_person)
                        else:
                            headcounts_estimated = 0  # Pas de techniciens estimés si aucune heure n'est nécessaire

                        # Créer une entrée pour le planning de remplacement
                        corrective_planning.append({
                            'Year': year,
                            'Subsystem': subsystem,
                            'Equipment': equipment,
                            'Element': element,
                            'Replacements': round(replacements, 2),
                            'Replacements Estimated': round(replacements_non_arrondi_per_year, 2),
                            'Total Hours': round(total_replacement_hours, 2),
                            'Total Hours Estimated': round(total_replacement_hours_estimated, 2),
                            'Staff Needed': staff_needed,  # Ajouter le nombre de techniciens nécessaires
                            'Reparable Cost': round(total_cost_with_reparability, 2),  # Ajout du coût total, arrondi à deux décimales
                            'Reparable Cost Estimated': round(total_repair_cost_estimated, 2),
                            'Total Cost': round(total_cost_without_reparability, 2),
                            'Total Cost Estimated': round(total_cost_estimated, 2),
                            'Currency': currency,  # Ajout de la devise correspondante
                            'Source': source,
                            'Material Type': material_type,
                            'Reparability' : reparability,
                            'headcounts' : headcounts,
                            'headcounts_estimated' : headcounts_estimated

                        })

            # Convertir le planning correctif en DataFrame
            corrective_planning_df = pd.DataFrame(corrective_planning)

            # Ajouter à la liste de plannings correctifs s'il y a des activités planifiées
            if not corrective_planning_df.empty:
                all_corrective_plannings.append(corrective_planning_df)            


        # Filtrer les DataFrames non vides avant la concaténation
        all_plannings_non_empty = [df for df in all_plannings if not df.empty]

        # S'assurer qu'il y a des DataFrames à concaténer
        if all_plannings_non_empty:
            combined_planning_df = pd.concat(all_plannings_non_empty, ignore_index=True)
        else:
            # Créer un DataFrame vide avec les colonnes nécessaires si aucun planning n'est présent
            combined_planning_df = pd.DataFrame(columns=['Location', 'Activity', 'Date', 'Duration', 'Technicians Needed',
                                                        'Equipment', 'Time of Day', 'Travel Time', 'Planning Year', 'subsystem'])
            
        # Ajouter la colonne 'maintenance_team' en fonction de 'planning_year'
        combined_planning_df['maintenance_team'] = combined_planning_df['Planning Year'].apply(
            lambda year: f"Team 1" if year == start_year else f"Team {int(str(year).replace('GEN', '')) + 1 if 'GEN' in str(year) else 0}")

        # Combiner toutes les informations de techniciens pour obtenir une vue complete
        combined_technicians_df = pd.concat(all_technicians, ignore_index=True)

        # Filtrer les DataFrames corrective non vides avant la concaténation
        all_corrective_plannings_non_empty = [df for df in all_corrective_plannings if not df.empty]

        # S'assurer qu'il y a des DataFrames corrective à concaténer
        if all_corrective_plannings_non_empty:
            combined_corrective_df = pd.concat(all_corrective_plannings_non_empty, ignore_index=True)
        else:
            # Créer un DataFrame vide avec les colonnes nécessaires si aucun planning corrective n'est présent
            combined_corrective_df = pd.DataFrame(columns=['Year', 'Subsystem', 'Equipment', 'Element', 'Replacements', 'Replacements Estimated',
                                    'Total Hours', 'Total Hours Estimated', 'Staff Needed', 'Total Cost', 'Total Cost Estimated',
                                    'Currency', 'Reparability', 'headcounts', 'headcounts_estimated'])


        # Planification des activités DEQ/VMI si des données existent
        deq_vmi_planning = []
        if deq_vmi_activities is not None:
            # Filtrer pour garder uniquement les activités avec strategy = 'Included'
            filtered_deq_vmi = deq_vmi_activities[deq_vmi_activities['strategy'] == 'Included']

            # Itérer sur chaque ligne filtrée
            for idx, row in filtered_deq_vmi.iterrows():
                subsystem = row['subsystem']
                equipment = row['equipment']
                activity_type = row['type']
                span_life = row['span_life']
                unit_cost = row['unit_cost']
                currency = row['currency']
                source = row['source']

                # Vérifier si span_life est inférieur ou égal à la durée du contrat
                if span_life <= contract_duration_years:
                    # Calculer le nombre de remplacements nécessaires
                    num_replacements = math.ceil(contract_duration_years / span_life)

                    # Commencer à planifier à partir de start_year + span_life
                    first_year = start_year + span_life - 1

                    for i in range(num_replacements):
                        year_of_planning = first_year + (i * span_life)

                        # Vérifier si l'année de planification dépasse la durée du contrat
                        if year_of_planning > start_year + contract_duration_years:
                            break

                        # Ajouter les données dans le planning
                        deq_vmi_planning.append({
                            'Year of Planning': year_of_planning,
                            'Subsystem': subsystem,
                            'Equipment': equipment,
                            'Type': activity_type,
                            'Span Life': span_life,
                            'Unit Cost': unit_cost,
                            'Currency': currency,
                            'Source': source
                        })

            # Convertir le planning en DataFrame
            deq_vmi_planning_df = pd.DataFrame(deq_vmi_planning)
        else:
            # Créer un DataFrame vide si aucune donnée n'existe
            deq_vmi_planning_df = pd.DataFrame(columns=[
                'Year of Planning', 'Subsystem', 'Equipment', 'Type', 'Span Life',
                'Unit Cost', 'Currency', 'Source'
            ])


        # Création de la feuille de synthèse "Synthesis"
        synthesis_data = []

        # Partie préventive
        for subsystem in combined_planning_df['subsystem'].unique():
            for shift in ['day', 'night']:
                filtered_df = combined_planning_df[(combined_planning_df['subsystem'] == subsystem) & (combined_planning_df['Time of Day'] == shift)]
                total_duration = filtered_df['Duration'].sum()

                # Récupérer les valeurs maximales des techniciens nécessaires
                max_day_technicians = combined_technicians_df.loc[combined_technicians_df['subsystem'] == subsystem, 'Day Technicians Needed'].max()
                max_night_technicians = combined_technicians_df.loc[combined_technicians_df['subsystem'] == subsystem, 'Night Technicians Needed'].max()

                synthesis_data.append({
                    'Subsystem': subsystem,
                    'Type': 'Preventive',
                    'Shift': shift,
                    'Total Preventive Duration': round(total_duration, 2),
                    'Day Technicians Calculated': max_day_technicians if shift == 'day' else None,
                    'Day Technicians Optimized': None,
                    'Night Technicians Calculated': max_night_technicians if shift == 'night' else None,
                    'Night Technicians Optimized': None,
                    'Paliative Hours (Corrective)': round(total_duration * paliative_hours_coef, 2)
                })

        # Calculer la somme annuelle et ensuite la moyenne des valeurs
        def empty_groupby_series(name):
            s = pd.Series(dtype='float64')
            s.name = name
            return s

        if combined_corrective_df.dropna(how='all').empty:
            total_hours_per_year = empty_groupby_series('Total Hours')
            total_reparable_cost_per_year = empty_groupby_series('Reparable Cost Estimated')
            total_cost_per_year = empty_groupby_series('Total Cost Estimated')
        else:
            # 7) Groupby robustes
            # dropna=False pour conserver les clés NaN éventuelles (sinon elles disparaissent)
            total_hours_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem'], dropna=False)['Total Hours']
                .sum(min_count=1)  # si toutes NaN -> NaN (ou .fillna(0) si on préfère 0)
            )

            total_reparable_cost_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem', 'Currency', 'Material Type'], dropna=False)['Reparable Cost Estimated']
                .sum(min_count=1)
            )

            total_cost_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem', 'Currency', 'Material Type'], dropna=False)['Total Cost Estimated']
                .sum(min_count=1)
            )

        # Partie corrective pour la synthèse
        if not combined_corrective_df.empty:
            if overhaul_display_from_year is not None:
                combined_corrective_df = combined_corrective_df[
                    combined_corrective_df['Year'] >= overhaul_display_from_year
                ]
            if overhaul_display_to_year is not None:
                combined_corrective_df = combined_corrective_df[
                    combined_corrective_df['Year'] <= overhaul_display_to_year
                ]

            # Recalculer les agrégations APRES filtrage
            total_hours_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem'], dropna=False)['Total Hours']
                .sum(min_count=1)
            )

            total_reparable_cost_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem', 'Currency', 'Material Type'], dropna=False)['Reparable Cost Estimated']
                .sum(min_count=1)
            )

            total_cost_per_year = (
                combined_corrective_df
                .groupby(['Year', 'Subsystem', 'Currency', 'Material Type'], dropna=False)['Total Cost Estimated']
                .sum(min_count=1)
            )

            for subsystem in combined_corrective_df['Subsystem'].unique():
                filtered_df = combined_corrective_df[combined_corrective_df['Subsystem'] == subsystem]
                # Calculer les valeurs qui ne dépendent pas de la devise
                avg_total_hours = total_hours_per_year.xs(subsystem, level='Subsystem').mean()
                max_staff_needed = filtered_df['Staff Needed'].max()
                max_headcounts = filtered_df['headcounts'].max()

                synthesis_data.append({
                    'Subsystem': subsystem,
                    'Type': 'Corrective',
                    'Shift': None,
                    'Yearly Total Hours (Corrective)': math.ceil(avg_total_hours),
                    'Staff Needed (Corrective)': max_staff_needed,
                    'Headcounts (Corrective)': max_headcounts
                })

                for (currency, material_type), group in filtered_df.groupby(['Currency', 'Material Type']):
                    avg_reparable_cost = total_reparable_cost_per_year.xs((subsystem, currency, material_type), level=['Subsystem', 'Currency', 'Material Type']).mean()
                    avg_total_cost = total_cost_per_year.xs((subsystem, currency, material_type), level=['Subsystem', 'Currency', 'Material Type']).mean()
                
                    synthesis_data.append({
                        'Subsystem': subsystem,
                        'Type': 'Corrective',
                        'Shift': None,
                        'Currency': currency,
                        'Material Type': material_type,
                        'Yearly Reparable Cost': round(avg_reparable_cost, 2),
                        'Yearly Total Cost': round(avg_total_cost, 2)    
                    })
            

        # Synthèse des données subcontracting
        if not filtered_subcontracting.empty:
            # Regrouper par subsystem, type, et currency pour calculer les sommes de yearly_cost
            subcontracting_summary = filtered_subcontracting.groupby(['subsystem', 'type', 'currency'])['yearly_cost'].sum().reset_index()

            # Ajouter les données dans synthesis_data
            for idx, row in subcontracting_summary.iterrows():
                synthesis_data.append({
                    'Subsystem': row['subsystem'],
                    'Type': row['type'],
                    'Shift': None,  # Pas de notion de shift pour subcontracting
                    'Currency': row['currency'],
                    'Yearly Cost (Subcontracting)': round(row['yearly_cost'], 2)
                })
        else:
            print("No subcontracting data included in the Synthesis.")


        # Vérifier si `deq_vmi_planning_df` contient des données
        if not deq_vmi_planning_df.empty:
            if overhaul_display_from_year is not None:
                deq_vmi_planning_df = deq_vmi_planning_df[
                    deq_vmi_planning_df['Year of Planning'] >= overhaul_display_from_year
                ]
            if overhaul_display_to_year is not None:
                deq_vmi_planning_df = deq_vmi_planning_df[
                    deq_vmi_planning_df['Year of Planning'] <= overhaul_display_to_year
                ]
                
            # 1. Filtrer les données pour "Preventive" et "Corrective"
            preventive_corrective_df = deq_vmi_planning_df[
                deq_vmi_planning_df['Type'].isin(['Preventive', 'Corrective'])
            ]
            
            # Calculer les moyennes des sommes annuelles pour chaque Subsystem, Currency, et Type
            grouped_pc = preventive_corrective_df.groupby(['Subsystem', 'Currency', 'Type'])
            for (subsystem, currency, activity_type), group in grouped_pc:
                # Calculer la somme des Unit Cost pour chaque année
                yearly_sums = group.groupby('Year of Planning')['Unit Cost'].sum()
                
                # Calculer la moyenne des sommes annuelles
                yearly_average_cost = yearly_sums.mean()
                
                # Ajouter les résultats dans synthesis_data
                synthesis_data.append({
                    'Subsystem': subsystem,
                    'Type': activity_type,
                    'Currency': currency,
                    'Yearly Total Cost': round(yearly_average_cost, 2)
                })

            # 2. Filtrer les données pour "Overhaul" et "Renewal"
            overhaul_renewal_df = deq_vmi_planning_df[
                deq_vmi_planning_df['Type'].isin(['Overhaul', 'Renewal'])
            ]
            
            # Calculer la somme totale des Unit Cost pour chaque Subsystem, Currency, et Type
            grouped_or = overhaul_renewal_df.groupby(['Subsystem', 'Currency', 'Type'])
            for (subsystem, currency, activity_type), group in grouped_or:
                # Calculer la somme totale des Unit Cost
                total_global_cost = group['Unit Cost'].sum()
                
                # Ajouter les résultats dans synthesis_data
                synthesis_data.append({
                    'Subsystem': subsystem,
                    'Type': activity_type,
                    'Currency': currency,
                    'Total Global Cost': round(total_global_cost, 2)
                })
        else:
            print("The `deq_vmi_planning_df` DataFrame is empty. No data to process for the synthesis.")

        # Créer un DataFrame avec les données combinées
        synthesis_df = pd.DataFrame(synthesis_data)

        # Sauvegarder les plannings dans un fichier Excel
        # Message interactif avant de choisir le fichier excel de sortie
        if not silent:
            show_message(
                "Output File Saving",
                "Please preferably name the output Excel file as 'Project_name_Output_planning_and_synthesis_Vref' and save it.")
        print("Please name and save the output Excel file after running the Maintenance_planning script...")

        # Utiliser une boîte de dialogue pour demander à l'utilisateur où enregistrer le fichier Excel
        output_file = filedialog_asksave(
            title="Save Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        # Vérifier si l'utilisateur a bien spécifié un emplacement pour enregistrer le fichier
        if not output_file:
            print("No save location selected. Exiting the program.")
            raise RuntimeError("User cancelled / exited")
        try:
            with pd.ExcelWriter(output_file) as writer:
                general_parameters.to_excel(writer, sheet_name='General Parameters', index=False)
                combined_planning_df.to_excel(writer, sheet_name='Planning Complet', index=False)
                combined_technicians_df.to_excel(writer, sheet_name='Technicians Needed Per Day', index=False)

                # Sauvegarder le planning des activités correctives
                combined_corrective_df.to_excel(writer, sheet_name='Corrective Planning', index=False)

                subcontracting_planning.to_excel(writer, sheet_name='Subcontracting Planning', index=False)
                deq_vmi_planning_df.to_excel(writer, sheet_name='DEQ_VMI_Planning', index=False)
                # Sauvegarder la synthèse
                synthesis_df.to_excel(writer, sheet_name='Synthesis', index=False)

            print(f"Activity and technician planning successfully saved to {output_file}")
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
    p = argparse.ArgumentParser(description="Maintenance planning")
    p.add_argument("--in", dest="input_file", help="Input Excel file")
    p.add_argument("--out", dest="output_file", help="Output Excel file")
    p.add_argument("--silent", action="store_true", help="Reduce popups (if any)")
    args = p.parse_args()
    main(args.input_file, args.output_file, silent=args.silent)        
    