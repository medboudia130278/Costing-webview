import pandas as pd
import math
import numpy as np
from tkinter import Tk, filedialog, messagebox

import os

def lire_parametres(file_path):
    """Lit les parametres generaux depuis la feuille 'parametres_generaux', 'General Parameters' ou 'Parameters'."""
    _xl = pd.ExcelFile(file_path)
    _candidates = ['parametres_generaux', 'General Parameters', 'Parameters']
    _sheet = next((s for s in _candidates if s in _xl.sheet_names), None)
    if _sheet is None:
        raise ValueError(f"Aucune feuille de parametres trouvee. Attendu parmi : {_candidates}")
    general_parameters = pd.read_excel(_xl, sheet_name=_sheet)
    general_parameters.columns = general_parameters.columns.map(lambda c: str(c).strip())
    for col in general_parameters.columns:
        if general_parameters[col].dtype == "object":
            general_parameters[col] = general_parameters[col].map(
                lambda v: v.strip() if isinstance(v, str) else v
            )
    # Reconnaît 'Nom' ou 'Name', 'Valeur' ou 'Value'
    nom_col    = 'Nom'    if 'Nom'    in general_parameters.columns else 'Name'
    valeur_col = 'Valeur' if 'Valeur' in general_parameters.columns else 'Value'
    params_dict = pd.Series(general_parameters[valeur_col].values, index=general_parameters[nom_col]).to_dict()
    return params_dict

def calcul_zones(params_dict):
    """Calcule toutes les zones d’inspection/maintenance en fonction des paramètres."""

    # Récupération des paramètres
    type_track_installation = params_dict.get('type_track_installation', None)
    depot_track_installation = params_dict.get('depot_track_installation', None)
    V_ballast_tamping = params_dict.get('V_ballast_tamping', None)

    cleaning_strategy = params_dict.get('cleaning_strategy', 'out_of_scope')
    V_cleaning_vehicle = params_dict.get('V_cleaning_vehicle', None)
    drainage_cleaning_strategy = params_dict.get('drainage_cleaning_strategy', 'out_of_scope')
    V_cleaning_vehicle_drainage = params_dict.get('V_cleaning_vehicle_drainage', None)
    # Récupération du système d'alimentation
    feeding_system = params_dict.get('feeding_system', 'none')

    measurement_strategy = params_dict.get('measurement_strategy', None)

    OCS1_measurement_strategy = params_dict.get('OCS1_measurement_strategy', 'excluded')
    third_rail_measurement_strategy = params_dict.get('3rd_rail_measurement_strategy', 'excluded')
    V_rolling_stock = params_dict.get('V_rolling_stock', None)  # Vitesse pour on_board
    V_winfram = params_dict.get('V_winfram', None)  # Vitesse pour winfram
    V_manual_tool_OCS1 = params_dict.get('V_manual_tool_OCS1', None)
    measurement_v_inspection = params_dict['measurement_v_inspection']  # vitesse pour measurement_vehicle

    L_total_single_track = params_dict['L_total_single_track']  # Longueur totale de la ligne en km (voie simple)
    depot_total_single_track = float(params_dict.get('depot_total_single_track', 0))
    L_portion = params_dict['L_portion']  # Longueur d'une portion de voie en m
    v_inspection = params_dict['V_foot_inspection']  # km/h
    
    railroad_v_inspection = params_dict['railroad_v_inspection']  # km/h
    V_grinding_Vehicle = params_dict['V_grinding_vehicle']  # km/h
    night_work_hours = params_dict['night_work_hours']  # heures
    day_work_hours = float(params_dict.get('day_work_hours', 0))
    T_hors_tension = params_dict['T_hors_tension']  # minutes
    T_remise_tension = params_dict['T_remise_tension']  # minutes

    # Nombre de diamond_crossing
    diamond_crossing_count = int(params_dict.get('diamond_crossing', 0))
    crossover_count = int(params_dict.get('crossover', 0))
    switch_count = int(params_dict.get('switch', 0))
    switch_depot_count = int(params_dict.get('switch_depot', 0))
    insulated_rail_joint_count = int(params_dict.get('insulated_rail_joint', 0))
    rail_expansion_joint_count = int(params_dict.get('rail_expansion_joint', 0))
    buffer_stop_count = int(params_dict.get('buffer_stop', 0))
    buffer_stop_depot_count = int(params_dict.get('buffer_stop_depot', 0))
    wheel_stop_count = int(params_dict.get('wheel_stop_depot', 0))
    insulated_rail_joint_depot_count = int(params_dict.get('insulated_rail_joint_depot', 0))

    section_insulator_count = int(params_dict.get('OCS1_section_insulator', 0))
    section_insulator_depot_count = int(params_dict.get('OCS1_section_insulator_depot', 0))
    communication_count = int(params_dict.get('OCS1_communication', 0))

    # Lecture des paramètres pour lubrication
    lubrication_strategy = params_dict.get('lubrication_strategy', 'None')
    lubricator_count = int(params_dict.get('lubricator', 0)) if lubrication_strategy == 'On track' else 0
    # Lecture du paramètre fastening device
    distance_fastening_cm = float(params_dict.get('distance_fastening_cm', 0))

    # Calcul de la longueur réelle de la ligne en double voie
    L_total_double_track = L_total_single_track / 2
    L_total_m = L_total_double_track * 1000
    N_portions_par_voie = math.ceil(L_total_m / L_portion)

    # Calcul du temps disponible
    T_hors_tension_h = T_hors_tension / 60
    T_remise_tension_h = T_remise_tension / 60
    D_preparation = T_hors_tension_h + T_remise_tension_h
    D_inspection_disponible = night_work_hours - D_preparation

    L_portion_km = L_portion / 1000

    def generer_zones(methode, vitesse, suffix=''):
        """Génère les zones pour une méthode d'inspection donnée."""
        D_inspection_portion = (L_portion_km / vitesse) * 2
        N_inspectable_nuit = math.floor(D_inspection_disponible / D_inspection_portion)
        N_zones = math.ceil(N_portions_par_voie / N_inspectable_nuit)
        zones_locales = []
        for i in range(N_zones):
            start_portion = i * N_inspectable_nuit
            end_portion = min((i + 1) * N_inspectable_nuit, N_portions_par_voie) - 1

            start_km = (start_portion * L_portion) / 1000
            end_km = ((end_portion + 1) * L_portion) / 1000

            zone_name = f'Zone {i + 1}'
            if suffix:
                zone_name += f' ({suffix})'

            zone_info = {
                'Zone': zone_name,
                'Track': 'Right',
                'Start kilometer point (km)': start_km,
                'End kilometer point (km)': end_km,
                'Number of segments': end_portion - start_portion + 1,
                'Inspection method': methode
            }
            zones_locales.append(zone_info)

            # Voie gauche
            zone_info_gauche = zone_info.copy()
            zone_info_gauche['Track'] = 'Left'
            zones_locales.append(zone_info_gauche)
        return zones_locales

    # Initialisation des zones pour third_rail
    third_rail_zones = []
    OCS1_zones = []
    if feeding_system == 'third_rail':
        V_foot_inspection_third_rail = params_dict['V_foot_inspection_3rd_rail']  # km/h
        V_expansion_joint_third_rail = params_dict['V_expansion_joint_3rd_rail']

        def generer_zones_third_rail(methode, vitesse, suffix=''):
            """Génère les zones pour le système third_rail."""
            D_inspection_portion = (L_portion_km / vitesse) * 2
            N_inspectable_nuit = math.floor(D_inspection_disponible / D_inspection_portion)
            N_zones = math.ceil(N_portions_par_voie / N_inspectable_nuit)
            zones_locales = []
            for i in range(N_zones):
                start_portion = i * N_inspectable_nuit
                end_portion = min((i + 1) * N_inspectable_nuit, N_portions_par_voie) - 1

                start_km = (start_portion * L_portion) / 1000
                end_km = ((end_portion + 1) * L_portion) / 1000

                zone_name = f'Zone {i + 1}'
                if suffix:
                    zone_name += f' ({suffix})'

                zone_info = {
                    'Zone': zone_name,
                    'Track': 'Right',
                    'Start kilometer point (km)': start_km,
                    'End kilometer point (km)': end_km,
                    'Number of segments': end_portion - start_portion + 1,
                    'Inspection method': methode
                }
                zones_locales.append(zone_info)

                zone_info_gauche = zone_info.copy()
                zone_info_gauche['Track'] = 'Left'
                zones_locales.append(zone_info_gauche)
            return zones_locales
        third_rail_zones = generer_zones_third_rail('On feet', V_foot_inspection_third_rail, suffix='Conductor')
        third_rail_expansion_joint_zones = generer_zones_third_rail('Maintenance expansion joint', V_expansion_joint_third_rail, suffix='Conductor_EJ')
        third_rail_zones += third_rail_expansion_joint_zones

    if feeding_system == 'OCS1':
        V_foot_inspection_OCS1 = float(params_dict['V_foot_inspection_OCS1'])
        V_tensionning_device_OCS1 = float(params_dict['V_tensionning_device_OCS1'])  # par exemple
        V_detailed_inspection_OCS1 = float(params_dict['V_detailed_inspection_OCS1'])
        def generer_zones_OCS1(methode, vitesse, suffix=''):
            """Génère les zones pour le système OCS1."""
            D_inspection_portion = (L_portion_km / vitesse) * 2
            N_inspectable_nuit = math.floor(D_inspection_disponible / D_inspection_portion)
            N_zones = math.ceil(N_portions_par_voie / N_inspectable_nuit)
            zones_locales = []
            for i in range(N_zones):
                start_portion = i * N_inspectable_nuit
                end_portion = min((i + 1) * N_inspectable_nuit, N_portions_par_voie) - 1

                start_km = (start_portion * L_portion) / 1000
                end_km = ((end_portion + 1) * L_portion) / 1000

                zone_name = f'Zone {i + 1}'
                if suffix:
                    zone_name += f' ({suffix})'

                zone_info = {
                    'Zone': zone_name,
                    'Track': 'Right',
                    'Start kilometer point (km)': start_km,
                    'End kilometer point (km)': end_km,
                    'Number of segments': end_portion - start_portion + 1,
                    'Inspection method': methode
                }
                zones_locales.append(zone_info)

                zone_info_gauche = zone_info.copy()
                zone_info_gauche['Track'] = 'Left'
                zones_locales.append(zone_info_gauche)
            return zones_locales

        # Générer des zones pour OCS1 foot inspection
        OCS1_zones_foot = generer_zones_OCS1('On feet OCS1', V_foot_inspection_OCS1, suffix='OCS1_foot')
        
        # Générer des zones pour tensionning_device OCS1 (exemple)
        OCS1_zones_tension_device = generer_zones_OCS1('Tensionning device OCS1', V_tensionning_device_OCS1, suffix='OCS1_TD')
        OCS1_zones_feeding_pillar = generer_zones_OCS1('Earthing test OCS1', V_manual_tool_OCS1, suffix='OCS1_earthing')
        # Générer les zones OCS1 de detailed inspection
        OCS1_zones_detailed_inspection = generer_zones_OCS1(
            methode="Detailed inspection OCS1",
            vitesse=V_detailed_inspection_OCS1,
            suffix="OCS1_DI"
        )

        # Concaténer toutes les zones OCS1
        OCS1_zones = OCS1_zones_foot + OCS1_zones_tension_device + OCS1_zones_feeding_pillar + OCS1_zones_detailed_inspection

    zones = []
    # Zones inspection à pied
    zones += generer_zones('On feet', v_inspection)
    # Zones inspection en véhicule
    zones += generer_zones('By railroad', railroad_v_inspection, suffix='Railroad')
    # Zones grinding
    zones += generer_zones('By grinding Vehicle', V_grinding_Vehicle, suffix='Grinding')

    # Zones tamping si voie sur ballast
    if type_track_installation == 'ballast' and V_ballast_tamping:
        zones += generer_zones('Ballast_tamping', V_ballast_tamping, suffix='Tamping')

    # Zones cleaning si cleaning_strategy != out_of_scope
    if cleaning_strategy != 'out_of_scope' and V_cleaning_vehicle:
        zones += generer_zones('Cleaning', V_cleaning_vehicle, suffix='Cleaning')

    # Zones cleaning drainage si drainage_cleaning_strategy != out_of_scope
    if drainage_cleaning_strategy != 'out_of_scope' and V_cleaning_vehicle_drainage:
        zones += generer_zones('Cleaning_drainage', V_cleaning_vehicle_drainage, suffix='drainage')

    if measurement_strategy == 'on_board' and V_rolling_stock:
        zones += generer_zones('Measurement_on_board', V_rolling_stock, suffix='Measurement_on_board')
    elif measurement_strategy == 'measurement_vehicle':
        zones += generer_zones('Measurement_vehicle', measurement_v_inspection, suffix='Measurement_vehicle')
    elif measurement_strategy == 'winfram' and V_winfram:
        zones += generer_zones('Measurement_winfram', V_winfram, suffix='Measurement_winfram')

    if feeding_system == 'third_rail':
        if third_rail_measurement_strategy != 'excluded':
            if third_rail_measurement_strategy == 'on_board' and V_rolling_stock:
                third_rail_zones += generer_zones_third_rail('Measurement_on_board', V_rolling_stock, suffix='Conductor_on_board')
            elif third_rail_measurement_strategy == 'measurement_vehicle' and measurement_v_inspection:
                third_rail_zones += generer_zones_third_rail('Measurement_vehicle', measurement_v_inspection, suffix='Conductor_vehicle')
            elif third_rail_measurement_strategy == 'winfram' and V_winfram:
                third_rail_zones += generer_zones_third_rail('Measurement_winfram', V_winfram, suffix='Conductor_winfram')

    if feeding_system == 'OCS1':  
        if OCS1_measurement_strategy == 'on_board' and V_rolling_stock:
            OCS1_zones += generer_zones_OCS1('Measurement_on_board', V_rolling_stock, suffix='OCS1_on_board')
        elif OCS1_measurement_strategy == 'measurement_vehicle' and measurement_v_inspection:
            OCS1_zones += generer_zones_OCS1('Measurement_vehicle', measurement_v_inspection, suffix='OCS1_vehicle')
        elif OCS1_measurement_strategy == 'winfram' and V_winfram:
            OCS1_zones += generer_zones_OCS1('Measurement_winfram', V_winfram, suffix='OCS1_winfram')
        elif OCS1_measurement_strategy == 'manual_tool' and V_manual_tool_OCS1:
            OCS1_zones += generer_zones_OCS1('Manual_measurement', V_manual_tool_OCS1, suffix='OCS1_manual')

    # Création des zones fixes (diamond_crossing, crossover, switch)
    def generer_zones_fixes(count, prefix):
        """Génère des zones fixes nommées prefix_1, prefix_2, ... 
           Avec deux lignes (droite et gauche) sans km (None)."""
        fixed_zones = []
        for i in range(1, count + 1):
            zone_name = f"{prefix}_{i}"
            zone_droite = {
                'Zone': zone_name,
                'Track': 'Right',
                'Start kilometer point (km)': None,
                'End kilometer point (km)': None,
            }
            zone_gauche = {
                'Zone': zone_name,
                'Track': 'Left',
                'Start kilometer point (km)': None,
                'End kilometer point (km)': None,
            }
            fixed_zones.append(zone_droite)
            fixed_zones.append(zone_gauche)
        return fixed_zones

    diamond_crossing_zones = generer_zones_fixes(diamond_crossing_count, "diamond_crossing")
    crossover_zones = generer_zones_fixes(crossover_count, "crossover")
    switch_zones = generer_zones_fixes(switch_count, "switch")

    # Pour switch_depot, une seule zone "switch_depot" mais deux lignes (Droite et Gauche).
    # Le paramètre switch_depot_count indique le nombre de Turnout.
    switch_depot_zones = []
    if switch_depot_count > 0:
        zone_name = "switch_depot"
        zone_droite = {
            'Zone': zone_name,
            'Track': 'Right',
            'Start kilometer point (km)': None,
            'End kilometer point (km)': None,
        }
        zone_gauche = {
            'Zone': zone_name,
            'Track': 'Left',
            'Start kilometer point (km)': None,
            'End kilometer point (km)': None,
        }
        switch_depot_zones.append(zone_droite)
        switch_depot_zones.append(zone_gauche)

    # IRJ (Insulated Rail Joint)
    # On divise insulated_rail_joint_count par 4 pour obtenir le nombre de zones
    IRJ_zone_count = insulated_rail_joint_count // 2 if insulated_rail_joint_count >= 2 else 0
    IRJ_zones = generer_zones_fixes(IRJ_zone_count, "IRJ")

    # REJ (Rail Expansion Joint)
    # On divise rail_expansion_joint_count par 2 pour obtenir le nombre de zones
    REJ_zone_count = rail_expansion_joint_count // 2 if rail_expansion_joint_count >= 2 else 0
    REJ_zones = generer_zones_fixes(REJ_zone_count, "REJ")

    # buffer_stop : diviser par 2 pour obtenir le nombre de zones buffer_stop
    buffer_stop_zone_count = buffer_stop_count // 2 if buffer_stop_count >= 2 else 0
    buffer_stop_zones = generer_zones_fixes(buffer_stop_zone_count, "buffer_stop")

    # Génération des zones lubrication si stratégie est "On track"
    zone_lubrication_zones = []
    if lubrication_strategy == 'On track' and lubricator_count > 0:
        zone_lubrication_zones = generer_zones_fixes(lubricator_count, "zone_lubrication")

    # Génération des zones communication et section insulator si "OCS1"
    communication_zones = []
    section_insulator_zones = []
    if feeding_system == 'OCS1' :
        if communication_count > 0 :
            communication_zones = generer_zones_fixes(communication_count, "OCS1_communication")
        if section_insulator_count > 0:
            section_insulator_zones = generer_zones_fixes(section_insulator_count, "OCS1_section_insulator")

    # buffer_stop_depot : une seule zone "depot" si >0
    depot_zones = []
    if buffer_stop_depot_count > 0 or insulated_rail_joint_depot_count > 0 or switch_depot_count > 0 or depot_track_installation or section_insulator_depot_count > 0:
        zone_name = "depot"
        zone_droite = {
            'Zone': zone_name,
            'Track': 'Right',
            'Start kilometer point (km)': None,
            'End kilometer point (km)': None,
        }
        zone_gauche = {
            'Zone': zone_name,
            'Track': 'Left',
            'Start kilometer point (km)': None,
            'End kilometer point (km)': None,
        }
        depot_zones.append(zone_droite)
        depot_zones.append(zone_gauche)

    # définition des zones d'inspections Track au dépot
    track_depot_foot_quantity = 0
    if depot_total_single_track > 0 and day_work_hours > 0 and v_inspection > 0:
        distance_per_day_km = v_inspection * day_work_hours
        track_depot_foot_quantity = math.ceil(depot_total_single_track / distance_per_day_km)
    else:
        track_depot_foot_quantity = 0

    # définition des zones d'inspections OCS1 au dépot
    OCS1_depot_detailed_quantity = 0
    OCS1_depot_foot_quantity = 0
    if feeding_system == 'OCS1':
        if depot_total_single_track > 0 and day_work_hours > 0 and night_work_hours > 0 and V_detailed_inspection_OCS1 > 0 and V_foot_inspection_OCS1 > 0:
            distance_per_day_km_detailed_OCS1 = V_detailed_inspection_OCS1 * night_work_hours
            distance_per_day_km_OCS1 = V_foot_inspection_OCS1 * day_work_hours
            OCS1_depot_detailed_quantity = math.ceil(depot_total_single_track / distance_per_day_km_detailed_OCS1)
            OCS1_depot_foot_quantity = math.ceil(depot_total_single_track / distance_per_day_km_OCS1)
        else:
            OCS1_depot_detailed_quantity = 0
            OCS1_depot_foot_quantity = 0

    # Calcul du nombre de zones pour le tamping ballast au dépôt
    track_depot_ballast_quantity = 0
    if depot_track_installation == 'ballast' and V_ballast_tamping > 0:
        distance_per_night_km = V_ballast_tamping * night_work_hours
        if depot_total_single_track > 0 and distance_per_night_km > 0:
            track_depot_ballast_quantity = math.ceil(depot_total_single_track / distance_per_night_km)
        else:
            track_depot_ballast_quantity = 0

    # On retourne toutes ces infos
    return (
        zones,
        type_track_installation, depot_track_installation,
        cleaning_strategy, drainage_cleaning_strategy,
        measurement_strategy, lubrication_strategy,
        diamond_crossing_zones, crossover_zones, switch_zones, switch_depot_zones, communication_zones,
        section_insulator_zones, switch_depot_count, section_insulator_count, communication_count, section_insulator_depot_count,
        IRJ_zones, IRJ_zone_count, REJ_zones, REJ_zone_count,
        buffer_stop_zones, buffer_stop_zone_count,
        depot_zones, buffer_stop_depot_count, wheel_stop_count, insulated_rail_joint_depot_count,
        zone_lubrication_zones, lubricator_count, third_rail_zones, OCS1_zones, distance_fastening_cm,
        track_depot_foot_quantity, OCS1_depot_foot_quantity, OCS1_depot_detailed_quantity, track_depot_ballast_quantity, D_inspection_disponible)

def construire_track_df(zones, type_track_installation, cleaning_strategy, drainage_cleaning_strategy,
                        measurement_strategy,lubrication_strategy,
                        diamond_crossing_zones, crossover_zones, switch_zones, switch_depot_zones, switch_depot_count,
                        IRJ_zones, IRJ_zone_count, REJ_zones, REJ_zone_count,
                        buffer_stop_zones, buffer_stop_zone_count,
                        depot_zones, buffer_stop_depot_count, wheel_stop_count, insulated_rail_joint_depot_count,
                        zone_lubrication_zones, lubricator_count, distance_fastening_cm,
                        track_depot_foot_quantity, track_depot_ballast_quantity):
    
    """Construit le DataFrame 'Track' en assignant chaque zone aux équipements concernés."""

    equipments = ['Track_foot', 'Track_railroad', 'Track_grinding']

    if type_track_installation == 'ballast':
        equipments.append('Track_ballast')

    if cleaning_strategy != 'out_of_scope':
        equipments.append('Track_cleaning')
    if drainage_cleaning_strategy != 'out_of_scope':
        equipments.append('Drainage_cleaning')
    # Ajouter measurement selon la stratégie
    if measurement_strategy == 'on_board':
        equipments.append('Track_measurement_on_board')
    elif measurement_strategy == 'measurement_vehicle':
        equipments.append('Track_measurement_vehicle')
    elif measurement_strategy == 'winfram':
        equipments.append('Track_measurement_winfram')

    # Ajout de l'équipement Turnout et Trnout_depot
    equipments.append('Turnout')
    equipments.append('Turnout_depot')

    # Insulated_rail_joint
    if IRJ_zone_count > 0:
        equipments.append('Insulated_rail_joint')
    if insulated_rail_joint_depot_count > 0:
        equipments.append('Insulated_rail_joint_depot')
    # Rail_expansion_joint
    if REJ_zone_count > 0:
        equipments.append('Rail_expansion_joint')
        
    if buffer_stop_zone_count > 0:
        equipments.append('buffer_stop')
    if buffer_stop_depot_count > 0:
        equipments.append('buffer_stop_depot')

    if wheel_stop_count > 0:   
        equipments.append('wheel_stop')

    # Ajout de lubricator si stratégie "On track"
    if lubrication_strategy == 'On track' and lubricator_count > 0:
        equipments.append('lubricator')
    
    # Ajout du fastening_device si distance_fastening_cm > 0
    if distance_fastening_cm > 0:
        equipments.append('fastening_device')

    # Ajout de Track_depot_foot et Track_depot_ballast
    if track_depot_foot_quantity > 0:
        equipments.append('Track_depot_foot')
    if track_depot_ballast_quantity > 0:
        equipments.append('Track_depot_ballast')

    track_data = {'Equipment': equipments}

    # les zones selon leur méthode d'inspection
    zones_categories = {
        'zones_foot': 'On feet',
        'zones_vehicle': 'By railroad',
        'zones_grinding': 'By grinding Vehicle',
        'zones_tamping': 'Ballast_tamping',
        'zones_cleaning': 'Cleaning',
        'zones_cleaning_drainage': 'Cleaning_drainage',
        'zones_measurement_on_board': 'Measurement_on_board',
        'zones_measurement_vehicle': 'Measurement_vehicle',
        'zones_measurement_winfram': 'Measurement_winfram'
    }
    zones_dict = {}
    for key, method in zones_categories.items():
        zones_dict[key] = [z['Zone'] for z in zones if z.get('Inspection method') == method]

    # Extraire les noms des zones fixes
    diamond_crossing_names = list({z['Zone'] for z in diamond_crossing_zones})
    crossover_names = list({z['Zone'] for z in crossover_zones})
    switch_names = list({z['Zone'] for z in switch_zones})
    switch_depot_names = list({z['Zone'] for z in switch_depot_zones})
    IRJ_names = list({z['Zone'] for z in IRJ_zones})
    REJ_names = list({z['Zone'] for z in REJ_zones})
    buffer_stop_names = list({z['Zone'] for z in buffer_stop_zones})
    depot_names = list({z['Zone'] for z in depot_zones})
    zone_lubrication_names = list({z['Zone'] for z in zone_lubrication_zones})

    all_zones = (zones_dict['zones_foot'] + zones_dict['zones_vehicle'] + zones_dict['zones_grinding'] +
                 zones_dict['zones_tamping'] + zones_dict['zones_cleaning'] + zones_dict['zones_cleaning_drainage'] +
                 zones_dict['zones_measurement_on_board'] + zones_dict['zones_measurement_vehicle'] + zones_dict['zones_measurement_winfram'] +
                 diamond_crossing_names + crossover_names + switch_names + switch_depot_names +
                 IRJ_names + REJ_names + buffer_stop_names + depot_names + zone_lubrication_names)

    # Créer un dictionnaire pour acceder aux points kilométriques par zone
    point_km_dict = {zone['Zone']: (zone.get('Start kilometer point (km)', 0), zone.get('End kilometer point (km)', 0)) for zone in zones}

    for zone in all_zones:
        # Par défaut, toutes les valeurs à 0
        values = [0]*len(equipments)

        # Affecter la bonne valeur à 1 selon la méthode
        if zone in zones_dict['zones_foot']:
            for i, eq in enumerate(equipments):
                if eq == "Track_foot":
                    values[i] = 1
        elif zone in zones_dict['zones_vehicle']:
            for i, eq in enumerate(equipments):
                if eq == "Track_railroad":
                    values[i] = 1
        elif zone in zones_dict['zones_grinding']:
            for i, eq in enumerate(equipments):
                if eq == "Track_grinding":
                    values[i] = 1
        elif zone in zones_dict['zones_tamping']:
            for i, eq in enumerate(equipments):
                if eq == "Track_ballast":
                    values[i] = 1
        elif zone in zones_dict['zones_cleaning']:
            for i, eq in enumerate(equipments):
                if eq == "Track_cleaning":
                    values[i] = 1
        elif zone in zones_dict['zones_cleaning_drainage']:
            for i, eq in enumerate(equipments):
                if eq == "Drainage_cleaning":
                    values[i] = 1
        elif zone in zones_dict['zones_measurement_on_board']:
            for i, eq in enumerate(equipments):
                if eq == "Track_measurement_on_board":
                    values[i] = 1
        elif zone in zones_dict['zones_measurement_vehicle']:
            for i, eq in enumerate(equipments):
                if eq == "Track_measurement_vehicle":
                    values[i] = 1
        elif zone in zones_dict['zones_measurement_winfram']:
            for i, eq in enumerate(equipments):
                if eq == "Track_measurement_winfram":
                    values[i] = 1

        # Pour les diamond_crossing, crossover, switch on place 4, 2, 1 sur Turnout
        if zone in diamond_crossing_names:
            for i, eq in enumerate(equipments):
                if eq == "Turnout":
                    values[i] = 4
        elif zone in crossover_names:
            for i, eq in enumerate(equipments):
                if eq == "Turnout":
                    values[i] = 2
        elif zone in switch_names:
            for i, eq in enumerate(equipments):
                if eq == "Turnout":
                    values[i] = 1
        # switch_depot = switch_depot_count Turnout
        elif zone in switch_depot_names:
            for i, eq in enumerate(equipments):
                if eq == "Turnout_depot":
                    values[i] = switch_depot_count
        # IRJ : chaque zone IRJ a 4 Insulated_rail_joint
        if zone in IRJ_names:
            # On ne définit pas de Turnout, mais 4 Insulated_rail_joint
            for i, eq in enumerate(equipments):
                if eq == "Insulated_rail_joint":
                    values[i] = 2

        # REJ : chaque zone REJ a 2 Rail_expansion_joint
        if zone in REJ_names:
            for i, eq in enumerate(equipments):
                if eq == "Rail_expansion_joint":
                    values[i] = 2
        # buffer_stop_zones : 2 buffer_stop par zone
        if zone in buffer_stop_names:
            values[equipments.index('buffer_stop')] = 2

        # depot zone : buffer_stop_depot_count buffer_stop
        if zone in depot_names and buffer_stop_depot_count > 0:
            values[equipments.index('buffer_stop_depot')] = buffer_stop_depot_count
        
        # Assigner insulated_rail_joint_depot_count à la zone "depot"
        if zone in depot_names and insulated_rail_joint_depot_count > 0:
            values[equipments.index('Insulated_rail_joint_depot')] = insulated_rail_joint_depot_count
        
         # Assigner wheel_stop_count à la zone "depot"
        if zone in depot_names and wheel_stop_count > 0:
            values[equipments.index('wheel_stop')] = wheel_stop_count

         # Assigner lubricator_count à la zone "zone_lubrication"
        if zone in zone_lubrication_names and lubricator_count > 0:
            values[equipments.index('lubricator')] = 1

        # Calcul du nombre de fastening Device pour les zones pertinentes
        if distance_fastening_cm > 0 and zone in point_km_dict:
            point_debut_km, point_fin_km = point_km_dict[zone]
            # Convertir la distance en cm
            longueur_voie_cm = (point_fin_km - point_debut_km) * 100000
            if longueur_voie_cm > 0:
                nombre_fastening = math.ceil(longueur_voie_cm / distance_fastening_cm)
                if (zone in zones_dict['zones_foot']) and (measurement_strategy in ['on_board', 'measurement_vehicle', 'winfram']):
                    if 'fastening_device' in equipments:
                        values[equipments.index('fastening_device')] = nombre_fastening
                elif zone in zones_dict['zones_measurement_on_board'] and measurement_strategy == 'on_board':
                    if 'fastening_device' in equipments:
                        values[equipments.index('fastening_device')] = nombre_fastening
                elif zone in zones_dict['zones_measurement_vehicle'] and measurement_strategy == 'measurement_vehicle':
                    if 'fastening_device' in equipments:
                        values[equipments.index('fastening_device')] = nombre_fastening
                elif zone in zones_dict['zones_measurement_winfram'] and measurement_strategy == 'winfram':
                    if 'fastening_device' in equipments:
                        values[equipments.index('fastening_device')] = nombre_fastening
   
        track_data[zone] = values

    if track_depot_foot_quantity > 0 and 'depot' in all_zones:
        
        depot_foot_index = equipments.index('Track_depot_foot')
        track_data['depot'][depot_foot_index] = track_depot_foot_quantity

    # Vérification et calcul pour la zone "depot"
    if track_depot_ballast_quantity > 0 and 'depot' in all_zones:
        depot_index = equipments.index('Track_depot_ballast')
        # Affecter la quantité calculée à la zone "depot"
        track_data['depot'][depot_index] = track_depot_ballast_quantity

    track_df = pd.DataFrame(track_data)
    return track_df

def show_message(title, message):
    """Affiche une boîte de message compatible Windows / macOS / Linux."""
    try:
        messagebox.showinfo(title, message)
    except Exception:
        print(f"{title}: {message}")

import argparse  # si pas déjà importé

def read_and_strip_all_text_cells(file_path: str, sheet_name: str, engine="openpyxl") -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
    except ValueError:
        print(f"La feuille '{sheet_name}' est inexistante ou invalide.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Lecture de '{sheet_name}' échouée ({e}). DataFrame vide utilisé.")
        return pd.DataFrame()

    if df is None or df.shape[1] == 0:
        return pd.DataFrame()

    # Strip des noms de colonnes
    df.columns = df.columns.map(lambda c: str(c).strip())

    # Strip des cellules texte pour toutes les colonnes
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)

    # Feuille visuellement vide
    if df.dropna(how="all").empty:
        return pd.DataFrame()

    return df

def subsystem_from_location(loc: str) -> str:
    s = str(loc)
    if "Conductor" in s:
        return "3rd_Rail"
    if ("OCS1_" in s) or ("ROCS_" in s):
        return "CAT"
    return "Track"

def build_equipements_premises(*, track_df, third_rail_df=None, OCS1_df=None):
    dfs = []
    matrices = [("Track", track_df)]

    if third_rail_df is not None:
        matrices.append(("3rd_Rail", third_rail_df))
    if OCS1_df is not None:
        matrices.append(("CAT", OCS1_df))

    for subsystem, mdf in matrices:
        if mdf is None or mdf.empty:
            continue

        tmp = mdf.copy()
        if "Equipment" not in tmp.columns:
            continue

        long = tmp.melt(id_vars=["Equipment"], var_name="Location", value_name="Quantity")

        # convertir quantity en numérique, ignorer le reste
        long["Quantity"] = pd.to_numeric(long["Quantity"], errors="coerce")
        long = long[long["Quantity"].fillna(0) > 0]

        long.insert(0, "subsystem", subsystem)
        long = long.rename(columns={"Equipment": "Equipment"})
        long = long[["subsystem", "Location", "Equipment", "Quantity"]]
        dfs.append(long)

    if dfs:
        return pd.concat(dfs, ignore_index=True)

    return pd.DataFrame(columns=["subsystem", "Location", "Equipment", "Quantity"])

def stack_sheets_in_order(dfs_in_order):
    # dfs_in_order = [(name, df), ...]
    kept = []
    missing = []
    for name, df in dfs_in_order:
        if df is None or df.empty:
            missing.append(name)
        else:
            kept.append(df)
    out = pd.concat(kept, ignore_index=True) if kept else pd.DataFrame()
    return out, missing

def core_process(input_file=None, output_file=None, *, silent=False):
    # Alias locaux pour éviter tout conflit avec 'tk' du module
    import tkinter as _tk
    from tkinter import filedialog as _fd

    # Créer un root temporaire UNIQUEMENT si on a besoin de boîtes de dialogue
    temp_root = None
    needs_dialog = (input_file is None or output_file is None)
    if needs_dialog and getattr(_tk, "_default_root", None) is None:
        temp_root = _tk.Tk()
        temp_root.withdraw()

    # Wrappers : si un chemin est fourni par l’UI, on l’utilise; sinon on ouvre une boîte
    def filedialog_askopen(**kwargs):
        if input_file is not None:
            return input_file
        return _fd.askopenfilename(**kwargs)

    def filedialog_asksave(**kwargs):
        if output_file is not None:
            return output_file
        return _fd.asksaveasfilename(**kwargs)

    try: 
        # Sélection du fichier d'entrée
        
        # Message interactif avant de choisir le fichier excel d'entrée
        if not silent:
            show_message(
                "Input file selection",
                "Please select the Excel file 'Project_name_Linear_assets_preliminary_inputs_with_ovh_renew_per_subsystem_Vref'.")
        print("Please select the input Excel file to run the linear_asset_planning...")
        
        file_path = filedialog_askopen(
            title="Select the input file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        
        if not file_path:
            print("No file selected. Program stopped.")
            return

        # Lecture des paramètres
        try:
            params_dict = lire_parametres(file_path)
        except Exception as e:
            print(f"Error while importing general parameters : {e}")
            return

        feeding_system = params_dict.get('feeding_system', 'none')
        third_rail_measurement_strategy = params_dict.get('3rd_rail_measurement_strategy', 'excluded')

        # Calcul des zones
        (zones,
        type_track_installation, depot_track_installation,
        cleaning_strategy, drainage_cleaning_strategy,
        measurement_strategy, lubrication_strategy,
        diamond_crossing_zones, crossover_zones, switch_zones, switch_depot_zones,
        communication_zones, section_insulator_zones,   
        switch_depot_count, section_insulator_count, communication_count, section_insulator_depot_count,
        IRJ_zones, IRJ_zone_count, REJ_zones, REJ_zone_count,
        buffer_stop_zones, buffer_stop_zone_count,
        depot_zones, buffer_stop_depot_count, wheel_stop_count, insulated_rail_joint_depot_count,
        zone_lubrication_zones, lubricator_count, third_rail_zones, OCS1_zones, distance_fastening_cm,
        track_depot_foot_quantity, OCS1_depot_foot_quantity, OCS1_depot_detailed_quantity,
        track_depot_ballast_quantity, D_inspection_disponible) = calcul_zones(params_dict)
        
        zones_combined = zones.copy()
        if feeding_system == 'third_rail':
            zones_combined += third_rail_zones
        if feeding_system == 'OCS1':
            zones_combined += OCS1_zones
        zones_df = pd.DataFrame(zones_combined)

        # On fusionne toutes les zones fixes dans un seul DataFrame Fixed_Zones
        # On les concatène verticalement
        fixed_zones_all = (diamond_crossing_zones + crossover_zones + switch_zones + 
                        switch_depot_zones + IRJ_zones + REJ_zones +
                        section_insulator_zones + communication_zones +
                        buffer_stop_zones + depot_zones + zone_lubrication_zones)
        fixed_zones_df = pd.DataFrame(fixed_zones_all)

        # Création du DataFrame Track_maintenance_activities vide
        track_maintenance_activities = pd.DataFrame(columns=[
            'subsystem', 'equipment', 'description', 'time_of_day', 
            'periodicity_weeks', 'duration_hours', 'outage_time_before', 
            'outage_time_after', 'technicians_needed'
        ])

        # Construction du DataFrame 'Track'
        track_df = construire_track_df(
            zones, type_track_installation, cleaning_strategy, drainage_cleaning_strategy, measurement_strategy, lubrication_strategy,
            diamond_crossing_zones, crossover_zones, switch_zones, switch_depot_zones, switch_depot_count,
            IRJ_zones, IRJ_zone_count, REJ_zones, REJ_zone_count,
            buffer_stop_zones, buffer_stop_zone_count,
            depot_zones, buffer_stop_depot_count, wheel_stop_count, insulated_rail_joint_depot_count,
            zone_lubrication_zones, lubricator_count, distance_fastening_cm,
            track_depot_foot_quantity, track_depot_ballast_quantity
        )

        # Vérifier si 'fastening_device' est dans les équipements
        if 'fastening_device' in track_df['Equipment'].values:
            # Extraire la ligne 'fastening_device' et transposer pour avoir les Zones comme index
            fastening_series = track_df.set_index('Equipment').loc['fastening_device']
            # Convertir en DataFrame
            fastening_df = fastening_series.reset_index()
            fastening_df.columns = ['Zone', 'Number of fastening devices']
            # Vérifier les duplications dans fastening_df (il ne devrait y en avoir aucune)
            duplicated_zones = fastening_df['Zone'][fastening_df['Zone'].duplicated()]
            if not duplicated_zones.empty:
                print("Attention : Zones dupliquées dans 'fastening_df'. Agrégation nécessaire.")
                # Agréger les duplications par somme
                fastening_df = fastening_df.groupby('Zone', as_index=False).sum()
        else:
            # Si 'fastening_device' n'est pas présent, créer un DataFrame avec 0
            fastening_df = pd.DataFrame({'Zone': zones_df['Zone'].unique(), 'Number of fastening devices': 0})
        # **Étape 3 : Fusionner zones_df avec fastening_df**
        # Effectuer une fusion sur la colonne 'Zone'
        zones_df = zones_df.merge(fastening_df, on='Zone', how='left')
        # Remplacer les NaN par 0 et convertir en entier
        zones_df['Number of fastening devices'] = zones_df['Number of fastening devices'].fillna(0).astype(int)

        # Vérifier si le système d'alimentation est 'third_rail' pour créer la feuille '3rd_Rail'
        if feeding_system == 'third_rail':
            
            # Créer un DataFrame pour '3rd_Rail' avec l'équipement '3rd_rail_foot'
            third_rail_equipments = ['3rd_rail_foot', 'third_rail_expansion_joint']
            if third_rail_measurement_strategy != 'excluded':
                third_rail_equipments.insert(1, 'third_rail_measurement')
            third_rail_data = {'Equipment': third_rail_equipments}
            
            for zone_info in third_rail_zones:
                zone_name = zone_info['Zone']
                inspection_method = zone_info['Inspection method']
                if zone_name not in third_rail_data:
                    third_rail_data[zone_name] = [0] * len(third_rail_equipments)
                
                # raccourci vers la liste
                values = third_rail_data[zone_name]

                if inspection_method == 'On feet':
                    values[third_rail_equipments.index('3rd_rail_foot')] = 1

                elif inspection_method == 'Maintenance expansion joint':
                    values[third_rail_equipments.index('third_rail_expansion_joint')] = 1

                else:                              # cas « Measurement »
                    if 'third_rail_measurement' in third_rail_equipments:
                        values[third_rail_equipments.index('third_rail_measurement')] = 1
            
            third_rail_df = pd.DataFrame(third_rail_data)

        if feeding_system == 'OCS1':
            # 1) Définir la liste des équipements OCS1 de base
            # (pour les zones OCS1 générées dynamiquement)
            OCS1_equipments = [
                'OCS1_foot',               # On feet OCS1
                'OCS1_measurement',        # Measurement_on_board, Measurement_vehicle, Measurement_winfram, Manual_measurement
                'OCS1_detailed',           # Detailed inspection OCS1
                'OCS1_tensionning_device', # Tensionning device OCS1
                'OCS1_feeding_pillar',
            ]
            if OCS1_depot_detailed_quantity > 0:
                OCS1_equipments.append('OCS1_depot_detailed')
            if OCS1_depot_foot_quantity > 0:
                OCS1_equipments.append('OCS1_depot_foot')
            if section_insulator_depot_count > 0:
                OCS1_equipments.append('depot_section_insulator')
            if communication_count > 0:
                OCS1_equipments.append('communication')
            if section_insulator_count > 0:
                OCS1_equipments.append('section_insulator')
                
            # 3) Initialiser le dictionnaire de DataFrame
            OCS1_data = {'Equipment': OCS1_equipments}
        
            # 4) Pour chaque zone OCS1, créer ou compléter la colonne zone_name
            for zone_info in OCS1_zones:
                zone_name = zone_info['Zone']
                inspection_method = zone_info['Inspection method']
                # S'assurer que la colonne zone_name existe
                if zone_name not in OCS1_data:
                    # on crée un tableau de 0 (len(OCS1_equipments))
                    OCS1_data[zone_name] = [0]*len(OCS1_equipments)
                # On récupère la liste de valeurs associée à la colonne zone_name
                values = OCS1_data[zone_name]
                # Affectation selon la méthode d'inspection
                if inspection_method == 'On feet OCS1':
                    idx_foot = OCS1_equipments.index('OCS1_foot')
                    values[idx_foot] = 1
                elif inspection_method in ['Measurement_on_board', 'Measurement_vehicle', 'Measurement_winfram', 'Manual_measurement']:
                    idx_meas = OCS1_equipments.index('OCS1_measurement')
                    values[idx_meas] = 1
                elif inspection_method == 'Detailed inspection OCS1':
                    idx_detailed = OCS1_equipments.index('OCS1_detailed')
                    values[idx_detailed] = 1
                elif inspection_method == 'Tensionning device OCS1':
                    idx_tension = OCS1_equipments.index('OCS1_tensionning_device')
                    values[idx_tension] = 1
                elif inspection_method == 'Earthing test OCS1':
                    idx_feeding = OCS1_equipments.index('OCS1_feeding_pillar')
                    values[idx_feeding] = 1
                # réassigner
                OCS1_data[zone_name] = values

            # 5) Pour les zones fixes et quantités de dépôt
            if 'depot' not in OCS1_data:
                OCS1_data['depot'] = [0]*len(OCS1_equipments)
            if OCS1_depot_detailed_quantity > 0:
                idx_depot_detailed = OCS1_equipments.index('OCS1_depot_detailed')
                OCS1_data['depot'][idx_depot_detailed] = OCS1_depot_detailed_quantity
            if OCS1_depot_foot_quantity > 0:
                idx_depot_foot = OCS1_equipments.index('OCS1_depot_foot')
                OCS1_data['depot'][idx_depot_foot] = OCS1_depot_foot_quantity

            if section_insulator_depot_count > 0 and 'depot_section_insulator' in OCS1_equipments:
                idx_section_depot = OCS1_equipments.index('depot_section_insulator')
                OCS1_data['depot'][idx_section_depot] = section_insulator_depot_count

            # => On suppose que communication_zones est une liste de zones fixes
            if communication_count > 0 and 'communication' in OCS1_equipments:
                idx_comm = OCS1_equipments.index('communication')
                for z_info in communication_zones:
                    z_name = z_info['Zone']
                    if z_name not in OCS1_data:
                        OCS1_data[z_name] = [0]*len(OCS1_equipments)
                    # Mettre 1
                    OCS1_data[z_name][idx_comm] = 1
            
            # Pareil pour section_insulator_count => section_insulator_zones => index = 'section_insulator'
            if section_insulator_count > 0 and 'section_insulator' in OCS1_equipments:
                idx_section = OCS1_equipments.index('section_insulator')
                for z_info in section_insulator_zones:
                    z_name = z_info['Zone']
                    if z_name not in OCS1_data:
                        OCS1_data[z_name] = [0]*len(OCS1_equipments)
                    OCS1_data[z_name][idx_section] = 1
            
            # 6) Convertir en DataFrame
            OCS1_df = pd.DataFrame(OCS1_data)

        # Zones Inspection -> uniquement Voie = Droite
        pp_zones = zones_df[zones_df["Track"].eq("Right")].copy()
        pp_zones["Location"] = pp_zones["Zone"]
        pp_zones["Position"] = (pp_zones["Start kilometer point (km)"] + pp_zones["End kilometer point (km)"]) / 2
        pp_zones["subsystem"] = pp_zones["Location"].map(subsystem_from_location)
        pp_zones = pp_zones[["subsystem", "Location", "Position"]]

        # Fixed_Zones -> uniquement Voie = Droite, Position vide
        pp_fixed = fixed_zones_df[fixed_zones_df["Track"].eq("Right")].copy()
        pp_fixed["Location"] = pp_fixed["Zone"]
        pp_fixed["Position"] = np.nan
        pp_fixed["subsystem"] = pp_fixed["Location"].map(subsystem_from_location)
        pp_fixed = pp_fixed[["subsystem", "Location", "Position"]]

        premises_positions_df = pd.concat([pp_zones, pp_fixed], ignore_index=True)

        equipements_premises_df = build_equipements_premises(
            track_df=track_df,
            third_rail_df=third_rail_df if feeding_system == "third_rail" else None,
            OCS1_df=OCS1_df if feeding_system == "OCS1" else None,
        )

        grinding_strategy = params_dict.get('grinding_strategy', 'buy')
        ballast_tamping_strategy = params_dict.get('ballast_tamping_strategy', 'buy')
        ultrasonic_inspection_strategy = params_dict.get('ultrasonic_inspection_strategy', 'buy')
        ultrasonic_measurment_strategy = params_dict.get('ultrasonic_measurment_strategy', 'manual_tool')
        day_work_hours = float(params_dict.get('day_work_hours', 0))
        T_hors_tension = params_dict['T_hors_tension']
        T_remise_tension = params_dict['T_remise_tension']
        third_rail_measurement_strategy = params_dict.get('3rd_rail_measurement_strategy', 'excluded')
        OCS1_measurement_strategy = params_dict.get('OCS1_measurement_strategy', 'excluded')

        # track_maintenance_activities est déjà créé, on va ajouter les activités

        activities = []

        # Fonction utilitaire pour déterminer technicians_needed selon buy/make
        def tech_needed(strategy, buy_val=1, make_val=2):
            return make_val if strategy == 'make' else buy_val

        # Fonction utilitaire pour déterminer equipment measurement (Activity 11)
        def measurement_equipment(meas_strat):
            if meas_strat == 'winfram':
                return 'Track_measurement_winfram'
            elif meas_strat == 'measurement_vehicle':
                return 'Track_measurement_vehicle'
            elif meas_strat == 'on_board':
                return 'Track_measurement_on_board'
            else:
                # Si pas défini, par défaut Track_measurement_on_board
                return 'Track_measurement_on_board'

        # Fonction utilitaire pour equipment ultrasonic (Activity 18)
        def ultrasonic_equipment(um_strat):
            if um_strat == 'manual_tool':
                return 'Track_foot'
            elif um_strat == 'measurement_vehicle':
                return 'Track_measurement_vehicle'
            elif um_strat == 'winfram':
                return 'Track_measurement_winfram'
            else:
                return 'Track_foot'  # Par défaut

        # technicians_needed ultrasonic
        def ultrasonic_tech(uis_strat):
            return tech_needed(uis_strat, buy_val=1, make_val=2)

        # Activity 1
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Track_foot',
            'description': 'foot inspection of the track line',
            'time_of_day': 'night',
            'periodicity_weeks': 26,
            'duration_hours': D_inspection_disponible,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 2
        })

        # Activity 2
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Track_railroad',
            'description': 'on board inspection of the track line',
            'time_of_day': 'night',
            'periodicity_weeks': 4,
            'duration_hours': D_inspection_disponible,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 1
        })

        # Activity 3
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Track_depot_foot',
            'description': 'foot inspection of the track depot',
            'time_of_day': 'day',
            'periodicity_weeks': 26,
            'duration_hours': day_work_hours,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 1
        })

        # Activity 4 (si Track_cleaning existe, donc cleaning_strategy != 'out_of_scope')
        if cleaning_strategy != 'out_of_scope':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_cleaning',
                'description': 'cleaning of the track line',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': tech_needed(cleaning_strategy, buy_val=1, make_val=2)
            })

        # Activity 5 (si Drainage_cleaning existe, drainage_cleaning_strategy != 'out_of_scope')
        if drainage_cleaning_strategy != 'out_of_scope':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Drainage_cleaning',
                'description': 'cleaning of the line track drainage',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': tech_needed(drainage_cleaning_strategy, buy_val=1, make_val=2)
            })

        grinding_strategy = params_dict.get('grinding_strategy', 'buy')
        # On suppose Track_grinding est créé si grinding_strategy != 'out_of_scope'
        if grinding_strategy != 'out_of_scope':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_grinding',
                'description': 'rail grinding of the line',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': tech_needed(grinding_strategy, buy_val=1, make_val=2)
            })

        # Activity 7 (si type_track_installation = 'ballast')
        ballast_tamping_strategy = params_dict.get('ballast_tamping_strategy', 'buy')
        if type_track_installation == 'ballast':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_ballast',
                'description': 'ballast_tamping of the line',
                'time_of_day': 'night',
                'periodicity_weeks': 520,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 1 if ballast_tamping_strategy == 'buy' else 2
            })

        # Activity 8 (si depot_track_installation = 'ballast')
        if depot_track_installation == 'ballast':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_depot_ballast',
                'description': 'ballast_tamping of the depot',
                'time_of_day': 'night',
                'periodicity_weeks': 520,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 1 if ballast_tamping_strategy == 'buy' else 2
            })

        # Activity 9 (type_track_installation = 'ballast')
        if type_track_installation == 'ballast':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_foot',
                'description': 'Phytosanitary spreading of the line',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })

        # Activity 10 (depot_track_installation = 'ballast')
        if depot_track_installation == 'ballast':
            activities.append({
                'subsystem': 'Track',
                'equipment': 'Track_depot_foot',
                'description': 'Phytosanitary spreading of the depot',
                'time_of_day': 'day',
                'periodicity_weeks': 52,
                'duration_hours': day_work_hours,
                'outage_time_before': 0,
                'outage_time_after': 0,
                'technicians_needed': 1
            })

        # Activity 11 (measurement_strategy)
        equipment_11 = measurement_equipment(measurement_strategy)
        activities.append({
            'subsystem': 'Track',
            'equipment': equipment_11,
            'description': 'Track measurement of the line',
            'time_of_day': 'night',
            'periodicity_weeks': 26,
            'duration_hours': D_inspection_disponible,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 2
        })

        # Activity 12, 13, 14 (lubrication_strategy = 'On track')
        if lubrication_strategy == 'On track':
            # Activity 12
            activities.append({
                'subsystem': 'Track',
                'equipment': 'lubricator',
                'description': 'Detailed inspection of lubricators',
                'time_of_day': 'night',
                'periodicity_weeks': 13,
                'duration_hours': 0.25,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
            # Activity 13
            activities.append({
                'subsystem': 'Track',
                'equipment': 'lubricator',
                'description': 'Semestrial inspection of lubricators',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': 0.33,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
            # Activity 14
            activities.append({
                'subsystem': 'Track',
                'equipment': 'lubricator',
                'description': 'Annual inspection of lubricators',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': 0.5,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })

        # Activity 15
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Rail_expansion_joint',
            'description': 'Maintenance of rail expansion joint',
            'time_of_day': 'night',
            'periodicity_weeks': 26,
            'duration_hours': 0.75,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 2
        })

        # Activity 16
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Insulated_rail_joint',
            'description': 'Maintenance of line insulated rail joint',
            'time_of_day': 'night',
            'periodicity_weeks': 26,
            'duration_hours': 0.5,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 2
        })

        # Activity 17
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Insulated_rail_joint_depot',
            'description': 'Maintenance of depot insulated rail joint',
            'time_of_day': 'day',
            'periodicity_weeks': 26,
            'duration_hours': 0.5,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 2
        })

        # Activity 18 (ultrasonic)
        equipment_18 = ultrasonic_equipment(ultrasonic_measurment_strategy)
        tech_18 = ultrasonic_tech(ultrasonic_inspection_strategy)
        activities.append({
            'subsystem': 'Track',
            'equipment': equipment_18,
            'description': 'Ultrasonic inspection of the line',
            'time_of_day': 'night',
            'periodicity_weeks': 26,
            'duration_hours': D_inspection_disponible,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': tech_18
        })

        # Activity 19
        tech_19 = ultrasonic_tech(ultrasonic_inspection_strategy)
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Track_depot_foot',
            'description': 'Ultrasonic inspection of the depot',
            'time_of_day': 'day',
            'periodicity_weeks': 52,
            'duration_hours': day_work_hours,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': tech_19
        })

        # Activity 20
        activities.append({
            'subsystem': 'Track',
            'equipment': 'buffer_stop',
            'description': 'Inspection and maintenance of line buffer stop',
            'time_of_day': 'night',
            'periodicity_weeks': 52,
            'duration_hours': 0.7,
            'outage_time_before': T_hors_tension,
            'outage_time_after': T_remise_tension,
            'technicians_needed': 2
        })

        # Activity 21
        activities.append({
            'subsystem': 'Track',
            'equipment': 'buffer_stop_depot',
            'description': 'Inspection and maintenance of depot buffer stop',
            'time_of_day': 'day',
            'periodicity_weeks': 52,
            'duration_hours': 0.7,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 1
        })

        # Activity 22
        activities.append({
            'subsystem': 'Track',
            'equipment': 'wheel_stop',
            'description': 'Inspection and maintenance of depot wheel stop',
            'time_of_day': 'day',
            'periodicity_weeks': 52,
            'duration_hours': 0.5,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 1
        })

        # Activity 23
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Turnout',
            'description': 'Routine inspection of line turnouts',
            'time_of_day': 'night',
            'periodicity_weeks': 2,
            'duration_hours': 0.2,
            'outage_time_before': T_hors_tension / 2,
            'outage_time_after': T_remise_tension / 2,
            'technicians_needed': 2
        })

        # Activity 24
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Turnout',
            'description': 'cleaning, lubrication and functional check of line turnouts',
            'time_of_day': 'night',
            'periodicity_weeks': 13,
            'duration_hours': 2 if params_dict.get('point_machine', 'excluded') == 'excluded' else 3,
            'outage_time_before': T_hors_tension / 2,
            'outage_time_after': T_remise_tension / 2,
            'technicians_needed': 2
        })

        # Activity 25
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Turnout_depot',
            'description': 'Routine inspection of depot turnouts',
            'time_of_day': 'day',
            'periodicity_weeks': 4,
            'duration_hours': 0.3,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 2
        })

        # Activity 26
        activities.append({
            'subsystem': 'Track',
            'equipment': 'Turnout_depot',
            'description': 'cleaning, lubrication and functional check of depot turnouts',
            'time_of_day': 'day',
            'periodicity_weeks': 26,
            'duration_hours': 3 if params_dict.get('point_machine', 'excluded') == 'excluded' else 4.5,
            'outage_time_before': 0,
            'outage_time_after': 0,
            'technicians_needed': 2
        })

        # Activity 27, 28, 29 (feeding_system == 'third_rail')
        if feeding_system == 'third_rail':
            # Activity 27
            activities.append({
                'subsystem': '3rd_Rail',
                'equipment': '3rd_rail_foot',
                'description': 'Foot inspection of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
            # Activity 28
            activities.append({
                'subsystem': '3rd_Rail',
                'equipment': 'third_rail_expansion_joint',
                'description': 'Maintenance of expansion joint of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
            # Activity 29
            if third_rail_measurement_strategy != 'excluded':
                activities.append({
                    'subsystem': '3rd_Rail',
                    'equipment': 'third_rail_measurement',
                    'description': 'Geometric measurement of the line feeding system',
                    'time_of_day': 'night',
                    'periodicity_weeks': 26,
                    'duration_hours': D_inspection_disponible,
                    'outage_time_before': T_hors_tension,
                    'outage_time_after': T_remise_tension,
                    'technicians_needed': 1 if third_rail_measurement_strategy == 'on_board' else 2
                })

        # Activity 30, ..., 38 (feeding_system == 'OCS1')
        if feeding_system == 'OCS1':
            # Activity 30
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_foot',
                'description': 'Foot inspection of the line feeding system',
                'time_of_day': 'day',
                'periodicity_weeks': 8,
                'duration_hours': day_work_hours,
                'outage_time_before': 0,
                'outage_time_after': 0,
                'technicians_needed': 2
            })
            # Activity 31
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_depot_foot',
                'description': 'foot inspection of the feeding system depot',
                'time_of_day': 'day',
                'periodicity_weeks': 13,
                'duration_hours': day_work_hours,
                'outage_time_before': 0,
                'outage_time_after': 0,
                'technicians_needed': 1
            })
            # Activity 32
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_detailed',
                'description': 'Detailed maintenance at height of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 3
            })
            # Activity 33
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_depot_detailed',
                'description': 'Detailed maintenance at height of the depot feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 3
            })
            # Activity 34
            if OCS1_measurement_strategy != 'excluded':
                activities.append({
                    'subsystem': 'CAT',
                    'equipment': 'OCS1_measurement',
                    'description': 'Height & stagger measurement and wire thickness of the line feeding system',
                    'time_of_day': 'night',
                    'periodicity_weeks': 52,
                    'duration_hours': D_inspection_disponible,
                    'outage_time_before': T_hors_tension,
                    'outage_time_after': T_remise_tension,
                    'technicians_needed': 3
                })
            # Activity 35
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_tensionning_device',
                'description': 'Wire tensors inspection and setting of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
            # Activity 36
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'section_insulator',
                'description': 'Section insulators inspection and alignment of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': 1,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 3
            })
            # Activity 37
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'depot_section_insulator',
                'description': 'Section insulators inspection and alignment of the depot feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': 0.5,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 3
            })
            # Activity 38
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'communication',
                'description': 'Communication inspection and alignment of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 26,
                'duration_hours': 1,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 3
            })
            # Activity 39
            activities.append({
                'subsystem': 'CAT',
                'equipment': 'OCS1_feeding_pillar',
                'description': 'Earthing resistance test of feeding pillars of the line feeding system',
                'time_of_day': 'night',
                'periodicity_weeks': 52,
                'duration_hours': D_inspection_disponible,
                'outage_time_before': T_hors_tension,
                'outage_time_after': T_remise_tension,
                'technicians_needed': 2
            })
        # Convertir la liste d'activités en DataFrame
        activities_df = pd.DataFrame(activities)
        # Fusionner activities_df avec track_maintenance_activities
        if track_maintenance_activities.empty:
            track_maintenance_activities = activities_df.copy()
        else:
            track_maintenance_activities = pd.concat([track_maintenance_activities, activities_df], ignore_index=True)
        track_maintenance_activities = track_maintenance_activities.fillna(0).infer_objects(copy=False)

        # Charger les feuilles spécifiques du fichier d'entrée
        _pg_candidates = ['parametres_generaux', 'General Parameters', 'Parameters']
        _pg_sheet = next((s for s in _pg_candidates if s in pd.ExcelFile(file_path, engine='openpyxl').sheet_names), 'parametres_generaux')
        parametres_generaux_df = read_and_strip_all_text_cells(file_path, _pg_sheet, engine='openpyxl')

        corrective_track_df = read_and_strip_all_text_cells(file_path, "corrective_Track")

        ovh_renew_track_df = read_and_strip_all_text_cells(file_path, "ovh_renew_Track")

        subcontract_activities_df = read_and_strip_all_text_cells(file_path, "subcontract_activities")
        
        days_off_df = pd.read_excel(file_path, sheet_name='days_off', engine='openpyxl')

        corrective_3rd_rail_df = None
        ovh_renew_3rd_rail_df = None
        corrective_CAT_df = None
        ovh_renew_CAT_df = None
        sheets = pd.ExcelFile(file_path, engine='openpyxl').sheet_names
        if feeding_system == 'third_rail':
            if 'corrective_3rd_Rail' in sheets:
                corrective_3rd_rail_df = read_and_strip_all_text_cells(file_path, 'corrective_3rd_Rail', engine='openpyxl')
            if 'ovh_renew_3rd_Rail' in sheets:
                ovh_renew_3rd_rail_df = read_and_strip_all_text_cells(file_path, 'ovh_renew_3rd_Rail', engine='openpyxl')

        if feeding_system == 'OCS1':
            if 'corrective_CAT' in sheets:
                corrective_CAT_df = read_and_strip_all_text_cells(file_path, 'corrective_CAT', engine='openpyxl')
                
            if 'ovh_renew_CAT' in sheets:
                ovh_renew_CAT_df = read_and_strip_all_text_cells(file_path, 'ovh_renew_CAT', engine='openpyxl')

        ovh_list = [("ovh_renew_Track", ovh_renew_track_df)]
        if feeding_system == "third_rail":
            ovh_list.append(("ovh_renew_3rd_Rail", ovh_renew_3rd_rail_df))
        if feeding_system == "OCS1":
            ovh_list.append(("ovh_renew_CAT", ovh_renew_CAT_df))
        # APS si un jour tu l’ajoutes : ovh_list.append(("ovh_renew_APS", ovh_renew_APS_df))

        ovh_renew_activities_df, ovh_missing = stack_sheets_in_order(ovh_list)

        corr_list = [("corrective_Track", corrective_track_df)]
        if feeding_system == "third_rail":
            corr_list.append(("corrective_3rd_Rail", corrective_3rd_rail_df))
        if feeding_system == "OCS1":
            corr_list.append(("corrective_CAT", corrective_CAT_df))
        # APS si un jour tu l’ajoutes : corr_list.append(("corrective_APS", corrective_APS_df))

        corrective_activities_df, corr_missing = stack_sheets_in_order(corr_list)

        if ovh_missing:
            print("[INFO] Missing ovh_renew sheets ignored:", ovh_missing)
        if corr_missing:
            print("[INFO] Missing corrective sheets ignored:", corr_missing)

        # Sélection du fichier de sortie
        # Message interactif avant de choisir le fichier excel d'entrée
        if not silent:
            show_message(
                "Saving the output file",
                "Please preferably name the output Excel file as follows 'Project_name_Linear_assets_output_Vref' and save it.")
        print("Please name and save the output Excel file after running the linear_asset_planning...")

        output_file = filedialog_asksave(
            title="Save the output file",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")]
        )

        # Vérifier si un fichier a été sélectionné
        if not output_file:
            print("No file selected for saving. Program stopped.")
            return

        # Ecriture des résultats dans le fichier Excel
        try:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                # Exporter les feuilles d'entrée
                parametres_generaux_df.to_excel(writer, sheet_name='parametres_generaux', index=False)
                days_off_df.to_excel(writer, sheet_name='days_off', index=False)
                subcontract_activities_df.to_excel(writer, sheet_name='subcontract_activities', index=False)     
                zones_df.to_excel(writer, sheet_name='Zones Inspection', index=False)
                track_df.to_excel(writer, sheet_name='Track', index=False)
                fixed_zones_df.to_excel(writer, sheet_name='Fixed_Zones', index=False)
                if feeding_system == 'third_rail':
                    third_rail_df.to_excel(writer, sheet_name='3rd_Rail', index=False)
                if feeding_system == 'OCS1':
                    OCS1_df.to_excel(writer, sheet_name='CAT', index=False)
                track_maintenance_activities.to_excel(writer, sheet_name='maintenance_activities', index=False)
                corrective_track_df.to_excel(writer, sheet_name='corrective_Track', index=False)
                ovh_renew_track_df.to_excel(writer, sheet_name='ovh_renew_Track', index=False)
                premises_positions_df.to_excel(writer, sheet_name="premises_positions", index=False)
                equipements_premises_df.to_excel(writer, sheet_name="equipements_premises", index=False)
                ovh_renew_activities_df.to_excel(writer, sheet_name="ovh_renew_activities", index=False)
                corrective_activities_df.to_excel(writer, sheet_name="corrective_activities", index=False)

                if corrective_3rd_rail_df is not None:
                    corrective_3rd_rail_df.to_excel(writer, sheet_name='corrective_3rd_Rail', index=False)
                if ovh_renew_3rd_rail_df is not None:
                    ovh_renew_3rd_rail_df.to_excel(writer, sheet_name='ovh_renew_3rd_Rail', index=False)
                if corrective_CAT_df is not None:
                    corrective_CAT_df.to_excel(writer, sheet_name='corrective_CAT', index=False)
                if ovh_renew_CAT_df is not None:
                    ovh_renew_CAT_df.to_excel(writer, sheet_name='ovh_renew_CAT', index=False)

            print(f"The file '{output_file}' has been successfully generated.")
            if not silent:
                show_message("Completed", "The file has been generated successfully.")
            
        except Exception as e:
            print(f"Error while saving the Excel file : {e}")
    finally:
        if temp_root is not None:
            try:
                temp_root.destroy()
            except Exception:
                pass

def main(input_file=None, output_file=None, *, silent=False):
    """Appel commun: depuis l’interface (avec chemins) ou en standalone (dialogs)."""
    return core_process(input_file=input_file, output_file=output_file, silent=silent)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Linear assets with subcontract + Ovh/Renewals")
    p.add_argument("--in",  dest="input_file",  help="Input Excel file")
    p.add_argument("--out", dest="output_file", help="Output Excel file")
    p.add_argument("--silent", action="store_true", help="Reduce pop-ups (if any are used)")
    args = p.parse_args()
    main(args.input_file, args.output_file, silent=args.silent)