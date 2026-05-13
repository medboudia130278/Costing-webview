"""
maintenance_and_inspection_planner.py
Entrée  :
  - Feuille 'Points'      : type / nom / pk_km
  - Feuille 'Parameters'  : Nom / Valeur
                       indispensables pour l'implantation :
                         v_day_kmh, v_night_kmh, t_response_min, grid_step_m
                       indispensables pour l'inspection :
                         v_inspect_veh, v_transp_veh,
                         inspect_freq, night_window, line_length
Sortie : Excel unique → 3 onglets
         Implantations_Jour, Implantations_Nuit, Inspection_Schedule
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import xlsxwriter
from math import floor
from collections import OrderedDict
import os, sys
from pulp import (LpProblem, LpVariable, lpSum,
                  LpMinimize, LpBinary, PULP_CBC_CMD)

# ────────────────────────── GUI utilitaires ──────────────────────────
def ask_open_file(msg):
    return filedialog.askopenfilename(
        title=msg, filetypes=[("Excel", "*.xlsx *.xls")])

def ask_save_file():
    return filedialog.asksaveasfilename(
        title="Save the output workbook",
        defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])

# ────────────────────────── Lecture Points + Parameters ──────────────
def load_input(path):
    try:
        xl = pd.ExcelFile(path)                  # ← ouverture unique
        pts = pd.read_excel(xl, sheet_name="Points")
        _prm_candidates = ['Parameters', 'parametres_generaux', 'General Parameters']
        _prm_sheet = next((s for s in _prm_candidates if s in xl.sheet_names), None)
        if _prm_sheet is None:
            raise RuntimeError(f"Sheet not found. Expected one of: {_prm_candidates}")
        prm_df  = pd.read_excel(xl, sheet_name=_prm_sheet)
        
    except Exception as e:
        raise RuntimeError(f"Unable to read file : {e}")
        

    nom_col    = 'Nom'    if 'Nom'    in prm_df.columns else 'Name'
    valeur_col = 'Valeur' if 'Valeur' in prm_df.columns else 'Value'
    if nom_col not in prm_df.columns or valeur_col not in prm_df.columns:
        raise RuntimeError("'Parameters' must contain 'Nom'/'Name' and 'Valeur'/'Value'.")

    prm = prm_df.set_index(nom_col)[valeur_col].apply(lambda v: pd.to_numeric(v, errors="coerce")).fillna(prm_df.set_index(nom_col)[valeur_col])
    needed = {"v_day_kmh", "v_night_kmh", "t_response_min", "grid_step_m",
              "v_inspect_veh", "v_transp_veh", "v_detail_inspect_veh","detail_inspect_freq",
              "inspect_freq", "night_window", "line_length", "num_track", "num_ultrason_veh", "num_clean_veh",
              "num_inspec_veh", "v_grind_veh", "v_transp_grind_veh", "v_ultrason_veh","v_transp_ultrason_veh","ultrason_freq",
              "grind_freq", "num_grind_veh", "v_clean_veh","v_transp_clean_veh","clean_freq", "capacity_nights_an","work_nights_week"}
    missing = needed - set(prm.index)
    if missing:
        raise RuntimeError(f"Missing parameters : {', '.join(missing)}")
        
    
    if pts.empty:
        raise RuntimeError("The sheet 'Points' is empty.")
        

    # --- lecture (ou fallback) de Velocity ----------------------------
    vel = load_velocity(xl, prm)    # ← appelle la fonction avec fallback

    return pts.reset_index(drop=True), prm, vel.reset_index(drop=True)


# ────────────────────────── Aide : parking le plus proche ────────────
def nearest_idx(arr, value):
    arr = np.asarray(arr)
    return int(np.abs(arr - value).argmin())

# ────────────────────── PARTIE 1 : IMPLANTATION DES ÉQUIPES ───────────

def add_responsibilities(df_bases, step_m, pk_max):
    """Découpage Voronoï 1-D (équipes responsables du point le plus proche)."""
    df = df_bases.sort_values("pk_km").reset_index(drop=True).copy()
    pks = df["pk_km"].to_numpy()
    nb_pts = int((pk_max*1000)//step_m) + 1
    grid = np.arange(nb_pts)*step_m/1000  # km

    prio = df["type"].map({"depot": 0, "station": 1}).to_numpy()
    owner = np.zeros_like(grid, dtype=int)
    for i, g in enumerate(grid):
        dist = np.abs(pks - g)
        eq   = np.flatnonzero(dist == dist.min())
        best = eq[prio[eq].argmin()] if len(eq) > 1 else eq[0]
        owner[i] = best

    # extraire intervalles
    starts = [None]*len(df); ends = [None]*len(df)
    start_idx = 0
    for i in range(1, len(grid)):
        if owner[i] != owner[i-1]:
            idx = owner[i-1]
            starts[idx] = grid[start_idx]
            ends[idx]   = grid[i]
            start_idx   = i
    idx = owner[-1]
    starts[idx] = grid[start_idx];  ends[idx] = grid[-1]

    df["resp_pk_start"] = np.round(starts,3)
    df["resp_pk_end"]   = np.round(ends,3)
    
    return df

def cover_for_speed(
        pts: pd.DataFrame,
        v_kmh: float,
        prm: pd.Series,
        label: str = ""
    ) -> pd.DataFrame:
    """
    Sélectionne la liste minimale de bases (dépôts + stations) pour couvrir
    toute la ligne dans le délai imposé.
    - v_kmh          : vitesse du scénario en km/h
    - label          : '', 'NUIT', etc. (pour le DEBUG)
    Retour : DataFrame trié par pk_km, avec resp_pk_start / resp_pk_end.
    """
    # -------- constantes générales --------
    d_max  = v_kmh * prm["t_response_min"] / 60          # rayon de couverture (km)
    pk_max = prm["line_length"]                          # longueur voie simple (km)
    step_m = int(prm["grid_step_m"])                     # pas de grille (m)

    # -------- grille à couvrir --------
    nb_pts = int((pk_max * 1000) // step_m) + 1
    grid   = np.arange(nb_pts) * step_m / 1000           # km

    # -------- dépôts (imposés) et stations candidates --------
    depots   = pts[pts["type"] == "depot"].copy()
    stations = pts[pts["type"] == "station"].copy()

    pk_dep = depots["pk_km"].to_numpy()
    pk_sta = stations["pk_km"].to_numpy()

    # Couverture initiale par les seuls dépôts
    covered_init = (np.abs(pk_dep[:, None] - grid[None, :]) <= d_max).any(axis=0)

    # Si des points restent découverts → set-cover avec les stations
    if not covered_init.all() and not stations.empty:
        grid_nc = grid[~covered_init]                              # points non couverts
        cover_sta = (np.abs(pk_sta[:, None] - grid_nc[None, :]) <= d_max).astype(int)

        prob = LpProblem("cover_stations", LpMinimize)
        x = LpVariable.dicts("x", range(len(pk_sta)), 0, 1, LpBinary)
        prob += lpSum(x.values())
        for j in range(len(grid_nc)):
            prob += lpSum(x[i] * cover_sta[i, j] for i in range(len(pk_sta))) >= 1
        prob.solve(PULP_CBC_CMD(msg=False))

        stations["implanter_equipe"] = [int(x[i].value()) for i in range(len(pk_sta))]
        stations = stations[stations["implanter_equipe"] == 1]
    else:
        stations = stations.iloc[0:0]       # vide si rien à ajouter

    # -------- fusion dépôts + stations retenues --------
    depots["implanter_equipe"] = 1
    bases = pd.concat([depots, stations]).sort_values("pk_km").reset_index(drop=True)

    # -------- DEBUG final : vérification de couverture --------
    pk_all = bases["pk_km"].to_numpy()
    covered_final = (np.abs(pk_all[:, None] - grid[None, :]) <= d_max).any(axis=0)
    uncovered_len = (~covered_final).sum() * (step_m / 1000)        # km non couverts

    print(f"[DEBUG {label}] v={v_kmh} km/h  d_max={d_max:.1f} km  "
          f"bases={len(bases)}  FINALgap={uncovered_len:.1f} km")

    # -------- calcul des responsabilités (Voronoï 1-D) --------
    bases = add_responsibilities(bases, step_m, pk_max)
    return bases

def build_implantations(pts, prm):
    jour = cover_for_speed(pts, prm["v_day_kmh"],   prm, label="DAY")
    nuit = cover_for_speed(pts, prm["v_night_kmh"], prm, label="NIGHT")
    return jour, nuit

def travel_time(pk_a: float, pk_b: float, vel: pd.DataFrame, col: str) -> float:
    """
    Renvoie le temps (h) pour parcourir la distance |pk_a → pk_b|
    en intégrant la vitesse *col* pour chaque segment Velocity traversé.
    """
    if pk_a == pk_b:
        return 0.0
    d_sign = 1 if pk_b > pk_a else -1
    pk1, pk2 = sorted([pk_a, pk_b])     # sens croissant pour l'algorithme
    subset = vel[(vel.pk_end > pk1) & (vel.pk_start < pk2)]

    tot_h = 0.0
    for _, seg in subset.iterrows():
        seg_start = max(pk1, seg.pk_start)
        seg_end   = min(pk2, seg.pk_end)
        dist = seg_end - seg_start
        v = seg[col]
        if v <= 0:
            raise ValueError(f"Invalid speed value (zero or negative) in {col} (PK {seg_start}-{seg_end}).")
        tot_h += dist / v
    return tot_h

# ───────────── PARTIE 2 : PLANNING D'INSPECTION DE NUIT ──────────────
def nightly_segment(pk_curr, park_idx, pk_parks, night_h, t_transp, v_i, pk_max):
    """Calcule la portion max inspectable depuis pk_curr (sens croissant)."""
    pk_start = pk_curr
    t_out = t_transp(pk_start, pk_parks[park_idx])
    if t_out >= night_h:
        raise ValueError("Insufficient night window (outbound travel).")

    # borne haute si pas de retour
    pk_high = pk_start + (night_h - t_out)*v_i
    pk_high = min(pk_high, pk_max)

    # recherche dichotomique pour inclure retour
    low, high = pk_start, pk_high
    while high - low > 0.01:
        mid = (low+high)/2
        p_ret = nearest_idx(pk_parks, mid)
        t_back  = t_transp(mid, pk_parks[p_ret])
        t_insp = (mid - pk_start)/v_i
        if t_out + t_insp + t_back <= night_h:
            low = mid
        else:
            high = mid
    pk_end = round(low,3)
    p_ret  = nearest_idx(pk_parks, pk_end)
    t_back = t_transp(pk_end, pk_parks[p_ret])
    t_insp = (pk_end - pk_start)/v_i
    return pk_start, pk_end, park_idx, p_ret, round(t_out,3), round(t_insp,3), round(t_back,3)

def check_coverage(df: pd.DataFrame,
                   tracks: list[dict],
                   pk_L: float,
                   eps: float,
                   label: str,
                   enforce_limit: bool):
    """
    • Vérifie si chaque voie est terminée.
    • Si incomplète et enforce_limit=True → lève RuntimeError.
    • Sinon, pose df.attrs["coverage_warning"] avec un message clair.
    """
    incompletes = [i for i, tr in enumerate(tracks)
                   if tr["pk_curr"] < pk_L - eps]

    if not incompletes:
        return                           # tout est couvert

    reste = max(pk_L - tracks[i]["pk_curr"] for i in incompletes)
    voies = ", ".join(f"T{idx+1}" for idx in incompletes)
    msg   = f"{label} : tracks {voies} incomplete (remaining ≥ {reste:.1f} km)"

    if enforce_limit:
        raise RuntimeError(msg)
    else:
        df.attrs["coverage_warning"] = msg

def load_velocity(xl: pd.ExcelFile, prm: pd.Series) -> pd.DataFrame:
    """
    Renvoie un DataFrame normalisé avec les colonnes:
        pk_start, pk_end,
        v_transp_veh, v_transp_grind_veh, v_transp_ultrason_veh,
        v_transp_clean_veh, v_transp_measu_veh
    S'il n'existe pas de feuille 'Velocity', on crée un segment unique
    couvrant toute la ligne avec les vitesses globales.
    """
    line_L = float(prm["line_length"])

    # --- plan B : segment unique issu de Parameters -------------------------
    def default_segment() -> pd.DataFrame:
        return pd.DataFrame([{
            "pk_start": 0.0,
            "pk_end":   line_L,
            "v_transp_veh":           prm["v_transp_veh"],
            "v_transp_grind_veh":     prm["v_transp_grind_veh"],
            "v_transp_ultrason_veh":  prm["v_transp_ultrason_veh"],
            "v_transp_clean_veh":     prm["v_transp_clean_veh"],
            "v_transp_measu_veh":     prm["v_transp_measu_veh"],
        }])

    if "Velocity" not in xl.sheet_names:
        return default_segment()

    vel = pd.read_excel(xl, sheet_name="Velocity")

    if vel.empty:
        return default_segment()

    expected = {
        "segment", "pk_start", "pk_end", "radius",
        "v_transp_veh", "v_transp_grind_veh", "v_transp_ultrason_veh",
        "v_transp_clean_veh", "v_transp_measu_veh"
    }
    if not expected.issubset(vel.columns):
        raise ValueError(
            "Invalid Velocity sheet : "
            "missing columns → "
            + ", ".join(expected - set(vel.columns))
        )

    # normalisation
    vel = (vel
           .rename(columns={"Pk_start": "pk_start", "Pk_end": "pk_end"})
           .sort_values("pk_start")
           .reset_index(drop=True))

    # contrôles rapides
    if vel["pk_start"].iloc[0] > 0 or vel["pk_end"].iloc[-1] < line_L:
        raise ValueError("Velocity segments do not cover the full range 0–line_length.")
    if (vel["pk_end"] < vel["pk_start"]).any():
        raise ValueError("At least one segment has pk_end < pk_start.")
    if (vel["pk_start"].shift(-1) < vel["pk_end"]).iloc[:-1].any():
        raise ValueError("Overlapping Velocity segments detected.")

    return vel


def build_inspection_schedule(
        pts: pd.DataFrame,
        prm: pd.Series,
        vel,
        *,
        v_kmh: float,           # vitesse du mode inspection (km/h)
        freq_sem: float,        # périodicité globale (semaines)
        veh_prefix: str,        # 'V'   pour standard, 'DV' pour détail…
        col_suffix: str, 
        label="Inspection",            
        enforce_limit=True                 
    ) -> pd.DataFrame:
    """
    Planifie une inspection linéaire multi-voies / multi-véhicules.
    Renvoie un DataFrame vide si num_inspec_veh = 0.
    Lève RuntimeError si la fréquence est intenable.
    Colonnes retournées :
        nuit • vehicule • voie • pk_start • pk_end • parking_depart •
        parking_arrivee • t_out_h • t_<col_suffix>_h • t_back_h
    """

    # ───── paramètres communs ──────────────────────────────────────────
    n_veh   = int(prm.get("num_inspec_veh", 1))
    NIGHTS_PER_WEEK = int(prm.get("work_nights_week", 7))   # défaut : 7

    cols    = ["nuit", "vehicule", "voie", "pk_start", "pk_end",
               "parking_depart", "parking_arrivee",
               "t_out_h", f"t_{col_suffix}_h", "t_back_h"]

    if n_veh < 1:
        return pd.DataFrame(columns=cols)

    pk_L      = float(prm["line_length"])           # km (voie simple)
    v_t       = float(prm["v_transp_veh"])          # km/h (transport)
    night_h   = float(prm["night_window"])          # h
    nights_ok = int(freq_sem * NIGHTS_PER_WEEK)                   # nuits maxi
    n_tracks  = int(prm.get("num_track", 2))        # voies simples

    t_transp = lambda a, b: travel_time(a, b, vel, "v_transp_veh")
    # ───── parkings utilisables ────────────────────────────────────────
    pk_parks = pts[pts["type"].isin(["depot", "pocket"])].sort_values("pk_km")
    if pk_parks.empty:
        raise RuntimeError("No depot or pocket track available to park the vehicle (Inspection).")

    pk_vals = pk_parks["pk_km"].to_numpy()
    names   = pk_parks["nom"].to_numpy()

    EPS = max(0.001, float(prm["grid_step_m"])/2000)   # 0,5 pas ou 1 m mini

    # état de progression de chaque voie
    tracks = [{"pk_curr": 0.0,
               "park_idx": nearest_idx(pk_vals, 0.0)} for _ in range(n_tracks)]

    rec, nuit = [], 1
    while True:
        # toutes les voies terminées ?
        if all(tr["pk_curr"] >= pk_L - EPS for tr in tracks):
            break
        # dépassement de fréquence ?
        if enforce_limit and nuit > nights_ok:
            raise RuntimeError(f"Calculated nights exceed the allowed limit ({nights_ok} nights)."
                               f"Frequency not achievable – {freq_sem} weeks × {NIGHTS_PER_WEEK} nights "
                               f"are insufficient with {n_veh} vehicle(s)."
            )

        progressed = False
        # voies triées par avancement croissant
        ordre = sorted(range(n_tracks),
                       key=lambda i: tracks[i]["pk_curr"]/pk_L)

        for v in range(n_veh):
            cible = next((i for i in ordre
                          if tracks[i]["pk_curr"] < pk_L - EPS), None)
            if cible is None:
                break       # plus de travail pour cette nuit

            tr = tracks[cible]
            try:
                pk_s, pk_e, p_dep, p_arr, t_out, t_run, t_back = nightly_segment(tr["pk_curr"], tr["park_idx"],
                    pk_vals, night_h, t_transp, v_kmh, pk_L)
            except ValueError:
                continue

            # sécurité : si l'avancée < EPS, on déclare la voie terminée
            if pk_e - pk_s < EPS:
                # pas d'avancée possible cette nuit : on considérera la voie incomplète
                # → on la laisse avec son pk_curr actuel
                tracks[cible]["park_idx"] = p_arr   # on peut au moins changer de parking
                continue

            progressed = True
            rec.append({
                "nuit":           nuit,
                "vehicule":       f"{veh_prefix}{v+1}",
                "voie":           f"T{cible+1}",
                "pk_start":       pk_s,
                "pk_end":         pk_e,
                "parking_depart": names[p_dep],
                "parking_arrivee":names[p_arr],
                "t_out_h":        t_out,
                f"t_{col_suffix}_h": t_run,
                "t_back_h":       t_back
            })

            tracks[cible]["pk_curr"]  = pk_e
            tracks[cible]["park_idx"] = p_arr

        if not progressed:
            # b) nuit perdue : impossible d'aller plus loin avec les paramètres actuels
            msg = (f"{label} : no further progress possible during night {nuit}. "
                "Night-time duration too short or insufficient parking locations.")
            if enforce_limit:
                raise RuntimeError(msg)
            else:
                df = pd.DataFrame(rec, columns=cols)
                df.attrs["coverage_warning"] = msg
                return df

        nuit += 1

    df = pd.DataFrame(rec, columns=cols)

    # 2) contrôle de couverture
    check_coverage(df, tracks, pk_L, EPS, label=label, enforce_limit=enforce_limit)

    return df

# ─────────────────── build_grind_schedule ───────────────────
def build_grind_schedule(pts: pd.DataFrame, prm: pd.Series, vel, label="Grinding", enforce_limit=True) -> pd.DataFrame:
    """
    Planifie le meulage des num_track voies simples avec num_grind_veh véhicules.
    Rend un DataFrame vide si num_grind_veh = 0.
    """
    n_veh = int(prm.get("num_grind_veh", 1))
    NIGHTS_PER_WEEK = int(prm.get("work_nights_week", 7))   # défaut : 7
    cols = ["nuit", "vehicule", "voie", "pk_start", "pk_end",
            "parking_depart", "parking_arrivee",
            "t_out_h", "t_grind_h", "t_back_h"]
    if n_veh < 1:
        return pd.DataFrame(columns=cols)

    # ---------- paramètres ----------
    pk_L    = float(prm["line_length"])
    v_g     = float(prm["v_grind_veh"])
    v_t     = float(prm["v_transp_grind_veh"])
    night_h = float(prm["night_window"])
    grind_freq = float(prm["grind_freq"])
    nights_ok = int(float(prm["grind_freq"]) * NIGHTS_PER_WEEK)
    n_tracks  = int(prm.get("num_track", 2))

    t_transp = lambda a, b: travel_time(a, b, vel, "v_transp_grind_veh")

    pk_parks = pts[pts["type"].isin(["depot","pocket"])].sort_values("pk_km")
    if pk_parks.empty:
        raise RuntimeError("No depot or pocket track available to park the vehicle (grinding).")
    pk_vals = pk_parks["pk_km"].to_numpy()
    names   = pk_parks["nom"].to_numpy()

    EPS = max(0.001, float(prm["grid_step_m"])/2000)

    # état de chaque voie
    tracks = [{"pk_curr": 0.0,
               "park_idx": nearest_idx(pk_vals, 0.0)} for _ in range(n_tracks)]

    rec  = [];  nuit = 1
    while True:
        if all(tr["pk_curr"] >= pk_L - EPS for tr in tracks):
            break
        if enforce_limit and nuit > nights_ok:
            raise RuntimeError(f"Calculated nights exceed the allowed limit ({nights_ok} nights)."
                               f"Unfeasible frequency – {grind_freq} weeks × {NIGHTS_PER_WEEK} nights "
                               f"are insufficient with {n_veh} vehicle(s)."
            )

        progressed = False
        # voies triées par avancement croissant
        prog = sorted(range(n_tracks),
                      key=lambda i: tracks[i]["pk_curr"]/pk_L)

        for v in range(n_veh):
            cible = next((i for i in prog
                          if tracks[i]["pk_curr"] < pk_L - EPS), None)
            if cible is None:
                break

            tr  = tracks[cible]
            try:
                pk_s, pk_e, p_dep, p_arr,t_out, t_gr, t_back = nightly_segment(
                    tr["pk_curr"], tr["park_idx"],
                    pk_vals, night_h, t_transp, v_g, pk_L)
            except ValueError:
                continue

            if pk_e - pk_s < EPS:
                # pas d'avancée possible cette nuit : on considérera la voie incomplète
                # → on la laisse avec son pk_curr actuel
                tracks[cible]["park_idx"] = p_arr   # on peut au moins changer de parking
                continue

            progressed = True
            rec.append(dict(
                nuit=nuit, vehicule=f"G{v+1}", voie=f"T{cible+1}",
                pk_start=pk_s, pk_end=pk_e,
                parking_depart=names[p_dep],
                parking_arrivee=names[p_arr],
                t_out_h=t_out, t_grind_h=t_gr, t_back_h=t_back
            ))

            tracks[cible]["pk_curr"]  = pk_e
            tracks[cible]["park_idx"] = p_arr

        if not progressed:
            # b) nuit perdue : impossible d'aller plus loin avec les paramètres actuels
            msg = (f"{label} : no further progress possible on night {nuit}. "
                "Night window too short or insufficient parking capacity.")
            if enforce_limit:
                raise RuntimeError(msg)
            else:
                df = pd.DataFrame(rec, columns=cols)
                df.attrs["coverage_warning"] = msg
                return df

        nuit += 1

    df = pd.DataFrame(rec, columns=cols)

    # 2) contrôle de couverture
    check_coverage(df, tracks, pk_L, EPS, label=label, enforce_limit=enforce_limit)

    return df


# ─────────────────── build_measu_schedule ───────────────────
def build_measu_schedule(pts: pd.DataFrame, prm: pd.Series, vel, label="Measurement", enforce_limit=True) -> pd.DataFrame:
    """
    Planifie la mesure des num_track voies simples avec num_measu_veh véhicules.
    Renvoie un DataFrame vide (en-tête seul) si num_measu_veh = 0.
    """
    n_measu_veh = int(prm.get("num_measu_veh", 1))
    NIGHTS_PER_WEEK = int(prm.get("work_nights_week", 7))   # défaut : 7
    cols = ["nuit", "vehicule", "voie", "pk_start", "pk_end",
            "parking_depart", "parking_arrivee",
            "t_out_h", "t_measure_h", "t_back_h"]
    if n_measu_veh < 1:
        return pd.DataFrame(columns=cols)

    # ---- paramètres spécifiques ----
    pk_L    = float(prm["line_length"])
    v_m     = float(prm["v_measu_veh"])
    v_t     = float(prm["v_transp_measu_veh"])
    night_h = float(prm["night_window"])
    measu_freq = float(prm["measu_freq"])
    nights_ok = int(float(prm["measu_freq"]) * NIGHTS_PER_WEEK)
    n_tracks  = int(prm.get("num_track", 2))

    t_transp = lambda a, b: travel_time(a, b, vel, "v_transp_measu_veh")
    # ---- parkings ----
    pk_parks = pts[pts["type"].isin(["depot", "pocket"])].sort_values("pk_km")
    if pk_parks.empty:
        raise RuntimeError("No depot or pocket track available to park the vehicle (measurement).")
    pk_vals = pk_parks["pk_km"].to_numpy()
    names   = pk_parks["nom"].to_numpy()

    EPS = max(0.001, float(prm["grid_step_m"]) / 2000)

    tracks = [{"pk_curr": 0.0,
               "park_idx": nearest_idx(pk_vals, 0.0)} for _ in range(n_tracks)]

    rec, nuit = [], 1
    while True:
        if all(tr["pk_curr"] >= pk_L - EPS for tr in tracks):
            break
        if enforce_limit and nuit > nights_ok:
            raise RuntimeError(f"Calculated nights exceed the allowed limit ({nights_ok} nights)."
                               f"Unfeasible frequency – {measu_freq} weeks × {NIGHTS_PER_WEEK} nights "
                               f"are insufficient with {n_measu_veh} vehicle(s)."
            )
     
        progressed = False
        # voies les moins avancées d'abord
        order = sorted(range(n_tracks),
                       key=lambda i: tracks[i]["pk_curr"]/pk_L)

        for v in range(n_measu_veh):
            cible = next((i for i in order
                          if tracks[i]["pk_curr"] < pk_L - EPS), None)
            if cible is None:
                break

            tr = tracks[cible]
            try:
                pk_s, pk_e, p_dep, p_arr, t_out, t_meas, t_back = nightly_segment(
                    tr["pk_curr"], tr["park_idx"],
                    pk_vals, night_h, t_transp, v_m, pk_L)
            except ValueError:
                continue

            if pk_e - pk_s < EPS:
                # pas d'avancée possible cette nuit : on considérera la voie incomplète
                # → on la laisse avec son pk_curr actuel
                tracks[cible]["park_idx"] = p_arr   # on peut au moins changer de parking
                continue

            progressed = True
            rec.append(dict(
                nuit=nuit, vehicule=f"M{v+1}", voie=f"T{cible+1}",
                pk_start=pk_s, pk_end=pk_e,
                parking_depart=names[p_dep],
                parking_arrivee=names[p_arr],
                t_out_h=t_out, t_measure_h=t_meas, t_back_h=t_back
            ))

            tracks[cible]["pk_curr"]  = pk_e
            tracks[cible]["park_idx"] = p_arr

        if not progressed:
            # b) nuit perdue : impossible d'aller plus loin avec les paramètres actuels
            msg = (f"{label} : no further progress possible on night {nuit}. "
                "Night window too short or insufficient parking capacity.")
            if enforce_limit:
                raise RuntimeError(msg)
            else:
                df = pd.DataFrame(rec, columns=cols)
                df.attrs["coverage_warning"] = msg
                return df

        nuit += 1

    df = pd.DataFrame(rec, columns=cols)

    # 2) contrôle de couverture
    check_coverage(df, tracks, pk_L, EPS, label=label, enforce_limit=enforce_limit)

    return df

# ─────────────────── build_ultrason_schedule ───────────────────
def build_ultrason_schedule(pts: pd.DataFrame, prm: pd.Series, vel, label="Ultrason", enforce_limit=True) -> pd.DataFrame:
    """
    Planifie le contrôle ultrasonique multi-voies / multi-véhicules.
    Renvoie un DataFrame vide si num_ultrason_veh = 0.
    """
    n_us_veh = int(prm.get("num_ultrason_veh", 1))
    NIGHTS_PER_WEEK = int(prm.get("work_nights_week", 7))   # défaut : 7
    cols = ["nuit", "vehicule", "voie", "pk_start", "pk_end",
            "parking_depart", "parking_arrivee",
            "t_out_h", "t_ultrason_h", "t_back_h"]
    if n_us_veh < 1:
        return pd.DataFrame(columns=cols)

    pk_L      = float(prm["line_length"])
    v_us      = float(prm["v_ultrason_veh"])
    v_t       = float(prm["v_transp_ultrason_veh"])
    night_h   = float(prm["night_window"])
    ultrason_freq = float(prm["ultrason_freq"])
    nights_ok = int(float(prm["ultrason_freq"]) * NIGHTS_PER_WEEK)
    n_tracks  = int(prm.get("num_track", 2))

    t_transp = lambda a, b: travel_time(a, b, vel, "v_transp_ultrason_veh")

    pk_parks = pts[pts["type"].isin(["depot", "pocket"])].sort_values("pk_km")
    if pk_parks.empty:
        raise RuntimeError("No depot or pocket track available to park the vehicle US.")
    pk_vals = pk_parks["pk_km"].to_numpy()
    names   = pk_parks["nom"].to_numpy()

    EPS = max(0.001, float(prm["grid_step_m"]) / 2000)

    tracks = [{"pk_curr": 0.0,
               "park_idx": nearest_idx(pk_vals, 0.0)} for _ in range(n_tracks)]

    rec, nuit = [], 1
    while True:
        if all(tr["pk_curr"] >= pk_L - EPS for tr in tracks):
            break
        if enforce_limit and nuit > nights_ok:
            raise RuntimeError(f"Calculated nights exceed the allowed limit ({nights_ok} nights)."
                               f"Unfeasible frequency – {ultrason_freq} weeks × {NIGHTS_PER_WEEK} nights "
                               f"are insufficient with {n_us_veh} vehicle(s)."
            )

        progressed = False
        order = sorted(range(n_tracks),
                       key=lambda i: tracks[i]["pk_curr"]/pk_L)

        for v in range(n_us_veh):
            cible = next((i for i in order
                          if tracks[i]["pk_curr"] < pk_L - EPS), None)
            if cible is None:
                break

            tr = tracks[cible]
            try:
                pk_s, pk_e, p_dep, p_arr, t_out, t_us, t_back = nightly_segment(
                    tr["pk_curr"], tr["park_idx"],
                    pk_vals, night_h, t_transp, v_us, pk_L)
            except ValueError:
                continue

            if pk_e - pk_s < EPS:
                # pas d'avancée possible cette nuit : on considérera la voie incomplète
                # → on la laisse avec son pk_curr actuel
                tracks[cible]["park_idx"] = p_arr   # on peut au moins changer de parking
                continue

            progressed = True
            rec.append(dict(
                nuit=nuit, vehicule=f"U{v+1}", voie=f"T{cible+1}",
                pk_start=pk_s, pk_end=pk_e,
                parking_depart=names[p_dep],
                parking_arrivee=names[p_arr],
                t_out_h=t_out, t_ultrason_h=t_us, t_back_h=t_back
            ))

            tracks[cible]["pk_curr"]  = pk_e
            tracks[cible]["park_idx"] = p_arr

        if not progressed:
            # b) nuit perdue : impossible d'aller plus loin avec les paramètres actuels
            msg = (f"{label} : no further progress possible on night {nuit}. "
                "Night window too short or insufficient parking capacity.")
            if enforce_limit:
                raise RuntimeError(msg)
            else:
                df = pd.DataFrame(rec, columns=cols)
                df.attrs["coverage_warning"] = msg
                return df

        nuit += 1

    df = pd.DataFrame(rec, columns=cols)

    # 2) contrôle de couverture
    check_coverage(df, tracks, pk_L, EPS, label=label, enforce_limit=enforce_limit)

    return df

# ─────────────── build_clean_schedule ────────────────
def build_clean_schedule(pts: pd.DataFrame, prm: pd.Series, vel, label="Cleaning", enforce_limit=True) -> pd.DataFrame:
    """
    Planifie le nettoyage voie-par-voie avec num_clean_veh véhicules.
    Renvoie un DataFrame vide (en-tête seul) si num_clean_veh = 0.
    """
    n_clean = int(prm.get("num_clean_veh", 1))
    NIGHTS_PER_WEEK = int(prm.get("work_nights_week", 7))   # défaut : 7
    cols = ["nuit", "vehicule", "voie", "pk_start", "pk_end",
            "parking_depart", "parking_arrivee",
            "t_out_h", "t_clean_h", "t_back_h"]
    if n_clean < 1:
        return pd.DataFrame(columns=cols)

    pk_L      = float(prm["line_length"])
    v_c       = float(prm["v_clean_veh"])
    v_t       = float(prm["v_transp_clean_veh"])
    night_h   = float(prm["night_window"])
    clean_freq = float(prm["clean_freq"])
    nights_ok = int(float(prm["clean_freq"]) * NIGHTS_PER_WEEK)
    n_tracks  = int(prm.get("num_track", 2))

    t_transp = lambda a, b: travel_time(a, b, vel, "v_transp_clean_veh")

    pk_parks = pts[pts["type"].isin(["depot", "pocket"])]\
                   .sort_values("pk_km")
    if pk_parks.empty:
        raise RuntimeError("No depot or pocket track available to park the cleaning vehicle.")
    pk_vals = pk_parks["pk_km"].to_numpy()
    names   = pk_parks["nom"].to_numpy()

    EPS = max(0.001, float(prm["grid_step_m"]) / 2000)

    tracks = [{"pk_curr": 0.0,
               "park_idx": nearest_idx(pk_vals, 0.0)} for _ in range(n_tracks)]

    rec, nuit = [], 1
    while True:
        if all(tr["pk_curr"] >= pk_L - EPS for tr in tracks):
            break
        if enforce_limit and nuit > nights_ok:
            raise RuntimeError(f"Calculated nights exceed the allowed limit ({nights_ok} nights)."
                               f"Unfeasible frequency – {clean_freq} weeks × {NIGHTS_PER_WEEK} nights "
                               f"are insufficient with {n_clean} vehicle(s)."
            )

        progressed = False
        ordre = sorted(range(n_tracks),
                       key=lambda i: tracks[i]["pk_curr"]/pk_L)

        for v in range(n_clean):
            cible = next((i for i in ordre
                          if tracks[i]["pk_curr"] < pk_L - EPS), None)
            if cible is None:
                break

            tr = tracks[cible]
            try:
                pk_s, pk_e, p_dep, p_arr, t_out, t_run, t_back = nightly_segment(
                    tr["pk_curr"], tr["park_idx"],
                    pk_vals, night_h, t_transp, v_c, pk_L)
            except ValueError:
                continue

            if pk_e - pk_s < EPS:
                # pas d'avancée possible cette nuit : on considérera la voie incomplète
                # → on la laisse avec son pk_curr actuel
                tracks[cible]["park_idx"] = p_arr   # on peut au moins changer de parking
                continue
            progressed = True
            rec.append(dict(
                nuit=nuit, vehicule=f"C{v+1}", voie=f"T{cible+1}",
                pk_start=pk_s, pk_end=pk_e,
                parking_depart=names[p_dep],
                parking_arrivee=names[p_arr],
                t_out_h=t_out, t_clean_h=t_run, t_back_h=t_back))

            tracks[cible]["pk_curr"]  = pk_e
            tracks[cible]["park_idx"] = p_arr

        if not progressed:
            # b) nuit perdue : impossible d'aller plus loin avec les paramètres actuels
            msg = (f"{label} : no further progress possible on night {nuit}. "
                "Night window too short or insufficient parking capacity.")
            if enforce_limit:
                raise RuntimeError(msg)
            else:
                df = pd.DataFrame(rec, columns=cols)
                df.attrs["coverage_warning"] = msg
                return df

        nuit += 1

    df = pd.DataFrame(rec, columns=cols)

    # 2) contrôle de couverture
    check_coverage(df, tracks, pk_L, EPS, label=label, enforce_limit=enforce_limit)

    return df


# ────────────────────────── EXPORT MULTI-ONGLETS ─────────────────────
def export_all(jour, nuit, insp_sched, grind_sched, measu_sched, us_sched,
               clean_sched,              # ← nouveau
               path):
    try:
        with pd.ExcelWriter(path, engine="xlsxwriter") as wr:
            jour .to_excel(wr, sheet_name="Day_Deployment",  index=False)
            nuit .to_excel(wr, sheet_name="Night_Deployment", index=False)
            insp_sched.to_excel(wr, sheet_name="Inspection_Schedule",index=False)
            grind_sched.to_excel(wr, sheet_name="Grinding_Schedule",  index=False)
            measu_sched .to_excel(wr, sheet_name="Measurement_Schedule", index=False)
            us_sched    .to_excel(wr, sheet_name="Ultrason_Schedule",   index=False)
            clean_sched .to_excel(wr, sheet_name="Cleaning_Schedule",    index=False)
    except Exception as e:
        raise RuntimeError(f"Impossible saving : {e}")
        

def add_synoptic_sheet(pts: pd.DataFrame, prm: pd.Series, writer: pd.ExcelWriter):
    """
    Ajoute un onglet 'Synoptic' :
      • n voies horizontales (n = num_track)
      • marqueurs : dépôts (♦), stations (●), pockets (■)
    """
    # 0) paramètres généraux ---------------------------------------------------
    pk_max   = float(prm["line_length"])
    n_tracks = int(prm.get("num_track", 2))          # défaut = 2 voies
    y_lines  = [-i for i in range(n_tracks)]         # 0, -1, -2, …

    # 1) nouveau sheet ---------------------------------------------------------
    sheet = writer.book.add_worksheet('Synoptic')
    sheet.write_row(0, 0, ["type", "pk_km", "y"])

    chart = writer.book.add_chart({'type': 'scatter', 'subtype': 'straight'})

    # 2) tracer les voies (lignes grises) --------------------------------------
    for y in y_lines:
        # on écrit deux points factices pour chaque ligne
        row = sheet.dim_rowmax + 1
        sheet.write_row(row,   0, ["_void", 0,      y])
        sheet.write_row(row+1, 0, ["_void", pk_max, y])

        chart.add_series({
            'categories': ['Synoptic', row,   1, row+1, 1],   # x = pk
            'values'    : ['Synoptic', row,   2, row+1, 2],   # y constants
            'name'      : f'Voie {y_lines.index(y)+1}',
            'line'      : {'color': '#A0A0A0'},
            'marker'    : {'type': 'none'},
        })

    # 3) dictionnaires marqueur/couleur ----------------------------------------
    markers = {'depot': 'diamond', 'station': 'circle', 'pocket': 'square'}
    colors  = {'depot': '#C00000', 'station': '#0070C0', 'pocket': '#00B050'}

    # 4) écrire et tracer chaque type ------------------------------------------
    current_row = sheet.dim_rowmax + 1
    for typ in ('depot', 'station', 'pocket'):
        subset = pts[pts['type'] == typ]
        if subset.empty:
            continue

        start = current_row
        for pk in subset['pk_km']:
            sheet.write_row(current_row, 0, [typ, pk, 0])   # pk sur voie 1 (y=0)
            current_row += 1
        end = current_row - 1

        chart.add_series({
            'categories': ['Synoptic', start, 1, end, 1],
            'values'    : ['Synoptic', start, 2, end, 2],
            'name'      : typ.capitalize(),
            'line'      : {'none': True},
            'marker'    : {
                'type':  markers[typ],
                'size':  7,
                'fill':  {'color': colors[typ]},
                'border': {'color': colors[typ]}
            }
        })

    # 5) axes et insertion ------------------------------------------------------
    chart.set_x_axis({'name': 'PK (km)', 'min': 0, 'max': pk_max})
    chart.set_y_axis({'visible': False})
    chart.set_legend({'position': 'bottom'})
    chart.set_size({'width': 1100, 'height': 70 + 70*n_tracks})

    sheet.insert_chart('E2', chart)

# ───────────────────────────────── MAIN ───────────────────────────────
def main():
    

    f_in = ask_open_file("Select the workbook Points/Parameters")
    if not f_in: return
    pts, prm, vel = load_input(f_in)

    jour, nuit = build_implantations(pts, prm)
    from collections import OrderedDict
    OPS = OrderedDict([
        # ——— Inspection standard ———
        ("Inspection", lambda pts, prm, vel, **kw:
            build_inspection_schedule(
                pts, prm, vel,
                v_kmh      = float(prm["v_inspect_veh"]),
                freq_sem   = float(prm["inspect_freq"]),
                veh_prefix = "V",         # V1, V2, …
                col_suffix = "inspect",   # t_inspect_h
                label="Inspection", **kw       
            )
        ),

        # ——— Inspection détail ———
        ("Inspection_Detail", lambda pts, prm, vel, **kw:
            build_inspection_schedule(
                pts, prm, vel,
                v_kmh      = float(prm["v_detail_inspect_veh"]),
                freq_sem   = float(prm["detail_inspect_freq"]),
                veh_prefix = "DV",        # DV1, DV2, …
                col_suffix = "detail",   # t_detail_h
                label="Inspection_Detail", **kw         
            )
        ),

        # ——— autres opérations déjà existantes ———
        ("Grinding", lambda p, pr, vel, **kw: 
               build_grind_schedule(p, pr, vel, **kw)),
        ("Measurement", lambda p,pr,vel, **kw: build_measu_schedule(p, pr,vel,  **kw)),
        ("Ultrason",    lambda p,pr,vel, **kw: build_ultrason_schedule(p, pr,vel,  **kw)),
        ("Cleaning",    lambda p,pr,vel, **kw: build_clean_schedule(p, pr,vel,  **kw)),
    ])

    schedules  = {}   # nom → DataFrame
    warnings   = []   # messages d'erreur

    for name, func in OPS.items():
        try:
            df = func(pts, prm, vel)                # tente de planifier

        except RuntimeError as err:            # fréquence intenable
            err_msg = f"✖ {name}: {err}"
            warnings.append(err_msg)
            df = func(pts, prm, vel, enforce_limit=False)

            # 4) Si la fonction a mis un avertissement de couverture,
            #    on le propage aussi dans warnings
            if "coverage_warning" in df.attrs:
                warn_msg = "⚠ " + df.attrs["coverage_warning"]
                if warn_msg != err_msg:           # ← évite le doublon
                    warnings.append(warn_msg)

        else:
            # appel sans exception : on regarde juste coverage_warning
            if "coverage_warning" in df.attrs:
                warnings.append("⚠ " + df.attrs["coverage_warning"])

        schedules[name] = df
            
    # --------------- tableau de synthèse ------------------
    summary = []
    CAPACITE_AN = int(prm.get("capacity_nights_an", 365))

    def add_row(name, df, num_veh, freq_sem):
        nuits_cycle = df["nuit"].nunique() if not df.empty else 0
        if num_veh == 0:           # opération désactivée
            nuits_by_vehicle = cycles_an = nuits_by_vehicle_an = 0
        else:
            nuits_by_vehicle = nuits_cycle / num_veh
            cycles_an         = 52 / freq_sem
            nuits_by_vehicle_an = nuits_by_vehicle * cycles_an

        commentaire = ""
        if nuits_by_vehicle_an > CAPACITE_AN:
            commentaire = "⚠ Capacity exceeded — add one vehicle or increase the frequency"

        # test couverture linéaire (attribut éventuellement posé par build_*_schedule)
        if "coverage_warning" in df.attrs:
            commentaire = (commentaire + " ; " if commentaire else "") + df.attrs["coverage_warning"]

        coverage_ok = "Yes" if "coverage_warning" not in df.attrs else "No"

        summary.append({
            "Schedule":                name,
            "Vehicules":               num_veh,
            "Freq_sem":                freq_sem,
            "Nights_cycle":             nuits_cycle,
            "Nights/veh_cycle":         round(nuits_by_vehicle, 1),
            "Cycles_an":               round(cycles_an, 2),
            "Nights/veh_an":            round(nuits_by_vehicle_an, 1),
            "Capacity_nights_an":       CAPACITE_AN,
            "Coverage_ok":             coverage_ok,
            "Comments":             commentaire
        })

    # appeler pour chaque planning
    add_row("Inspection",          schedules["Inspection"],          prm["num_inspec_veh"],   prm["inspect_freq"])
    add_row("Detail Inspection",   schedules["Inspection_Detail"],   prm["num_inspec_veh"],   prm["detail_inspect_freq"])
    add_row("Grinding",            schedules["Grinding"],            prm["num_grind_veh"],    prm["grind_freq"])
    add_row("Measurement",         schedules["Measurement"],         prm["num_measu_veh"],    prm["measu_freq"])
    add_row("Ultrason",            schedules["Ultrason"],            prm["num_ultrason_veh"], prm["ultrason_freq"])
    add_row("Cleaning",            schedules["Cleaning"],            prm["num_clean_veh"],    prm["clean_freq"])

    df_syn = pd.DataFrame(summary)

    f_out = ask_save_file()
    if not f_out: return
    # ─── Export multi-onglets ───
    with pd.ExcelWriter(f_out, engine="xlsxwriter") as wr:
        jour .to_excel(wr, sheet_name="Day_Deployment", index=False)
        nuit .to_excel(wr, sheet_name="Night_Deployment", index=False)

        schedules["Inspection"]        .to_excel(wr, sheet_name="Inspection_Schedule",        index=False)
        schedules["Inspection_Detail"] .to_excel(wr, sheet_name="Detail_Inspection_Schedule", index=False)
        schedules["Grinding"]   .to_excel(wr, sheet_name="Grinding_Schedule",    index=False)
        schedules["Measurement"].to_excel(wr, sheet_name="Measurement_Schedule", index=False)
        schedules["Ultrason"]   .to_excel(wr, sheet_name="Ultrason_Schedule",    index=False)
        schedules["Cleaning"]   .to_excel(wr, sheet_name="Cleaning_Schedule",    index=False)
        df_syn.to_excel(wr, sheet_name="Synthesis", index=False)
        add_synoptic_sheet(pts, prm, wr)

    # ─── Résumé final ───
    def nights(df): return df["nuit"].nunique() if not df.empty else 0

    msg = (
        f"• Day_Deployment   : {len(jour)} bases\n"
        f"• Night_Deployment   : {len(nuit)} bases\n"
        f"• Inspection_Schedule  : {nights(schedules['Inspection'])} nights\n"
        f"• Detail_Inspection_Schedule : {nights(schedules['Inspection_Detail'])} nights\n"
        f"• Grinding_Schedule    : {nights(schedules['Grinding'])} nights\n"
        f"• Measurement_Schedule : {nights(schedules['Measurement'])} nights\n"
        f"• Ultrason_Schedule    : {nights(schedules['Ultrason'])} nights\n"
        f"• Cleaning_Schedule    : {nights(schedules['Cleaning'])} nights\n\n"
        f"Enregistré : {f_out}"
    )

    msg += "\n\nNights / vehicle / an :\n"
    for _, row in df_syn.iterrows():
        msg += f"- {row['Schedule']}: {row['Nights/veh_an']} nights\n"

    if warnings:
        warnings = list(dict.fromkeys(warnings))
        msg += "\n⚠︎ Impossible schedules :\n" + "\n".join(warnings)

    messagebox.showinfo("Completed", msg)

if __name__ == "__main__":
    main()
