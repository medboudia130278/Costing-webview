"""
Outil de dimensionnement des effectifs opérateurs + Génération de planning annuel.

🔄 MISE À JOUR – intégration **training_days, flexibility_days, sickness_days**
───────────────────────────────────────────────────────────────────────────
Le classeur Excel comporte désormais 4 feuilles :
  1. **Resultats**      – récapitulatif du dimensionnement.
  2. **Planning**       – roster quotidien (présence & shifts), incluant :
       • Vacances   (absent « VAC »)
       • Formation  (absent « TRAIN »)
       • Flexibilité (absent « FLEX »)
       • Maladie    (absent « SICK »)
  3. **Vacances**       – blocs de congés par opérateur.
  4. **Hours_report**   – synthèse annuelle : jours assignés, heures travaillées, jours VAC/TRAIN/FLEX/SICK.
"""

from __future__ import annotations
from collections import defaultdict
import pandas as pd
from math import ceil
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple
import random

# ---------------------------------------------------------------------------
# Conversions robustes
# ---------------------------------------------------------------------------

def to_int(value: Any, param_name: str) -> int:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise ValueError(f"The parameter '{param_name}' is missing.")
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        s = value.strip().replace(',', '.')
        if s == '':
            raise ValueError(f"The parameter '{param_name}' is empty.")
        try:
            return int(float(s))
        except ValueError:
            raise ValueError(f"The parameter '{param_name}' must be an integer (value : '{value}').")
    raise ValueError(f"The parameter '{param_name}' has an unexpected type : {type(value)}")


def to_float(value: Any, param_name: str) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise ValueError(f"The parameter '{param_name}' is missing.")
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(',', '.')
        if s == '':
            raise ValueError(f"The parameter '{param_name}' is empty.")
        try:
            return float(s)
        except ValueError:
            raise ValueError(f"The parameter '{param_name}' must be numeric (value : '{value}').")
    raise ValueError(f"The parameter '{param_name}' has an unexpected type : {type(value)}")

# ---------------------------------------------------------------------------
# Lecture des paramètres
# ---------------------------------------------------------------------------

def load_parameters() -> tuple[Dict[str, Any] | None, str | None]:
    
    file_path = filedialog.askopenfilename(
        title="Select the input file",
        filetypes=[('Excel Files', '*.xlsx *.xls'), ('All files', '*.*')],
    )
    if not file_path:
        messagebox.showinfo('Information', 'No file selected. Program will now exit.')
        return None, None
    try:
        xls = pd.ExcelFile(file_path)
        sheet_name = 'Parametres' if 'Parametres' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as exc:
        messagebox.showerror('Error', f"Unable to read the file : {exc}")
        return None, None

    nom_col    = 'Nom'    if 'Nom'    in df.columns else 'Name'
    valeur_col = 'Valeur' if 'Valeur' in df.columns else 'Value'
    if nom_col not in df.columns or valeur_col not in df.columns:
        messagebox.showerror('Format error', "The file must contain two columns : 'Nom'/'Name' and 'Valeur'/'Value'.")
        return None, None

    df[nom_col] = df[nom_col].astype(str).str.strip()
    params = pd.Series(df[valeur_col].values, index=df[nom_col]).to_dict()
    return params, file_path

def load_ops_per_shift(params: Dict[str, Any], num_shifts: int) -> list[int]:
    ops_list: list[int] = []
    for i in range(1, num_shifts + 1):
        key = f'operators_per_shift_{i}'
        if key not in params:
            raise ValueError(
                f"The parameter '{key}' is missing then number_of_shifts = {num_shifts}."
            )
        ops_list.append(to_int(params[key], key))
    return ops_list

# ---------------------------------------------------------------------------
# Calcul effectif requis
# ---------------------------------------------------------------------------

def compute_required_operators(
    params: Dict[str, Any],
    ops_per_shift_list: list[int],
    max_working_hours: float
) -> tuple[int, int, float]:
    
    working_days = to_float(params.get('working_days_per_week'), 'working_days_per_week')
    vacation_weeks = to_float(params.get('vacation_weeks_per_year'), 'vacation_weeks_per_year')
    training_days = to_int(params.get('training_days', 10), 'training_days')
    flexibility_days = to_int(params.get('flexibility_days', 5), 'flexibility_days')
    sickness_days = to_int(params.get('sickness_days', 10), 'sickness_days')
    hours_per_day = to_float(params.get('hours_per_day', 8), 'hours_per_day')

    max_hours_param = params.get('max_working_hours')

    if (
        max_hours_param is None
        or (isinstance(max_hours_param, float) and pd.isna(max_hours_param))
        or (isinstance(max_hours_param, str) and max_hours_param.strip() == '')
    ):
        max_hours_cap = float('inf')
    else:
        max_hours_cap = to_float(max_hours_param, 'max_working_hours')
    annual_required_shift_coverage = 52 * 7 * sum(ops_per_shift_list)
    annual_available_shifts_per_operator = (
        working_days * (52 - vacation_weeks) - (training_days + flexibility_days + sickness_days)
    )
    print(f"max_hours_param brut : {max_hours_param}")
    print(f"max_hours_cap after conversion : {max_hours_cap}")
    # ⏱️ Capacité limitée par les heures
    if max_hours_cap == float('inf'):
        hours_limit_shifts = float('inf')   # capacité illimitée
    else:
        hours_limit_shifts = int(max_hours_cap // hours_per_day)
    effective_capacity = min(annual_available_shifts_per_operator, hours_limit_shifts)

    if annual_available_shifts_per_operator <= 0:
        raise ValueError('Annual operator capacity is zero or negative — please check the parameters.')

    required_operators = ceil(annual_required_shift_coverage / effective_capacity)
    return required_operators, annual_required_shift_coverage, effective_capacity
    
# ---------------------------------------------------------------------------
# Absence scheduling helpers
# ---------------------------------------------------------------------------

def generate_vacation_schedule(num_ops: int, year: int, vacation_weeks: int, consecutive_weeks: int) -> Dict[int, Set[date]]:
    if vacation_weeks == 0:
        return {op: set() for op in range(num_ops)}

    blocks = vacation_weeks // consecutive_weeks
    rem = vacation_weeks % consecutive_weeks
    if rem:
        blocks += 1

    start_weeks_cycle: List[int] = list(range(0, 52, consecutive_weeks))
    slot_index = 0
    vac_schedule: Dict[int, List[int]] = {op: [] for op in range(num_ops)}

    while any(len(vac_schedule[op]) < blocks for op in range(num_ops)):
        for op in range(num_ops):
            if len(vac_schedule[op]) >= blocks:
                continue
            start_week = start_weeks_cycle[slot_index % len(start_weeks_cycle)]
            vac_schedule[op].append(start_week)
            slot_index += 1

    vacations: Dict[int, Set[date]] = {op: set() for op in range(num_ops)}
    jan1 = date(year, 1, 1)
    for op, starts in vac_schedule.items():
        for idx, week_start in enumerate(starts):
            length = consecutive_weeks if (idx < blocks - 1 or rem == 0) else rem
            for day in range(length * 7):
                d = jan1 + timedelta(days=week_start * 7 + day)
                if d.year == year:
                    vacations[op].add(d)
    return vacations


def generate_single_day_off_schedule(num_ops: int, year: int, n_days: int, seed_offset: int) -> Dict[int, Set[date]]:
    """Retourne dict op -> set(date) pour n_days isolés spacés régulièrement.
    seed_offset assure une répartition différente entre TRAIN/FLEX/SICK.
    """
    if n_days == 0:
        return {op: set() for op in range(num_ops)}
    days_dict: Dict[int, Set[date]] = {}
    for op in range(num_ops):
        random.seed(op + seed_offset)  # reproductible
        chosen = set()
        # Échantillonnage sans chevauchement sur 365 jours
        possible_days = list(range(365))
        random.shuffle(possible_days)
        for idx in possible_days:
            d = date(year, 1, 1) + timedelta(days=idx)
            if d.year == year:
                chosen.add(d)
                if len(chosen) >= n_days:
                    break
        days_dict[op] = chosen
    return days_dict


def vacation_blocks_from_dates(dates_set: Set[date]) -> List[Tuple[date, date]]:
    if not dates_set:
        return []
    sorted_dates = sorted(dates_set)
    blocks: List[Tuple[date, date]] = []
    block_start = prev = sorted_dates[0]
    for d in sorted_dates[1:]:
        if (d - prev).days == 1:
            prev = d
        else:
            blocks.append((block_start, prev))
            block_start = prev = d
    blocks.append((block_start, prev))
    return blocks

# ---------------------------------------------------------------------------
# Génération planning annuel
# ---------------------------------------------------------------------------

def daterange(start: date, end: date):
    for n in range((end - start).days + 1):
        yield start + timedelta(n)


def assign_shift(initial_shift: int, week_index: int, num_shifts: int) -> int:
    return (initial_shift + week_index) % num_shifts + 1


def build_planning_df(params: Dict[str, Any],
    year: int,
    num_shifts: int,
    ops_per_shift_list: list[int],
    required_ops: int,
    vacation_weeks: int,
    consecutive_vac_weeks: int,
    training_days: int,
    flex_days: int,
    sick_days: int,
    hours_per_day: float,
    max_working_hours: float,
) -> Tuple[pd.DataFrame, Dict[str, Dict[int, Set[date]]]]:
    """Construit le planning et garantit >= operators_per_shift par jour/shift."""

    # ---------- 1. Génère toutes les absences ---------- #
    start_date = date(year, 1, 1)
    end_date = date(year, 12, 31)
    all_dates = list(daterange(start_date, end_date))

    working_days = to_int(params.get('working_days_per_week'), 'working_days_per_week')

    vac  = generate_vacation_schedule(required_ops, year, vacation_weeks, consecutive_vac_weeks)
    train = generate_single_day_off_schedule(required_ops, year, training_days, 42)
    flex  = generate_single_day_off_schedule(required_ops, year, flex_days, 84)
    sick  = generate_single_day_off_schedule(required_ops, year, sick_days, 126)
    abs_dict_all = {'VAC': vac, 'TRAIN': train, 'FLEX': flex, 'SICK': sick}
    abs_all_sets = {op: vac[op] | train[op] | flex[op] | sick[op] for op in range(required_ops)}

    offsets     = [i % 7 for i in range(required_ops)]
    init_shifts = [i % num_shifts for i in range(required_ops)]
    last_work = {f'Op_{i+1:03d}' : None for i in range(required_ops)}
    streak = {k: 0 for k in last_work}
    assigned = {k: 0 for k in last_work}

    # ---------- 2. Affectation nominale ---------- #
    records: List[Tuple[date, int, str]] = []
    pres_map: Dict[Tuple[date, int], Set[str]] = defaultdict(set)
    for cur_date in all_dates:
        for sh in range(1, num_shifts + 1):
            key    = (cur_date, sh)
            target = ops_per_shift_list[sh - 1]

            while len(pres_map[key]) < target:
                # Trie dynamique : les - chargés d'abord
                for op in sorted(range(required_ops),
                                 key=lambda x: (assigned[f'Op_{x+1:03d}'], x)):
                    op_id = f'Op_{op+1:03d}'

                    # --- filtres d'inéligibilité ---
                    if (assigned[op_id] + 1) * hours_per_day > max_working_hours:
                        continue
                    if op_id in pres_map[key] or cur_date in abs_all_sets[op]:
                        continue
                    if any(op_id in pres_map.get((cur_date, s), set())
                           for s in range(1, num_shifts + 1)):
                        continue  # déjà affecté sur un autre shift ce jour
                    day_idx = (cur_date - start_date).days
                    if (day_idx - offsets[op]) % 7 >= working_days:
                        continue  # week‑end personnel
                    prev = last_work[op_id]
                    if prev and (cur_date - prev).days == 1 \
                             and streak[op_id] >= working_days:
                        continue  # dépasserait la limite de jours consécutifs

                    # --- affectation ---
                    pres_map[key].add(op_id)
                    records.append((cur_date, sh, op_id))
                    assigned[op_id] += 1
                    last_work[op_id] = cur_date
                    streak[op_id] = (streak[op_id] + 1
                                     if prev and (cur_date - prev).days == 1 else 1)
                    break  # on ressort du for pour tester à nouveau le while
                else:
                    # Aucun opérateur éligible : on abandonne ce shift pour ce jour
                    break
    
    uncovered_shifts = 0  # Compteur de shifts non couverts                
    # ───────── 4. Complément (roue de secours) ──────────
    # Garde la logique d'origine ; devrait très peu agir après équilibre initial
    for cur_date in all_dates:
        for sh in range(1, num_shifts + 1):
            key = (cur_date, sh)
            while len(pres_map[key]) < ops_per_shift_list[sh - 1]:
                added = False
                for op in sorted(range(required_ops),
                                 key=lambda x: (assigned[f'Op_{x+1:03d}'], x)):
                    op_id = f'Op_{op+1:03d}'
                    if (assigned[op_id] + 1) * hours_per_day > max_working_hours:
                        continue
                    if op_id in pres_map[key] or cur_date in abs_all_sets[op]:
                        continue
                    if any(op_id in pres_map.get((cur_date, s), set())
                           for s in range(1, num_shifts + 1)):
                        continue
                    day_idx = (cur_date - start_date).days
                    if (day_idx - offsets[op]) % 7 >= working_days:
                        continue
                    prev = last_work[op_id]
                    if prev and (cur_date - prev).days == 1 \
                             and streak[op_id] >= working_days:
                        continue

                    pres_map[key].add(op_id)
                    records.append((cur_date, sh, op_id))
                    assigned[op_id] += 1
                    last_work[op_id] = cur_date
                    streak[op_id] = (streak[op_id] + 1
                                     if prev and (cur_date - prev).days == 1 else 1)
                    added = True
                    break
                if not added:
                    uncovered_shifts += 1  # Incrément du compteur
                    break  # impossible de compléter, on sort

    # ───────── 5. DataFrame final ──────────
    planning_df = (pd.DataFrame(records,
                                columns=['Date', 'Shift', 'Operator_ID'])
                     .sort_values(['Date', 'Shift'])
                     .reset_index(drop=True))

    return planning_df, abs_dict_all, uncovered_shifts

# ---------------------------------------------------------------------------
# Rapport heures & jours
# ---------------------------------------------------------------------------

def build_hours_report(planning_df: pd.DataFrame, abs_dict_all: Dict[str, Dict[int, Set[date]]], required_ops: int, hours_per_day: float,) -> pd.DataFrame:
    work_days = planning_df.groupby('Operator_ID').size().rename('Assigned_Days')
    hours = work_days * hours_per_day
    hours.name = 'Hours_Worked'

    vac_days_series = pd.Series({f'Op_{op+1:03d}': len(abs_dict_all['VAC'].get(op, set())) for op in range(required_ops)}, name='Vacation_Days')
    train_days_series = pd.Series({f'Op_{op+1:03d}': len(abs_dict_all['TRAIN'].get(op, set())) for op in range(required_ops)}, name='Training_Days')
    flex_days_series = pd.Series({f'Op_{op+1:03d}': len(abs_dict_all['FLEX'].get(op, set())) for op in range(required_ops)}, name='Flexibility_Days')
    sick_days_series = pd.Series({f'Op_{op+1:03d}': len(abs_dict_all['SICK'].get(op, set())) for op in range(required_ops)}, name='Sickness_Days')

    report = (
        pd.concat([work_days, hours, vac_days_series, train_days_series, flex_days_series, sick_days_series], axis=1)
        .fillna(0)
        .astype({
            'Assigned_Days': int,
            'Hours_Worked': int,
            'Vacation_Days': int,
            'Training_Days': int,
            'Flexibility_Days': int,
            'Sickness_Days': int,
        })
        .reset_index()
        .rename(columns={'index': 'Operator_ID'})
    )
    return report

# ---------------------------------------------------------------------------
# Sauvegarde Excel multi-feuilles
# ---------------------------------------------------------------------------

def save_to_excel(
    output_path: Path,
    summary_df: pd.DataFrame,
    planning_df: pd.DataFrame,
    vacation_df: pd.DataFrame,
    hours_df: pd.DataFrame,
) -> None:
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='Results', index=False)
            planning_df.to_excel(writer, sheet_name='Planning', index=False)
            vacation_df.to_excel(writer, sheet_name='Vacation', index=False)
            hours_df.to_excel(writer, sheet_name='Hours_report', index=False)
        messagebox.showinfo('Success', f"File saved :\n{output_path.resolve()}")
    except Exception as exc:
        messagebox.showerror('Error', f"Unable to save the file : {exc}")

# ---------------------------------------------------------------------------
# Flux principal
# ---------------------------------------------------------------------------

def main() -> None:
    params, _ = load_parameters()
    if params is None:
        return

    try:
        year = to_int(params.get('Year'), 'Year')
        num_shifts = to_int(params.get('number_of_shifts'), 'number_of_shifts')
        ops_per_shift_list = load_ops_per_shift(params, num_shifts)
        
        # Limite d'heures par opérateur (∞ si absent)
        max_hours_param = params.get('max_working_hours')

        # Test robuste
        if (
            max_hours_param is not None
            and not (isinstance(max_hours_param, float) and pd.isna(max_hours_param))
            and str(max_hours_param).strip() != ''
        ):
            max_working_hours = to_float(max_hours_param, 'max_working_hours')
        else:
            max_working_hours = float('inf')

        vacation_weeks = to_int(params.get('vacation_weeks_per_year', 0), 'vacation_weeks_per_year')
        consecutive_vac_weeks = to_int(params.get('consecutive_vacation_weeks', 1), 'consecutive_vacation_weeks')
        training_days = to_int(params.get('training_days', 10), 'training_days')
        flex_days = to_int(params.get('flexibility_days', 10), 'flexibility_days')
        sick_days = to_int(params.get('sickness_days', 10), 'sickness_days')
        hours_per_day = to_float(params.get('hours_per_day', 8), 'hours_per_day')
    except ValueError as e:
        messagebox.showerror('Error', str(e))
        return

    try:
        required_ops_calc, coverage, capacity = compute_required_operators(params, ops_per_shift_list, max_working_hours)

    except ValueError as e:
        messagebox.showerror('Calculation error', str(e))
        return

    try:
        total_ops_param = params.get('total_operators')   # champ facultatif
        if total_ops_param is not None and str(total_ops_param).strip() != '':
            required_ops = to_int(total_ops_param, 'total_operators')
        else:
            required_ops = required_ops_calc
    except ValueError as e:
        messagebox.showerror('Error', str(e))
        return

    # message d'alerte si l'utilisateur en demande moins que le calcul
    if required_ops < required_ops_calc:
        messagebox.showwarning(
            'Avertissement',
            f'The total_operators ({required_ops}) is lower than the theoretical requirement '
            f'({required_ops_calc}). The schedule may be understaffed.')

    planning_df, abs_dict_all, uncovered_shifts = build_planning_df(params,
        year,
        num_shifts,
        ops_per_shift_list,
        required_ops,
        vacation_weeks,
        consecutive_vac_weeks,
        training_days,
        flex_days,
        sick_days,
        hours_per_day,
        max_working_hours,
    )

    hours_df = build_hours_report(planning_df, abs_dict_all, required_ops, hours_per_day)

    # Vacances – blocs pour feuille Vacances uniquement
    vacation_records: List[Tuple[str, date, date, int]] = []
    for op_id, dates_set in abs_dict_all['VAC'].items():
        for start, end in vacation_blocks_from_dates(dates_set):
            vacation_records.append((f'Op_{op_id+1:03d}', start, end, (end - start).days + 1))
    vacation_df = pd.DataFrame(vacation_records, columns=['Operator_ID', 'Start_Date', 'End_Date', 'Days'])

    summary_df = pd.DataFrame({
        'Calculation date': [datetime.now().strftime('%Y-%m-%d %H:%M')],
        'required operators': [required_ops_calc],
        'Coverage (shifts/an)': [coverage],
        'Operator capacity (shifts/an)': [capacity],
        'Vacation in weeks/op': [vacation_weeks],
        'Consecutive blocks': [consecutive_vac_weeks],
        'Training_Days/op': [training_days],
        'Flexibility_Days/op': [flex_days],
        'Sickness_Days/op': [sick_days],
        'Max_hours/op': [None if max_working_hours == float("inf") else max_working_hours],
        'Shifts non covered': [uncovered_shifts],
    })

    for i, v in enumerate(ops_per_shift_list, 1):
        summary_df[f'Op/Shift_{i}'] = [v]

    
    default_name = f'planning_operators_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    file_path = filedialog.asksaveasfilename(
        defaultextension='.xlsx',
        initialfile=default_name,
        filetypes=[('Excel Files', '*.xlsx'), ('All files', '*.*')],
        title='Save the output file',
    )
    if not file_path:
        messagebox.showinfo('Information', 'No output file selected. Program will now exit.')
        return

    save_to_excel(Path(file_path), summary_df, planning_df, vacation_df, hours_df)


if __name__ == '__main__':
    main()
