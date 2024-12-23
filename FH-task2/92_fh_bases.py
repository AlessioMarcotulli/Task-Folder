#This code returns 4 sheets in which for every base, the amount of fleet hours is counted.
# standard libraries
# from src.tools import clear_directory
import re
from typing import List, Any
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pathlib import Path
import datetime
import shutil
from copy import deepcopy
from scipy.interpolate import interp1d
#####################################################################

# Funzione per estrarre solo le lettere iniziali
def extract_letters(aircraft):
    match = re.match(r'[A-Za-z]+', aircraft)
    return match.group(0) if match else ''

def generate_excel_fh_bases(sim_path, input_file, output_file):
    df_aircraft_base_position = pd.read_excel(sim_path, sheet_name='aircraft_base_position', index_col=0)
    df_fh = pd.read_excel(sim_path, sheet_name='FH', index_col=0)
    df_fh_init = pd.read_excel(input_file, sheet_name='aircrafts', index_col=0)

    seed = 2024

    # Create dict with zip()
    # aircraft_list = df_aircraft_base_position.index.tolist()
    # aircraft_to_fh = dict(zip(aircraft_list, input_file))
    # print(aircraft_to_fh)

    # Estrai la colonna desiderata (ad esempio 'FH') in un dizionario
    aircraft_to_fh = df_fh_init['initial_fh'].to_dict()

    # Debug: stampa il dizionario
    print("Dizionario aircraft_to_fh:", aircraft_to_fh)

    if 'Total' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Total'])

    if 'Totale AC' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Totale AC'])

    if 'Totale AT' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Totale AT'])

    # Lista dei mesi e basi nel file Excel
    months = df_fh.columns[:12]  # Primi 12 mesi
    bases = df_aircraft_base_position.values[:, :12]  # Basi nei primi 12 mesi
    basi_uniche = sorted(set(bases.flatten()))  # Basi uniche ordinate

    # Inizializzazione dei dizionari
    dizionario_output = {}  # Dizionario con i dati di output per ogni base

    # Iterazione sugli aerei
    for aircraft in df_aircraft_base_position.index:
        aircraft_data = pd.Series(
            index=['FH init', 'Configuration'] + months.tolist() + ['FH flown', 'FH final', 'Base from', 'Base to'],
            dtype="object"
        )
        stayed_in_one_base = True
        prev_base = None
        last_recorded_base = None  # Traccia la base di provenienza effettiva

        # Inizializza i dati iniziali dell'aereo
        aircraft_data['FH init'] = aircraft_to_fh.get(aircraft, 0)  # Inizializza come numero
        aircraft_data['Configuration'] = extract_letters(aircraft)
        aircraft_data[months] = "-"

        for month in months:
            current_base = df_aircraft_base_position.loc[aircraft, month]
            fh_value = (
                df_fh.loc[aircraft, month]
                if aircraft in df_fh.index and month in df_fh.columns
                else "-"
            )

            # Se è il primo mese, imposta la base iniziale
            if prev_base is None:
                prev_base = current_base
                last_recorded_base = current_base  # Imposta la base iniziale come base di provenienza

            # Se la base cambia
            if current_base != prev_base:
                stayed_in_one_base = False

                # Completa i dati per la base precedente
                fh_flown = pd.to_numeric(aircraft_data[months], errors="coerce").sum()
                aircraft_data['FH flown'] = fh_flown

                # Gestisci FH init se ci sono più valori
                if isinstance(aircraft_data['FH init'], str):
                    # Se FH init è una stringa (cioè contiene più valori), dividiamo e sommiamo
                    fh_init_values = [int(x) for x in aircraft_data['FH init'].split(', ')]
                    fh_init_numeric = sum(fh_init_values)
                else:
                    # Se FH init è un singolo valore numerico, lo trattiamo come tale
                    fh_init_numeric = aircraft_data['FH init']

                aircraft_data['FH final'] = int(fh_init_numeric + fh_flown)
                aircraft_data['Base from'] = "-" if last_recorded_base == prev_base else last_recorded_base
                aircraft_data['Base to'] = current_base

                # Salva i dati nella base precedente, gestendo la concatenazione di FH init e FH final
                if prev_base in dizionario_output:
                    if aircraft in dizionario_output[prev_base].index:
                        # Se l'aereo è già presente, concatenare i nuovi valori di FH init e FH final
                        existing_data = dizionario_output[prev_base].loc[aircraft].copy()
                        existing_data['FH init'] = f"{existing_data['FH init']}, {aircraft_data['FH init']}"
                        existing_data['FH final'] = f"{existing_data['FH final']}, {aircraft_data['FH final']}"
                        for month_to_update in months:
                            if aircraft_data[month_to_update] != "-":
                                existing_data[month_to_update] = aircraft_data[month_to_update]
                        dizionario_output[prev_base].loc[aircraft] = existing_data
                    else:
                        dizionario_output[prev_base].loc[aircraft] = aircraft_data
                else:
                    dizionario_output[prev_base] = pd.DataFrame(columns=aircraft_data.index)
                    dizionario_output[prev_base].loc[aircraft] = aircraft_data

                # Ripristina per la nuova base
                last_recorded_base = prev_base  # Aggiorna la base di provenienza
                aircraft_data['FH init'] = aircraft_data['FH final']  # Aggiungi il valore finale al tracking di FH init
                aircraft_data[months] = "-"  # Resetta i valori mensili
                prev_base = current_base

            # Aggiorna il valore del mese corrente
            if current_base == prev_base:
                aircraft_data[month] = fh_value

        # Finalizza i dati per l'ultima base
        fh_flown = pd.to_numeric(aircraft_data[months], errors="coerce").sum()
        aircraft_data['FH flown'] = fh_flown

        # Gestisci FH init per l'ultima base
        if isinstance(aircraft_data['FH init'], str):
            fh_init_values = [int(x) for x in aircraft_data['FH init'].split(', ')]
            fh_init_numeric = sum(fh_init_values)
        else:
            fh_init_numeric = aircraft_data['FH init']

        aircraft_data['FH final'] = int(fh_init_numeric + fh_flown)
        aircraft_data['Base from'] = "-" if stayed_in_one_base else last_recorded_base
        aircraft_data['Base to'] = "-"

        # Salva i dati nella base finale
        if prev_base in dizionario_output:
            if aircraft in dizionario_output[prev_base].index:
                # Se l'aereo è già presente, concatenare i nuovi valori di FH init e FH final
                existing_data = dizionario_output[prev_base].loc[aircraft].copy()
                existing_data['FH init'] = f"{existing_data['FH init']}, {aircraft_data['FH init']}"
                existing_data['FH final'] = f"{existing_data['FH final']}, {aircraft_data['FH final']}"
                for month_to_update in months:
                    if aircraft_data[month_to_update] != "-":
                        existing_data[month_to_update] = aircraft_data[month_to_update]
                dizionario_output[prev_base].loc[aircraft] = existing_data
            else:
                dizionario_output[prev_base].loc[aircraft] = aircraft_data
        else:
            dizionario_output[prev_base] = pd.DataFrame(columns=aircraft_data.index)
            dizionario_output[prev_base].loc[aircraft] = aircraft_data

        #Generate excel
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            for base in sorted(dizionario_output.keys()):
                df = dizionario_output[base]
                df.index.name = "Aircrafts"
                df.to_excel(writer, sheet_name=f"{base}_FH", startrow=2, index=True)

                # Ottenere il workbook e il worksheet
                workbook = writer.book
                worksheet = writer.sheets[f"{base}_FH"]

                # Scrittura del titolo
                title = f"{base.replace('_', ' ').upper()}"
                worksheet.merge_range('G1:M1', title, workbook.add_format({
                    'bold': True,
                    'font_size': 24,
                    'align': 'center',
                    'valign': 'vcenter'
                }))

                anno = f"{seed}"
                worksheet.merge_range('A2:B2', anno, workbook.add_format({
                    'bold': True,
                    'font_size': 16,
                    'align': 'center',
                    'valign': 'vcenter'
                }))

                # Calcolare la somma totale per ogni configurazione
                config_sum = df.groupby('Configuration')['FH flown'].sum()
                bold_format = workbook.add_format({'bold': True})
                # Trova la colonna "fh flown"
                fh_flown_col = df.columns.get_loc("FH flown")  # Trova l'indice della colonna "fh flown"

                last_row = len(df) + 2  # Calcolare dove termina la tabella
                sum_row = last_row + 2  # Sposta la somma di due righe più in basso

                # Scrivere la somma totale per "fh flown"
                sum_formula = f"=SUM({chr(65 + fh_flown_col+1)}3:{chr(65 + fh_flown_col+1)}{last_row+1})"
                worksheet.write(sum_row, fh_flown_col+1, sum_formula, bold_format)  # Scrive la somma in grassetto
                worksheet.write(sum_row, fh_flown_col + 3, "Total FH flown in this base", bold_format)

                # Aggiungere la somma per ogni configurazione subito sotto la somma totale
                config_row_start = sum_row + 1  # Aggiungi la somma configurazioni direttamente sotto la somma totale
                for idx, (config, total_fh) in enumerate(config_sum.items()):
                    worksheet.write(config_row_start + idx, fh_flown_col+1, total_fh,
                                    bold_format)  # Scrive la somma delle ore di volo per configurazione
                    worksheet.write(config_row_start + idx, fh_flown_col + 3, f"Total FH flown {config}",
                                    bold_format)  # Scrive il testo "Total FH flown configX"

    print(f"File Excel in {output_file}")



###################################################################

sim_path = 'C:/Users/Utente/Desktop/Tesi/file/Task/FH-task2/pafam_optimization_results_2024_11_13_16_27.xlsx'
output_file = 'FH_bases_ultimateBASE1_pafam_optimization_results.xlsx'
input_file = 'C:/Users/Utente/Desktop/Tesi/file/Task/FH-task2/fleet_25_11baie6f.xlsx'
generate_excel_fh_bases(sim_path, input_file, output_file)