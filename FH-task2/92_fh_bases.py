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

    # List of all the bases and months in the excel
    months = df_fh.columns[1:]  # Tutti i mesi tranne il primo
    bases = df_aircraft_base_position.values[:, 1:]  # Tutte le basi, saltando la prima colonna

    # Group the unique bases of the sheet aircraft_base_position
    basi_uniche = sorted(set(bases.flatten()))

    # Dictionary that contains all the Dataframe for each base
    dizionario_output = {}
    # Dizionario per memorizzare i valori FH final degli aerei che cambiano base
    fh_final_dict = {}

    for base in basi_uniche:
        # Create for each base a list of the aircrafts that flown in at least once during the year
        aircrafts_in_base = df_aircraft_base_position[df_aircraft_base_position.eq(base).any(axis=1)].index.tolist()

        # Define additional columns to be added
        additional_columns_before = ['FH init', 'Configuration']
        additional_columns_after = ['FH flown', 'FH final', 'Base from', 'Base to']
        months_to_include = df_fh.columns[:12]  # Limitiamo l'elaborazione ai primi 12 mesi

        # Combine columns: before, original, after
        all_columns = additional_columns_before + months_to_include.tolist() + additional_columns_after

        # Create an empty dataframe for each base
        output_df = pd.DataFrame(index=aircrafts_in_base, columns=all_columns)
        output_df.index.name = 'aircraft'

        # Initialize the additional columns
        output_df[additional_columns_before] = 0
        output_df[additional_columns_after] = 0

        # Populate the dataframe with the data of the FH sheet for each base
        for aircraft in aircrafts_in_base:
            if aircraft in df_fh.index:
                # Prendi solo i dati dei primi 12 mesi
                fh_row = df_fh.loc[aircraft, months_to_include]

                output_df.loc[aircraft, 'FH init'] = 0  # Valore iniziale predefinito
                # Assegna FH init dai valori del dizionario, se l'aereo ha cambiato base durante l'anno
                if aircraft in fh_final_dict:
                    output_df.loc[aircraft, 'FH init'] = fh_final_dict[aircraft]
                else:
                    output_df.loc[aircraft, 'FH init'] = aircraft_to_fh[aircraft]

                # Configurazione dell'aereo
                output_df.loc[aircraft, 'Configuration'] = extract_letters(aircraft)

                # Variabili di transizione
                prev_base = None
                stayed_in_one_base = True  # Assumiamo che l'aereo non cambi base

                for month in months_to_include:
                    current_base = df_aircraft_base_position.loc[aircraft, month]

                    if month in df_aircraft_base_position.columns:  # check that the month appears in both sheet
                        if df_aircraft_base_position.loc[aircraft, month] == base:
                            # If the aircraft flown in the corrispondig base, add to the dataframe the corrisponding row of the fh sheet
                            output_df.loc[aircraft, month] = fh_row[month]
                        else:
                            # If aircraft didn't fly on that base we add a "-"
                            output_df.loc[aircraft, month] = "-"

                    if current_base == base:  # Se l'aereo è nella base corrente
                        if prev_base and prev_base != base:
                            # Cambio di base: aggiorna Base from
                            output_df.loc[aircraft, 'Base from'] = prev_base
                        elif prev_base is None:
                            # L'aereo parte da questa base
                            output_df.loc[aircraft, 'Base from'] = "-"

                        # Se siamo all'ultimo mese dell'anno
                        if month == months_to_include[-1]:
                            output_df.loc[aircraft, 'Base to'] = "-"

                    elif prev_base == base:
                        # L'aereo lascia la base
                        output_df.loc[aircraft, 'Base to'] = current_base
                        stayed_in_one_base = False

                        # Calcola FH flown e FH final
                        fh_flown = pd.to_numeric(output_df.loc[aircraft, months_to_include], errors='coerce').sum()
                        output_df.loc[aircraft, 'FH flown'] = fh_flown
                        fh_final = fh_flown + output_df.loc[aircraft, 'FH init']
                        output_df.loc[aircraft, 'FH final'] = fh_final

                        # Memorizza FH final nel dizionario
                        fh_final_dict[aircraft] = fh_final

                    prev_base = current_base

                # Se l'aereo è rimasto nella stessa base per tutto l'anno
                if stayed_in_one_base:
                    fh_flown = pd.to_numeric(output_df.loc[aircraft, months_to_include], errors='coerce').sum()
                    output_df.loc[aircraft, 'FH flown'] = fh_flown
                    fh_final = fh_flown + output_df.loc[aircraft, 'FH init']
                    output_df.loc[aircraft, 'FH final'] = fh_final

                # Dopo aver popolato la riga con i dati per tutti i mesi
                # Controlla se tutti i mesi nella riga contengono solo '-'
                if all(output_df.loc[aircraft, months_to_include] == "-"):
                    output_df.drop(index=aircraft, inplace=True)
                    print(
                        f"L'aereo {aircraft} ha solo '-' nei primi 12 mesi, rimosso dal foglio della base {base}.")
                    continue

            else:
                print(f"Attenzione: {aircraft} non trovato in df_fh")

    # if 'Gen.1' in df_fh.columns:
    #     df_fh = df_fh.drop(columns=['Gen.1'])
    #
    # # List of all the bases and months in the excel
    # months = df_fh.columns[1:]
    # bases = df_aircraft_base_position.values[:, 1:]
    #
    # # Group the unique bases of the sheet aircraft_base_position
    # basi_uniche = sorted(set(bases.flatten()))
    #
    # # Dictionary that contains all the Dataframe for each base
    # dizionario_output = {}
    #
    # for base in basi_uniche:
    #     # Create for each base a list of the aircrafts that flown in at least once during the year
    #     aircrafts_in_base = df_aircraft_base_position[df_aircraft_base_position.eq(base).any(axis=1)].index.tolist()
    #
    #     # Define additional columns to be added
    #     additional_columns_before = ['FH init', 'Configuration']
    #     additional_columns_after = ['FH flown', 'FH final', 'Base from', 'Base to']
    #
    #     # Combine columns: before, original, after
    #     all_columns = additional_columns_before + df_fh.columns.tolist() + additional_columns_after
    #
    #
    #     # Create an empty dataframe for each base
    #     output_df = pd.DataFrame(index=aircrafts_in_base, columns=all_columns)
    #     output_df.index.name = 'aircraft'
    #
    #     # Initialize the additional columns
    #     output_df[additional_columns_before] = 0
    #     output_df[additional_columns_after] = 0
    #
    #     # Populate the dataframe with the data of the FH sheet for each base
    #     for aircraft in aircrafts_in_base:
    #         if aircraft in df_fh.index:
    #             # Logic to populate extra col
    #             if aircraft in df_fh_init.index:
    #                 output_df.loc[aircraft, 'FH init'] = df_fh_init.loc[aircraft, 'initial_fh']
    #             else:
    #                 output_df.loc[aircraft, 'FH init'] = 0
    #
    #             output_df.loc[aircraft, 'Configuration'] = aircraft[:2]
    #
    #             # Extract the corresponding row of the sheet
    #             fh_row = df_fh.loc[aircraft]
    #
    #             # Variabili per tracciare transizioni
    #             prev_base = None  # Base precedente
    #             base_from_set = False  # Flag per indicare se "Base from" è già stato riempito
    #             base_to_set = False  # Flag per indicare se "Base to" è già stato riempito
    #             stayed_in_one_base = True  # Assumiamo che l'aereo non cambi base finché non troviamo una transizione
    #
    #             # This is a cycle that iterates for each month in order to check if the aircraft flown in the base of interest
    #             for month in df_fh.columns:
    #
    #                 current_base = df_aircraft_base_position.loc[aircraft, month]  # Base in cui si trova l'aereo nel mese attuale
    #
    #                 if current_base == base:
    #                     # If aircraft doesnt come from another base, set "Base from"
    #                     if prev_base is None and not base_from_set:
    #                         output_df.loc[aircraft, 'Base from'] = "-"
    #                         base_from_set = True
    #
    #                     # If aircraft doesnt come from another base, set "Base from"
    #                     elif prev_base and prev_base != base and not base_from_set:
    #                         output_df.loc[aircraft, 'Base from'] = prev_base
    #                         base_from_set = True
    #                         stayed_in_one_base = False  # base changed flag
    #
    #                     prev_base = base  # update the previous base
    #
    #                     # If aircraft end the year in this base, "Base to"  "-"
    #                     if month == df_fh.columns[-1]:
    #                         output_df.loc[aircraft, 'Base to'] = "-"
    #
    #                 # If the aircraft changed base during the year
    #                 elif prev_base == base and not base_to_set:
    #                     output_df.loc[aircraft, 'Base to'] = current_base
    #                     base_to_set = True
    #                     stayed_in_one_base = False  # base changed flag
    #
    #                 prev_base = current_base  # update the previous base
    #
    #                 # If the aircraft stayed in the same base for the whole year
    #                 if stayed_in_one_base:
    #                     output_df.loc[aircraft, 'Base from'] = "-"
    #                     output_df.loc[aircraft, 'Base to'] = "-"
    #
    #                 if month in df_aircraft_base_position.columns:  # check that the month appears in both sheet
    #                     if df_aircraft_base_position.loc[aircraft, month] == base:
    #                         # If the aircraft flown in the corrispondig base, add to the dataframe the corrisponding row of the fh sheet
    #                         output_df.loc[aircraft, month] = fh_row[month]
    #                     else:
    #                         # If aircraft didn't fly on that base we add a "-"
    #                         output_df.loc[aircraft, month] = "-"
    #
    #                 # Compute the total flown hours in the base
    #                 fh_flown = pd.to_numeric(output_df.loc[aircraft, df_fh.columns], errors='coerce').sum()  # `errors='coerce'` trasforma valori non numerici in NaN
    #                 output_df.loc[aircraft, 'FH flown'] = fh_flown
    #
    #                 # Compute the final FH in the that base
    #                 fh_final = fh_flown + output_df.loc[aircraft, 'FH init']
    #                 output_df.loc[aircraft, 'FH final'] = fh_final
    #
    #         else:
    #             print(f"Attenzione: {aircraft} non trovato in df_fh")
            
           
        # Add the DataFrame to the dictionary
        dizionario_output[base] = output_df

    # Create an Excel output file with all the sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for base, df in dizionario_output.items():
            nome_foglio = f"{base}_FH"  # Rename the sheet
            df.to_excel(writer, sheet_name=nome_foglio, index=True, header=True)

    # Remove the bold
    # wb = load_workbook(output_file)
    # for sheet_name in wb.sheetnames:
    #     ws = wb[sheet_name]
    #     for cell in ws[1]:  # Prima riga (header)
    #         cell.font = Font(bold=False)
    #     for row in ws.iter_rows(min_row=2):  # Colonna A (index)
    #         row[0].font = Font(bold=False)

    #wb.save(output_file)
    print(f"File Excel in {output_file}")



###################################################################

sim_path = 'C:/Users/Utente/Desktop/Tesi/file/dev_Alessio/dev_Alessio/files/outputs_mixed_fleet/task2/task 2 - y1/pafam_optimization_results_2024_11_13_16_27.xlsx'
output_file = 'FH_bases_robbb13_pafam_optimization_results.xlsx'
input_file = 'C:/Users/Utente/Desktop/Tesi/file/dev_Alessio/dev_Alessio/files/input_data_mixed_fleet/aircrafts/fleet_25_11baie6f.xlsx'
generate_excel_fh_bases(sim_path, input_file, output_file)