#This code returns 4 sheets in which for every base, the amount of fleet hours is counted.
from typing import List, Any
from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import pandas as pd
import numpy as np

#####################################################################


def generate_excel_fh_bases(sim_path, input_file, output_file):
    df_aircraft_base_position = pd.read_excel(sim_path, sheet_name='aircraft_base_position', index_col=0)
    df_fh = pd.read_excel(sim_path, sheet_name='FH', index_col=0)
    df_fh_init = pd.read_excel(input_file, sheet_name='aircrafts', index_col=0)

    if 'initial_fh' not in df_fh_init.columns:
        raise ValueError("Col 'initial_fh' not in the input file!")

    if 'Total' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Total'])

    if 'Totale AC' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Totale AC'])

    if 'Totale AT' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Totale AT'])

    if 'Gen.1' in df_fh.columns:
        df_fh = df_fh.drop(columns=['Gen.1'])     

    # List of all the bases and months in the excel
    months = df_fh.columns[1:]  # Tutti i mesi tranne il primo
    bases = df_aircraft_base_position.values[:, 1:]  # Tutte le basi, saltando la prima colonna

    # Group the unique bases of the sheet aircraft_base_position
    basi_uniche = sorted(set(bases.flatten()))

    # Dictionary that contains all the Dataframe for each base
    dizionario_output = {}

    for base in basi_uniche:
        # Create for each base a list of the aircrafts that flown in at least once during the year
        aircrafts_in_base = df_aircraft_base_position[df_aircraft_base_position.eq(base).any(axis=1)].index.tolist()

        # Define additional columns to be added
        additional_columns_before = ['FH init', 'Configuration'] 
        additional_columns_after = ['FH new y', 'FH flown', 'Base from', 'Base to'] 

        # Combine columns: before, original, after
        all_columns = additional_columns_before + df_fh.columns.tolist() + additional_columns_after 


        # Create an empty dataframe for each base
        output_df = pd.DataFrame(index=aircrafts_in_base, columns=all_columns)
        output_df.index.name = 'aircraft'

        # Initialize the additional columns
        output_df[additional_columns_before] = 0
        output_df[additional_columns_after] = 0

        # Populate the dataframe with the data of the FH sheet for each base
        for aircraft in aircrafts_in_base:
            if aircraft in df_fh.index:
                # Logic to populate extra col
                if aircraft in df_fh_init.index:  # Assicurati che gli aerei siano allineati
                    output_df.loc[aircraft, 'FH init'] = df_fh_init.loc[aircraft, 'initial_fh']
                else:
                    output_df.loc[aircraft, 'FH init'] = 0  # Valore di default se non trovato

                output_df.loc[aircraft, 'Configuration'] = aircraft[:2]

                # Custom logic to populate 'extra_col3', ..., 'extra_col6' (after the months)
                # output_df.loc[aircraft, 'FH new y'] = some_custom_logic3(aircraft, base) #TODO capire come fare
                # output_df.loc[aircraft, 'Base from'] = some_custom_logic5(aircraft, base)
                # output_df.loc[aircraft, 'Base to'] = some_custom_logic6(aircraft, base)

                # Extract the corresponding row of the sheet
                fh_row = df_fh.loc[aircraft]
                # This is a cycle that iterates for each month in order to check if the aircraft flown in the base of interest    
                for month in df_fh.columns:
                    if month in df_aircraft_base_position.columns:  # check that the month appears in both sheet
                        if df_aircraft_base_position.loc[aircraft, month] == base:
                            # If the aircraft flown in the corrispondig base, add to the dataframe the corrisponding row of the fh sheet
                            output_df.loc[aircraft, month] = fh_row[month]
                        else:
                            # If aircraft didn't fly on that base we add a "-"
                            output_df.loc[aircraft, month] = "-"

                fh_flown = pd.to_numeric(output_df.loc[aircraft, df_fh.columns], errors='coerce').sum()  # `errors='coerce'` trasforma valori non numerici in NaN
                output_df.loc[aircraft, 'FH flown'] = fh_flown  
                         
            else:
                print(f"Attenzione: {aircraft} non trovato in df_fh")

        
        # Add the DataFrame to the dictionary
        dizionario_output[base] = output_df

    # Create an Excel output file with all the sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for base, df in dizionario_output.items():
            nome_foglio = f"{base}_FH"  # Rename the sheet
            df.to_excel(writer, sheet_name=nome_foglio, index=True, header=True)

    # Remove the bold
    wb = load_workbook(output_file)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for cell in ws[1]:  # Prima riga (header)
            cell.font = Font(bold=False)
        for row in ws.iter_rows(min_row=2):  # Colonna A (index)
            row[0].font = Font(bold=False)

    wb.save(output_file)
    print(f"File Excel in {output_file}")

###################################################################

sim_path = "C:/Users/Utente/Desktop/Tesi/file/dev_Alessio/dev_Alessio/files/outputs_mixed_fleet/task 2 - y1/pafam_optimization_results_2024_11_13_16_27.xlsx"
output_file = 'FH_bases_pafam_optimization_results_2024_11_13_16_27.xlsx'
input_file = 'C:/Users/Utente/Desktop/Tesi/file/dev_Alessio/dev_Alessio/files/input_data_mixed_fleet/aircrafts/fleet_25_11baie6f.xlsx'
generate_excel_fh_bases(sim_path, input_file, output_file)