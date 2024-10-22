#This code returns a table of datas in which for every base counts how many aircraft flew in that base during the year.

# Prima cosa leggo i dati dall'excel

from typing import List, Any
import os
import pandas as pd
import numpy as np

aircraft_output_dict = {}

simulation_path = "C:/Users/Utente/Desktop/Tesi/simulations/NEWDELAY/pafam_optimization_results_anno5_bayes_NEWDELAY.xlsx"
df_aircraft_base_position = pd.read_excel(simulation_path, sheet_name = 'aircraft_base_position')
print(df_aircraft_base_position.head())

#Salvo sul dizionario una lista con gli elementi estratti dalla colonna 'aircraft' del dataframe 
aircraft_output_dict['aircraft'] = df_aircraft_base_position['aircraft'].tolist()
#print(aircraft_output_dict)

# La prima colonna è quella degli aerei
aircraft = df_aircraft_base_position['aircraft'].values
print(aircraft)

# Crea un DataFrame vuoto per i risultati
months = df_aircraft_base_position.columns[1:]  # Tutti i mesi
bases = df_aircraft_base_position.values[:, 1:]  # Tutte le basi
risultati = pd.DataFrame(columns=['Base'] + list(months))

# Raccogli le basi uniche
basi_uniche = sorted(set(bases.flatten()))
print(basi_uniche)

# Inizializza una lista per i risultati
result_list = []

# Inizializzo il DataFrame dei risultati con le basi uniche
for base in basi_uniche:
    counter = {'Base': base}
        
    # Conto gli aerei per ogni mese
    for month in months:
        counter[month] = (df_aircraft_base_position[month] == base).sum()

    # Aggiungo il dizionario alla lista dei risultati
    result_list.append(counter)

# Converto la lista di dizionari in un DataFrame
risultati = pd.DataFrame(result_list)

# Mostra i risultati
print(risultati)        

# Salvo i risultati in un nuovo file Excel
risultati.to_excel('pafam_optimization_results_anno5_bayes_NEWDELAY_aerei_per_base.xlsx', sheet_name= 'aircraft_to_base', index=False)