import math
import pandas as pd
import matplotlib as plt
import os as os
from datetime import datetime
import re
import openpyxl
from openpyxl import load_workbook


"""
eLogbook

multi-file analysis
Count number of macroactivities
calculate length
calculate mean, median, min, max, IDs of min/max

"""


def time_difference(timestamp1, timestamp2):
    # Convert the timestamps to datetime objects
    date1 = datetime.strptime(timestamp1, '%Y-%m-%d %H:%M:%S')
    date2 = datetime.strptime(timestamp2, '%Y-%m-%d %H:%M:%S') 
    
    # # Calculate the difference between the two dates
    diff = date2 - date1
    
    return (diff.total_seconds() / 3600,str(diff))


folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\drive-download-20230207T104056Z-001"
sheetname = "Observations - Anomalies"
header = 1

files = os.listdir(folder_path)

# list_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\Checklists"
# lists = os.listdir(list_path)
# checklists = {}
# for file in lists:
#     listfile_path = os.path.join(list_path, file)
#     if file.endswith('.csv'):
#         cl = pd.read_csv(listfile_path, names = ["Macroactivity"])
#         checklists[file] = cl

dfs = {}
df_filtered = {}
length_df = {}

for file in files:
    file_path = os.path.join(folder_path, file)
    if file.endswith('.xls'):
        df = pd.read_excel(file_path,sheetname,header = header)
        # df = df.replace(r'\n(?!\s)', ' ', regex=True).replace(r'\n', '', regex=True)
        
        df = df.applymap(lambda x: re.sub(r'(\s)\n', r'\1', x) if isinstance(x, str) else x)
        df = df.applymap(lambda x: re.sub(r'\n(\s)', r' \1', x) if isinstance(x, str) else x)
        df = df.replace(r'\r\n|\r|\n', ' ', regex=True)
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].str.rstrip()
        
        
        "match the keyactivities"       
        
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 02 POINT FIXE AVANT 1ER VOL PFA', 'POSTE 02 PFA (Point Fixe Avant 1er vol)')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 02 TRAVAUX SYSTEMATIQUES après 1er point fixe', 'POSTE 02 TRAVAUX SYSTEMATIQUES APRES 1er POINT FIXE')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 02 VAV (VISITE AVANT VOL)', 'POSTE 02 VAV (Visite avant vol)')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 03 VOL T1','POSTE 03 VOL T1 (K1)')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 06 MAINTENACE VAL', 'POSTE 06 MAINTENANCE VAL')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 06 VAH', 'POSTE 06 VAH (Visite Avant Habillage)')
        df["Travail à faire - Décision"] = df["Travail à faire - Décision"].replace('POSTE 09 Présentation machine services officiels / TRANSFERT LH', 'POSTE 09 PRESENTATION MACHINE SERVICES OFFICIELS / TRANSFERT LH')
        
        ""
        
        
        
        dfs[file] = df
        dfs[file]['Date'] = pd.to_datetime(dfs[file]['Date'])
        dfs[file]['Date'] = dfs[file]['Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
        dfs[file]['Fait (Date)'] = pd.to_datetime(dfs[file]['Fait (Date)'])
        dfs[file]['Fait (Date)'] = dfs[file]['Fait (Date)'].dt.strftime('%Y-%m-%d %H:%M:%S')
        dfs[file]['Unnamed: 15'] = pd.to_datetime(dfs[file]['Unnamed: 15'])
        dfs[file]['Unnamed: 15'] = dfs[file]['Unnamed: 15'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        is_nan = dfs[file][['Date', 'Auteur','N° de tampon','Réponse','Date.1','Auteur.1','N° de tampon.1']].isna().all(axis=1)
        df_filtered[file] = dfs[file][is_nan]
        
        columnnames = ["Macroactivity", "length_calc_hours", "length_calc_days","Start", "End"]
        length_df[file] = pd.DataFrame(columns = columnnames)
        
        for index, row in df_filtered[file].iterrows():

            start = str(row['Fait (Date)'])
            end = str(row['Unnamed: 15'])
            name = str(row["Travail à faire - Décision"])
            if start == "nan" or end == "nan":
                pass
            else:
                length_calc = time_difference(start,end)
                length_df_small = pd.DataFrame({
                    "Macroactivity": [name],
                    "length_calc_hours": [length_calc[0]],
                    "length_calc_days": [length_calc[1]],
                    "Start": [start],
                    "End": [end]
                    })
            
                length_df[file] = pd.concat([length_df[file], length_df_small])
                length_df[file] = length_df[file].reset_index(drop=True)


results = []

for key in length_df:
    result = length_df[key]["Macroactivity"].value_counts().reset_index()
    result.columns = ['Activity', 'count']
    results.append(result)

final_result = pd.concat(results, ignore_index=True)

final_result = final_result.groupby(['Activity'], as_index=False).sum()

columnnames = ["Macroactivity", "mean_hours","median","min_hours","max_hours", "ID"]
mean_df = pd.DataFrame(columns = columnnames)

for key in length_df:
    for i in final_result['Activity']:
        mean = length_df[key][length_df[key]['Macroactivity'] == i]
        if mean.empty:
            pass
        else:
            mean_activ = mean["length_calc_hours"].mean()
            median_activ = mean["length_calc_hours"].median()  
            min_activ = mean["length_calc_hours"].min()            
            max_activ = mean["length_calc_hours"].max()
            
            mean_df_small = pd.DataFrame({
                "Macroactivity": [i],
                "mean_hours": [mean_activ],
                "median": [median_activ],
                "min_hours":[min_activ],
                "max_hours":[max_activ],
                "ID": [key]
                })
        
            mean_df = pd.concat([mean_df, mean_df_small])
            
    print ("{0} done!".format(key))

mean_df = mean_df.reset_index(drop=True)
        
columnnames = ["Macroactivity", "mean_hours", "mean [days]","median","median [days]","min_hours", "min [days]","ID_min","max_hours", "max [days]","ID_max"]
results_average = pd.DataFrame(columns = columnnames)

for i in final_result['Activity']:
    mean = mean_df[mean_df['Macroactivity'] == i]
    mean_activ = mean["mean_hours"].mean()
    median_activ = mean["median"].median()
    min_activ = mean["min_hours"].min()

    min_index = mean['min_hours'].idxmin()
    min_id = mean.at[min_index, 'ID']            
    max_activ = mean["max_hours"].max()
    max_index = mean['min_hours'].idxmax()
    max_id = mean.at[max_index, 'ID'] 
    
    results_average_small = pd.DataFrame({
        "Macroactivity": [i],
        "mean_hours": [mean_activ],
        "mean [days]": [mean_activ/24],
        "median": [median_activ],
        "median [days]":[median_activ/24],
        "min_hours":[min_activ],
        "min [days]": [min_activ/24],
        "ID_min": [min_id],
        "max_hours":[max_activ],
        "max [days]": [max_activ/24],
        "ID_max": [max_id],        
        })
    results_average = pd.concat([results_average, results_average_small])
    
    print ("{0} done!".format(i))
    
       
    

merged_df = pd.merge(results_average, final_result, left_on='Macroactivity', right_on='Activity')

merged_df.drop('Activity', axis=1, inplace=True)


merged_df.to_excel("average_Macroavtivities2.xlsx", index=False) 















