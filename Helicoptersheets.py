import math
import pandas as pd
import matplotlib as plt
import os as os
from datetime import datetime

"""
Helicopter sheet

multi-file analysis

tbd.

"""



folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\Helicopter sheets"
sheetname = "2172"
header = 4


def extract_rows(df, keyword):
    # Get the index of the keyword
    idx = df[df['Unnamed: 0'] == keyword].index[0]
    
    top = df
    # Extract all rows after the keyword
    bottom = df.loc[idx+1:]
    
    # Remove the extracted rows from the original dataframe
    top.drop(df.index[idx:], inplace=True)
    
    bottom.columns = bottom.iloc[0]
    bottom = bottom.drop(bottom.index[0])
    
    return bottom, top

'Frozenreasons'

data = {'number': ["15","20","05",
                   "24","25","26",
                   "19","02","01","03","27","13",
                   "28","17",
                   "29","14","09",
                   "100","101","102","103","104","105","106"
                   ], 
        
        'origin': ["Design", "Design", "Design",
                   "Geplant","Geplant","Geplant",
                   "Kunde","Kunde","Kunde","Kunde","Kunde","Kunde",
                   "Programm","Programm",
                   "Zertif/ Qualif","Zertif/ Qualif","Zertif/ Qualif",
                   "Blocked Reason","Blocked Reason","Blocked Reason","Blocked Reason","Blocked Reason","Blocked Reason","Blocked Reason"], 
        
        'title': ["Fehldes Design","Sonstiges","Zusätzlicher Flugversuch",
                  
                  "Training / Messe / Demo", "Zertifizierung / Qualifikation","zusätzlicher Flugversuch",
                  
                  "Entwicklung gemäß Customer Sheet","Fehlende Information vom Kunden","Fehlender Kunde / Re-Allokation",
                  "Produktions stopp durch Kunden gewünscht", "Sonstiges","Änderungswunsch",
                  
                  "Sonstiges","Training / Messe / Demo",
                  
                  "Sonstiges", "Zertifizierung", "während Serienproduktion",
                  
                  "Qualitätsproblem Lieferant / Nacharbeit ausstehend", "Dynamisches Center", "Interne Fertigung (ohne DC und Blätter PLB)",
                  "Blätter (PLB)", "Service Center (VIP, Lakierung, Composite)", "Einkauf", "Produdukt Industrialisierung"
            ]}

frozenreasons = pd.DataFrame(data)

'read in all the files and split them into top / bottom'

files = os.listdir(folder_path)

hcsheet_original = {}
hcsheet_top = {}
hcsheet_bottom = {}

for file in files:
    file_path = os.path.join(folder_path, file)
    if file.endswith('.xlsx'):
        sheetname = file.rsplit(".", 1)[0]
        df = pd.read_excel(file_path,sheetname,header = header)
        df = df = df.astype(str)
        hcsheet_original[file] = df
        
        split = extract_rows(df,"Reasons for Delay / Frozen Period:")
        hcsheet_bottom[file] = split[0]
        hcsheet_top[file] = split[1]
        
        hcsheet_top[file] = hcsheet_top[file].drop(hcsheet_top[file].index[-1])
        
               
        
        print ("{0} done!".format(file))


columnnames = ["Thema", "Datum", "erstellt durch", "Station", "Delay", "Frozen", "Frozen reason", "ID"]
FL = pd.DataFrame(columns = columnnames)


column_order = ["Thema", "Datum", "erstellt durch", "Station", "Delay", "Frozen", "Frozen reason"]

for key in hcsheet_bottom:
    if set(column_order) == set(hcsheet_bottom[key].columns):
        pass
    else:
        hcsheet_bottom[key] = hcsheet_bottom[key].reindex(columns=["Thema", "Datum", "erstellt durch", "Station", "Delay","Frozen","Frozen reason", "Tage"])
        for index, row in hcsheet_bottom[key].iterrows():
            if row['Frozen'] == 'x':
                hcsheet_bottom[key].loc[index, 'Frozen'] = row['Tage']
            else:
                hcsheet_bottom[key].loc[index, 'Delay'] = str(row['Tage'])
        
        # drop the 'Tage' column
        hcsheet_bottom[key].drop('Tage', axis=1, inplace=True)
     
    FL_small = hcsheet_bottom[key][hcsheet_bottom[key]['Station'].isin(['CoC', 'S14', 'S15', 'S16', 'S17', '14', '15', '16', '17'])]
    FL_small.insert(7,'ID', key)
    FL = pd.concat([FL, FL_small])
    
changes_in_FL = FL
changes_in_FL['Frozen'] = pd.to_numeric(changes_in_FL['Frozen'], errors='coerce').fillna(0).astype(int)
changes_in_FL['Delay'] = pd.to_numeric(changes_in_FL['Delay'], errors='coerce').fillna(0).astype(int)

changes_in_FL.to_excel("Changes_in_Flightline.xlsx", index=False) 



















