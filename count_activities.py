import math
import pandas as pd
import matplotlib as plt
import os as os
from datetime import datetime
import plotly.express as px
import re
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import textwrap



"""
eLogbook

multi-file analysis
Count number of macroactivities
calculate length
calculate mean, median, min, max, IDs of min/max

"""

# folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\NH90"
folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\drive-download-20230207T104056Z-001"

Datefile = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\H225_Dates.xlsx"
datefile = pd.read_excel(Datefile, "Data")

header = 1

files = os.listdir(folder_path)




def time_difference(timestamp1, timestamp2):
    # Convert the timestamps to datetime objects
    date1 = datetime.strptime(timestamp1, '%Y-%m-%d %H:%M:%S')
    date2 = datetime.strptime(timestamp2, '%Y-%m-%d %H:%M:%S') 
    
    # # Calculate the difference between the two dates
    diff = date2 - date1
    
    return (diff.total_seconds() / 3600,str(diff))


date_str = datetime.today().strftime('%Y%m%d')
# folder_name = date_str + '_Complete_analysis_NH90'
folder_name = date_str + '_Complete_analysis_H225'


# Create the new folder
if not os.path.exists(folder_name):
    os.mkdir(folder_name)

'prepare and structure the elogbook files'

dfs = {}
logs = {}
df_filtered = {}
length_df = {}
length_df_minor = {}


for file in files:
    file_path = os.path.join(folder_path, file)
    if file.endswith('.xls'):
        xx = pd.read_excel(file_path)
        if xx.columns[0] == "Filtre":
            sheetname = "Observations - Anomalies"
            sheetname2 = "RecordLogs"
            df = pd.read_excel(file_path,sheetname,header = header)
            df2 = pd.read_excel(file_path,sheetname2,header = header)
            
            
            logs[file] = df2
            logs[file]['Unnamed: 3'] = pd.to_datetime(logs[file]['Unnamed: 3'])
            logs[file]['Date'] = pd.to_datetime(logs[file]['Date'])
            
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
            
            mask = df["Travail à faire - Décision"].isin(['POSTE 06 VAH (Visite Avant Habillage)', 'POSTE 07 HABILLAGE / VOL CONTRÔLE HABILLAGE'])
            subset = df.loc[mask]
            start = subset['Fait (Date)'].min()
            end = subset['Unnamed: 15'].max()
            
            new_row = {"Travail à faire - Décision": 'POSTE 06 VAH & POSTE 07 HABILLAGE COMBINED', 'Fait (Date)': start, 'Unnamed: 15': end}
            
            df = df.append(new_row, ignore_index=True)
            
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
            
            'minors'
            
            macros = df_filtered[file]["Travail à faire - Décision"].tolist()
            index_list = df_filtered[file].index.tolist()
            
            columnnames = ["Macroactivity","Minoractivity", "length_calc_hours", "length_calc_days", "Start", "End"]
            length_df_minor[file] = pd.DataFrame(columns = columnnames)
            
            dic = {}
            o = 0
            m = 0
            for j in index_list:
                if o+1 < len(index_list):
                    i = "{0}".format(macros[o])
                    dic[i] = {}
                    dic[i][j] = dfs[file].iloc[index_list[o]+1:index_list[o+1]]
                    o += 1
                    
                else:
                    i = "{0}".format(macros[o])
                    dic[i] = {}
                    dic[i][j] = dfs[file].iloc[index_list[o]+1:]        
    
    
            for outer_key, inner_dict in dic.items():
                for inner_key, dic[outer_key][inner_key] in inner_dict.items():
                    for index, row in dic[outer_key][inner_key].iterrows():
    
                        start = str(row['Date'])
                        end = str(row['Unnamed: 15'])
                        name = str(row["Travail à faire - Décision"])
                        testing = str(row["Limited change"])
                        response = str(row["Réponse"])
                        if start == "nan" or end == "nan":
                            pass
                        else:
                            length_calc = time_difference(start,end)
                            length_df_small = pd.DataFrame({
                                "Macroactivity": [outer_key],
                                "Minoractivity": [name],
                                "Answer": [response],
                                "length_calc_hours": [length_calc[0]],
                                "length_calc_days": [length_calc[1]],
                                "Start": [start],
                                "End": [end],
                                "test_activitiy": [testing]
                                })
                            
                            length_df_minor[file] = pd.concat([length_df_minor[file], length_df_small])                    
            
            
            columnnames = ["Macroactivity", "length_calc_hours", "length_calc_days","Start", "End"]
            length_df[file] = pd.DataFrame(columns = columnnames)
            
            for index, row in df_filtered[file].iterrows():
    
                start = str(row['Fait (Date)'])
                end = str(row['Unnamed: 15'])
                name = str(row["Travail à faire - Décision"])
                testing = str(row["Limited change"])
                if start == "nan" or end == "nan":
                    pass
                else:
                    length_calc = time_difference(start,end)
                    length_df_small = pd.DataFrame({
                        "Macroactivity": [name],
                        "length_calc_hours": [length_calc[0]],
                        "length_calc_days": [length_calc[1]],
                        "Start": [start],
                        "End": [end],
                        "test_activitiy": [testing]
                        })
                
                    length_df[file] = pd.concat([length_df[file], length_df_small])
                    length_df[file] = length_df[file].reset_index(drop=True)
        
            print ("{0} done!".format(file))
        
        elif xx.columns[0] == "Filter":
            sheetname = "Discrepancies"
            df = pd.read_excel(file_path,sheetname,header = header)
            # df = df.replace(r'\n(?!\s)', ' ', regex=True).replace(r'\n', '', regex=True)
            
            df = df.applymap(lambda x: re.sub(r'(\s)\n', r'\1', x) if isinstance(x, str) else x)
            df = df.applymap(lambda x: re.sub(r'\n(\s)', r' \1', x) if isinstance(x, str) else x)
            df = df.replace(r'\r\n|\r|\n', ' ', regex=True)
            df["To be done"] = df["To be done"].str.rstrip()
            
            
            "match the keyactivities"       
            
            df["To be done"] = df["To be done"].replace('POSTE 02 POINT FIXE AVANT 1ER VOL PFA', 'POSTE 02 PFA (Point Fixe Avant 1er vol)')
            df["To be done"] = df["To be done"].replace('POSTE 02 TRAVAUX SYSTEMATIQUES après 1er point fixe', 'POSTE 02 TRAVAUX SYSTEMATIQUES APRES 1er POINT FIXE')
            df["To be done"] = df["To be done"].replace('POSTE 02 VAV (VISITE AVANT VOL)', 'POSTE 02 VAV (Visite avant vol)')
            df["To be done"] = df["To be done"].replace('POSTE 03 VOL T1','POSTE 03 VOL T1 (K1)')
            df["To be done"] = df["To be done"].replace('POSTE 06 MAINTENACE VAL', 'POSTE 06 MAINTENANCE VAL')
            df["To be done"] = df["To be done"].replace('POSTE 06 VAH', 'POSTE 06 VAH (Visite Avant Habillage)')
            df["To be done"] = df["To be done"].replace('POSTE 09 Présentation machine services officiels / TRANSFERT LH', 'POSTE 09 PRESENTATION MACHINE SERVICES OFFICIELS / TRANSFERT LH')
            
            ""
            
            
            
            dfs[file] = df
            dfs[file]['Date'] = pd.to_datetime(dfs[file]['Date'])
            dfs[file]['Date'] = dfs[file]['Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
            dfs[file]['Done (Date)'] = pd.to_datetime(dfs[file]['Done (Date)'])
            dfs[file]['Done (Date)'] = dfs[file]['Done (Date)'].dt.strftime('%Y-%m-%d %H:%M:%S')
            dfs[file]['Unnamed: 15'] = pd.to_datetime(dfs[file]['Unnamed: 15'])
            dfs[file]['Unnamed: 15'] = dfs[file]['Unnamed: 15'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            is_nan = dfs[file][['Date', 'Author','Stamp-ID','Answer','Date.1','Author.1','Stamp-ID.1']].isna().all(axis=1)
            df_filtered[file] = dfs[file][is_nan]
            
            'minors'
            
            macros = df_filtered[file]["To be done"].tolist()
            index_list = df_filtered[file].index.tolist()
            
            columnnames = ["Macroactivity","Minoractivity", "length_calc_hours", "length_calc_days", "Start", "End"]
            length_df_minor[file] = pd.DataFrame(columns = columnnames)
            
            dic = {}
            o = 0
            m = 0
            for j in index_list:
                if o+1 < len(index_list):
                    i = "{0}".format(macros[o])
                    dic[i] = {}
                    dic[i][j] = dfs[file].iloc[index_list[o]+1:index_list[o+1]]
                    o += 1
                    
                else:
                    i = "{0}".format(macros[o])
                    dic[i] = {}
                    dic[i][j] = dfs[file].iloc[index_list[o]+1:]        
    
    
            for outer_key, inner_dict in dic.items():
                for inner_key, dic[outer_key][inner_key] in inner_dict.items():
                    for index, row in dic[outer_key][inner_key].iterrows():
    
                        start = str(row['Date'])
                        end = str(row['Unnamed: 15'])
                        name = str(row["To be done"])
                        testing = str(row["Limited change"])
                        response = str(row["Answer"])
                        
                        if start == "nan" or end == "nan":
                            pass
                        else:
                            length_calc = time_difference(start,end)
                            length_df_small = pd.DataFrame({
                                "Macroactivity": [outer_key],
                                "Minoractivity": [name],
                                "Answer": [response],
                                "length_calc_hours": [length_calc[0]],
                                "length_calc_days": [length_calc[1]],
                                "Start": [start],
                                "End": [end],
                                "test_activitiy": [testing]
                                })
                            
                            length_df_minor[file] = pd.concat([length_df_minor[file], length_df_small])                    
            
            
            columnnames = ["Macroactivity", "length_calc_hours", "length_calc_days","Start", "End"]
            length_df[file] = pd.DataFrame(columns = columnnames)
            
            for index, row in df_filtered[file].iterrows():
    
                start = str(row['Done (Date)'])
                end = str(row['Unnamed: 15'])
                name = str(row["To be done"])
                testing = str(row["Limited change"])
                if start == "nan" or end == "nan":
                    pass
                else:
                    length_calc = time_difference(start,end)
                    length_df_small = pd.DataFrame({
                        "Macroactivity": [name],
                        "length_calc_hours": [length_calc[0]],
                        "length_calc_days": [length_calc[1]],
                        "Start": [start],
                        "End": [end],
                        "test_activitiy": [testing]
                        })
                
                    length_df[file] = pd.concat([length_df[file], length_df_small])
                    length_df[file] = length_df[file].reset_index(drop=True)
        
            print ("{0} done!".format(file))



'count the macroactivities'
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


merged_df.to_excel(os.path.join(folder_name,"average_Macroavtivities.xlsx"), index=False) 


'count the microactivities'

columnnames = ["Macroactivity","Minoractivity", "Answer", "length_calc_hours", "length_calc_days","Start", "End"]
length_df_minor_big = pd.DataFrame(columns = columnnames)

for key in length_df_minor:
    length_df_minor_big = pd.concat([length_df_minor_big, length_df_minor[key]])
    
length_df_minor_big.to_excel(os.path.join(folder_name,"all_activs.xlsx"), index=False) 




results_minor = []
for key in length_df_minor:
    result_minor = length_df_minor[key]["Minoractivity"].value_counts().reset_index()
    result_minor.columns = ['Activity', 'count']
    results_minor.append(result_minor)

final_result_minor = pd.concat(results_minor, ignore_index=True)

final_result_minor = final_result_minor.groupby(['Activity'], as_index=False).sum()

final_result_minor.to_excel(os.path.join(folder_name,"counter_Microavtivities.xlsx"), index=False) 


'calculate number of rows from macros'
columnnames = ["file", "Number of macro tasks"]
Macro_numbers = pd.DataFrame(columns = columnnames)

for key, df in length_df.items():
    df_list_small = pd.DataFrame({
        "file": [key],
        "Number of macro tasks": [len(df)]        
        })
    
    Macro_numbers = pd.concat([Macro_numbers, df_list_small])
    
'calculate number of rows from micros'
columnnames = ["file", "Number of micro tasks"]
Micro_numbers = pd.DataFrame(columns = columnnames)

for key, df in length_df_minor.items():
    df_list_small = pd.DataFrame({
        "file": [key],
        "Number of micro tasks": [len(df)]        
        })
    
    Micro_numbers = pd.concat([Micro_numbers, df_list_small])

merged_numbers = pd.merge(Macro_numbers, Micro_numbers, left_on='file', right_on='file')

merged_numbers['Number of macro tasks'] = merged_numbers['Number of macro tasks'].astype(int)
merged_numbers['Number of micro tasks'] = merged_numbers['Number of micro tasks'].astype(int)

'write summary df on number of activities in the file'

mean_macro = merged_numbers["Number of macro tasks"].mean()
median_macro = merged_numbers["Number of macro tasks"].median()
mini_macro = merged_numbers["Number of macro tasks"].min()
mini_index_macro = merged_numbers["Number of macro tasks"].idxmin()
mini_id_macro = merged_numbers.at[mini_index_macro, 'file']            
maxi_macro = merged_numbers["Number of macro tasks"].max()
maxi_index_macro = merged_numbers['Number of macro tasks'].idxmax()
maxi_id_macro = merged_numbers.at[maxi_index_macro, 'file'] 

mean_micro = merged_numbers["Number of micro tasks"].mean()
median_micro = merged_numbers["Number of micro tasks"].median()
mini_micro = merged_numbers["Number of micro tasks"].min()
mini_index_micro = merged_numbers["Number of micro tasks"].idxmin()
mini_id_micro = merged_numbers.at[mini_index_micro, 'file']            
maxi_micro = merged_numbers["Number of micro tasks"].max()
maxi_index_micro = merged_numbers['Number of micro tasks'].idxmax()
maxi_id_micro = merged_numbers.at[maxi_index_micro, 'file'] 


summary_numbers = pd.DataFrame({
    "mean": [mean_macro,mean_micro],
    "median": [median_macro,median_micro],
    "min":[mini_macro,mini_micro],
    "ID_min": [mini_id_macro,mini_id_micro],
    "max":[maxi_macro,maxi_micro],
    "ID_max": [maxi_id_macro,maxi_id_micro]       
    }, index=["Macroactivities", "Microactivities"])


with pd.ExcelWriter(os.path.join(folder_name, "Summary.xlsx")) as writer:
    summary_numbers.to_excel(writer, sheet_name='Summary', index=True)
    merged_numbers.to_excel(writer, sheet_name='Data', index=False)
    


'get the GR and Flight activity infos'
subfolder_name = "test_activities"
subfolder_path = os.path.join(folder_name, subfolder_name)
if not os.path.exists(subfolder_path):
    os.mkdir(subfolder_path)

test_dic = {}

columnnames = []
test_activs_df = pd.DataFrame(columns = columnnames)
   
for key in length_df:
    test_dic[key] = {}
    unique_tests = length_df[key]['test_activitiy'].unique().tolist()
    unique_tests = [x for x in unique_tests if str(x) != 'nan']
    
    test = length_df[key][length_df[key]['test_activitiy'].str.contains('VSO')]
    if test.empty:
        VSO_end = 'No VSO'
    else:
        VSO_end = test["End"].max()
        
    

    
    GR_list = []
    GR_list_before_VSO = []
    GR_list_after_VSO = []
    FL_list = []
    FL_list_before_VSO = []
    FL_list_after_VSO = []
    
    for string in unique_tests:
        if string.startswith('GR'):
            GR_list.append(string)
        elif string.startswith('FL'):
            FL_list.append(string) 
    
    filename = "{0}_GR_activities.xlsx".format(key)
    filename = os.path.join(subfolder_path, filename)
    with pd.ExcelWriter(filename) as writer:
        for i in GR_list:
            test_dic[key][i] = length_df[key][length_df[key]['test_activitiy'] == i]
            test_dic[key][i].to_excel(writer, sheet_name='{0}'.format(i), index=True)
            if VSO_end == "No VSO":
                pass
            else:
                date =  test_dic[key][i]["Start"].iloc[0]           
                if date < VSO_end:
                    GR_list_before_VSO.append(i)
                else:
                    GR_list_after_VSO.append(i)
                   
                    
        
    filename = "{0}_FL_activities.xlsx".format(key)
    filename = os.path.join(subfolder_path, filename)
    with pd.ExcelWriter(filename) as writer:
        for i in FL_list:
            test_dic[key][i] = length_df[key][length_df[key]['test_activitiy'] == i]
            test_dic[key][i].to_excel(writer, sheet_name='{0}'.format(i), index=True)
            if VSO_end == "No VSO":
                pass
            else:
                date =  test_dic[key][i]["Start"].iloc[0]           
                if date < VSO_end:
                    FL_list_before_VSO.append(i)
                else:
                    FL_list_after_VSO.append(i)
                    
    if len(GR_list_before_VSO) == 0 and   len(GR_list_after_VSO) == 0:            
        test_activs_df_small = pd.DataFrame({
            "Total Number of Ground Runs": [len(GR_list)],
            "Number of Ground Runs before VSO": [float('NaN')],
            "Number of Ground Runs after VSO": [float('NaN')],
            "Total Number of Flight tests": [len(FL_list)],
            "Number of Flight Tests before VSO": [float('NaN')],
            "Number of Flight Tests after VSO": [float('NaN')]}
            , index=[key])
    else:
        test_activs_df_small = pd.DataFrame({
            "Total Number of Ground Runs": [len(GR_list)],
            "Number of Ground Runs before VSO": [len(GR_list_before_VSO)],
            "Number of Ground Runs after VSO": [len(GR_list_after_VSO)],
            "Total Number of Flight tests": [len(FL_list)],
            "Number of Flight Tests before VSO": [len(FL_list_before_VSO)-1],
            "Number of Flight Tests after VSO": [len(FL_list_after_VSO)]}
            , index=[key])
    
    test_activs_df = pd.concat([test_activs_df, test_activs_df_small])


gr_mean = test_activs_df["Total Number of Ground Runs"].mean()
fl_mean = test_activs_df["Total Number of Flight tests"].mean()

gr_mean_before_VSO = test_activs_df[ "Number of Ground Runs before VSO"].mean()
fl_mean_before_VSO = test_activs_df["Number of Flight Tests before VSO"].mean()

gr_mean_after_VSO = test_activs_df["Number of Ground Runs after VSO"].mean()
fl_mean_after_VSO = test_activs_df["Number of Flight Tests after VSO"].mean()


new_row = pd.DataFrame({"Total Number of Ground Runs": gr_mean, "Number of Ground Runs before VSO": gr_mean_before_VSO, "Number of Ground Runs after VSO": gr_mean_after_VSO,
                        "Total Number of Flight tests": fl_mean, "Number of Flight Tests before VSO": fl_mean_before_VSO,"Number of Flight Tests after VSO": fl_mean_after_VSO,
                        }, index=["mean"])
test_activs_df = test_activs_df.append(new_row)


flight_activities_df = {}

for key in test_dic:
    columnnames = []
    flight_activities_df[key] = pd.DataFrame(columns = columnnames)
    for i in test_dic[key]:
        activity_start = test_dic[key][i]['Start'].min()
        activity_end = test_dic[key][i]['End'].max()
        
        length = time_difference(activity_start,activity_end)[1]
        
        activities_df_small = pd.DataFrame({
            "Activity": [i],
            "Start": [activity_start],
            "End": [activity_end],
            "Length": [length]})
        
        
        flight_activities_df[key] = pd.concat([flight_activities_df[key], activities_df_small])

        flight_activities_df[key] = flight_activities_df[key].reset_index(drop=True)
        
filename = "Summary_testactivities.xlsx"
filename = os.path.join(subfolder_path, filename)
with pd.ExcelWriter(filename) as writer:
    test_activs_df.to_excel(writer, index=True)
    for key in flight_activities_df:
        strip = textwrap.wrap(key, 25)[0]
        flight_activities_df[key].to_excel(writer, sheet_name='{0}'.format(strip), index = True)




'Find the PFA Start and End DATE'

columnnames = ["length_GR", "start_GR", "end_GR","length_FL","start_FL", "end_FL"]
Ground_test_duration = pd.DataFrame(columns = columnnames)

for key in length_df:
    search_value = 'POSTE 02 PFA (Point Fixe Avant 1er vol)'
    result = length_df[key][length_df[key]['Macroactivity'] == search_value]
    
    PFA_start_date = result['Start'].iloc[0]
    PFA_end_date = result['End'].iloc[0]
    
    index_PFA = result.index[0]
    
    'peinture check'
    # check = length_df[key].loc[:index_PFA, :][length_df[key].loc[:index_PFA, 'Macroactivity'].str.contains('peinture')]

    
    # if check.empty:
    #     start_GR = datefile[datefile['file'] == key]
    #     start_GR = str(start_GR['VEP Date'].iloc[-1])
    # else: 
    #     start_GR = check['End'].iloc[-1]
    
    coc = datefile[datefile['file'] == key]
    coc = str(coc['CoC Date'].iloc[-1])
    
    start_GR = datefile[datefile['file'] == key]
    start_GR = str(start_GR['VEP Date'].iloc[-1])  
        
    Duration_GroundTests = time_difference(start_GR,PFA_start_date)[0] / 24
    
    
    new_df = flight_activities_df[key][~flight_activities_df[key]['Activity'].str.contains('CTL|VSO|T21') & flight_activities_df[key]['Activity'].str.startswith('FL')]
    new_df = new_df[new_df['End'] <=coc ]
    
    first_flight = new_df['Start'].iloc[0]
    Last_flight = new_df['End'].iloc[-1]
    
    # Duration_FlightTests = TSS - PFA_end_date
    # Duration_FlightTests = time_difference(PFA_end_date,Last_flight)[0] / 24
    
    Duration_FlightTests = time_difference(first_flight,Last_flight)[0] / 24
    
    
    
    Ground_test_duration_small = pd.DataFrame({
        "length_GR": [Duration_GroundTests],
        "start_GR": [start_GR],
        "end_GR": [PFA_start_date],
        "length_FL": [Duration_FlightTests],
        "start_FL": [PFA_end_date],
        "end_FL": [Last_flight]
        }
        , index=[key])
    
    
    Ground_test_duration = pd.concat([Ground_test_duration, Ground_test_duration_small])

mean_gr = Ground_test_duration['length_GR'].mean()
min_gr = Ground_test_duration['length_GR'].min()
min_gr_id = Ground_test_duration['length_GR'].idxmin()

mean_fl = Ground_test_duration['length_FL'].mean()
min_fl = Ground_test_duration['length_FL'].min()
min_fl_id = Ground_test_duration['length_FL'].idxmin()

new_row = pd.DataFrame({"length_GR": mean_gr, "length_FL": mean_fl},
                        index=["mean"])
Ground_test_duration = Ground_test_duration.append(new_row)

new_row = pd.DataFrame({"length_GR": min_gr,"min_GR_ID":min_gr_id, "length_FL": min_fl, "min_FL_ID":min_fl_id },
                        index=["min"])
Ground_test_duration = Ground_test_duration.append(new_row)

filename = "GR_FL_average.xlsx"
filename = os.path.join(subfolder_path, filename)
with pd.ExcelWriter(filename) as writer:
    Ground_test_duration.to_excel(writer, index=True)
    

'Calculate total lead time'

columnnames = ["total_LT", "VEP", "CoC"]
total_LT_df = pd.DataFrame(columns = columnnames)

for key in length_df:
    # search_value = 'POSTE 08 VISITE VSO'
    # result = length_df[key][length_df[key]['Macroactivity'] == search_value]
    # END_of_VSO = result['End'].iloc[0]
    
    CoC = datefile[datefile['file'] == key]
    CoC = str(CoC['CoC Date'].iloc[-1])

    VEP = datefile[datefile['file'] == key]
    VEP = str(VEP['VEP Date'].iloc[-1]) 
    
    total_LT = time_difference(VEP,CoC)[0] / 24
    
    total_LT_df_small = pd.DataFrame({
        "total_LT": [total_LT],
        "VEP": [VEP],
        "CoC": [CoC]
        }
        , index=[key])
    
    
    total_LT_df = pd.concat([total_LT_df, total_LT_df_small])


mean_LT = total_LT_df['total_LT'].mean()
new_row = pd.DataFrame({"total_LT": mean_LT},
                        index=["mean"])

total_LT_df = total_LT_df.append(new_row)

filename = "Total_LT.xlsx"
filename = os.path.join(folder_name, filename)
with pd.ExcelWriter(filename) as writer:
    total_LT_df.to_excel(writer, index=True)


'Gantt charts for flying activities'

for key in flight_activities_df:   
    fig = px.timeline(flight_activities_df[key], x_start='Start', x_end='End', y='Activity')
    fig.update_layout(
        width=1800,  # set the width to 1800 pixels
        height=1800,  # set the height to 600 pixels
        margin=dict(l=50, r=50, t=50, b=50),  # set the margins
        annotations=[
            dict(
                x=row['End'], 
                y=row['Activity'], 
                xref='x', 
                yref='y', 
                text= row['Length'],
                showarrow=False, 
                font=dict(size=14)
            ) for _, row in flight_activities_df[key].iterrows()
        ],
    title={
            'text': "{0}".format(key),  # set the title text
            'y': 0.98,  # set the y position of the title
            'x': 0.5,  # set the x position of the title
            'xanchor': 'center',  # set the x anchor of the title
            'yanchor': 'top'  # set the y anchor of the title
        }
    )
    
    fig.update_xaxes(
        type='date',
        tickformat='Start of %CW%V'
    )
    a = key
    micro_gantt = "{0}_Gantt_minor.png".format(a)
    micro_gantt = os.path.join(subfolder_path, micro_gantt)
    fig.write_image(micro_gantt, scale=2)


columnnames = ["total_discrep", 
               "discrep_started_before_VEP", "discrep_closed_before_VEP",
               "discrep_started_before_endGR", "discrep_closed_before_endGR",
               "discrep_started_before_endFL", "discrep_closed_before_endFL",
               "discrep_started_before_CoC", "discrep_closed_before_CoC"]
discrep_df = pd.DataFrame(columns = columnnames)

for key in length_df:
    total_discrep = len(length_df[key])  
    
    VEP_date = datefile[datefile['file'] == key]
    VEP_date = str(VEP_date['VEP Date'].iloc[-1]) 
    
    started_before_VEP = length_df[key][length_df[key]['Start'] < VEP_date]     
    discrep_started_before_VEP = len(started_before_VEP)  
    
    closed_before_VEP = length_df[key][length_df[key]['End'] < VEP_date]
    discrep_closed_before_VEP = len(closed_before_VEP)
    


    started_before_VEP = length_df[key][length_df[key]['Start'] < VEP_date]     

    endGR = Ground_test_duration.loc[key, 'end_GR']

    started_before_endGR = length_df[key][length_df[key]['Start'] < endGR]     
    discrep_started_before_endGR = len(started_before_endGR)  
    
    closed_before_endGR = length_df[key][length_df[key]['End'] < endGR]
    discrep_closed_before_endGR = len(closed_before_endGR)

    
    endFL = Ground_test_duration.loc[key, 'end_FL']

    started_before_endFL = length_df[key][length_df[key]['Start'] < endFL]     
    discrep_started_before_endFL = len(started_before_endFL)  
    
    closed_before_endFL = length_df[key][length_df[key]['End'] < endFL]
    discrep_closed_before_endFL = len(closed_before_endFL)


    CoC_date = datefile[datefile['file'] == key]
    CoC_date = str(CoC_date['CoC Date'].iloc[-1])
    
    started_before_CoC = length_df[key][length_df[key]['Start'] < CoC_date]     
    discrep_started_before_CoC = len(started_before_CoC)  
    
    closed_before_CoC = length_df[key][length_df[key]['End'] < CoC_date]
    discrep_closed_before_CoC = len(closed_before_CoC)
    
    
    discrep_df_small = pd.DataFrame({
        "total_discrep": [total_discrep], 
        "discrep_started_before_VEP": [discrep_started_before_VEP],
        "discrep_closed_before_VEP": [discrep_closed_before_VEP],
        "discrep_started_before_endGR": [discrep_started_before_endGR], 
        "discrep_closed_before_endGR": [discrep_closed_before_endGR],
        "discrep_started_before_endFL": [discrep_started_before_endFL], 
        "discrep_closed_before_endFL": [discrep_closed_before_endFL],
        "discrep_started_before_CoC": [discrep_started_before_CoC],
        "discrep_closed_before_CoC": [discrep_closed_before_CoC]
        }
        , index=[key])   
    
    discrep_df = pd.concat([discrep_df, discrep_df_small])
    

filename = "Discrepencies_overview.xlsx"
filename = os.path.join(folder_name, filename)
with pd.ExcelWriter(filename) as writer:
    discrep_df.to_excel(writer, index=True)


# 'structure logs'
# GR_activs = {}  # create new dictionary for GR activities
# FL_activs = {}  # create new dictionary for FL activities

# for key in logs:
#     # get index of last "SUM" entry in column B
#     last_sum_idx = logs[key]['Unnamed: 0'].last_valid_index()
#     # slice the dataframe to select only the rows before the last "SUM" entry
#     df_before_sum = logs[key].iloc[:last_sum_idx]
#     # extract rows with "GR" in column A
#     gr_rows =df_before_sum[df_before_sum['Unnamed: 0'] == 'GR']
#     # write them into the corresponding dataframe in GR_activs
#     GR_activs[key] = gr_rows
#     GR_activs[key]['Duration'] =  GR_activs[key]['Duration']
    
#     # extract rows with "FL" in column A
#     fl_rows = df_before_sum[df_before_sum['Unnamed: 0'] == 'FL']
#     # write them into the corresponding dataframe in FL_activs
#     FL_activs[key] = fl_rows
