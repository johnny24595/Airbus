import math
import pandas as pd
import matplotlib.pyplot as plt
import os as os
from datetime import datetime
import plotly.express as px
import re
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import textwrap
from collections import defaultdict
import numpy as np
import plotIKVStyle
import json

"""__________________________________________________________________________
"""
'Change these'


# 'H160 setup'

# model = "H160"
# folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\H160_export"
# filename = "H160_CONFORMITE.csv"

# search_value = 'Pick-Up (FAL) - MECHANIK'                                                               # Pick up mechanic from FAL
# search_value2 = 'Beanstandung zur Bodenlauf und Flugfreigabe'                                           # NCs from FAL
# search_value3 = 'Bodenlauffreigabe'                                                                     # GR release
# search_value4 = 'Bodenlaufprogramm'                                                                     # GR activity
# search_value6 = 'Flugfreigabe'                                                                          # FT release
# search_value7 = 'Einflugprogramm durchführen'                                                           # FT activity

# keyword_completelogbooks = 'Bodenlaufprogramm'

# name_filter = "AHD Civil"


'H135/H145 setup'

model = "H135_H145"
folder_path = r"C:\Users\jsiegber\OneDrive - Capgemini\05_Projekte\03_Airbus_Helicopters_Flightline\50_Data_&_Analytics\H135_H145_2022_export"
filename = "135_145.csv"

search_value = 'Pick-Up (FAL) - MECHANIK'                                                               # Pick up mechanic from FAL
search_value2 = 'Beanstandung zur Bodenlauf und Flugfreigabe'                                           # NCs from FAL
search_value3 = 'Bodenlauffreigabe'                                                                     # GR release
search_value4 = 'Bodenlaufprogramm'                                                                     # GR activity
search_value6 = 'Flugfreigabe'                                                                          # FT release
search_value7 = 'Einflugprogramm durchführen'                                                           # FT activity

keyword_completelogbooks = 'Bodenlaufprogramm'

name_filter = "AHD Civil"

"""__________________________________________________________________________
"""

'methods'

def calculate_days(start_time, end_time):
    delta = end_time - start_time
    hours = delta.total_seconds() / (3600*24)
    return hours

"""__________________________________________________________________________
"""

'setup'
file = os.path.join(folder_path, filename)

df = pd.read_csv(file)

df['Discrepancy_stampedby_Processed'] = pd.to_datetime(df['Discrepancy_stampedby_Processed'])
df['Discrepancy_createdby_Processed'] = pd.to_datetime(df['Discrepancy_createdby_Processed'])
df['Task_stampedby_Processed'] = pd.to_datetime(df['Task_stampedby_Processed'])
df['TaskToBeDone_processedDate'] = pd.to_datetime(df['TaskToBeDone_processedDate'])

unique_IDs = df['Discrepancy_logbookElogbookId'].unique().tolist()

unique_customers = df['LogBook_customer'].unique().tolist()

unique_names = df['name'].unique().tolist()  

unique_types = df['LogBook_aircraftTypeVersion'].unique().tolist()  


'files'
date_str = datetime.today().strftime('%Y%m%d')
folder_name = date_str + '_skywise_analysis_{0}'.format(model)

# Create the new folder
if not os.path.exists(folder_name):
    os.mkdir(folder_name)

Summary_file = "Summary_{0}_2022.xlsx".format(model)
Summary_file = os.path.join(folder_name, Summary_file)

mean_file = "Mean_of_main_tasks_{0}.xlsx".format(model)
mean_file = os.path.join(folder_name, mean_file)
    
discrep_task_summary_file = "Discrep&Tasks_{0}.xlsx".format(model)
discrep_task_summary_file = os.path.join(folder_name, discrep_task_summary_file)

discrep_task_summary_file_2022 = "Discrep&Tasks_{0}_2022.xlsx".format(model)
discrep_task_summary_file_2022 = os.path.join(folder_name, discrep_task_summary_file_2022)

types = "Types_{0}_2022.xlsx".format(model)
types = os.path.join(folder_name, types)

customers = "customers_{0}_all.xlsx".format(model)
customers = os.path.join(folder_name, customers)

"""__________________________________________________________________________

'split the logbooks'
   __________________________________________________________________________
"""

logbooks = {}

for i in unique_IDs:
    logbooks[i] = df[df['Discrepancy_logbookElogbookId'] == i]
    logbooks[i] = logbooks[i].sort_values(by='Discrepancy_sequence').reset_index(drop=True)
    logbooks[i] = logbooks[i].reset_index(drop=True) 
    
   
'count number of different aircraft types'
type_counter = {category: 0 for category in unique_types}

for df in logbooks.values():
    # Get the value in the first row of column A
    first_a_value = df.loc[0, 'LogBook_aircraftTypeVersion']
    # Update the counter for the appropriate category
    if first_a_value in unique_types:
        type_counter[first_a_value] += 1   
    
'count activities' 
activitiy_list = []

for key in logbooks:
    unique_activs = logbooks[key]['Discrepancy_description'].unique().tolist()
    activitiy_list.append(unique_activs)
    
counter_dict = defaultdict(int)

for lst in activitiy_list:
    for activity in lst:
        counter_dict[activity] += 1
        
counter_df = pd.DataFrame(counter_dict.items(), columns=['activity', 'counter'])
counter_df = counter_df.sort_values(by=['counter'], ascending=False)

'count customers'
customer_counter_complete = {category: 0 for category in unique_customers}

for df in logbooks.values():
    # Get the value in the first row of column A
    first_a_value = df.loc[0, 'LogBook_customer']
    # Update the counter for the appropriate category
    if first_a_value in unique_customers:
        customer_counter_complete[first_a_value] += 1

customer_counter_complete = pd.DataFrame(customer_counter_complete.items())
with pd.ExcelWriter(customers) as writer:
    customer_counter_complete.to_excel(writer, index=True)


'Discrep & tasks summary'
columnnames = ["Macro_tasks", "Micro_tasks","actype","file"]
discrep_task_summary = pd.DataFrame(columns = columnnames) 
            
for key in logbooks: 
    actype = logbooks[key]['LogBook_aircraftTypeVersion'].iloc[0]           
    unique_discrep = logbooks[key]['Discrepancy_sequence'].unique().tolist()
    unique_discrep = len(unique_discrep)
    
    unique_tasks = logbooks[key]['Task_taskElogbookId'].unique().tolist()
    unique_tasks = len(unique_tasks)
    
    discrep_task_summary_small = pd.DataFrame({
        "Macro_tasks": [unique_discrep],
        "Micro_tasks": [unique_tasks],
        "actype": [actype],
        "file": [key]
        })
    
    discrep_task_summary = pd.concat([discrep_task_summary,discrep_task_summary_small])   
    discrep_task_summary = discrep_task_summary.reset_index(drop=True) 

mean_macro = discrep_task_summary["Macro_tasks"].mean() 
mean_micro = discrep_task_summary["Micro_tasks"].mean() 

new_row = pd.DataFrame({"Macro_tasks": mean_macro, "Micro_tasks": mean_micro}, index=["mean"])
discrep_task_summary = discrep_task_summary.append(new_row) 

with pd.ExcelWriter(discrep_task_summary_file) as writer:
    discrep_task_summary.to_excel(writer, index=True)


"""__________________________________________________________________________

'find the complete logbooks with the defined keyword'
   __________________________________________________________________________
"""

complete_logbooks = {}

for key in logbooks:
    if keyword_completelogbooks in logbooks[key]['Discrepancy_description'].values:
        complete_logbooks[key] = logbooks[key]
        complete_logbooks[key] = complete_logbooks[key].reset_index(drop=True) 


'count number of different aircraft types complete logbooks'
type_counter_complete = {category: 0 for category in unique_types}

for df in complete_logbooks.values():
    # Get the value in the first row of column A
    first_a_value = df.loc[0, 'LogBook_aircraftTypeVersion']
    # Update the counter for the appropriate category
    if first_a_value in unique_types:
        type_counter_complete[first_a_value] += 1

'count number of different categories complete elogbooks'
category_counters = {category: 0 for category in unique_names}

for df in complete_logbooks.values():
    # Get the value in the first row of column A
    first_a_value = df.loc[0, 'name']
    # Update the counter for the appropriate category
    if first_a_value in unique_names:
        category_counters[first_a_value] += 1    


        
"""__________________________________________________________________________

Calculate
   __________________________________________________________________________
"""



'Calculate the Discrepancy and task time'
for key in complete_logbooks:
    # complete_logbooks[key]['Discrepancy_time'] = complete_logbooks[key]['Discrepancy_stampedby_Processed'] - complete_logbooks[key]['Discrepancy_createdby_Processed']
    complete_logbooks[key]['Discrepancy_time'] = complete_logbooks[key].apply(lambda x: calculate_days(x['Discrepancy_createdby_Processed'], x['Discrepancy_stampedby_Processed']), axis=1)
    # complete_logbooks[key]['Task_time'] = complete_logbooks[key]['Task_stampedby_Processed'] - complete_logbooks[key]['TaskToBeDone_processedDate']
    complete_logbooks[key]['Task_time'] = complete_logbooks[key].apply(lambda x: calculate_days(x['TaskToBeDone_processedDate'], x['Task_stampedby_Processed']), axis=1)

    
"""

Filters

"""

'filter only on keyword'
civil_logbooks = {}

for key in complete_logbooks:
    if name_filter in logbooks[key]['name'].values:
        civil_logbooks[key] = complete_logbooks[key]
        
           
'filter 2022 A/C'
logbooks_2022 = {}

for key in civil_logbooks: 
    if ((civil_logbooks[key]['Discrepancy_stampedby_Processed'] >= pd.Timestamp('2022-01-01')) & (civil_logbooks[key]['Discrepancy_stampedby_Processed'] <= pd.Timestamp('2022-12-31'))).any():
        logbooks_2022[key] = civil_logbooks[key]

            
           
'count number of different aircraft types filtered logbooks'
type_counter_filtered = {category: 0 for category in unique_types}

for df in logbooks_2022.values():
    # Get the value in the first row of column A
    first_a_value = df.loc[0, 'LogBook_aircraftTypeVersion']
    # Update the counter for the appropriate category
    if first_a_value in unique_types:
        type_counter_filtered[first_a_value] += 1           

type_counter_filtered = pd.DataFrame(type_counter_filtered.items())
with pd.ExcelWriter(types) as writer:
    type_counter_filtered.to_excel(writer, index=True)



"""__________________________________________________________________________

Analysis
   __________________________________________________________________________
"""


'Discrep & tasks summary'
columnnames = ["Macro_tasks", "Micro_tasks","actype","file"]
discrep_task_summary = pd.DataFrame(columns = columnnames) 
            
for key in logbooks_2022: 
    actype = logbooks_2022[key]['LogBook_aircraftTypeVersion'].iloc[0]           
    unique_discrep = logbooks_2022[key]['Discrepancy_sequence'].unique().tolist()
    unique_discrep = len(unique_discrep)
    
    unique_tasks = logbooks_2022[key]['Task_taskElogbookId'].unique().tolist()
    unique_tasks = len(unique_tasks)
    
    discrep_task_summary_small = pd.DataFrame({
        "Macro_tasks": [unique_discrep],
        "Micro_tasks": [unique_tasks],
        "actype": [actype],
        "file": [key]
        })
    
    discrep_task_summary = pd.concat([discrep_task_summary,discrep_task_summary_small])   
    discrep_task_summary = discrep_task_summary.reset_index(drop=True) 

mean_macro = discrep_task_summary["Macro_tasks"].mean() 
mean_micro = discrep_task_summary["Micro_tasks"].mean() 

new_row = pd.DataFrame({"Macro_tasks": mean_macro, "Micro_tasks": mean_micro}, index=["mean"])
discrep_task_summary = discrep_task_summary.append(new_row) 

with pd.ExcelWriter(discrep_task_summary_file_2022) as writer:
    discrep_task_summary.to_excel(writer, index=True)

discrep_task_summary.to_excel("Discrep&Tasks_{0}_2022.xlsx".format(model), index=True) 


'search and calculate length of key activities'

columnnames = ["Pick_up_FAL","NCs from FAL", "GR release", "GR activity","FT release", "FT activity"]
main_activity_summary = pd.DataFrame(columns = columnnames) 
     
for key in logbooks_2022:
    actype = logbooks_2022[key]['LogBook_aircraftTypeVersion'].iloc[0]
    
    
    result1 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value]
    if result1.empty:
        pick_up_FAL = float('NaN')
    else:
        pick_up_FAL = result1['Discrepancy_time'].iloc[0]
    
    result2 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value2]
    if result2.empty:
        Beanstandungen = float('NaN')
    else: 
        Beanstandungen = result2['Discrepancy_time'].iloc[0]
        
    result3 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value3]
    if result3.empty:
        Bodenlauffreigabe = float('NaN')
    else:
        Bodenlauffreigabe = result3['Discrepancy_time'].iloc[0]
    
    result4 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value4]
    if result4.empty:
        Bodenlaufprogramm = float('NaN')
    else:
        Bodenlaufprogramm = result4['Discrepancy_time'].iloc[0]
    
    result6 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value6]
    if result6.empty:
        Flugfreigabe = float('NaN')
    else:
        Flugfreigabe = result6['Discrepancy_time'].iloc[0]  
    
    result7 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value7]
    if result7.empty:
        Einflugprogramm = float('NaN')
    else:
        Einflugprogramm = result7['Discrepancy_time'].iloc[0]
    
    
    main_activity_summary_small = pd.DataFrame({
        "actype": [actype],
        "file": [key],
        "Pick_up_FAL": [pick_up_FAL],
        "NCs from FAL": [Beanstandungen],
        "GR release": [Bodenlauffreigabe],
        "GR activity": [Bodenlaufprogramm],
        "FT release": [Flugfreigabe],
        "FT activity": [Einflugprogramm]})
    
    main_activity_summary = pd.concat([main_activity_summary,main_activity_summary_small])   
    main_activity_summary = main_activity_summary.reset_index(drop=True) 
            
            
mean_pickupFAL = main_activity_summary["Pick_up_FAL"].mean() 
mean_NCs = main_activity_summary["NCs from FAL"].mean() 
mean_GRrelease = main_activity_summary["GR release"].mean() 
mean_GR = main_activity_summary["GR activity"].mean() 
mean_FLrelease = main_activity_summary["FT release"].mean() 
mean_FL = main_activity_summary["FT activity"].mean() 


new_row = pd.DataFrame({"Pick_up_FAL": mean_pickupFAL, "NCs from FAL": mean_NCs, "GR release": mean_GRrelease,
                        "GR activity": mean_GR, "FT release": mean_FLrelease,"FT activity": mean_FL,
                        }, index=["mean"])
main_activity_summary = main_activity_summary.append(new_row)   

main_activity_summary.to_excel(Summary_file, index=True) 

columnnames = ['file']
main_activs = pd.DataFrame(columns = columnnames)     
main_discrep = {}   
main_discr = pd.DataFrame(columns = columnnames)
 
for key in logbooks_2022:
  
    result4 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value4]
    VEP_Date = result4['Discrepancy_createdby_Processed'].iloc[0]
    main_discrep[key] = logbooks_2022[key][logbooks_2022[key]['TaskToBeDone_processedDate'] == VEP_Date]
    
    main_unique_discrepancies = main_discrep[key]['Discrepancy_description'].unique().tolist() 
    
    for i in main_unique_discrepancies:
        uniquer_df = main_discrep[key][main_discrep[key]['Discrepancy_description'] == i]    
        unique_tasks = uniquer_df['TaskToBeDone_description'].unique().tolist() 
        
        start_d = VEP_Date
        end_d = uniquer_df['Task_stampedby_Processed'].max()
        duration_d = calculate_days(start_d,end_d)  
        
        main_discr_small = pd.DataFrame({
            "Activity": [i],
            "Start": [start_d],
            "End": [end_d], 
            "Duration": [duration_d]},
            index =[key])
        
        main_discr = pd.concat([main_discr,main_discr_small])        
        
        
        for j in unique_tasks:
            unique_task_df = uniquer_df[uniquer_df['TaskToBeDone_description'] == j]
            start = VEP_Date
            end = unique_task_df['Task_stampedby_Processed'].max()   
            duration = calculate_days(start,end)
            main_activity_summary_small = pd.DataFrame({
                "file": [key],
                "{0}_{1} [days]".format(i,j): [duration]})
    
            main_activs = pd.concat([main_activs,main_activity_summary_small])
     
    
    print ("{0} done!".format(key))
main_activs.to_excel(mean_file, index=True)    
 

'gantt'   
columnnames = ['file']   
 
for key in logbooks_2022:
    completed = pd.DataFrame(columns = columnnames)  
    result4 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value4]
    VEP_Date = result4['Discrepancy_createdby_Processed'].iloc[0]
    type_ac = logbooks_2022[key]["LogBook_aircraftTypeVersion"].iloc[0]
    
    unique_discrepancies = logbooks_2022[key]['Discrepancy_description'].unique().tolist() 
    main_unique_discrepancies = main_discrep[key]['Discrepancy_description'].unique().tolist() 
    
    for i in unique_discrepancies:
        uniquer_df = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == i]
        
        start_d = uniquer_df['TaskToBeDone_processedDate'].min()
        end_d = uniquer_df['Task_stampedby_Processed'].max()
        duration_d = calculate_days(start_d,end_d)  
        
        main_discr_small = pd.DataFrame({
            "Activity": [i],
            "Start": [start_d],
            "End": [end_d], 
            "Duration": [duration_d]},
            index =[key])
        
        completed = pd.concat([completed,main_discr_small])  
    
    completed['Activity'] = completed['Activity'].str.slice(stop=30)    
    colors = ['red' if a in main_unique_discrepancies else 'blue' for a in completed['Activity'].tolist()]
    'GANTT charts'
    fig = px.timeline(completed, x_start='Start', x_end='End', y='Activity', color=colors)
    fig.update_layout(
     width=2000,  # set the width to 1800 pixels
     height=2000,  # set the height to 600 pixels
     margin=dict(l=50, r=50, t=50, b=50),  # set the margins
     annotations=[
         dict(
             x=row['End'], 
             y=row['Activity'], 
             xref='x', 
             yref='y', 
             text='{:.1f} days'.format(row['Duration']), 
             showarrow=False, 
             font=dict(size=14),
             borderpad=4  # add some padding to the annotation box
         ) for _, row in completed.iterrows()
     ],
     title={
        'text': "{0}_{1}".format(filename,type_ac),  # set the title text
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
    
    gantt = "Gantt_complete_{0}.png".format(key)
    gantt = os.path.join(folder_name, gantt)
    fig.write_image(gantt, scale=1)   
        
        
    print ("{0} done!".format(key))
    
    
'do the open activities analysis'

for key in logbooks_2022:
    type_ac = logbooks_2022[key]["LogBook_aircraftTypeVersion"].iloc[0]
    result4 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value4]
    VEP_Date = result4['Discrepancy_createdby_Processed'].iloc[0]
    End_GR = result4['Discrepancy_stampedby_Processed'].max()
    
    result7 = logbooks_2022[key][logbooks_2022[key]['Discrepancy_description'] == search_value7]
    End_FT = result7['Discrepancy_stampedby_Processed'].max()
    
    highlight_dates = [VEP_Date, End_GR, End_FT]
    logbooks_2022[key]
    logbooks_2022[key] = logbooks_2022[key].sort_values('TaskToBeDone_processedDate')
    logbooks_2022[key] = logbooks_2022[key].reset_index(drop=True) 
    
    # Find the minimum and maximum timestamps
    min_time = logbooks_2022[key]['TaskToBeDone_processedDate'].min()
    max_time = logbooks_2022[key]['Task_stampedby_Processed'].max()
    
    time_range = pd.date_range(start=min_time, end=max_time, freq='D')
    
    # For each time point in the time range, count how many activities are open
    activity_counts = []
    for time_point in time_range:
        count = ((logbooks_2022[key]['TaskToBeDone_processedDate'] <= time_point) & (logbooks_2022[key]['Task_stampedby_Processed'] > time_point)).sum()
        activity_counts.append(count)
    
    # Plot the results
    
    fig, ax = plt.subplots()
    ax.plot(time_range, activity_counts)
    ax.set_xlabel('Calendar week')
    ax.set_ylabel('Number of open tasks')
    week_numbers = time_range.to_series().dt.isocalendar().week
    
    ax.set_xticklabels(week_numbers[::5].unique())  # set the tick labels
    
    ax.axvline(x=VEP_Date, color='red', linestyle='--', label = "VEP")
    ax.axvline(x=End_GR, color='green', linestyle='--',label = "End_GR")
    ax.axvline(x=End_FT, color='blue', linestyle='--',label = "End_FT")
    
    plt.title("{0}_{1}".format(key,type_ac))
    fig_name = "{0}_{1}.png".format(key,type_ac)
    fig_name = os.path.join(folder_name, fig_name)
    fig.savefig(fig_name,bbox_inches='tight')
    plt.show()
    




# completed = completed.sort_values('Start')
# completed = completed.reset_index(drop=True) 

# # Find the minimum and maximum timestamps
# min_time = completed['Start'].min()
# max_time = completed['End'].max()

# time_range = pd.date_range(start=min_time, end=max_time, freq='D')

# # For each time point in the time range, count how many activities are open
# activity_counts = []
# for time_point in time_range:
#     count = ((completed['Start'] <= time_point) & (completed['End'] > time_point)).sum()
#     activity_counts.append(count)

# # Plot the results
# plt.plot(time_range, activity_counts)
# plt.xlabel('Time')
# plt.ylabel('Number of open activities')
# plt.show()












# # Create a new dataframe with one column for each minute
# index = pd.date_range(min_time, max_time, freq='D')
# columns = ['minute_' + str(i) for i in range(len(index))]
# activity_df = pd.DataFrame(index=index, columns=columns)
# activity_df = activity_df.fillna(0)

# # Iterate over the rows in the original dataframe and set the corresponding minute columns in the new dataframe to 1 if the activity was open during that minute
# for _, row in completed.iterrows():
#     start = row['Start']
#     end = row['End']
#     activity = row['Activity']
#     activity_df.loc[start:end, 'minute_':] = np.where(activity_df.loc[start:end, 'minute_':] == 0, activity, activity_df.loc[start:end, 'minute_':])

# # Sum the values in each minute column to get the number of activities open at each minute
# activity_count = activity_df.sum()

# # Plot the results using a line graph or bar chart
# plt.bar(activity_count.index, activity_count.values)
# plt.xticks(rotation=90)
# plt.show()









































