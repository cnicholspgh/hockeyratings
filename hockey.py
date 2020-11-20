import os, xlsxwriter, xlrd, openpyxl
import pandas as pd
import numpy as np

#print("Current Working Directory " , os.getcwd())
os.chdir('/Users/cnichols/Desktop/python/hockey')
#print("Current Working Directory " , os.getcwd())

#CALLS FANTRAX CSV FILE LOADS INTO DATA FRAME
file = 'fantasy.csv'
xl = pd.read_csv(file, header=0)
df = pd.DataFrame(xl)

#CALLS CORSI CSV FILE LOADS INTO DATA FRAME
#MUST CHANGE IN CORSI.CSV SHEET - CF, CF Rel, FF, FF REL, OZ SHARE
corsi_data = 'corsi.csv'
cors = pd.read_csv(corsi_data, header=0)
corsi = pd.DataFrame(cors)


#identifies headers of the columns for both lists
corsi_headers = list(cors.columns)
fantrax_headers = list(df.columns)

#LOADS DATAFRAME INTO LIST 
nd = df.values.tolist()
cor = corsi.values.tolist()

#merges data frame where names('Player') match
mergedStuff = pd.merge(df, corsi, on=['Player'], how='inner')
mergedStuff.head()
merge = mergedStuff.values.tolist()
combined_master = pd.DataFrame(mergedStuff)

for col in combined_master.columns:
  print(col)

#List created for output location
master = []

#Pulls from combined_master DF and puts it in new organzied list
for index, row in combined_master.iterrows():
    master_merge = [row.Player, row.Team, row.FRK, row.Age_x, row.Status, row.Salary, row.Contract, row.GP, row.TFP, row.FPG, row.CorsiFor, row.CFRel, row.Ffor, row.FFRel, row.ozshare]
    master.append(master_merge)
    
#puts new organized list in DF - will use this master to export
master_df = pd.DataFrame(master)
master_df.columns = ['Player', 'Team', 'Fantrax Rank', 'Age', 'Status', 'Salary', 'Contract', 'Games Played', 'Total Fantasy Points', 'FP/Game','CF','CF Rel', 'FF','FF Rel','OZ Share']

#Creates Excel Sheet
writer = pd.ExcelWriter('hockey_output.xlsx', engine='xlsxwriter')

#Creates and sorts Agent list by Games played and CF
free_agents = pd.DataFrame(master_df[master_df['Status'].str.contains('FA|W ')])
free_final = free_agents.sort_values(by=['Games Played','CF'], ascending=False)
print(free_final)

#Creates and sorts Taken list by Games played and CF
taken_players = pd.DataFrame(master_df[~master_df['Status'].str.contains('FA|W ')])
taken_final = taken_players.sort_values(by=['CF','FP/Game'], ascending=False)
print(taken_final)

# Convert the dataframe to an XlsxWriter Excel object.
free_final.to_excel(writer, sheet_name='Free Agents', index=False)

# Convert the dataframe to an XlsxWriter Excel object.
taken_final.to_excel(writer, sheet_name='Rostered Players', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()




