# -*- coding:/ utf-8 -*-
"""
Created on Tue Jul 23 12:07:20 2019
This piece of software is bound by The MIT License (MIT)
Copyright (c) 2019 Prashank Kadam
Code written by : Prashank Kadam
User name - ADM-PKA187
Email ID : prashank.kadam@maersktankers.com
Created on - Mon Jul 29 09:17:59 2019
version : 1.1
"""  

# Importing the required libraries
import pandas as pd
import os
from pathlib import PureWindowsPath

def SisterVesselExtract(df, model_name):

    flag = ''
    df_final = pd.DataFrame(columns = ['Model name', 'Vessel name', 'IMO number', 'Sister vessel', 'Sister vessel IMO'])
    df_inter_1 = pd.DataFrame(columns = ['Model name', 'Vessel name', 'IMO number', 'Sister vessel'])
    df_inter_2 = pd.DataFrame(columns = ['Sister vessel IMO'])
    df_sister = pd.DataFrame(columns = ['Model name', 'Vessel name', 'IMO number', 'Sister vessel'])
    df_sister_imo = pd.DataFrame(columns = ['Sister vessel IMO'])

    for index, row in df.iterrows():
    
        if 'IMO number' in str(row['Unnamed: 0']): 
            imo_number = row['Value']
            for x in row[2:]:
                df_sister_imo.loc[0] = ({'Sister vessel IMO':str(x)})
                df_inter_2 = pd.concat([df_inter_2, df_sister_imo])
            continue
       
        if 'Vessel name' in str(row['Unnamed: 0']):
            vessel_name = row['Value'] 
            for x in row[2:]:                                          
                   df_sister.loc[0] = ({'Model name':model_name, 'Vessel name':vessel_name, 'IMO number':imo_number, 'Sister vessel':x})
                   df_inter_1 = pd.concat([df_inter_1, df_sister])
            flag = 'X'
                
        if flag == 'X':
            break                
    
    df_final = pd.concat([df_inter_1, df_inter_2], axis=1, sort=False)
    
    return df_final               

# Finding the path to the different excel sheets:
# Initializing a list to store the paths of various excel files        
paths = []

# Declaring the output dataframe:
df_final_data = pd.DataFrame(columns=['Model name', 'Vessel name', 'IMO number', 'Sister vessel'])

# Looping over the files in specified path 
for root, dirs, files in os.walk("M:\Fuel Opt\VesselPerformance\Vessel Information\Vessel Data\Static Data\Prashank\Sister_vessel"):
    for file in files:        
        # If the file ends with excel extensions we append its path to the path list
        if file.endswith(".xlsx"):
             s = os.path.join(root, file)
             paths.append(s)
        if file.endswith(".xls"):
             s = os.path.join(root, file)
             paths.append(s)         
        if file.endswith(".xlsm"):
             s = os.path.join(root, file)
             paths.append(s)             
            
# Looping over all the fetched paths:
for f in paths:
    
    p = PureWindowsPath(f)
    model_name = p.stem
    
    # Loading the excel on each path to a dataframe 
    df = pd.read_excel(f, 'General data')
    
    # Calling the extraction function 
    df1 = SisterVesselExtract(df, model_name)
    
    # If the file is processed correctly, we concatenate the extracted
    # data to the final dataframe
    df_final_data = pd.concat([df_final_data,df1])
        
    # Now we remove the file post processing from the current folder
    os.remove(f)
    
# After all the data has been extracted into a single dataframe we export the 
# dataframe to an excel sheet        
df_final_data.to_excel('M:\Fuel Opt\VesselPerformance\Vessel Information\Vessel Data\Static Data\Prashank\Sister_vessel\\final.xlsx')
