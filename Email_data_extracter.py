# -*- coding: utf-8 -*-
"""
Created on Tue Jul 23 12:07:20 2019
This piece of software is bound by The MIT License (MIT)
Copyright (c) 2019 Prashank Kadam
Code written by : Prashank Kadam
Email ID : prashank.kadam@maersktankers.com
version : 1.0
"""
# Importing the required libraries:
import pandas as pd
import os
import math

# Defining a function to extract data from the excel sheet,
# this function will run in a loop for each excel sheet and the 
# output extracted data for  each excel will then be merged with
# the final excel sheet
def emailExtract(static_df):
    
    # Declaring a variable which will later on flag out the invalid data
    invalid = ''       
    
    # Now we check if the number of columns in the dataframe matches our
    # standard format which is 4 columns, if the number number of columns
    # is less other than 4 the invalid flag is set and the function exits
    if len(static_df.columns) != 4:
        invalid = 'X'
        return invalid
    
    # Redefining the names of the columns in the dataframe:
    static_df.columns = ['Key', 'Value_1', 'Value_2', 'Value_3']
    
    # Declaring flags - Here we set the values of a few flags as the data 
    # format varies for a large number of files, if non-initialized variables 
    # are used, interpreter may throw an exception
    c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16 =\
    0, 0, 0, 0 ,0 ,0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    Idle_aux_HFO, Load_aux_HFO, Disc_aux_HFO, Hot_aux_HFO, Cold_aux_HFO,\
    Idle_aux_MDO, Load_aux_MDO, Disc_aux_MDO, Hot_aux_MDO, Cold_aux_MDO,\
    Idle_blr_HFO, Load_blr_HFO, Disc_blr_HFO, Hot_blr_HFO, Cold_blr_HFO,\
    Idle_blr_MDO, Load_blr_MDO, Disc_blr_MDO, Hot_blr_MDO, Cold_blr_MDO,\
    Idle_inr_HFO, Load_inr_HFO, Disc_inr_HFO, Hot_inr_HFO, Cold_inr_HFO,\
    Idle_inr_MDO, Load_inr_MDO, Disc_inr_MDO, Hot_inr_MDO, Cold_inr_MDO =\
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
    
    # Initializing the vessel name with a constant in case the vessel name 
    # value is not entered
    vessel_name = 'NaF'    
    
    # Looping over the dataframe - We now loop iver the dataframe and parse the 
    # whole data frame only once for optimal performance. For each iteration,
    # we extract the necessary data and once that is obtained, we continue the 
    # loop
    for index, row in static_df.iterrows():
  
  # Checking for the vessel name in the dataframe key
      if 'Vessel name' in str(row['Key']):
          
          # Here we check if the vessel name is empty or not, we are using a
          # float concersion operator due to the incosistencies in data
          try: 
              if math.isnan(float(row['Value_1'])) == True:
                  invalid = 'X'
                  return invalid
          except:
              # Assigning the vessel name to a local variable
              vessel_name = row['Value_1']
  
  # Checking the fuel consumption at Idle anchorage/drifting 
      if 'Idle' in str(row['Key']):
          continue
  # Assigning the auxillary consumption(HFO/MDO) to local variables
      if c1 == 0 and row['Key'] == 'Auxillary Engine':  
          Idle_aux_HFO = row['Value_2']
          Idle_aux_MDO = row['Value_3']
          c1 = 1
          continue    
  # Assigning the boiler consumption(HFO/MDO) to local variables  
      if c2 == 0 and row['Key'] == 'Boiler':
          Idle_blr_HFO = row['Value_2']
          Idle_blr_MDO = row['Value_3']
          c2 = 1
          continue
  # Assigning the Inert gas generator consumption(HFO/MDO) to local variables    
      if c3 == 0 and row['Key'] == 'Inertgas generator':
          Idle_inr_HFO = row['Value_2']
          Idle_inr_MDO = row['Value_3']
          c3 = 1
          continue
          
  # Checking the fuel consumption at Loading
      if 'Loading' in str(row['Key']):
          continue 
  # Assigning the auxillary consumption(HFO/MDO) to local variables  
      if c4 == 0 and row['Key'] == 'Auxillary Engine':  
          Load_aux_HFO = row['Value_2']
          Load_aux_MDO = row['Value_3']
          c4 = 1
          continue
  # Assigning the boiler consumption(HFO/MDO) to local variables  
      if c5 == 0 and row['Key'] == 'Boiler':
          Load_blr_HFO = row['Value_2']
          Load_blr_MDO = row['Value_3']
          c5 = 1
          continue
  # Assigning the Inert gas generator consumption(HFO/MDO) to local variables  
      if c6 == 0 and row['Key'] == 'Inertgas generator':
          Load_inr_HFO = row['Value_2']
          Load_inr_MDO = row['Value_3']
          c6 = 1   
  
  # Checking the fuel consumption at Discharge
      if 'Discharge' in str(row['Key']):
          continue 
  # Assigning the auxillary consumption(HFO/MDO) to local variables  
      if c7 == 0 and row['Key'] == 'Auxillary Engine':  
          Disc_aux_HFO = row['Value_2']
          Disc_aux_MDO = row['Value_3']
          c7 = 1
          continue
  # Assigning the boiler consumption(HFO/MDO) to local variables  
      if c8 == 0 and row['Key'] == 'Boiler':
          Disc_blr_HFO = row['Value_2']
          Disc_blr_MDO = row['Value_3']
          c8 = 1
          continue
  # Assigning the Inert gas generator consumption(HFO/MDO) to local variables  
      if c9 == 0 and row['Key'] == 'Inertgas generator':
          Disc_inr_HFO = row['Value_2']
          Disc_inr_MDO = row['Value_3']
          c9 = 1 
          
  # Skipping the unwanted consumptions until the hot wash consumption is iterated
      if 'Hot' not in str(row['Key']) and c10 == 0:
          continue
  
  # Checking the fuel consumption at Hot Wash
      if 'Hot' in str(row['Key']):
          c10 = 1
          continue 
  # Assigning the auxillary consumption(HFO/MDO) to local variables  
      if c11 == 0 and row['Key'] == 'Auxillary Engine':  
          Hot_aux_HFO = row['Value_2']
          Hot_aux_MDO = row['Value_3']
          c11 = 1
          continue
  # Assigning the boiler consumption(HFO/MDO) to local variables  
      if c12 == 0 and row['Key'] == 'Boiler':
          Hot_blr_HFO = row['Value_2']
          Hot_blr_MDO = row['Value_3']
          c12 = 1
          continue
  # Assigning the Inert gas generator consumption(HFO/MDO) to local variables  
      if c13 == 0 and row['Key'] == 'Inertgas generator':
          Hot_inr_HFO = row['Value_2']
          Hot_inr_MDO = row['Value_3']
          c13 = 1       
  
  # Checking the fuel consumption at Cold Wash
      if 'cold' in str(row['Key']):
          continue 
  # Assigning the auxillary consumption(HFO/MDO) to local variables  
      if c14 == 0 and row['Key'] == 'Auxillary Engine':  
          Cold_aux_HFO = row['Value_2']
          Cold_aux_MDO = row['Value_3']
          c14 = 1
          continue
  # Assigning the auxillary consumption(HFO/MDO) to local variables  
      if c15 == 0 and row['Key'] == 'Boiler':
          Cold_blr_HFO = row['Value_2']
          Cold_blr_MDO = row['Value_3']
          c15 = 1
          continue
  # Assigning the Inert gas generator consumption(HFO/MDO) to local variables  
      if c16 == 0 and row['Key'] == 'Inertgas generator':
          Cold_inr_HFO = row['Value_2']
          Cold_inr_MDO = row['Value_3']
          c16 = 1   
  # After parsing through all the data, break the loop
      break 
  
    # After parsing the whole dataframe we check if the vessel name is has been 
    # changed, if it is still the default value, we raise the invalid flag
    if vessel_name == 'NaF':
        invalid = 'X'
        return invalid
    
    # If the invalid flag is not set, we will append all the extracted data
    # to a single dataframe which will be returned by this function
    if invalid != 'X':
        
        # Creating a new dataframe having the structure of the output structure
        static_final_df = pd.DataFrame(columns=['Vessel', 'Equipment', 'Type', 'Idle', 'Loading', 'Discharge', 'Hot', 'Cold'])
        
        # Clearing all the data 
        static_final_df.iloc[0:0]
      
        #Appending the data extracted from the excel sheet to a dataframe:
        static_final_df.loc[0] = ({'Vessel':vessel_name, 'Equipment':'Auxillary Engine', 'Type':'HFO', 'Idle':Idle_aux_HFO, 'Loading':Load_aux_HFO, 'Discharge':Disc_aux_HFO, 'Hot':Hot_aux_HFO, 'Cold':Cold_aux_HFO})
        static_final_df.loc[1] = ({'Vessel':vessel_name, 'Equipment':'Auxillary Engine', 'Type':'MDO', 'Idle':Idle_aux_MDO, 'Loading':Load_aux_MDO, 'Discharge':Disc_aux_MDO, 'Hot':Hot_aux_MDO, 'Cold':Cold_aux_MDO})
        static_final_df.loc[2] = ({'Vessel':vessel_name, 'Equipment':'Boiler', 'Type':'HFO', 'Idle':Idle_blr_HFO, 'Loading':Load_blr_HFO, 'Discharge':Disc_blr_HFO, 'Hot':Hot_blr_HFO, 'Cold':Cold_blr_HFO})
        static_final_df.loc[3] = ({'Vessel':vessel_name, 'Equipment':'Boiler', 'Type':'MDO', 'Idle':Idle_blr_MDO, 'Loading':Load_blr_MDO, 'Discharge':Disc_blr_MDO, 'Hot':Hot_blr_MDO, 'Cold':Cold_blr_MDO})
        static_final_df.loc[4] = ({'Vessel':vessel_name, 'Equipment':'Inertgas generator', 'Type':'HFO', 'Idle':Idle_inr_HFO, 'Loading':Load_inr_HFO, 'Discharge':Disc_inr_HFO, 'Hot':Hot_inr_HFO, 'Cold':Cold_inr_HFO})
        static_final_df.loc[5] = ({'Vessel':vessel_name, 'Equipment':'Inertgas generator', 'Type':'MDO', 'Idle':Idle_inr_MDO, 'Loading':Load_inr_MDO, 'Discharge':Disc_inr_MDO, 'Hot':Hot_inr_MDO, 'Cold':Cold_inr_MDO})
      
        return static_final_df  

# Finding the path to the different excel sheets:
# Initializing a list to store the paths of various excel files        
paths = []

# Declaring the output dataframe:
df_final = pd.DataFrame(columns=['Vessel', 'Equipment', 'Type', 'Idle', 'Loading', 'Discharge', 'Hot', 'Cold'])

# Looping over the files in specified path 
for root, dirs, files in os.walk("M:\Fuel Opt\VesselPerformance\Vessel Information\Vessel Data\Static Data\Attachment_data"):
    for file in files:
        
        # If the file ends with excel extensions we append its path to the path list
        if file.endswith(".xlsx"):
             s = os.path.join(root, file)
             paths.append(s)
        if file.endswith(".xls"):
             s = os.path.join(root, file)
             paths.append(s)          
            
# Looping over all the fetched paths:
for f in paths:
    
    # Loading the excel on each path to a dataframe 
    df = pd.read_excel(f)
    
    # Calling the extraction function 
    df1 = emailExtract(df)
    try:
        # If the function returns an invalid flag, we continue the loop
        if df1 == 'X':
            continue
    except:
        # If the file is processed correctly, we concatenate the extracted
        # data to the final dataframe
        df_final = pd.concat([df_final,df1])
        
        # Now we remove the file post processing from the current folder
        os.remove(f)

# After all the data has been extracted into a single dataframe we export the 
# dataframe to an excel sheet        
df_final.to_excel('M:\Fuel Opt\VesselPerformance\Vessel Information\Vessel Data\Static Data\Attachment_data\\final.xlsx')