# -*- coding:/ utf-8 -*-
"""
Created on Tue Jul 23 12:07:20 2019
This piece of software is bound by The MIT License (MIT)
Copyright (c) 2019 Prashank Kadam
Code written by : Prashank Kadam
User name - ADM-PKA187
Email ID : prashank.kadam@maersktankers.com
Created on - Tue Jul 30 10:00:14 2019
version : 1.0
"""

# Importing the required libraries
import pandas as pd
import xml.etree.ElementTree as et

# Importing the standard semark light format for sea and port reports inorder to ger the column
# names of all the standard columns into runtime
df_init_sea = pd.read_excel('semark_light.xlsx', 'Sea')
df_init_port = pd.read_excel('semark_light.xlsx', 'Port')

# Importing the xml converted data into a dataframe
df_target = pd.read_excel('test_1.xlsx')

# Taking a subset of only the required columns from the dataframe
df_target = df_target[['ImoNumber', 'VesselName', 'ReportTime', 'Longitude', 'Port',
                       'Location', 'Latitude', 'VoyageNo', 'ObservedDistance', 'FWDDraft', 'LOG_DISTANCE',
                       'WindForce', 'SeaDir', 'SwellHeight', 'CurrentDirection', 'WindDirection', 'SeaHeight',
                       'SwellDir', 'SeaState', 'Swell', 'Current', 'FuelType', 'AuxEngineConsumption',
                       'BoilerEngineConsumption', 'Units', 'Received', 'Consumption', 'SeaTemp', 'VesselCondition']]

# Initializing the rows list in which we will append the mapped data
rows = []

# Looping over the dataframe to map each row to the corresponding columns in the other dataframe
for index, row in df_target.iterrows():
    # Kindly note that the time complexity of the below code higher than the optimum as the same data values are
    # repeated for 4 rows in a succession but since the time is not an issue in our particular use case we can
    # refrain from adding further validations to complicate the code
    s_imo = row['ImoNumber']
    s_vesselname = row['VesselName']
    s_time = row['ReportTime']
    s_longitutde = row['Longitude']
    s_port = row['Port']
    s_latitude = row['Latitude']
    s_voyage = row['VoyageNo']
    s_obvdis = row['ObservedDistance']
    s_draught = row['FWDDraft']
    s_dist = row['LOG_DISTANCE']
    s_wind = row['WindForce']
    s_seadir = row['SeaDir']
    s_swellhgt = row['SeaDir']
    s_curdir = row['CurrentDirection']
    s_windir = row['WindDirection']
    s_seahgt = row['SeaHeight']
    s_swelldir = row['SwellDir']
    s_seastate = row['SeaState']
    s_swell = row['Swell']
    s_curr = row['Current']
    s_units = row['Units']
    s_seatemp = row['SeaTemp']
    s_vesscon = row['VesselCondition']

    # Filling in the corresponding fields for the repective fuel types:

    if row['FuelType'] == 'IFO':
        s_hshfo_ae = row['AuxEngineConsumption']
        s_hshfo_blr = row['BoilerEngineConsumption']
        s_hshfo_me = row['Consumption']

    elif row['FuelType'] == 'LSF':
        s_lshfo_ae = row['AuxEngineConsumption']
        s_lshfo_blr = row['BoilerEngineConsumption']
        s_lshfo_me = row['Consumption']

    elif row['FuelType'] == 'LSG':
        s_lsmdo_ae = row['AuxEngineConsumption']
        s_lsmdo_blr = row['BoilerEngineConsumption']
        s_lsmdo_me = row['Consumption']

    elif row['FuelType'] == 'MGO':
        s_hsmdo_ae = row['AuxEngineConsumption']
        s_hsmdo_blr = row['BoilerEngineConsumption']
        s_hsmdo_me = row['Consumption']

        # Since MGO is the last fuel type for a particular report, we append the mapped values to the
        # other dataframe
        rows.append({'Vessel_Name': s_vesselname, 'Report_Date': s_time, 'IMO_NO': s_imo,
                     'Main Engine Fuel Consumption (H.S.HFO)': s_hshfo_me,
                     'Main Engine Fuel Consumption (L.S.HFO)': s_lshfo_me,
                     'Main Engine Fuel Consumption (H.S.MDO)': s_hsmdo_me,
                     'Main Engine Fuel Consumption (L.S.MDO)': s_lsmdo_me,
                     'Boiler Consumption (H.S.HFO)': s_hshfo_blr,
                     'Boiler Consumption (L.S.HFO)': s_lshfo_blr,
                     'Boiler Consumption (H.S.MDO)': s_hsmdo_blr,
                     'Boiler Consumption (L.S.MDO)': s_lsmdo_blr,
                     'Auxiliary Engine (Diesel Generator ) (H.S.HFO)': s_hshfo_ae,
                     'Auxiliary Engine (Diesel Generator ) (L.S.HFO)': s_lshfo_ae,
                     'Auxiliary Engine (Diesel Generator ) (H.S.MDO)': s_hsmdo_ae,
                     'Auxiliary Engine (Diesel Generator ) (L.S.MDO)': s_lsmdo_ae,
                     'Vessel State( Loaded\Ballast)': s_vesscon, 'True Wind Direction ': s_windir,
                     'True Wave Direction': s_seadir, 'True Swell Direction': s_swelldir})

# Creating the final dataframe with the mapped data and using the column values mapped from the semark sheet
df_final = pd.DataFrame(rows, columns=df_init_sea.columns)

# Exporting the data to excel sheet
df_final.to_excel('final.xlsx', index=False)

######################################################################################################
# The below piece of code is for xml to pandas dataframe conversion.
# Kindly note that all the fields have not yet been added to the dictionary
# xtree = et.parse("vess_test.xml")

# xroot = xtree.getroot()
#
# df_cols = ["VesselName", "ReportTime", "Longitude", "Port", "IMO_NUMBER"]
# rows = []
#
# for node in xroot:
#     # s_name = node.attrib.get("name")
#     s_vess_name = node.find("VesselName").text if node is not None else None
#     s_report_time = node.find("ReportTime").text if node is not None else None
#     s_longitude = node.find("Longitude").text if node is not None else None
#     s_port = node.find("Port").text if node is not None else None
#     s_imo = node.find("IMO_NUMBER").text if node is not None else None
#     # s_location = node.find("Location").text if node is not None else None
#
#     rows.append({"VesselName": s_vess_name, "ReportTime": s_report_time, "Longitude":s_longitude,
#                  "Port": s_port, "IMO_NUMBER": s_imo})
#
# out_df = pd.DataFrame(rows, columns=df_cols)
#
# print(out_df.head(10))
