# -*- coding:/ utf-8 -*-
"""
Created on Tue Jul 23 12:07:20 2019
This piece of software is bound by The MIT License (MIT)
Copyright (c) 2019 Prashank Kadam
Code written by : Prashank Kadam
User name - ADM-PKA187
Email ID : prashank.kadam@maersktankers.com
Created on - Tue Jul 30 09:12:00 2019
version : 1.0
"""  

import pyodbc 
import pandas as pd

conn = pyodbc.connect('Driver={########################};'
                      'Server=##########################;'
                      'Database=########################;'
                      'uid=####################;'
                      'pwd=######################;'
                      'Trusted_Connection=###;')

cursor = conn.cursor()

#Getting all the tables in the database
#for table_name in cursor.tables(tableType='TABLE'):
    #print(table_name)

# Coverting the SELECT query into a pandas dataframe
sql = 'SELECT TOP 10 * FROM db.table'
data = pd.read_sql(sql, conn)

print(data.columns)

data.to_excel(r'C:\Users\ADM-PKA187\Desktop\Dastabase\db_test.xlsx')

    
    
 
