# -*- coding:/ utf-8 -*-
"""
This piece of software is bound by The MIT License (MIT)
Copyright (c) 2019 Prashank Kadam
Code written by : Prashank Kadam
User name - ADM-PKA187
Email ID : prashank.kadam@maersktankers.com
Created on - Thu Nov 7 2019
version : 1.0
"""

# We use gmplot, google's open source library to plot scatter plots on google maps
# installing gmplot - $ pip install gmplot

# Importing the required libraries:
import gmplot
import pandas as pd
from math import sin, cos, sqrt, atan2, radians

# Importing the data from excel sheet
# This data contains the latitude, logitude and the vessel name of the vessel whose
#  voyage needs to be traced
df = pd.read_excel('Lat_long.xlsx')

# Selecting the list of unique vessels from the dataframe
vessels = df['Vessel_name'].unique()

# Plotting the center of the intitial perspective from which our map should open
#  in the browser along with the required zoom
gmap3 = gmplot.GoogleMapPlotter(df['Latitude'].mean(),
                                df['Longitude'].mean(), 5)

# Looping over all individual vessel
for n in vessels:

    # Fetching the data for the vessel in the current iteration of the dataframe
    df_temp = df[df['Vessel_name'] == n]

    # Converting a pandas series to a list so that it can be given as an input to
    #  the gmap plot
    latitude_list = df_temp['Latitude'].to_list()
    longitude_list = df_temp['Longitude'].to_list()

    # scatter method of map object
    # scatter points on the google map
    gmap3.scatter(latitude_list, longitude_list, 'cornflowerblue',
                  size=10000, marker=False)

    # Plot method Draw a line in
    # between given coordinates
    gmap3.plot(latitude_list, longitude_list,
               'cornflowerblue', edge_width=2.5)

    # approximate radius of earth in km
    R = 6373.0

    # Initialization of the final distance variable
    final_dist = 0

    # Initialization of the lat and long temporary variables
    lat_temp = 'X'
    long_temp = 'X'

    # Looping over all the data for the vessel in the current iteration
    for index, row in df_temp.iterrows():

        # If the loop is in the first iteration, we set the values of the temporary variables
        # based on the lat and long values of the first iteration and continue the loop
        if lat_temp == 'X' and long_temp == 'X':
            lat_temp = row['Latitude']
            long_temp = row['Longitude']
            continue

        # Initializing the lat and long variable which we will be using for calculation and
        # converting them into radians
        lat1 = radians(row['Latitude'])
        lon1 = radians(row['Longitude'])
        lat2 = radians(lat_temp)
        lon2 = radians(long_temp)

        # Calculating the difference between the lats and longs of the two points
        dlon = lon2 - lon1
        dlat = lat2 - lat1

        # Calculating the actual distance between the two positions i.e. in the previous
        # and the current iteration of the loop
        a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
        c = 2 * atan2(sqrt(a), sqrt(1 - a))

        distance = R * c

        # Adding the distance calculated in the current iteration to the final distance
        final_dist = final_dist + distance

        # Setting the values of temporary lat and long for the next iteration
        lat_temp = row['Latitude']
        long_temp = row['Longitude']

        # Adding a visual tool tip to the point plotted containing the vessel name
        gmap3.marker(row['Latitude'], row['Longitude'], title=n)

    print(n, '->', final_dist, 'km')
# Currently we use the free version of google maps, this will show a developer
# watermark on the map. Inorder to remove the watermark, you need to purchase the
# API credential from google cloud platform and add the API key to the code
# snippet below
# gmap3.apikey = 'Enter API' #Google cloud platform

# Drawing the map to an HTML file, not that this file will be stored in the same
# location as this code file. Opening the HTML file will open your map in the
# browser
gmap3.draw("geoplot.html")

