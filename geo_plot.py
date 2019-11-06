# import gmplot package 
import gmplot
import pandas as pd

df = pd.read_excel('Lat_long.xlsx')

latitude_list = df['Latitude'].to_list()
longitude_list = df['Longitude'].to_list()

gmap3 = gmplot.GoogleMapPlotter(df['Latitude'].mean(),
                                df['Longitude'].mean(), 5)

# scatter method of map object  
# scatter points on the google map 
gmap3.scatter(latitude_list, longitude_list, '# FF0000',
              size=40, marker=True)

# Plot method Draw a line in 
# between given coordinates 
gmap3.plot(latitude_list, longitude_list,
           'cornflowerblue', edge_width=2.5)

gmap3.draw("map13.html")
