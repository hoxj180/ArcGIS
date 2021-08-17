#!/usr/bin/env python
# coding: utf-8

# In[ ]:


##IMPORTANT: to run this code, you must input the paths to your ArcGIS and excel files where noted below. 
##Your ArcGIS folder should look something like '..../ArcGIS/Default.gdb/' unless you are putting it in a geodatabase (gdb)
##other than the Default. For both paths, do not include the file name. End both paths with a '/'.

import numpy as np
import pandas as pd
import os
import arcpy
from arcpy import env
import time

startTime = time.time()
#clear previous tables
arcpy.env.overwriteOutput = True 

##ATTENTION: replace ###### the path to your ArcGIS folder below.
ArcDir = r'######'

#ATTENTION: replace ###### with the path to your excel files below.
homeDir = r'######'

supplyLines = os.path.join(ArcDir,'supplyLines')
arcCoords = os.path.join(ArcDir,'arcCoords')
files = [supplyLines,arcCoords] #testvalue?

for file in files: 
    if arcpy.Exists(file):
        arcpy.Delete_management(file)
        
#initialize geodatabase

tradeDataPath = os.path.join(homeDir, 'TestTradeData.xlsx')
capitalsPath = os.path.join(homeDir, 'Capitals.xlsx')

tradeData = pd.read_excel(tradeDataPath)
capitals = pd.read_excel(capitalsPath) 

columnHeadings = {'Origin Lat':[],'Origin Long':[],'Dest Lat':[],'Dest Long':[],
                  'Strategic Commodity':[], 'Value':[],'Quantity (kg)':[]}
supplyData = pd.DataFrame(columnHeadings) 
for i in range(0,len(tradeData)): #scanning each city in the trade data
    Quantity = tradeData.iloc[i,7]
    Strat_Comm = tradeData.iloc[i,11]
    Value = tradeData.iloc[i,12]
    
    for j in range(0,len(capitals)): #scanning the city in the capital database for each trade data city
        if tradeData.iloc[i,5] == capitals.iloc[j,0]: #assign coordinates if the capital and trade data partner city line up. 
            partnerLat = capitals.iloc[j,2]
            partnerLong = capitals.iloc[j,3]
            country = capitals.iloc[j,0]
        if tradeData.iloc[i,10] == capitals.iloc[j,0]: #assign coordinates if the capital and trade data reporter city line up. 
            reporterLat = capitals.iloc[j,2]
            reporterLong = capitals.iloc[j,3]
            country = capitals.iloc[j,0]
    if tradeData.iloc[i,2] == 'Export': #put coordinates in correct column depending on whether the goods are exported or imported. 
        iCountry = tradeData.iloc[i,5]
        eCountry = tradeData.iloc[i,10]
        row = [{'Dest Lat':partnerLat,'Dest Long':partnerLong, 'Origin Lat': reporterLat, 'Origin Long': reporterLong, 
               'Strategic Commodity':Strat_Comm, 'Value':Value, 'Quantity (kg)': Quantity, 'Importer': iCountry,
               'Exporter': eCountry}]
    elif tradeData.iloc[i,2] == 'Import': 
        iCountry = tradeData.iloc[i,10]
        eCountry = tradeData.iloc[i,5]
        row = [{'Origin Lat': reporterLat, 'Origin Long': reporterLong, 'Dest Lat':partnerLat,
               'Dest Long':partnerLong,'Strategic Commodity':Strat_Comm, 'Value':Value, 'Quantity (kg)': Quantity,
               'Importer': iCountry, 'Exporter': eCountry}]
    supplyData = supplyData.append(row)
print("dataframe: ", supplyData)
#potentially use ExceltoTable to shorten processing time

#create numpy array and export to ArcGIS
npSupplyData = supplyData.as_matrix(['Origin Lat','Origin Long', 'Dest Lat', 'Dest Long', 'Strategic Commodity',
                                    'Value', 'Quantity (kg)', 'Importer', 'Exporter']) #pandas 0.18.1
#npSupplyData = supplyData.to_numpy() #pandas 1.2.4

#creating structured array to export to ArcGIS

dts = np.dtype({'names': ['Origin Lat', 'Origin Long', 'Dest Lat', 'Dest Long',
                          'Strategic Commodity', 'Value', 'Quantity (kg)','Importer','Exporter'], 
                'formats': ['f8', 'f8', 'f8', 'f8','S30','f8','f8','S20','S20']})

coordtable = np.rec.fromrecords(npSupplyData, dtype=dts)
#print('arcCoords: ', arcCoords)

#creating feature layer of supply lines
out_lines = os.path.join(ArcDir,'supplyLines')

#convert array into geodatabase table
arcpy.da.NumPyArrayToTable(coordtable, arcCoords)

#create lines
lines = arcpy.XYToLine_management(arcCoords, out_lines,'Origin_Long','Origin_Lat','Dest_Long','Dest_Lat','GEODESIC')

#join table with XY to Line
arcpy.JoinField_management("supplyLines","OID",arcCoords,"OBJECTID",
                           ["Exporter","Importer","Strategic_Commodity","Value","Quantity__kg_"])

#assign symbology to attributes
#for some reason this command is going through but not changing the symbology. I can manipulate other tables'
#symbology this way no problem. I wonder if you have to save it first or something. 

#establish workspace environment
#env.workspace = homeDir

#create symbology layering - to be improved
   # subfolder = 'ArcGIS Project/'
   # mapName = 'Test.mxd'
   # mapLayer = os.path.join(homeDir,subfolder,mapName)
   # print("mapLayer: ", mapLayer)
   # mxd = arcpy.mapping.MapDocument("CURRENT") 
   # df = arcpy.mapping.ListDataFrames(mxd, "*")[0]  
   # lyrPath = arcpy.mapping.Layer(os.path.join(ArcDir,'Test_Value.lyr'))
   # layer = arcpy.mapping.AddLayer(df, lyrPath)
   # #arcpy.ApplySymbologyFromLayer_management("supplyLines",'USDandSCsymbology.lyr')
   # arcpy.ApplySymbologyFromLayer_management("supplyLines",lyrPath)

totalTime = time.time() - startTime
print('Script finished in ' + str(totalTime) +" s")

