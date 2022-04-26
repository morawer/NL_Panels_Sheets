import os
import pandas as pd

#Sheets sizes
width = 841
height = 595

A4 = [width, height]

#Destination folder of PDF sheets and Excel files created
pathDestination = "/home/dani/Projects/NL_Panels_Sheets/destination"

#The info input is trough of excel file
excel = "/home/dani/Projects/NL_Panels_Sheets/excel/Overview panels ESP 0426.xlsx"
df = pd.read_excel(excel, sheet_name="Database 20220421")

#If destination folder does not exist, the folder is created automatically
if not os.path.exists(pathDestination):
    os.makedirs(pathDestination)

#In "SEGUIMIENTO_PEDIDOS.xlsm" file we get in variables the differents columns values we need.
mo = df["Ref order no"].values
co = df["CO España"].values
itemNumber = df["Item no"].values
itemName = df["Item name"].values
po = df["Order no PO"].values
date = df["New Date ES"].values
qty = df["Ordered qty"].values

for line in range(len(mo)):
    print(itemName[line] + ' >>> ' + str(date[line]))

