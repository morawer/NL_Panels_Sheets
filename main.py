import os
import pandas as pd
from reportlab.pdfgen import canvas
import openpyxl

#Sheets sizes
width = 841
height = 595
A4 = [width, height]

#Declare variables
moNew = 0

def msg_error(co, line, mo):
    print("===========================================================================")
    print("   |>>> ERROR AL CREAR ==> CO: " + str(co[line]) + "_ MO:" + str(mo[line]) + " <<<|  ")
    print("===========================================================================")

#Destination folder of PDF sheets and Excel files created
pathDestination = "/home/dani/Projects/NL_Panels_Sheets/destination/"

#The info input is trough of excel file
excel = "/home/dani/Projects/NL_Panels_Sheets/excel/Overview panels ESP 0426.xlsx"
df = pd.read_excel(excel, sheet_name="Database 20220421")




#If destination folder does not exist, the folder is created automatically
if not os.path.exists(pathDestination):
    os.makedirs(pathDestination)

#In "Overview panels ESP 0426.xlsx" file we get in variables the differents columns values we need.
mo = df["Ref order no"].values
co = df["CO España"].values
itemNumber = df["Item no"].values
itemName = df["Item name"].values
po = df["Order no PO"].values
date = df["New Date ES"].values
qty = df["Ordered qty"].values
week = df["Week"].values

for line in range(len(mo)):
    
    if mo[line] != moNew:
        try:
            pdf_panelsNL = canvas.Canvas(
                pathDestination + str(mo[line]) + "_" + "PanelsNL" + ".pdf", pagesize=A4)
            pdf_panelsNL.setFont('Helvetica-Bold', 72)
            pdf_panelsNL.drawCentredString(
                width/2, height - 100, 'PAN-PUE-INT')
            pdf_panelsNL.drawCentredString(
                width/2, height - 220, "PO: " + str(po[line]))
            pdf_panelsNL.drawCentredString(width/2, height - 340, "MO: " + str(mo[line]))
            pdf_panelsNL.drawCentredString(width/2, height - 460, 'HOLANDA')
            pdf_panelsNL.save()

            moNew = mo[line]
        except:
            msg_error(co, line, mo)

newWeek = True
weekNumber = 0

for line in range(len(mo)):
    if newWeek is True:
        #Create the excel
        excel_labels = openpyxl.load_workbook()
        sheet_labels = excel_labels.active
        #Write the title in every column
        cell_A1 = sheet_labels.cell(row=1, column=1)
        cell_A1.value = 'Item Nr'
        cell_A2 = sheet_labels.cell(row=1, column=2)
        cell_A2.value = 'Descripción'
        cell_A3 = sheet_labels.cell(row=1, column=3)
        cell_A3.value = 'MO'
        cell_A4 = sheet_labels.cell(row=1, column=4)
        cell_A4.value = 'PO'
        cell_A5 = sheet_labels.cell(row=1, column=4)
        cell_A5.value = 'CO España'
    if weekNumber == 0 or newWeek is True:
        for qty in qty[line]:
            rowTotal = rowTotal + qty + line
            cell_itemNumber = sheet_labels.cell(row=1 + rowTotal, column=1)
            cell_itemNumber.value = itemNumber[line]
            cell_description = sheet_labels.cell(row=1 + rowTotal, column=2)
            cell_description.value = itemName[line]
            
        
            
    
    



        
