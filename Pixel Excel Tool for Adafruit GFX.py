#a helpful article I used as reference for openpyxl: https://www.geeksforgeeks.org/reading-excel-file-using-python/
#Adafruit GFX website: https://learn.adafruit.com/adafruit-gfx-graphics-library/overview

import openpyxl

#Insert exact excel file location in the quotes
path = "" 

#Opening a workbook and creating a workbook object
wb_obj = openpyxl.load_workbook(path)


sheet_obj = wb_obj.active

#Get max rows and columns
row = sheet_obj.max_row
column = sheet_obj.max_column

#print number of rows and columns in the excel file 
print("Total Rows:", row)
print("Total Columns:", column)
y = 0 
for y in range(1, row + 1):
   
    print("\nPixels in row ",y - 1 )
    for x in range(1, column + 1):
        
        cell_obj = sheet_obj.cell(row = y, column = x)
        color = cell_obj.fill.bgColor
        
        if color.rgb != '00000000' : 
            print("On: (", y -1 ,",", x - 1,")") 
            
        else:
            print("Off: (", y - 1,",", x - 1, ")")
            
print("")            
for y in range(1, row + 1):
    for x in range(1, column + 1):
        cell_obj = sheet_obj.cell(row = y, column = x)
        color = cell_obj.fill.bgColor
        if color.rgb != '00000000' : 
            print("drawPixel(",x - 1,",",y - 1,",",1,")",";") 
            
     
    
