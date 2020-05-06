import tkinter as tk
from tkinter import filedialog
import pandas as pd
from fuzzywuzzy import fuzz 
from fuzzywuzzy import process
import xlwt
from xlwt import Workbook

wb = Workbook()
# Sheet 1 is the resultant sheet with all matches
sheet1 = wb.add_sheet('Sheet 1')

#Sheet 2 outputs only the correct matches
sheet2 = wb.add_sheet('Sheet 2')

#Cell Formatting for output excel file
style_heading = xlwt.easyxf('font: name Times New Roman, bold on')
style_success = xlwt.easyxf(
'pattern: pattern solid, fore_colour light_green;'
)

#-----------GUI to take input excel file-----------#
root= tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global excel_file
    global products   
    global products_data,item_code,description
    
    import_file_path = filedialog.askopenfilename()
    products = pd.read_excel (import_file_path)
    products_data = pd.DataFrame(products,columns=['Distributor Item code','Company Dscription'])
    products_data = products_data.dropna(thresh=1)

    item_code=pd.DataFrame(products,columns=['Distributor Item code'])
    item_code = item_code.dropna(thresh=1)

    description = pd.DataFrame(products,columns=['Company Dscription'])
    description = description.dropna(thresh=1)

browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)
root.mainloop()

#--------------GUI to take input excel file---------------#




#-------Main Algorithm Starts here--------#
item_code_list=item_code.values.tolist()
description_list=description.values.tolist()

size_item_code_list=[]
size_description_list=[]

for ele in item_code_list:    
    size_item_code_list.append(ele[0].split()[len(ele[0].split())-1])
    
for ele in description_list:
    newstring=""
    if len(ele[0].split())==1:
        strin=ele[0].split()[0]
    else:
        strin=ele[0].split()[1]
    for ch in strin:
        if ch=='X':
            newstring+='x'
        else:
            newstring+=ch
    strin=newstring
    size_description_list.append(strin)
    
global row,row_2
row=1
row_2=1

#Column headings for the output file
sheet1.write(0,0,"Distributor Item Code",style_heading)
sheet1.write(0,1,"Company Description",style_heading)
sheet1.write(0,2,"Comparision Score",style_heading)
sheet2.write(0,0,"Distributor Item Code",style_heading)
sheet2.write(0,1,"Company Description",style_heading)
sheet2.write(0,2,"Comparision Score",style_heading)

print("Executing File...")

for i in range(len(item_code_list)):
    fr=fuzz.ratio((size_item_code_list[i]),(size_description_list[i])),
    fpr=fuzz.partial_ratio((size_item_code_list[i]),(size_description_list[i])),
    
    #Thershold set to 80
    #If fuzzywuzzy algorithm returns a score more than or equal to 80, the columns match
    #This threshold is decided by trial and error on the file
    if fr[0]>=80:
        sheet1.write(row,0,item_code_list[i],style_success)
        sheet1.write(row,1,description_list[i],style_success)
        sheet1.write(row,2,fr[0],style_success)
        
        sheet2.write(row_2,0,item_code_list[i])
        sheet2.write(row_2,1,description_list[i])
        sheet2.write(row_2,2,fr[0])
        row_2 = row_2 + 1

    else:
        sheet1.write(row,0,item_code_list[i])
        sheet1.write(row,1,description_list[i])
        sheet1.write(row,2,fr[0])
        
    row = row+1
    
print("File named Comparision_Output.xls saved | Check Sheet 2 for all correct matches")
wb.save('Comparision_Output.xls')

print()
print("Result:")
print("Out of " ,row-2, " entries, " ,row_2-1, " are correct Item Code and Description matches")

