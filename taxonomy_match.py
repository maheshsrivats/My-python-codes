# Python code showing all the ratios together,  
# make sure you have installed fuzzywuzzy module 
  
from fuzzywuzzy import fuzz 
from fuzzywuzzy import process 
import pandas as pd
import csv
import xlwt 
from xlwt import Workbook 
  
wb = Workbook()
row = 1
sheet1 = wb.add_sheet('Sheet 1') 

#def item_with_rates(filename):
filename = "E://delete//test.xlsx"
data = pd.read_excel(filename)
length = len(data)
x={}
a=[0] *length
s1=[0] *length
s2 = [0] *length 
#print(a)


#print (data["itemdesc"][0])
# for i in range(length):
#     #print(data["Itemcode"][i],data["Location"][i],data["Lower_bound"][i])
#     if (data["Itemcode"][i+1]) is None:
#         exit
# # Give the location of the file 
# loc =  ("E:\delete\test.xlsx")
  
# # To open Workbook 
# wb = xlrd.open_workbook(loc) 
# #sheet = wb.sheet_by_index(0) 
  
# # For row 0 and column 0 
# #print(sheet.cell_value(0, 0)) 
#item is ours 
for i in range(length):
    s1 = data["Ship_desc"][i]
    row=row+1
    for j in range(length):
        s2[j] = data["Contruction_desc"][j]
        c= fuzz.ratio(s1, s2[j])
        a[j] = c
        #print(c,j)
    x=a.index(max(a))
    #print(x)
    print(s1,"::::",s2[x],"::::",a[x])
    sheet1.write(row, 1, s1)
    sheet1.write(row, 2, s2[x])
    sheet1.write(row, 3, a[x])
    #output file
    wb.save('ship_trial_1.xls') 

         
        