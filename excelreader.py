import openpyxl as op
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('File.xlsx', sheet_name='Sheet1')

list_of_Code = df['Row 3']
for i in list_of_Code.index:
    print(list_of_Code[i])
wb = op.load_workbook("C:\\Users\\asus\\Desktop\\New folder\\File.xlsx")
sh = wb.active
row_position=2
for i in list_of_Code.index:
    c = sh.cell(row=row_position, column=4)
    c.value = 33
    row_position+=1
wb.save("C:\\Users\\asus\\Desktop\\New folder\\modified_File.xlsx")
