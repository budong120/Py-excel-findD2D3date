import openpyxl
import re

def findD2D3Date(sheet,s):
    l=[]
    for i in sheet:
        if i.value == s:
            l.append(i.column-2)
    return l

sourcePath=r'D:\\Verizon\\Roster\\202206Jun.xlsx'
wb=openpyxl.load_workbook(sourcePath)
activeSheet=wb.active
rosterScope=activeSheet['B7':'AG12']
for i in rosterScope:
    print(findD2D3Date(i,"D2"))
    print(findD2D3Date(i,"D3"))
