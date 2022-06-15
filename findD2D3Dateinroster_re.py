import openpyxl
import re

dictOfMonth={'January':1, 'February':2, 'March':3, 'April':4, 'May':5, 'June':6, 'July':7, 'August':8, 'September':9, 'October':10, 'November':11, 'December':12, }


def findD2D3Date(sheet,s):
    l=[yearMonth,]
    regexD2D3=re.compile(r'.*(%s)'%s)
    nameofroster=sheet[0].value
    for i in sheet:
        if regexD2D3.match(i.value):
            l.append(i.column-2)
    n=len(l)
    return (nameofroster,s,n,l)

sourcePath=r'D:\\tmp\\202204April.xlsx'
wb=openpyxl.load_workbook(sourcePath)
activeSheet=wb.active
rosterScope=activeSheet['B7':'AG12']
yearOfSheet=wb['January']
yearOfActive=yearOfSheet['AH4']
monthOfAcitve=activeSheet['B4']
yearMonth=str(yearOfActive.value)+'/'+str(dictOfMonth[monthOfAcitve.value])+'/'

for i in rosterScope:
    print(findD2D3Date(i,"D2"))
    print(findD2D3Date(i,"D3"))