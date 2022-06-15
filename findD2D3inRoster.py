import openpyxl

def findD2(sheet,s):
    l=[]
    for i in sheet:
        if i.value == s:
            l.append(i.column-2)
    return l

sourcePath=r'D:\\py\\202206Jun.xlsx'
wb=openpyxl.load_workbook(sourcePath)
#print(wb.sheetnames)
sheet= wb['May']
column1=sheet['B7':'AG12']
#print(column1)
for j in column1:
    print(findD2(j,'D2'))
    print(findD2(j,'D3'))
#    for i in j:
        #print(i.coordinate,i.value,i.row,i.column)
        


    pass

#print(findD2(column1))