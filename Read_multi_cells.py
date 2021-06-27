import openpyxl  

  

wb = openpyxl.load_workbook('marks.xlsx')  

  

sheet = wb.active  

#  

cells = sheet['A1','B7']  

# cells behave like range operator  

for i1,i2 in cells:  

    print("{0:8} {1:8}".format(i1.value,i2.value))  
