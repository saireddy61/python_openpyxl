from openpyxl import Workbook

from openpyxl.styles import Font

wb=Workbook()

wb['Sheet'].title='Data1'

sh1=wb.active

sh1['A1'].value='Name Of The Institution'

sh1.merge_cells('A1:H1')

sh1['A1'].font=Font(name='Cambria', bold=True, size=15)

sh1['A2'].value='Details Of ______________ Class'

sh1.merge_cells('A2:H2')

sh1['A2'].font=Font(name='Cambria', bold=True, size=13, underline='single')

sh1['A3'].value='S.No'

sh1['A3'].font=Font(name='Cambria', bold=True, size=12)

sh1['B3'].value='Student Name'

sh1['B3'].font=Font(name='Cambria', bold=True, size=12)

sh1['C3'].value='Father Name'

sh1['C3'].font=Font(name='Cambria', bold=True, size=12)

sh1['D3'].value='Admin No'

sh1['D3'].font=Font(name='Cambria', bold=True, size=12)

sh1['E3'].value='Joining Data'

sh1['E3'].font=Font(name='Cambria', bold=True, size=12)

sh1['F3'].value='Fee Status'

sh1['F3'].font=Font(name='Cambria', bold=True, size=12)

sh1['G3'].value='Caste'

sh1['G3'].font=Font(name='Cambria', bold=True, size=12)

sh1['H3'].value='Rank'

sh1['H3'].font=Font(name='Cambria', bold=True, size=12)

wb.save("Data_Form.xlsx")
