from openpyxl import Workbook
from openpyxl.styles import PatternFill

wb = Workbook()
ws = wb.active
fil = PatternFill(start_color='FFFF00',end_color='FFFF00',fill_type='solid')
#fill2 = PatternFill(start_color='000000',end_color='000000',fill_type='solid')

ws['B1'].fill = fil
wb.save('test2.xlsx')