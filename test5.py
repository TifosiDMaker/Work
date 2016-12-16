from openpyxl import Workbook

wb = Workbook()
ws = wb.active

for i in range(100):
    ws.append({'B':'2'})

wb.save('hahahaha.xlsx')