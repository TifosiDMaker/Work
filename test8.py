from openpyxl import Workbook

wb = Workbook()
ws = wb.active

r1 = [3,5,6,67,3,34,6,3,5]
r2 = [3,5,6,3,4,6,4,3,6]
r3 = [4,6,345,6,45,2,5]
a = set(r1)
for row in zip(r1,r2,r3):
        ws.append(row)
wb.save('test.xlsx')