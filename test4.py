from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.append(['df','dds','dsf','fh'])
wb.save("hahaha.xlsx")