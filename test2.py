from openpyxl import Workbook
from openpyxl import load_workbook
import os
from os.path import join

wb = load_workbook('d:/Tifosi/Code/test.xlsx')
ws = wb.active

print(wb.sheetnames)
ws['A1'] = 2
i = 1
for root, dirs, files in os.walk('d:/Tifosi/12月/OT5865/OT5865-泰安特种车-中译英-手册、说明书/External Review/en-US/零部件图册'):
    for name in files:
        ws['B' + str(i)].value = name
        i = int(i)
        i += 1


