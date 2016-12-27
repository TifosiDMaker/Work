import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from docx import Document
import re
import time
start = time.time()

#excel part
wb = Workbook()
ws = wb.active
fil = PatternFill(start_color='FFFF00', end_color='FFFF00',fill_type='solid')

#os part
dest_dir = input('请输入外部审校文件所在路径。\n>').replace("\\", "/")

def copy(range):
    return str(t.cell(range, 2).text)

def del_blank(text):
    return text != ''

def clean_tag(text2):
    return re.sub('<'r'/?[a-z]{0,3}[0-9]{0,5}/?''>', '', text2)

for root, dirs, files in os.walk(dest_dir):
    pass

for name in files:
    t = Document(os.path.join(dest_dir, name)).tables[0]
    ws.append({'B': name})
    ws['B' + str(ws.max_row)].fill = fil
    j = list(range(len(t.rows)))
    r1 = list(map(clean_tag, filter(del_blank,map(copy, j))))
    r2 = [n for n in range(len(t.rows))]
    for row in zip(r2, r1):
        ws.append(row)

wb.save(dest_dir + '/删重文件.xlsx')

print (time.time() - start)
