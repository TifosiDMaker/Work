import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from docx import Document
import re
import time

start = time.time()

#os part
dest_dir = input('请输入外部审校文件所在路径.\n>')
dest_dir = dest_dir.replace("\\", "/")

#excel part
wb = Workbook()
ws = wb.active
fil = PatternFill(start_color='FFFF00', end_color='FFFF00',fill_type='solid')

#traverse files in dest_dir
for root, dirs, files in os.walk(dest_dir):
    for name in files:
        dest_path = os.path.join(root, name)
        doc = Document(dest_path)
        t = doc.tables[0]
        ws.append({'B': name})
        ws['B' + str(ws.max_row)].fill = fil

        for j in range(len(t.rows)):
            if t.cell(j, 2).text == "":
                continue
            strr = str(t.cell(j, 2).text)
            strr = re.sub('<'r'/?[a-z]{0,3}[0-9]{0,5}/?''>', '', strr)
            ws.append({'B': strr})
wb.save(dest_dir + '/删重文件.xlsx')
end = time.time()
cost = end - start
print (cost)
