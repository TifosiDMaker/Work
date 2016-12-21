import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from docx import Document
import re

#os part
dest_dir = input('请输入外部审校文件所在路径.\n>')
dest_dir = dest_dir.replace("\\","/")
root = dest_dir
for root, dirs, files in os.walk(dest_dir):
    pass
name = files[0]
dest_path = os.path.join(root, name)

#excel part
wb = Workbook()
ws = wb.active
fil = PatternFill(start_color='FFFF00',end_color='FFFF00',fill_type='solid')

#word part
doc = Document(dest_path)


#traverse files in dest_dir
i = 1
for root, dirs, files in os.walk(dest_dir):
    for name in files:
        dest_path = os.path.join(root, name)
        doc = Document(dest_path)
        t = doc.tables[0]
        ws.append({'B':name})
        k = 0
        for haha in ws.rows:
            k += 1
        ws['B' + str(k)].fill = fil

        for j in range(len(t.rows)):
            if t.cell(j, 2).text == "":
                continue
            strr = str(t.cell(j, 2).text)
            strr = re.sub('<'r'/?[a-z]{0,3}[0-9]{0,5}/?''>','',strr)
            ws.append({'B': strr})
            k += 1
wb.save(dest_dir + '/删重文件.xlsx')