import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from docx import Document

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
fill = PatternFill(start_color='FFFF00',end_color='FFFF00',fill_type='solid')

#word part
doc = Document(dest_path)
t = doc.tables[0]

#traverse files in dest_dir
i = 1
for root, dirs, files in os.walk(dest_dir):
    for name in files:
        ws.append({'B':name})
        k = 0
        for haha in ws.rows:
            k = k + 1
        ws['B1'].fill
        for j in range(len(t.rows)):
            ws.append({'B': t.cell(j, 2).text})
wb.save('test.xlsx')