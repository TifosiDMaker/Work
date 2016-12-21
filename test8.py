from docx import Document
from docx import table

doc = Document('d:/test.docx')
t = doc.tables[0]
print(t.Column[0])