import os
import win32com.client

dest_dir = input('请输入外部审校文件所在路径.\n>')
dest_dir = dest_dir.replace("\\","/")

word = win32com.client.gencache.EnsureDispatch('Word.Application')
word.Visible = 0
p = 0


for root, dirs, files in os.walk(dest_dir):
    for name in files:
        file_dir = os.path.join(root, name)
        file_dir = file_dir.replace("/","\\\\")
        thisdoc = word.Documents.Open(FileName=file_dir)
        #print(thisdoc.BuiltInDocumentProperties(win32com.client.constants.wdPropertyPages))
            #print(p)
        word.Documents.Open(FileName=file_dir)
        word.ActiveDocument.Repaginate()
        print(word.ActiveDocument.BuiltInDocumentProperties(win32com.client.constants.wdPropertyPages))
        thisdoc.Close()
