import os
import win32com.client

dest_dir = input('请输入外部审校文件所在路径.\n>')
dest_dir = dest_dir.replace("\\","/")

word = win32com.client.gencache.EnsureDispatch('Word.Application')
word.Visible = 0

for root, dirs, files in os.walk(dest_dir):
    for name in files:
        file_dir = os.path.join(root, name)
        file_dir = file_dir.replace("/","\\\\")
        thisdoc = word.Documents.Open(FileName=file_dir)
        for rcount in range(thisdoc.Tables(1).Rows.Count):
            if thisdoc.Tables(1).Cell(int(rcount), 1).Shading.BackgroundPatternColorIndex != 8:
                #thisdoc.Tables(1).Rows(1).ConvertToText
                thisdoc.Tables(1).Cell(int(rcount), 3).Range.Text = ""
                #rcount -= 1
        thisdoc.Save()
        thisdoc.Close()
