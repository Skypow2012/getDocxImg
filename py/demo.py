import os
import sys
import time
from win32com import client
 
def doc_to_docx(path):
    if os.path.splitext(path)[1] == ".doc":
        word = client.Dispatch('Word.Application')
        doc = word.Documents.Open(path)  # 目标路径下的文件
        print('开始转换', path)
        doc.SaveAs(os.path.splitext(path)[0]+".docx")  # 转化后路径下的文件
        doc.Close()
        print('转换结束', os.path.splitext(path)[0]+".docx")
        word.Quit()

path = os.path.abspath(sys.argv[1])
# path = "C:/Users/maxmon/Desktop/demo/东北凤仙花.doc"#填写文件路径
doc_to_docx(path)