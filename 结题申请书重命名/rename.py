# 由于收取到的结题申请书命名过于杂乱、格式不一，且不符合规则，所以将其重命名为项目名称并统一存为.docx格式。

import os
from win32com import client
from docx import Document
import shutil

# 设置输入文件路径(当前脚本所在的路径)和输出文件路径(脚本所在的路径后创建名为'outfile'的文件夹)
current_dir = os.getcwd()
input_path = current_dir
new_path = 'outfile'
output_path = os.path.join(current_dir, new_path)
if not os.path.exists(output_path):
    os.mkdir(output_path)

# 打开Word程序将.doc另存为.docx，因为python-docx库仅能处理.docx文件
word_app = client.Dispatch('Word.Application')
word_app.Visible = False
file_list = os.listdir(input_path)
# 重命名.doc文件到输出文件夹
for file_name in file_list:
    if file_name.endswith('.doc'):
        file_path = os.path.join(input_path, file_name)
        doc = word_app.Documents.Open(file_path)
        new_file_path = os.path.join(output_path, file_name + 'x')
        doc.SaveAs(new_file_path, FileFormat=16)
        doc.Close()
    # 将.docx文件复制到输出文件夹
    if file_name.endswith('.docx'):
            shutil.copy(os.path.join(input_path, file_name), output_path)
word_app.Quit()

# 读取word中的某一个单元格的内容
file_list = os.listdir(output_path)
for file_name in file_list:
    if file_name.endswith('.docx'):
        file_path = os.path.join(output_path, file_name)
        file_docx = Document(file_path)
        for table in file_docx.tables[0:1]:
            project_text = table.rows[1].cells[2].text

# 使用读取到的内容重命名该word文件
            new_file_name = project_text + '.docx'
            new_file_path = os.path.join(output_path, new_file_name)
            os.rename(file_path, new_file_path)
            print(file_name + ' renamed to ' + new_file_name)