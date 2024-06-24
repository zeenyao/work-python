import os
from win32com import client
from docx import Document
import shutil


current_dir = os.getcwd()
input_path = current_dir
new_path = "outfile"
output_path = os.path.join(current_dir, new_path)
if not os.path.exists(output_path):
    os.mkdir(output_path)


word_app = client.Dispatch("Word.Application")
word_app.Visible = False
file_list = os.listdir(input_path)
for file_name in file_list:
    if file_name.endswith('.doc'):
        file_path = os.path.join(input_path, file_name)
        doc = word_app.Documents.Open(file_path)
        new_file_path = os.path.join(output_path, file_name + 'x')
        doc.SaveAs(new_file_path, FileFormat=16)
        doc.Close()
    if file_name.endswith('.docx'):
            shutil.copy(os.path.join(input_path, file_name), output_path)
word_app.Quit()


file_list = os.listdir(output_path)
for file_name in file_list:
    if file_name.endswith(".docx"):
        file_path = os.path.join(output_path, file_name)
        file_docx = Document(file_path)
        for table in file_docx.tables[0:1]:
            project_text = table.rows[1].cells[2].text
            print(f"file: {file_name}，projectname: {project_text}")


            new_file_name = project_text + ".docx"
            new_file_path = os.path.join(output_path, new_file_name)
            os.rename(file_path, new_file_path)
            print(f"file'{file_name}' rename '{new_file_name}'")