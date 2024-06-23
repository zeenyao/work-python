import os
import pdfplumber

# 指定目录路径
directory_path = r'鸟撞调查/'

# 指定文本区域
region = (200, 440, 380, 520)
'''# 获取文本坐标
import pdfplumber
path = r'D:/Python/PythonProjects/niaozhuang/姚泽恩.pdf'
with pdfplumber.open(path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"第 {i + 1} 页的文本对象和坐标:")
        for obj in page.extract_words():
            print(f"文本: {obj['text']} | 坐标: {obj['x0']}, {obj['top']}, {obj['x1']}, {obj['bottom']}")
'''
# 读取目录中的所有文件，并检查文件名称是否以.pdf结尾，以确保只处理pdf文件
file_list = os.listdir(directory_path)
for file_name in file_list:
    if file_name.endswith('.pdf'):

        # 构建完整路径
        file_path = os.path.join(directory_path, file_name)
        
        # 使用 pdfplumber 打开 PDF 文件，并提取指定区域的文本
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                region_text = page.within_bbox(region).extract_text()
                extracted_text = region_text
        
        # 使用提取到的文本作为新文件名
        new_file_name = extracted_text[:20] + '.pdf'  # 这里只取文本的前20个字符作为文件名
        
        # 构建新的文件路径
        new_file_path = os.path.join(directory_path, new_file_name)
        
        # 重命名文件
        os.rename(file_path, new_file_path)
        print(f"File '{file_name}' 已重命名为 '{new_file_name}'")
