import os
from docx import Document
from docx.shared import Pt


# 自定义自动化调整样式的函数

def auto_word_style(file_path):
    doc = Document(file_path)
    for parag in doc.paragraphs:
        for run in parag.runs:
            run.font.name = '宋体'  # 设置你的字体
            run.font.size =Pt(12)   # 设置你的字号
        parag.paragraph_format.line_spacing = 1.5  # 设置你的行间距
    doc.save(file_path)


# 遍历指定文件夹中的所有文件

def   by_filename_auto_word_style(folder_path):


    if not os.path.exists(folder_path) :
        print("文件夹路径不能为空或文件夹路径无效")
    else:
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.docx'):
                 file_path = os.path.join(folder_path, file_name)
                 auto_word_style(file_path)
                 print('正在处理：'+file_name+'文件')
        print('执行完成')