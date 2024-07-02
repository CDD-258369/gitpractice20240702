# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 16:31:42 2024

@author: Administrator
"""

import os
from PIL import Image
import pytesseract
from docx import Document

# 确保已经安装了pytesseract和Pillow库
# pip install pytesseract Pillow python-docx

# 设置Tesseract-OCR的安装路径
# 请确保该路径正确指向您的Tesseract安装位置
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# D盘的图片目录和Word文档输出目录
image_dir = 'D:\\Images'
output_dir = 'D:\\WordDocuments'

# 确保输出目录存在
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 创建Word文档
doc = Document()

# 遍历图片目录中的所有文件
for filename in os.listdir(image_dir):
    # 检查文件扩展名，确保是图片文件
    if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
        # 构建图片的完整路径
        img_path = os.path.join(image_dir, filename)
        
        # 使用OCR识别图片中的文本
        # 指定使用中文简体语言包
        text = pytesseract.image_to_string(Image.open(img_path), lang='chi_sim')
        
        # 将识别的文本添加到Word文档
        # 如果文本不为空，则添加到文档中
        if text.strip():
            doc.add_paragraph(text)

# 保存Word文档
output_filename = os.path.join(output_dir, 'ExtractedText.docx')
doc.save(output_filename)

print(f'Document saved to {output_filename}')
print("congratulations for the first trying")
print("keep going good girl")
