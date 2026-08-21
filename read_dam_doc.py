
# -*- coding: utf-8 -*-
from docx import Document
import sys

doc = Document(r"C:\Users\32891\Desktop\新建水库大坝施工期渗压 - 沉降智能联动管控方案初稿.docx")

print("=== 标题 ===")
for paragraph in doc.paragraphs:
    print(paragraph.text)

print("\n=== 表格 ===")
for table in doc.tables:
    for row in table.rows:
        print(" | ".join([cell.text for cell in row.cells]))
