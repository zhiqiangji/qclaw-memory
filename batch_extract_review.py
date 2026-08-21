
# -*- coding: utf-8 -*-
import os
from docx import Document

# 文件列表
files = [
    "终稿_宋纯昭.docx",
    "终稿_张朔宁.docx", 
    "终稿_李莹莹.docx",
    "终稿_苗永洁.docx",
    "终稿_陈俊南.docx",
    "终稿_马祥铭.docx"
]

desktop_path = r"C:\Users\32891\Desktop"
workspace_path = r"C:\Users\32891\.qclaw\workspace"

print("开始批量提取文档内容...\n")

for file_name in files:
    file_path = os.path.join(desktop_path, file_name)
    if os.path.exists(file_path):
        try:
            print(f"正在处理: {file_name}")
            doc = Document(file_path)
            
            # 提取前50段文字
            preview_text = []
            for i, para in enumerate(doc.paragraphs[:100]):
                text = para.text.strip()
                if text:
                    preview_text.append(text)
            
            # 统计表格数量
            table_count = len(doc.tables)
            
            # 保存预览
            output_name = file_name.replace(".docx", "_预览.txt")
            output_path = os.path.join(workspace_path, output_name)
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write("="*80 + "\n")
                f.write(f"文档预览: {file_name}\n")
                f.write("="*80 + "\n\n")
                f.write(f"表格数量: {table_count}\n\n")
                f.write("文档内容预览:\n")
                f.write("\n".join(preview_text))
            
            print(f"  ✓ 表格数量: {table_count}")
            print(f"  ✓ 已保存预览: {output_name}\n")
            
        except Exception as e:
            print(f"  ✗ 处理失败: {e}\n")
    else:
        print(f"✗ 文件不存在: {file_name}\n")

print("批量提取完成！")
