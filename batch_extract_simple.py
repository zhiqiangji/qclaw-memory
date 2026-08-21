
# -*- coding: utf-8 -*-
import os
from docx import Document
import sys

# 设置输出编码
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

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

print("Starting batch extraction...\n")

results = []

for file_name in files:
    file_path = os.path.join(desktop_path, file_name)
    if os.path.exists(file_path):
        try:
            print(f"Processing: {file_name}")
            doc = Document(file_path)
            
            # 提取前100段文字
            preview_text = []
            title = ""
            for i, para in enumerate(doc.paragraphs[:150]):
                text = para.text.strip()
                if text:
                    if not title and len(text) < 50:
                        title = text
                    preview_text.append(text)
            
            # 统计表格数量
            table_count = len(doc.tables)
            
            # 保存预览
            output_name = file_name.replace(".docx", "_预览.txt")
            output_path = os.path.join(workspace_path, output_name)
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write("="*80 + "\n")
                f.write(f"Document Preview: {file_name}\n")
                f.write("="*80 + "\n\n")
                f.write(f"Title: {title}\n")
                f.write(f"Table Count: {table_count}\n\n")
                f.write("Content Preview:\n")
                f.write("\n".join(preview_text[:100]))
            
            results.append({
                "file": file_name,
                "title": title,
                "tables": table_count
            })
            
            print(f"  OK - Tables: {table_count}\n")
            
        except Exception as e:
            print(f"  ERROR: {e}\n")
            results.append({
                "file": file_name,
                "error": str(e)
            })
    else:
        print(f"File not found: {file_name}\n")

print("\nExtraction Summary:")
print("-"*50)
for r in results:
    if "error" in r:
        print(f"{r['file']}: ERROR - {r['error']}")
    else:
        print(f"{r['file']}: '{r['title']}' - {r['tables']} tables")

print("\nBatch extraction complete!")
