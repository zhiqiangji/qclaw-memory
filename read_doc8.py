
# -*- coding: utf-8 -*-
import sys
import zipfile

# 重定向输出到UTF-8
sys.stdout.reconfigure(encoding='utf-8')

file_path = r'C:\Users\32891\Downloads\2026_4_22185621_smt.doc'

print("=== Trying to read as ZIP (docx) ===")
try:
    zf = zipfile.ZipFile(file_path)
    print("It's actually a DOCX file! (zip format)")
    print("\nContents:")
    for name in zf.namelist():
        print(name)
    
    # 读取document.xml
    if 'word/document.xml' in zf.namelist():
        print("\n=== document.xml found ===")
        content = zf.read('word/document.xml').decode('utf-8')
        # 简单提取文本（去除XML标签）
        import re
        text = re.sub(r'<[^>]+>', '', content)
        # 清理
        text = re.sub(r'\s+', ' ', text)
        print("\n=== Extracted Text (first 3000 chars) ===")
        print(text[:3000])
        
except zipfile.BadZipFile:
    print("Not a ZIP/DOCX file")
    
    # 尝试直接读取前4KB
    print("\n=== Reading first 4096 bytes ===")
    with open(file_path, 'rb') as f:
        data = f.read(4096)
        # 尝试用各种编码解码
        for enc in ['utf-8', 'gbk', 'gb18030', 'latin1']:
            try:
                text = data.decode(enc, errors='replace')
                print(f"\n--- {enc} ---")
                print(text[:500])
            except:
                pass
