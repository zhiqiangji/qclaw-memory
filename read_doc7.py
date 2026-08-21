
# -*- coding: utf-8 -*-
import sys
import os

# 重定向输出到UTF-8
sys.stdout.reconfigure(encoding='utf-8')

try:
    import olefile
    ole = olefile.OleFileIO(r'C:\Users\32891\Downloads\2026_4_22185621_smt.doc')
    
    # 列出所有流
    print("=== OLE Streams ===")
    for entry in ole.listdir():
        print("/".join(entry))
    
    if ole.exists('WordDocument'):
        print("\n=== WordDocument Stream Found ===")
        # 尝试读取WordDocument流
        stream = ole.openstream('WordDocument')
        data = stream.read()
        
        # 尝试提取文本（简单方法）
        text = data.decode('utf-8', errors='ignore')
        # 简单清理
        text = text.replace('\x00', '')
        # 打印前2000字符
        print("\n=== Extracted Text (first 2000 chars) ===")
        print(text[:3000])
        
    ole.close()
        
except ImportError:
    print("olefile not found, trying to install...")
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "olefile"])
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
