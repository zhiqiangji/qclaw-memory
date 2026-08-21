
# -*- coding: utf-8 -*-
import sys
import time

# 重定向输出到UTF-8
sys.stdout.reconfigure(encoding='utf-8')

try:
    import win32com.client as win32
    
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    
    doc = word.Documents.Open(r"C:\Users\32891\Downloads\2026_4_22185621_smt.doc")
    
    text = doc.Content.Text
    print(text[8000:])  # 打印从第8000字符开始的剩余部分
    
    doc.Close()
    word.Quit()
    
except Exception as e:
    print(f"Error: {e}")
