
# -*- coding: utf-8 -*-
import sys
import time

# 重定向输出到UTF-8
sys.stdout.reconfigure(encoding='utf-8')

try:
    import win32com.client as win32
    
    print("Starting Word...")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    
    print("Opening document...")
    doc = word.Documents.Open(r"C:\Users\32891\Downloads\2026_4_22185621_smt.doc")
    
    print("=== DOCUMENT TEXT ===")
    # 获取文档文本
    text = doc.Content.Text
    print(text[:8000])  # 打印前8000字符
    
    print("\n=== SHAPES/FIGURES ===")
    try:
        for i, shape in enumerate(doc.InlineShapes):
            print(f"Shape {i+1}: {shape.Type}")
    except:
        pass
        
    print("\n=== TABLES ===")
    try:
        for i, table in enumerate(doc.Tables):
            print(f"\n--- Table {i+1} ---")
            for row in range(1, min(table.Rows.Count+1, 6)):  # 最多5行
                row_text = []
                for col in range(1, min(table.Columns.Count+1, 11)):  # 最多10列
                    cell = table.Cell(row, col)
                    row_text.append(cell.Range.Text.strip())
                print("|".join(row_text))
    except Exception as e:
        print(f"Table error: {e}")
    
    doc.Close()
    word.Quit()
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
