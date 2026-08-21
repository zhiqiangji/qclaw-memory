import sys
import subprocess

# 使用 Win32 COM 方式读取 .doc 文件
# 或者使用 antiword / wvWare 等第三方工具
# 更稳妥的方式：先检查是否安装了 python-docx，它也可以读取 .doc 但有限制
# 更好的方案：使用 python 的 olefile 或直接 win32com

# 尝试用更简单的方式：用 wordpad 或写字板导出为 txt
def extract_doc_win32(filepath, outputpath):
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(filepath)
        doc.SaveAs(outputpath, FileFormat=2)  # wdFormatText = 2
        doc.Close()
        word.Quit()
        return True
    except:
        return False

# 方式2: 使用 python-docx 读取
def extract_with_python_docx(filepath, outputpath):
    try:
        from docx import Document
        doc = Document(filepath)
        text = []
        for para in doc.paragraphs:
            text.append(para.text)
        
        # 也提取表格
        for i, table in enumerate(doc.tables):
            text.append(f"\n=== 表格 {i+1} ===")
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                text.append(" | ".join(cells))
        
        with open(outputpath, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(text))
        return True
    except:
        return False

if __name__ == "__main__":
    filepath = r"C:\Users\32891\Desktop\6 物理与电子信息学院转专业工作实施方案.doc"
    outputpath = r"C:\Users\32891\Desktop\6_物理与电子信息学院转专业工作实施方案.txt"
    
    # 先尝试 python-docx
    if extract_with_python_docx(filepath, outputpath):
        print(f"python-docx 成功: {outputpath}")
        sys.exit(0)
    
    # 再尝试 win32com
    if extract_doc_win32(filepath, outputpath):
        print(f"win32com 成功: {outputpath}")
        sys.exit(0)
    
    print("所有方法都失败")
    sys.exit(1)
