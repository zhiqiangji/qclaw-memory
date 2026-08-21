
import zipfile
import xml.etree.ElementTree as ET
import sys
import os

def read_docx_text(file_path):
    text_parts = []
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(file_path, 'r') as zip:
        xml_content = zip.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        for t in tree.findall('.//w:t', ns):
            if t.text:
                text_parts.append(t.text)
    return ''.join(text_parts)

def get_docx_info(file_path):
    info = {}
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/'
    }
    with zipfile.ZipFile(file_path, 'r') as zip:
        # 读取文档属性
        try:
            props_xml = zip.read('docProps/core.xml')
            tree = ET.fromstring(props_xml)
            title = tree.find('.//dc:title', ns)
            creator = tree.find('.//dc:creator', ns)
            info['标题'] = title.text if title is not None and title.text else '无标题'
            info['作者'] = creator.text if creator is not None and creator.text else '未知作者'
        except:
            info['标题'] = '无标题'
            info['作者'] = '未知作者'
    return info

if __name__ == '__main__':
    download_dir = r'C:\Users\32891\Downloads'
    output_dir = r'C:\Users\32891\.qclaw\workspace'
    files = [f for f in os.listdir(download_dir) if f.startswith('2026_4_') and f.endswith('.docx')]
    files.sort()
    for f in files:
        full_path = os.path.join(download_dir, f)
        info = get_docx_info(full_path)
        text = read_docx_text(full_path)
        output_file = os.path.join(output_dir, f.replace('.docx', '.txt'))
        with open(output_file, 'w', encoding='utf-8') as out:
            out.write(f"文件: {f}\n")
            out.write(f"标题: {info['标题']}\n")
            out.write(f"作者: {info['作者']}\n")
            out.write("="*80 + "\n\n")
            out.write(text)
        print(f"已读取: {f} -> {os.path.basename(output_file)} ({len(text)}字符)")
