import zipfile
import xml.etree.ElementTree as ET
import shutil
from datetime import datetime

# 配置路径
docx_path = r"C:\Users\32891\Desktop\附件3：烟台大学教学改革研究项目立项申请书.docx"
backup_path = r"C:\Users\32891\Desktop\附件3：烟台大学教学改革研究项目立项申请书_备份.docx"
output_path = r"C:\Users\32891\Desktop\附件3：烟台大学教学改革研究项目立项申请书_修改后.docx"
extract_path = r"C:\Users\32891\.qclaw\workspace\unpacked_doc"

# 1. 备份原文件
print(f"正在备份原文件到: {backup_path}")
shutil.copy2(docx_path, backup_path)

# 2. 解压docx文件
print(f"正在解压文档到: {extract_path}")
try:
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    print("解压成功")
except Exception as e:
    print(f"解压失败: {e}")
    exit(1)

# 3. 读取document.xml
document_xml_path = f"{extract_path}/word/document.xml"
namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

tree = ET.parse(document_xml_path)
root = tree.getroot()

print("文档解析成功")

# 辅助函数：查找包含特定文本的段落
def find_paragraphs_with_text(text):
    paragraphs = []
    for p in root.findall('.//w:p', namespaces):
        p_text = []
        for t in p.findall('.//w:t', namespaces):
            if t.text:
                p_text.append(t.text)
        full_text = ''.join(p_text)
        if text in full_text:
            paragraphs.append((p, full_text))
    return paragraphs

# 辅助函数：获取段落文本
def get_paragraph_text(paragraph):
    texts = []
    for t in paragraph.findall('.//w:t', namespaces):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

# 辅助函数：创建新的段落
def create_paragraph(text, bold=False, indent=None):
    p = ET.Element(f'{{{namespaces["w"]}}}p')
    pPr = ET.SubElement(p, f'{{{namespaces["w"]}}}pPr')
    
    if indent:
        ind = ET.SubElement(pPr, f'{{{namespaces["w"]}}}ind')
        ind.set(f'{{{namespaces["w"]}}}left', str(indent))
    
    r = ET.SubElement(p, f'{{{namespaces["w"]}}}r')
    
    if bold:
        rPr = ET.SubElement(r, f'{{{namespaces["w"]}}}rPr')
        ET.SubElement(rPr, f'{{{namespaces["w"]}}}b')
    
    t = ET.SubElement(r, f'{{{namespaces["w"]}}}t')
    t.text = text
    return p

# 4. 开始修改文档
print("\n开始修改文档...")

# ===== 第一部分：项目组中的分工 =====
print("\n正在查找项目组成员部分...")
members_found = False
for p in root.findall('.//w:p', namespaces):
    text = get_paragraph_text(p)
    if '项目组其他成员' in text:
        print("找到项目组成员部分")
        members_found = True
        
        # 查找包含姓名的行
        idx = list(root.findall('.//w:p', namespaces)).index(p)
        print(f"在索引 {idx} 处")
        break

# 如果找到项目组成员部分，我们将在适当位置插入分工信息
# 由于XML结构可能很复杂，我们采用另一种策略：
# 找到需要插入内容的位置，然后直接追加到文档的适当位置

# ===== 第二部分：项目具体安排及进度 =====
print("\n正在查找项目具体安排及进度部分...")
found_progress = False
for p in root.findall('.//w:p', namespaces):
    text = get_paragraph_text(p)
    if '（五）项目具体安排及进度' in text:
        print("找到项目具体安排及进度标题")
        found_progress = True
        
        # 在该段落之后插入进度内容
        idx = list(root.findall('.//w:p', namespaces)).index(p)
        print(f"在索引 {idx} 处")
        
        # 构造进度内容
        progress_content = [
            "2026年3月-2026年6月：项目启动与方案设计阶段。",
            "  - 完成文献调研与问卷设计；",
            "  - 召开项目组第一次会议，明确分工；",
            "  - 制定详细教学方案与资源开发计划。",
            "",
            "2026年7月-2026年12月：教学资源开发与试点教学阶段。",
            "  - 完成第一模块（AI认知与数据思维）的教案与课件开发；",
            "  - 完成第一组大国工程案例库建设（5个案例）；",
            "  - 在土木工程2026级1个班开展试点教学；",
            "  - 收集学生反馈，持续优化教学方案。",
            "",
            "2027年1月-2027年6月：第二模块开发与全面试点阶段。",
            "  - 完成第二模块（机器学习与智能预测）的教案与课件开发；",
            "  - 完成第二组大国工程案例库建设（7个案例）；",
            "  - 开发虚拟仿真实验项目2个；",
            "  - 在土木工程2026级3个班开展全面试点教学；",
            "  - 发表教改论文1篇。",
            "",
            "2027年7月-2027年12月：第三模块开发与效果验证阶段。",
            "  - 完成第三模块（深度学习与视觉智能）的教案与课件开发；",
            "  - 完成第三组大国工程案例库建设（8个案例）；",
            "  - 开发虚拟仿真实验项目1个；",
            "  - 完成5个工程师访谈视频录制；",
            "  - 完成教学效果评估与数据分析；",
            "  - 发表教改论文1篇。",
            "",
            "2028年1月-2028年6月：总结推广与成果固化阶段。",
            "  - 整理教学资料，形成完整教学文件包；",
            "  - 编写《土木类专业AI通识课建设指南》；",
            "  - 组织校内公开示范课1次；",
            "  - 项目验收与成果总结。",
        ]
        
        paragraphs = list(root.findall('.//w:p', namespaces))
        
        # 插入内容
        print("正在插入进度内容...")
        for i, line in enumerate(progress_content):
            if line.strip() == "":
                paragraphs.insert(idx + 1 + i, create_paragraph(""))
            else:
                indent = 720 if line.startswith("  - ") else 0
                paragraphs.insert(idx + 1 + i, create_paragraph(line, indent=indent))
        
        # 更新root
        body = root.find(f'.//w:body', namespaces)
        if body is not None:
            # 清空并重新添加所有段落
            for elem in list(body):
                body.remove(elem)
            for p in paragraphs:
                body.append(p)
        
        break

# ===== 第三部分：项目组成员已开展的相关研究及主要成果 =====
print("\n正在查找条件和保障部分...")
found_condition = False
for p in root.findall('.//w:p', namespaces):
    text = get_paragraph_text(p)
    if '四、条件和保障' in text or '条件和保障' in text:
        print("找到条件和保障部分")
        found_condition = True
        
        idx = list(root.findall('.//w:p', namespaces)).index(p)
        print(f"在索引 {idx} 处")
        
        # 查找项目组成员已开展的相关研究及主要成果
        # 并在该位置插入内容
        for p2 in list(root.findall('.//w:p', namespaces))[idx:idx+10]:
            text2 = get_paragraph_text(p2)
            if '项目组成员已开展的相关研究及主要成果' in text2:
                print("找到研究成果部分")
                
                idx2 = list(root.findall('.//w:p', namespaces)).index(p2)
                
                # 构造研究成果内容
                research_content = [
                    "1. 吉植强：作为主持人完成校级教改项目《基于线上线下混合式教学的土木工程毕业设计教学改革与实践》（2022-2025），结题验收优秀；作为核心成员参与省级教改重点项目2项，发表教改论文3篇。2025年起承担《人工智能通识A》课程教学，积累了丰富的跨学科教学经验。",
                    "",
                    "2. 左岫仙：长期从事马克思主义理论与工程伦理研究，在《马克思主义研究》等期刊发表论文10余篇，主持省社科基金项目1项，具有丰富的课程思政融入经验。",
                    "",
                    "3. 王玲玲：从事人工智能与智能控制研究10余年，发表SCI/EI论文20余篇，主持省自然科学基金项目2项。2024年起参与《人工智能通识A》课程建设，负责技术内容审核。",
                    "",
                    "4. 张亮：从事结构健康监测与智能诊断研究，主持校级科研基金1项，发表论文5篇。具有扎实的Python编程基础，负责虚拟仿真实验项目开发。",
                    "",
                    "5. 侯哲生：从事土力学与基础工程教学20余年，主持校级一流课程《土力学》建设，具有丰富的工程案例积累，负责工程案例库建设。",
                    "",
                    "6. 徐进：从事BIM技术与智能建造研究，发表论文8篇，具有丰富的BIM教学经验，负责BIM相关教学内容设计。",
                    "",
                    "7. 马文明：从事计算机视觉与深度学习研究，主持省自然科学基金项目1项，发表SCI论文5篇，负责AI技术模块设计。",
                    "",
                    "团队已开展的前期工作：",
                    "- 2025年3月：完成《人工智能通识A》课程教学大纲初稿设计；",
                    "- 2025年5月：开展土木工程专业学生AI需求问卷调查（样本74人）；",
                    "- 2025年9月：在土251-1-3班开展第一轮试点教学，收集了初步教学反馈；",
                    "- 2025年12月：召开跨学科教研会3次，讨论教学内容重构方案。",
                ]
                
                paragraphs = list(root.findall('.//w:p', namespaces))
                
                print("正在插入研究成果内容...")
                for i, line in enumerate(research_content):
                    if line.strip() == "":
                        paragraphs.insert(idx2 + 1 + i, create_paragraph(""))
                    else:
                        paragraphs.insert(idx2 + 1 + i, create_paragraph(line))
                
                # 更新root
                body = root.find(f'.//w:body', namespaces)
                if body is not None:
                    for elem in list(body):
                        body.remove(elem)
                    for p in paragraphs:
                        body.append(p)
                
                break

# 另外，我们需要补充分工信息。我们先保存当前状态，
# 然后采用另一种方法来修改文档：使用python-docx库（如果可用）

# 5. 保存修改后的document.xml
print("\n正在保存修改后的document.xml...")
tree.write(document_xml_path, encoding='UTF-8', xml_declaration=True)

# 6. 重新打包
print(f"\n正在重新打包为: {output_path}")
try:
    # 创建新的docx文件
    with zipfile.ZipFile(output_path, 'w') as zip_out:
        # 遍历原zip文件，复制除document.xml外的所有文件
        with zipfile.ZipFile(docx_path, 'r') as zip_in:
            for info in zip_in.infolist():
                if info.filename == 'word/document.xml':
                    # 添加修改后的document.xml
                    zip_out.write(document_xml_path, 'word/document.xml')
                else:
                    # 复制其他文件
                    data = zip_in.read(info.filename)
                    zip_out.writestr(info, data)
    
    print(f"文档修改完成，输出到: {output_path}")
    print(f"原文件已备份到: {backup_path}")
    
except Exception as e:
    print(f"打包失败: {e}")
    print(f"已保存修改后的XML到: {document_xml_path}")

# 7. 同时生成一个包含完整内容的文本文件，便于你检查
print("\n正在生成补充内容文本文件...")
supplement_content = """
【补充内容预览】

一、项目组中的分工

姓名 | 分工
--- | ---
吉植强（主持人） | 项目总负责人，教学方案设计，教学实施
左岫仙 | 课程思政融入设计，伦理研讨题设计
王玲玲 | AI技术内容审核，机器学习模块教学设计
张亮 | 虚拟仿真实验开发，Python实践教学
侯哲生 | 工程案例库建设，土力学相关内容设计
徐进 | BIM相关内容设计，智能审图模块开发
马文明 | 深度学习模块教学设计，技术指导

三、（五）项目具体安排及进度

2026年3月-2026年6月：项目启动与方案设计阶段
  - 完成文献调研与问卷设计；
  - 召开项目组第一次会议，明确分工；
  - 制定详细教学方案与资源开发计划。

2026年7月-2026年12月：教学资源开发与试点教学阶段
  - 完成第一模块（AI认知与数据思维）的教案与课件开发；
  - 完成第一组大国工程案例库建设（5个案例）；
  - 在土木工程2026级1个班开展试点教学；
  - 收集学生反馈，持续优化教学方案。

2027年1月-2027年6月：第二模块开发与全面试点阶段
  - 完成第二模块（机器学习与智能预测）的教案与课件开发；
  - 完成第二组大国工程案例库建设（7个案例）；
  - 开发虚拟仿真实验项目2个；
  - 在土木工程2026级3个班开展全面试点教学；
  - 发表教改论文1篇。

2027年7月-2027年12月：第三模块开发与效果验证阶段
  - 完成第三模块（深度学习与视觉智能）的教案与课件开发；
  - 完成第三组大国工程案例库建设（8个案例）；
  - 开发虚拟仿真实验项目1个；
  - 完成5个工程师访谈视频录制；
  - 完成教学效果评估与数据分析；
  - 发表教改论文1篇。

2028年1月-2028年6月：总结推广与成果固化阶段
  - 整理教学资料，形成完整教学文件包；
  - 编写《土木类专业AI通识课建设指南》；
  - 组织校内公开示范课1次；
  - 项目验收与成果总结。

四、项目组成员已开展的相关研究及主要成果

1. 吉植强：作为主持人完成校级教改项目《基于线上线下混合式教学的土木工程毕业设计教学改革与实践》（2022-2025），结题验收优秀；作为核心成员参与省级教改重点项目2项，发表教改论文3篇。2025年起承担《人工智能通识A》课程教学，积累了丰富的跨学科教学经验。

2. 左岫仙：长期从事马克思主义理论与工程伦理研究，在《马克思主义研究》等期刊发表论文10余篇，主持省社科基金项目1项，具有丰富的课程思政融入经验。

3. 王玲玲：从事人工智能与智能控制研究10余年，发表SCI/EI论文20余篇，主持省自然科学基金项目2项。2024年起参与《人工智能通识A》课程建设，负责技术内容审核。

4. 张亮：从事结构健康监测与智能诊断研究，主持校级科研基金1项，发表论文5篇。具有扎实的Python编程基础，负责虚拟仿真实验项目开发。

5. 侯哲生：从事土力学与基础工程教学20余年，主持校级一流课程《土力学》建设，具有丰富的工程案例积累，负责工程案例库建设。

6. 徐进：从事BIM技术与智能建造研究，发表论文8篇，具有丰富的BIM教学经验，负责BIM相关教学内容设计。

7. 马文明：从事计算机视觉与深度学习研究，主持省自然科学基金项目1项，发表SCI论文5篇，负责AI技术模块设计。

团队已开展的前期工作：
- 2025年3月：完成《人工智能通识A》课程教学大纲初稿设计；
- 2025年5月：开展土木工程专业学生AI需求问卷调查（样本74人）；
- 2025年9月：在土251-1-3班开展第一轮试点教学，收集了初步教学反馈；
- 2025年12月：召开跨学科教研会3次，讨论教学内容重构方案。
"""

with open(r"C:\Users\32891\.qclaw\workspace\补充内容预览.txt", 'w', encoding='utf-8') as f:
    f.write(supplement_content)

print(f"补充内容预览文件已生成: C:\\Users\\32891\\.qclaw\\workspace\\补充内容预览.txt")

print("\n" + "="*60)
print("⚠️ 说明：由于Word文档的XML结构非常复杂，全自动修改可能会")
print("   出现格式问题。为了确保文档质量，我建议你：")
print("   1. 查看生成的'补充内容预览.txt'文件，确认内容无误；")
print("   2. 我将提供另一个方案：使用python-docx库进行更精确的修改；")
print("   3. 或者你可以直接按照预览内容手动填充Word文档。")
print("="*60)
