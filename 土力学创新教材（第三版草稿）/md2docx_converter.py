#!/usr/bin/env python3
"""
土力学创新教材 — Markdown → DOCX 高质量转换脚本
支持：LaTeX公式(MathML)、中文字体、紧凑表格、目录、标题层级
"""

import os
import sys
import re
import subprocess
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ========== 配置区 ==========
TEXTBOOK_DIR = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材（第三版草稿）"
OUTPUT_DOCX = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材.docx"
TEMP_MD = "/tmp/土力学_整合教材.md"

# 章节文件顺序（无附录和知识图谱，用于生成纯净目录）
CHAPTER_ORDER = [
    "序.md",
    "前言.md",
    "大纲.md",
    "第1章_绪论.md",
    "第2章_土的物理性质及分类.md",
    "第3章_土中应力.md",
    "第4章_土的渗透性与渗流.md",
    "第5章_土的压缩性与地基沉降.md",
    "第6章_土的抗剪强度.md",
    "第7章_土压力理论.md",
    "第8章_地基承载力.md",
    "第9章_土坡稳定分析.md",
    "第10章_土的动力特性.md",
    "第11章_岩土工程前沿专题.md",
    "第12章_综合应用与课程群知识图谱.md",
    "知识图谱.md",
    "附录.md",
]

# 中文 Word 字体候选（按优先级）
ZH_FONTS = ["宋体", "SimSun", "Noto Serif CJK SC", "AR PL UMing CN"]
EN_FONT = "Times New Roman"
MONO_FONT = "Courier New"

# ========== 工具函数 ==========

def set_font(run, zh_font, en_font, size_pt, bold=False, color=None):
    """设置中英文字体、字号、颜色"""
    run.font.name = en_font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    # 中文字体（通过 rPr）
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), zh_font)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:cs"), en_font)

def is_table_row_empty(row):
    """判断表格行是否全空"""
    return all(cell.text.strip() == "" for cell in row.cells)

def set_table_compact(table):
        """设置表格紧凑：去掉额外间距"""
        tbl = table._tbl
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl.insert(0, tblPr)
        # 去掉单元格间距
        tblCellMar = OxmlElement("w:tblCellMar")
        for side in ["top", "left", "bottom", "right"]:
            node = OxmlElement(f"w:{side}")
            node.set(qn("w:w"), "50")
            node.set(qn("w:type"), "dxa")
            tblCellMar.append(node)
        existing = tblPr.find(qn("w:tblCellMar"))
        if existing is not None:
            tblPr.remove(existing)
        tblPr.append(tblCellMar)

        # 逐单元格设置垂直居中
        for row in table.rows:
            row.height = None  # 自动高度
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for para in cell.paragraphs:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    for run in para.runs:
                        run.font.size = Pt(9)

def remove_paragraph_spacing(para):
    """去掉段落额外间距"""
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(3)

def process_paragraph(para, zh_font, en_font):
    """处理段落：设置字体+紧凑"""
    remove_paragraph_spacing(para)
    for run in para.runs:
        set_font(run, zh_font, en_font, 10.5)

def process_heading(para, zh_font, en_font):
    """处理标题段落"""
    remove_paragraph_spacing(para)
    level = 0
    pStyle = para.style
    if pStyle:
        sid = pStyle.name or ""
        m = re.search(r"Heading(\d+)", sid)
        if m:
            level = int(m.group(1))

    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(6)

    for run in para.runs:
        is_bold = run.font.bold
        # 字号：标题层级越大，字越小
        size_map = {1: 18, 2: 16, 3: 14, 4: 13, 5: 12}
        size = size_map.get(level, 12)
        color_map = {1: (0,0,0), 2: (0,0,0), 3: (0,0,0)}
        color = color_map.get(level, (0, 0, 0))
        set_font(run, zh_font, en_font, size, bold=True, color=color)

def find_best_zh_font():
    """从候选列表中找到系统已安装的中文字体"""
    import platform
    system = platform.system()
    if system != "Darwin":
        return "宋体"

    try:
        result = subprocess.run(
            ["system_profiler", "SPFontsDataType", "-json"],
            capture_output=True, text=True, timeout=10
        )
        import json
        data = json.loads(result.stdout)
        available = set()
        for font_list in data.get("SPFontsDataType", []):
            for f in font_list:
                available.add(f.get("_name", ""))
    except Exception:
        available = set()

    for f in ZH_FONTS:
        if f in available:
            return f
    return "宋体"  # macOS 默认中文字体

# ========== 核心流程 ==========

def step1_merge_markdown():
    """步骤1：合并所有 Markdown 文件"""
    print("📋 步骤1：合并 Markdown 文件...")

    parts = []
    parts.append("# 土力学创新教材\n\n")
    parts.append("*本教材面向新工科背景下土木工程专业本科教学*\n\n")
    parts.append("---\n\n")

    found = []
    missing = []
    for fname in CHAPTER_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if os.path.exists(fpath):
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()
            parts.append(f"\n\n<!-- FILE: {fname} -->\n\n")
            parts.append(content)
            found.append(fname)
        else:
            missing.append(fname)

    if missing:
        print(f"  ⚠️  跳过（未找到）：{missing}")

    merged = "".join(parts)

    with open(TEMP_MD, "w", encoding="utf-8") as f:
        f.write(merged)

    size_kb = os.path.getsize(TEMP_MD) / 1024
    print(f"  ✅ 合并完成：{len(found)}/{len(CHAPTER_ORDER)} 个文件，{size_kb:.1f} KB → {TEMP_MD}")
    return True

def step2_pandoc_convert():
    """步骤2：用 Pandoc 转 DOCX"""
    print("\n🔄 步骤2：Pandoc 转 DOCX...")

    # 检测可用中文字体
    zh_font = find_best_zh_font()
    print(f"  🔤 检测到中文字体：{zh_font}")

    cmd = [
        "pandoc", TEMP_MD,
        "-o", OUTPUT_DOCX,
        "--from=markdown",
        "--to=docx",
        # 公式：用 MathML（Word 原生支持，渲染最清晰）
        "--mathml",
        # 目录
        "--table-of-contents",
        "--toc-depth=3",
        # 章节编号
        "--number-sections",
        # 高亮风格
        "--highlight-style=pygments",
        # 中文语言
        "-V", "lang=zh-CN",
        # 中文字体（通过 reference.docx 效果更好，见 step3）
        "-V", f"mainfont={EN_FONT}",
        "-V", f"sansfont={EN_FONT}",
        "-V", f"monofont={MONO_FONT}",
        # 表格预设
        "--wrap=preserve",
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ❌ Pandoc 失败：{result.stderr}")
        return False

    size_kb = os.path.getsize(OUTPUT_DOCX) / 1024
    print(f"  ✅ Pandoc 转换完成：{size_kb:.1f} KB")
    return True

def step3_post_process():
    """步骤3：Python-docx 后处理（字体/表格/间距）"""
    print(f"\n🛠️ 步骤3：后处理 DOCX（字体/表格/紧凑化）...")

    doc = Document(OUTPUT_DOCX)

    zh_font = find_best_zh_font()
    print(f"  🔤 使用中文字体：{zh_font}")

    # 统计
    stats = {"段落": 0, "标题": 0, "表格": 0, "图片": 0}

    for element in doc.element.body:
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag

        if tag == "p":
            para = element  # lxml element
            # 判断是否是标题
            pPr = para.find(qn("w:pPr"))
            is_heading = False
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    sid = pStyle.get(qn("w:val")) or ""
                    if "Heading" in sid or "TOC" in sid:
                        is_heading = True

            if is_heading:
                stats["标题"] += 1
            else:
                stats["段落"] += 1

        elif tag == "tbl":
            stats["表格"] += 1

        elif tag == "drawing":
            stats["图片"] += 1

    # 重新遍历 doc.paragraphs 和 doc.tables 处理格式
    for para in doc.paragraphs:
        sid = para.style.name if para.style else ""
        if "Heading" in sid or "TOC" in sid:
            process_heading(para, zh_font, EN_FONT)
        else:
            process_paragraph(para, zh_font, EN_FONT)

    # 处理表格紧凑化
    for i, table in enumerate(doc.tables):
        set_table_compact(table)
        stats["表格"] = stats.get("表格", 0)
    # 重新计数（上面已经统计过了）
    stats["表格"] = len(doc.tables)

    # 页面设置：上下左右各 2.5cm，边距适度
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    # 页眉：书名
    header = section.header
    header.is_linked_to_previous = False
    hdr_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hdr_para.text = "《土力学创新教材》"
    hdr_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr_para.runs[0].font.size = Pt(9) if hdr_para.runs else None
    hdr_para.runs[0].font.color.rgb = RGBColor(128, 128, 128) if hdr_para.runs else None

    doc.save(OUTPUT_DOCX)

    print(f"  ✅ 后处理完成！")
    print(f"     段落: {stats['段落']} | 标题: {stats['标题']} | 表格: {stats['表格']} | 图片: {stats['图片']}")
    return True

def step4_create_reference_docx():
    """步骤4：生成自定义 reference.docx，控制 Word 默认样式"""
    print("\n📐 步骤4：生成自定义 reference.docx（精细控制样式）...")

    ref_path = "/tmp/pandoc_reference.docx"
    ref_doc = Document()

    # 标题1：大号居中黑体
    h1 = ref_doc.styles["Heading 1"]
    h1.font.name = EN_FONT
    h1.font.size = Pt(18)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rFonts_h1 = h1.element.find(qn("w:rPr"))
    if rFonts_h1 is None:
        rFonts_h1 = h1.element.makeelement(qn("w:rPr"), {})
    # 设置中文字体
    pPr = h1.element.find(qn("w:pPr"))
    rPr = h1.element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        h1.element.append(rPr)
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), zh_font)
    rFonts.set(qn("w:ascii"), EN_FONT)
    rFonts.set(qn("w:hAnsi"), EN_FONT)
    rPr.append(rFonts)

    # 标题2/3/4
    for lvl, size in [(2, 16), (3, 14), (4, 13)]:
        h = ref_doc.styles[f"Heading {lvl}"]
        h.font.name = EN_FONT
        h.font.size = Pt(size)
        h.font.bold = True
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)

    # 正文
    body = ref_doc.styles["Normal"]
    body.font.name = EN_FONT
    body.font.size = Pt(10.5)

    ref_doc.save(ref_path)
    print(f"  ✅ reference.docx → {ref_path}")
    return ref_path

def run_full_pipeline():
    """完整流水线"""
    print("=" * 60)
    print("📖 土力学创新教材 Markdown → DOCX 转换工具")
    print("=" * 60)
    print(f"📂 源目录: {TEXTBOOK_DIR}")
    print(f"📄 输出文件: {OUTPUT_DOCX}")
    print()

    # 检查源文件
    existing = [f for f in CHAPTER_ORDER if os.path.exists(os.path.join(TEXTBOOK_DIR, f))]
    print(f"📁 找到 {len(existing)}/{len(CHAPTER_ORDER)} 个章节文件")
    if not existing:
        print("❌ 源目录为空，终止！")
        return False

    # Step 1: 合并
    if not step1_merge_markdown():
        return False

    # Step 2: Pandoc 转换
    if not step2_pandoc_convert():
        return False

    # Step 3: 后处理
    if not step3_post_process():
        return False

    # 完成
    import datetime
    size_mb = os.path.getsize(OUTPUT_DOCX) / (1024 * 1024)
    print()
    print("=" * 60)
    print(f"✅ 转换完成！")
    print(f"   📄 文件：{OUTPUT_DOCX}")
    print(f"   📦 大小：{size_mb:.2f} MB")
    print(f"   🕐 时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("💡 打开方式：")
    print(f"   open \"{OUTPUT_DOCX}\"")
    print("=" * 60)
    return True

if __name__ == "__main__":
    ok = run_full_pipeline()
    sys.exit(0 if ok else 1)
