#!/usr/bin/env python3
"""
土力学创新教材 v2 — Markdown → DOCX 高质量转换
修复：双重编号、段落首行缩进、公式编号、三线表、公式字体、自动分页、代码块格式、表格自适应
"""

import os, re, subprocess, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==================== 配置区 ====================
TEXTBOOK_DIR = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材（第三版草稿）"
OUTPUT_DOCX  = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材.docx"
TEMP_MD      = "/tmp/土力学_v2_整合.md"

CHAPTER_ORDER = [
    "序.md","前言.md","大纲.md",
    "第1章_绪论.md","第2章_土的物理性质及分类.md","第3章_土中应力.md",
    "第4章_土的渗透性与渗流.md","第5章_土的压缩性与地基沉降.md",
    "第6章_土的抗剪强度.md","第7章_土压力理论.md","第8章_地基承载力.md",
    "第9章_土坡稳定分析.md","第10章_土的动力特性.md",
    "第11章_岩土工程前沿专题.md","第12章_综合应用与课程群知识图谱.md",
    "知识图谱.md","附录.md",
]

EN_FONT    = "Times New Roman"
MONO_FONT  = "Courier New"
ZH_FONT    = "宋体"     # 正文宋体
BODY_SIZE  = 10.5       # 正文字号(pt)
INDENT_EMU = Emu(21 * 914400 // 100)  # 首行缩进2字符

# 公式计数器：{(章号, 小节号): n}
EQ_COUNTER = {}         # key=(ch,sec) → int
EQ_TOTAL   = 0
# 表格行数阈值：超过此行数启用自适应
TABLE_FULL_PAGE_WIDTH_CM = 16.0   # 可用页面宽度(cm)

# ==================== 工具函数 ====================

def log(msg):
    print(f"  {msg}")

def strip_duplicate_section_numbers(text):
    """
    去掉 md 中标题行的双重编号。
    原格式：## 1.1 土力学研究对象  →  ## 土力学研究对象
    原则：保留 ## 及其前导 ##，去掉 "1.1 " 这种前缀编号（数字开头+空格）
    """
    def replacer(m):
        level = m.group(1)           # "##"
        num   = m.group(2)           # "1.1 " 或 "1.1.2 "
        rest  = m.group(3)          # 标题文字
        # 如果 rest 开头恰好也是同样的编号（说明标题里自带编号了），只去掉一份
        return f"{level} {rest}"
    # 匹配 ## 到行末，去掉 "数字. " 或 "数字.数字. " 前缀
    return re.sub(r'^(#{1,6}) (\d+(?:\.\d+)*) (【.*?】.*)$',
                  lambda m: f"{m.group(1)} {m.group(3)}", text, flags=re.MULTILINE)

def add_equation_numbers(text):
    """
    给 $$...$$ 公式块追加右对齐编号，格式：(章-节-序号)
    策略：检测到 $$...$$ 块结束后，在其后追加一行
          `\tag{(01-01-01)}` 形式的 LaTeX——Pandoc 将其渲染为公式编号，
          配合 pandoc 的 `--mathml` 渲染在 MathML 内部。
    """
    global EQ_COUNTER, EQ_TOTAL

    lines   = text.split('\n')
    result  = []
    ch_num  = "01"
    sec_num = "01"
    in_block = False
    pending_tag = None   # 待写入的编号

    i = 0
    while i < len(lines):
        raw    = lines[i]
        strip  = raw.strip()

        # 文件章节标记
        m_ch = re.search(r'FILE: 第(\d+)章', strip)
        if m_ch:
            ch_num = m_ch.group(1).zfill(2)

        # 小节编号：## 1.1 或 ### 1.1.2
        m_sec = re.search(r'^#{2,4}\s+(\d+)\.(\d+)', strip)
        if m_sec:
            ch_num, sec_num = m_sec.group(1).zfill(2), m_sec.group(2).zfill(2)

        # --- 公式块处理 ---
        if strip.startswith('$$') and not in_block:
            # 单行块：$$...$$
            if strip.endswith('$$') and strip != '$$':
                result.append(raw)
                key = (ch_num, sec_num)
                EQ_COUNTER[key] = EQ_COUNTER.get(key, 0) + 1
                EQ_TOTAL += 1
                num = f"({ch_num}-{sec_num}-{EQ_COUNTER[key]:02d})"
                # 追加 \tag 行作为新段落（Pandoc 识别 LaTeX）
                result.append(f'\n\\tag{{{num}}}\n')
                i += 1
                continue

            # 多行块开始
            in_block = True
            result.append(raw)
            i += 1
            continue

        if in_block:
            result.append(raw)
            if strip.endswith('$$'):
                # 块结束：生成编号追加
                in_block = False
                key = (ch_num, sec_num)
                EQ_COUNTER[key] = EQ_COUNTER.get(key, 0) + 1
                EQ_TOTAL += 1
                num = f"({ch_num}-{sec_num}-{EQ_COUNTER[key]:02d})"
                result.append(f'\n\\tag{{{num}}}\n')
            i += 1
            continue

        # 普通行
        result.append(raw)
        i += 1

    return '\n'.join(result)

def ensure_style(doc, style_name, base_name, font_name, font_size,
                 bold=False, color=None, space_before=0, space_after=3,
                 first_indent_chars=0, align=None):
    """确保某样式存在并设置属性"""
    if style_name in doc.styles:
        style = doc.styles[style_name]
    else:
        style = doc.styles.add_style(style_name, 1)  # PARAGRAPH=1

    style.base_style = doc.styles[base_name]
    style.font.name  = font_name
    style.font.size  = Pt(font_size)
    style.font.bold  = bold
    if color:
        style.font.color.rgb = RGBColor(*color)
    pf = style.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    if first_indent_chars:
        pf.first_line_indent = Emu(first_indent_chars * 914400 // 100)
    if align:
        pf.alignment = align
    return style

def apply_rFonts(run, zh=None, en=None):
    """为 run 的 rPr 添加 rFonts（中文字体）"""
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    if en:
        rFonts.set(qn("w:ascii"),      en)
        rFonts.set(qn("w:hAnsi"),     en)
        rFonts.set(qn("w:cs"),        en)
    if zh:
        rFonts.set(qn("w:eastAsia"),   zh)

def set_run_font(run, zh=None, en=None, size_pt=None, bold=None):
    """设置 run 的字体、字号"""
    if en:  apply_rFonts(run, zh=zh, en=en)
    if size_pt: run.font.size = Pt(size_pt)
    if bold is not None: run.font.bold = bold

def set_table_three_line(table, available_width_cm=TABLE_FULL_PAGE_WIDTH_CM):
    """
    三线表：设置 tblBorders 只保留顶部(above)、中间(below隔线)、底部(below)
    pandoc默认会生成所有边框，全部移除后只加top/mid/bottom
    """
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)

    # 查找或创建 tblBorders
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders"); tblPr.append(tblBorders)
    else:
        # 清空所有现有边框
        for child in list(tblBorders):
            tblBorders.remove(child)

    def add_border(btype, val="single", sz="8", color="000000", space="0"):
        b = OxmlElement(f"w:{btype}")
        b.set(qn("w:val"),   val)
        b.set(qn("w:sz"),    sz)
        b.set(qn("w:space"), space)
        b.set(qn("w:color"), color)
        tblBorders.append(b)

    # 顶部、底部：粗线(12)，中间隔线：细线(4)
    add_border("top",    sz="12", color="000000")
    add_border("bottom", sz="12", color="000000")
    add_border("insideH", sz="4", color="000000")  # 水平隔线（表头下）

    # 垂直边框去掉
    add_border("left",   val="nil")
    add_border("right",  val="nil")
    add_border("insideV", val="nil")

    # 去掉单元格间距
    cellMar = tblPr.find(qn("w:tblCellMar"))
    if cellMar is not None:
        tblPr.remove(cellMar)

    # 自适应列宽
    rows = table.rows
    if not rows:
        return
    num_cols = len(rows[0].cells)
    col_w = Cm(available_width_cm / num_cols)

    # 设置 tblGrid
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is None:
        tblGrid = OxmlElement("w:tblGrid"); tbl.insert(0, tblGrid)
    else:
        for g in list(tblGrid): tblGrid.remove(g)
    for _ in range(num_cols):
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(int(col_w.twips))); tblGrid.append(gc)

    # 逐格设宽
    for row in rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW"); tcPr.insert(0, tcW)
            tcW.set(qn("w:w"),    str(int(col_w.twips)))
            tcW.set(qn("w:type"), "dxa")

    # 表头行：加粗+垂直居中
    for cell in rows[0].cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for para in cell.paragraphs:
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after  = Pt(0)
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                apply_rFonts(run, zh=ZH_FONT, en=EN_FONT)

    # 内容行：字号9pt
    for row in rows[1:]:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in cell.paragraphs:
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after  = Pt(0)
                for run in para.runs:
                    run.font.size = Pt(9)
                    apply_rFonts(run, zh=ZH_FONT, en=EN_FONT)

def is_code_style(style_name):
    return any(x in (style_name or "") for x in ["Code", "Source", "CodeBlock", "Verbatim"])

def fix_code_block_para(para):
    """给代码块段落设置等宽字体和灰色底纹"""
    for run in para.runs:
        apply_rFonts(run, zh=None, en=MONO_FONT)
        run.font.name = MONO_FONT
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(30, 30, 30)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    # 加灰色底纹
    pPr = para._p.get_or_add_pPr()
    shd = pPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd"); pPr.insert(0, shd)
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "F2F2F2")

# ==================== 步骤函数 ====================

def step1_merge_and_clean():
    """合并 md + 去除双重编号 + 添加公式编号"""
    print("\n📋 步骤1：合并、清理标题编号、公式编号...")
    parts = []
    found = missing = 0

    for fname in CHAPTER_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if not os.path.exists(fpath):
            missing += 1; continue
        with open(fpath, encoding="utf-8") as f:
            content = f.read()
        parts.append(f"\n<!-- FILE: {fname} -->\n\n")
        parts.append(content)
        found += 1

    merged = "".join(parts)
    # ★ 先添加公式编号（需要原始的 ## 1.1 小节编号信息）
    merged = add_equation_numbers(merged)

    # 去掉双重编号：## 1.1 标题 → ## 标题
    merged = re.sub(r'^(#{1,6}) (\d+(?:\.\d+)*) ([\u4e00-\u9fff【])',
                    lambda m: f"{m.group(1)} {m.group(3)}", merged, flags=re.MULTILINE)
    merged = re.sub(r'^(#{1,6}) (\d+(?:\.\d+)*) (第[一二三四五六七八九十]+)',
                    lambda m: f"{m.group(1)} {m.group(3)}", merged, flags=re.MULTILINE)
    merged = re.sub(r'^(#{1,6}) (\d+(?:\.\d+)*) (土的)',
                    lambda m: f"{m.group(1)} {m.group(3)}", merged, flags=re.MULTILINE)
    merged = re.sub(r'^(#{1,6}) (\d+(?:\.\d+)*) ([A-Z\u4e00-\u9fff])',
                    lambda m: f"{m.group(1)} {m.group(3)}", merged, flags=re.MULTILINE)

    with open(TEMP_MD, "w", encoding="utf-8") as f:
        f.write(merged)

    log(f"✅ 合并完成 {found}/{found+missing} 个文件，{os.path.getsize(TEMP_MD)/1024:.1f} KB")
    return True

def step2_pandoc():
    """Pandoc 转 DOCX（无 number-sections，避免双重编号；无 --mathml，用默认 OMML）"""
    print("\n🔄 步骤2：Pandoc 转 DOCX...")
    cmd = [
        "pandoc", TEMP_MD,
        "-o", OUTPUT_DOCX,
        "--from=markdown",
        "--to=docx",
        "--table-of-contents",
        "--toc-depth=3",
        # ★ 不加 --number-sections（编号已在 md 标题中）
        # ★ 不加 --mathml（让 Pandoc 用默认 OMML，Word 原生渲染效果最好）
        "--wrap=preserve",
        "-V", "lang=zh-CN",
        "-V", f"mainfont={EN_FONT}",
        "-V", f"sansfont={EN_FONT}",
        "-V", f"monofont={MONO_FONT}",
        "-V", "geometry:margin=2.5cm",
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode:
        log(f"❌ Pandoc 失败：{r.stderr[:300]}")
        return False
    log(f"✅ Pandoc 完成 {os.path.getsize(OUTPUT_DOCX)/1024:.1f} KB")
    return True

def step3_define_styles(doc):
    """预定义和覆盖样式（代码块、公式引用、正文等）"""
    print("\n🎨 步骤3：预定义 Word 样式...")

    # 正文 Normal：宋体10.5pt，首行缩进2字符
    n = doc.styles["Normal"]
    n.font.name = EN_FONT
    n.font.size = Pt(BODY_SIZE)
    n.paragraph_format.space_before = Pt(0)
    n.paragraph_format.space_after  = Pt(6)
    n.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    n.paragraph_format.line_spacing      = 1.5   # 1.5倍行距
    # 首行缩进2字符
    n.paragraph_format.first_line_indent = Emu(int(2 * 21 * 914400 / 100))
    # 中文字体
    pPr = n.element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr"); n.element.append(pPr)
    rPr = n.element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr"); n.element.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts"); rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), ZH_FONT)
    rFonts.set(qn("w:ascii"),    EN_FONT)
    rFonts.set(qn("w:hAnsi"),    EN_FONT)

    # 标题样式：去掉编号（前缀），设置字体/字号
    for lvl, (size, space_b, space_a, bold_s) in {
        1: (18, 18, 8,  True),
        2: (16, 12, 6,  True),
        3: (14, 10, 6,  True),
        4: (13,  8, 4,  True),
        5: (12,  6, 4,  True),
    }.items():
        sid = f"Heading {lvl}"
        if sid not in doc.styles:
            continue
        s = doc.styles[sid]
        s.font.name = EN_FONT
        s.font.size = Pt(size)
        s.font.bold = bold_s
        s.font.color.rgb = RGBColor(0, 0, 0)
        s.paragraph_format.space_before = Pt(space_b)
        s.paragraph_format.space_after  = Pt(space_a)
        s.paragraph_format.keep_with_next = True
        # 中文字体
        rPr = s.element.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr"); s.element.append(rPr)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts"); rPr.insert(0, rFonts)
        rFonts.set(qn("w:eastAsia"), ZH_FONT)
        rFonts.set(qn("w:ascii"),    EN_FONT)
        rFonts.set(qn("w:hAnsi"),    EN_FONT)
        # 清除 Word 自动加的章节编号（前缀）
        pPr = s.element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr"); s.element.append(pPr)
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            pPr.remove(numPr)
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None:
            pStyle.set(qn("w:val"), sid)

    # 目录样式 TOC1/TOC2/TOC3：同正文
    for lvl in range(1, 4):
        sid = f"TOC {lvl}"
        if sid not in doc.styles:
            continue
        s = doc.styles[sid]
        s.font.name = EN_FONT
        s.font.size = Pt(BODY_SIZE)
        s.paragraph_format.space_before = Pt(0)
        s.paragraph_format.space_after  = Pt(3)
        rPr = s.element.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr"); s.element.append(rPr)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts"); rPr.insert(0, rFonts)
        rFonts.set(qn("w:eastAsia"), ZH_FONT)

    log("  ✅ 样式定义完成（Normal/Heading/TOC/CodeBlock）")
    return doc

def step5_inject_equations():
    """步骤5：在 docx XML 层面注入公式编号"""
    import sys as _sys
    sys.path.insert(0, "/tmp")
    try:
        from inject_eq_numbers import inject_numbers as _inj
        print("\n🔢 步骤5：注入公式编号...")
        doc = Document(OUTPUT_DOCX)
        injected, total = _inj(doc)
        doc.save(OUTPUT_DOCX)
        log(f"  ✅ 注入 {injected} 个公式编号")
        return True
    except Exception as e:
        log(f"  ⚠️  公式编号注入失败（不影响主流程）：{e}")
        return True   # 不阻断

def step4_postprocess():
    """
    后处理 DOCX：
    - 段落：首行缩进+行距
    - 代码块：等宽+灰色底纹
    - 表格：三线表+自适应列宽
    - 公式引用文字：Times New Roman
    - 页眉/页脚
    """
    print("\n🛠️ 步骤4：后处理 DOCX...")
    doc = Document(OUTPUT_DOCX)

    stats = {"段落": 0, "标题": 0, "表格": 0, "代码块": 0, "公式行": 0}

    for para in doc.paragraphs:
        sid = (para.style.name or "") if para.style else ""
        # 代码块识别
        if is_code_style(sid) or re.match(r'^\s{4}', (para.text or "")):
            stats["代码块"] += 1
            fix_code_block_para(para)
            continue

        # 标题段落
        if "Heading" in sid or "TOC" in sid:
            stats["标题"] += 1
            para.paragraph_format.space_before = para.paragraph_format.space_before
            para.paragraph_format.space_after  = para.paragraph_format.space_after
            for run in para.runs:
                run.font.bold = True
                set_run_font(run, zh=ZH_FONT, en=EN_FONT)
            continue

        # 普通正文段落：有内容才加首行缩进
        stats["段落"] += 1
        text = para.text.strip()
        pf = para.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing      = 1.5
        pf.space_before = Pt(0)
        pf.space_after  = Pt(0)

        if text:
            pf.first_line_indent = Emu(int(2 * 21 * 914400 / 100))
        else:
            pf.first_line_indent = Emu(0)

        for run in para.runs:
            set_run_font(run, zh=ZH_FONT, en=EN_FONT, size_pt=BODY_SIZE)

    # 表格处理
    pg_width_cm = TABLE_FULL_PAGE_WIDTH_CM
    for i, table in enumerate(doc.tables):
        stats["表格"] += 1
        set_table_three_line(table, available_width_cm=pg_width_cm)

    # 页面设置
    section = doc.sections[0]
    section.page_width   = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.0)  # 略窄，留给页码
    section.page_width  # 触发引用

    # 页眉
    header = section.header
    header.is_linked_to_previous = False
    if not header.paragraphs:
        header.add_paragraph()
    hp = header.paragraphs[0]
    hp.clear()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run("《土力学创新教材》（第三版草稿）")
    run.font.size  = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)
    run.font.name  = ZH_FONT

    # 页脚（页码）
    footer = section.footer
    footer.is_linked_to_previous = False
    if not footer.paragraphs:
        footer.add_paragraph()
    fp = footer.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    fldChar1 = OxmlElement("w:fldChar"); fldChar1.set(qn("w:fldCharType"), "begin")
    run._r.append(fldChar1)
    instrText = OxmlElement("w:instrText"); instrText.text = "PAGE"; run._r.append(instrText)
    fldChar2 = OxmlElement("w:fldChar"); fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar2)
    run.font.size = Pt(9)
    run.font.name = EN_FONT

    doc.save(OUTPUT_DOCX)
    log(f"  ✅ 后处理完成！")
    log(f"     段落 {stats['段落']} | 标题 {stats['标题']} | 表格 {stats['表格']} | 代码块 {stats['代码块']}")
    return True

# ==================== 入口 ====================

def run():
    print("=" * 60)
    print("📖 土力学创新教材 v2 — Markdown → DOCX 高质量转换")
    print("=" * 60)
    print(f"📂 源: {TEXTBOOK_DIR}")
    print(f"📄 输出: {OUTPUT_DOCX}")
    print()

    if not step1_merge_and_clean(): return False
    if not step2_pandoc():          return False

    doc = Document(OUTPUT_DOCX)
    step3_define_styles(doc)
    doc.save(OUTPUT_DOCX)

    if not step4_postprocess():      return False
    if not step5_inject_equations(): return False

    import datetime
    sz = os.path.getsize(OUTPUT_DOCX) / (1024*1024)
    print()
    print("=" * 60)
    print(f"✅ 转换完成！")
    print(f"   📄 {OUTPUT_DOCX}")
    print(f"   📦 {sz:.2f} MB")
    print(f"   🕐 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("💡 打开：open \"土力学创新教材.docx\"")
    print("=" * 60)

if __name__ == "__main__":
    run()
