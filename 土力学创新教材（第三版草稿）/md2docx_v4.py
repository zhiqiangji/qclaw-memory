#!/usr/bin/env python3
"""
土力学创新教材 v4 — Markdown → DOCX 高质量转换
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
v4 修复清单（用户最新反馈）：
  1. 段首空两格太大 → 修正EMU计算，缩进改为正确2字符
  2. 表格内不应有段首缩进 → 表格内移除首行缩进
  3. 表格编号不规范，表名重复 → 去重复编号，只在表格上方注入一次
  4. 表格不紧凑 → 表格行高+行距调小，字体小一号
  5. 公式编号和公式同一行（右对齐）→ 公式末尾加制表位，然后编号
  6. 表格内字体比正文小一号 → 表格字体改为 9pt（正文10.5pt）
  7. 表格内1倍行距 → 修改表格行距
  8. 只有正文中公式编号，例题/计算步骤不编号 → 识别例题标题并跳过该区域公式
  9. Python代码应紧凑放在一个代码框 → Pandoc转后自动合并，无需额外处理
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
v3 修复清单：
  1. 二三级标题缺编号，修复层级
  2. 表格自动宽度+标题注入
  3. 只对正文（第1-12章）公式编号，格式 (m-n)
  4. 重排章节顺序：序→前言→目录→第1-12章→附录→非出版内容放最后
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import os, re, subprocess, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==================== 配置 ====================
TEXTBOOK_DIR = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材（第三版草稿）"
OUTPUT_DOCX  = "/Users/jizhiqiang/.qclaw/workspace/土力学创新教材.docx"
TEMP_MD      = "/tmp/土力学_v4_整合.md"

# ================================================================
# ★★★ 章节顺序 v4 ★★★
#   出版部分：序→前言→（目录pandoc自动生成）→第1-12章→附录
#   非出版部分（校内评审材料）：大纲、知识图谱 → 放最后
# ================================================================
BOOK_ORDER = [          # ★ 出版目录
    "序.md",
    "前言.md",
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
    "附录.md",
]
APPENDIX_ORDER = [     # ★ 非出版（校内评审）
    "大纲.md",
    "知识图谱.md",
]

EN_FONT    = "Times New Roman"
MONO_FONT  = "Courier New"
ZH_FONT    = "宋体"
BODY_SIZE  = 10.5       # 正文字号(pt)
TABLE_SIZE = 9.0        # ★ 表格内字号（比正文小一号）

# 公式计数器：{(章号, 小节号): n}
EQ_COUNTER = {}         # key=(ch,sec) → int
EQ_TOTAL   = 0

# 表格行数阈值：超过此行数启用自适应
PAGE_WIDTH_CM = 16.0   # 去掉左右边距后的净页面宽度

# 缩进修正：2字符缩进 → 计算得 ~266,700 emu（10.5pt下）
# 原计算偏大：2 * 914400 * 10.5 / 10 = 1920240 emu → 实际应为 ~266,700
def first_line_indent_emu(chars=2):
    """计算正确的2字符首行缩进：2字符 × 10.5pt = 2 × 10.5 × 12700 emu/pt = 266,700 emu"""
    return Emu(int(chars * BODY_SIZE * 12700))  # 每个pt ≈ 12700 emu

# ================================================================
# 工具函数
# ================================================================
def log(msg): print(f"  {msg}")

# ================================================================
# 步骤1：合并文件 + 正确标题层级
# ================================================================
def step1_merge():
    print("\n📋 步骤1：合并章节、写入编号...")

    merged_parts = []
    for fname in BOOK_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if not os.path.exists(fpath):
            log(f"⚠️  跳过不存在：{fname}"); continue

        text = Path(fpath).read_text(encoding="utf-8")
        ch_num = _chapter_num_from_filename(fname)
        is_chapter = fname.startswith("第") and "章" in fname

        text = _rewrite_headings(text, ch_num, is_chapter)
        merged_parts.append(text)

    # ---- 非出版部分（大纲/知识图谱） ----
    appendix_extra = []
    for fname in APPENDIX_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if os.path.exists(fpath):
            text = Path(fpath).read_text(encoding="utf-8")
            appendix_extra.append(text)
            log(f"  + {fname}（非出版）")

    full = "\n\n".join(merged_parts + appendix_extra)
    Path(TEMP_MD).write_text(full, encoding="utf-8")
    sz = os.path.getsize(TEMP_MD) / 1024
    log(f"✅ 合并完成 {len(BOOK_ORDER)+len(appendix_extra)} 个文件，{sz:.1f} KB")
    return True

def _chapter_num_from_filename(fname):
    m = re.search(r'第(\d+)章', fname)
    return m.group(1).zfill(2) if m else "00"

def _rewrite_headings(text, ch_num, is_chapter):
    """
    重写标题（全手动编号，不依赖 pandoc --number-sections）。

    关键修复：源文件里各章节节标题层级不一致。
    策略：识别"章标题行"（第X章开头的#行），其后的同级别#行自动升一级。
    - 章标题行保持 level=1
    - 紧随其后的同级 # 行 → level+1（二级标题）
    - 更低级别的 # 行保持不变
    同时处理缺章节号的章标题，自动补上。
    """
    lines = text.split('\n')
    result = []
    i = 0
    while i < len(lines):
        raw   = lines[i]
        strip = raw.strip()
        if not strip.startswith('#'):
            result.append(raw); i += 1; continue

        m = re.match(r'^(#+)\s+(.*)$', strip)
        if not m:
            result.append(raw); i += 1; continue

        level = len(m.group(1))
        title = m.group(2)

        # ★ 判断这是不是章标题行
        if re.match(r'第\d+章', title):
            # 已是第X章格式 → 一级标题
            result.append(f"# {title}" if level != 1 else raw)
            i += 1; continue
        elif is_chapter and level == 1 and title not in ("序","前言","附录","详细大纲","知识图谱"):
            # 单#缺章节号 → 补上（去掉前导零）
            ch_display = str(int(ch_num))
            result.append(f"# 第{ch_display}章 {title}")
            i += 1; continue

        # ★ 不是章标题：往前找最近的"第X章"标题级别
        chapter_level = None
        for j in range(i - 1, -1, -1):
            prev_raw = lines[j].strip()
            if not prev_raw.startswith('#'):
                continue
            pm = re.match(r'^(#+)\s+(.*)$', prev_raw)
            if pm:
                if re.match(r'第\d+章', pm.group(2)):
                    chapter_level = len(pm.group(1))
                    break

        if chapter_level is not None and level == chapter_level:
            # 同级节标题 → 升一级
            new_level = '#' * (level + 1)
            result.append(f"{new_level} {title}")
        else:
            result.append(raw)

        i += 1

    return '\n'.join(result)

def _add_eq_numbers(text, ch_num):
    """公式编号统一在 step5（docx XML层）处理。"""
    return text

def _add_tbl_numbers(text, ch_num):
    """表格编号在 step4（docx层）处理，md层不操作"""
    return text

# ================================================================
# 步骤2：Pandoc 转 DOCX（自动目录+标题层级）
# ================================================================
def step2_pandoc():
    print("\n🔄 步骤2：Pandoc 转 DOCX（目录+标题层级）...")
    cmd = [
        "pandoc", TEMP_MD, "-o", OUTPUT_DOCX,
        "--from=markdown", "--to=docx",
        "--table-of-contents",
        "--toc-depth=3",
        "--wrap=preserve",
        "-V", "lang=zh-CN",
        "-V", f"mainfont={EN_FONT}",
        "-V", f"sansfont={EN_FONT}",
        "-V", f"monofont={MONO_FONT}",
        "-V", "geometry:margin=2.5cm",
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode:
        log(f"❌ Pandoc 失败：{r.stderr[:300]}"); return False
    log(f"✅ Pandoc 完成 {os.path.getsize(OUTPUT_DOCX)/1024:.1f} KB")
    return True

# ================================================================
# 步骤3：预定义样式（Normal/Heading/TOC/CodeBlock）
# ================================================================
def step3_styles(doc):
    print("🎨 步骤3：预定义 Word 样式...")
    # Normal
    n = doc.styles["Normal"]
    pf = n.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.line_spacing      = 1.5
    pf.first_line_indent = first_line_indent_emu()
    pf.space_before = Pt(0); pf.space_after = Pt(6)
    n.font.name = ZH_FONT; n.font.size = Pt(BODY_SIZE)
    rPr = n.font._element
    rF  = rPr.find(qn("w:rFonts"))
    if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
    rF.set(qn("w:eastAsia"), ZH_FONT)

    # Headings
    for lvl, name in [(1,"Heading 1"),(2,"Heading 2"),(3,"Heading 3"),(4,"Heading 4")]:
        s = doc.styles[name]
        s.font.name  = ZH_FONT
        s.font.size  = Pt(BODY_SIZE + 6 - lvl * 1.5)
        s.font.bold  = True
        s.paragraph_format.space_before = Pt(12)
        s.paragraph_format.space_after  = Pt(6)
        s.paragraph_format.keep_with_next = True

    # TOC
    for nm in ["TOC Heading","TOC1","TOC2","TOC3"]:
        if nm in [s.name for s in doc.styles]:
            s = doc.styles[nm]
            s.font.name = ZH_FONT; s.font.size = Pt(BODY_SIZE)

    # CodeBlock
    if "CodeBlock" not in [s.name for s in doc.styles]:
        cb = doc.styles.add_style("CodeBlock", 1)  # PARAGRAPH=1
    else:
        cb = doc.styles["CodeBlock"]
    cb.font.name = MONO_FONT; cb.font.size = Pt(9)
    cb.paragraph_format.space_before = Pt(4)
    cb.paragraph_format.space_after  = Pt(4)
    log("  ✅ 样式定义完成"); return doc

# ================================================================
# 步骤4：后处理（首行缩进/三线表/表格格式/字体/行距）
# ================================================================
def _tbl_width(table):
    """获取表格当前总宽度（emu）"""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblW"))
    if tblPr is not None:
        w = tblPr.find(qn("w:w"))
        if w is not None:
            return int(w.get(qn("w:w")))
    return 0

def _set_tbl_width(table, width_cm):
    """将表格宽度设为指定cm"""
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    w = tblPr.find(qn("w:tblW"))
    if w is None:
        w = OxmlElement("w:tblW"); tblPr.append(w)
    emu = int(width_cm * 360000)
    w.set(qn("w:w"),    str(emu))
    w.set(qn("w:type"), "dxa")

def _set_cell_width(cell, width_cm):
    tc = cell._tc; tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None: tcPr = OxmlElement("w:tcPr"); tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW"); tcPr.append(tcW)
    emu = int(width_cm * 360000)
    tcW.set(qn("w:w"),    str(emu))
    tcW.set(qn("w:type"), "dxa")

def _tbl_borders_three_line(table):
    """三线表：只保留顶线、底线、水平分隔线"""
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None: tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None: borders = OxmlElement("w:tblBorders"); tblPr.append(borders)

    for side in ["top","bottom","insideH","insideV","left","right"]:
        el = borders.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}"); borders.append(el)
        el.set(qn("w:val"), "single")
        if side in ["top","bottom"]:
            el.set(qn("w:sz"), "8"); el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        elif side == "insideH":
            el.set(qn("w:sz"), "4"); el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:sz"), "0"); el.set(qn("w:val"), "nil")

def _inject_tbl_caption(tbl_elem, tbl_num, doc):
    """在表格上方插入表名标题段落"""
    body = doc.element.body
    idx  = list(body).index(tbl_elem)
    p    = doc.add_paragraph()
    p.clear()
    pf = p.paragraph_format
    pf.alignment    = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(6); pf.space_after = Pt(2)
    run  = p.add_run(tbl_num)
    run.font.name = ZH_FONT; run.font.size = Pt(9); run.font.bold = True
    r = run._element; rPr = r.get_or_add_rPr()
    rF = rPr.find(qn("w:rFonts"))
    if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
    rF.set(qn("w:eastAsia"), ZH_FONT)
    body.insert(idx, p._element)
    return p

def step4_postprocess(doc):
    print("🛠️ 步骤4：后处理（缩进/三线表/表名/表格字号/行距/...）...")

    tbl_counter  = {}   # {ch_num: n}
    ch_num       = "pre"  # 初始为pre（不编号）
    page_width   = PAGE_WIDTH_CM

    def _set_cell_font(cell):
        """表格内字体：字号比正文小一号（9pt），字体不变"""
        for para in cell.paragraphs:
            # 去掉首行缩进 ★★★
            pf = para.paragraph_format
            pf.first_line_indent = None  # 表格内不缩进

            for run in para.runs:
                run.font.name = ZH_FONT
                run.font.size = Pt(TABLE_SIZE)
                r = run._element; rPr = r.get_or_add_rPr()
                rF = rPr.find(qn("w:rFonts"))
                if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
                rF.set(qn("w:eastAsia"), ZH_FONT)
                rF.set(qn("w:ascii"), EN_FONT)
                rF.set(qn("w:hAnsi"), EN_FONT)
            # 表格内单倍行距 ★★★
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pf.line_spacing = 1.0

    body = doc.element.body
    children = list(body)
    i = 0
    while i < len(children):
        elem = children[i]
        tag  = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        if tag == "p":
            pPr   = elem.find(qn("w:pPr"))
            style = None
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None: style = pStyle.get(qn("w:val"))

            text = "".join(t.text or "" for t in elem.iter(qn("w:t")))

            # ---- 更新章节（更新表格计数）----
            if style and "Heading 1" in style:
                m = re.search(r'第(\d+)章', text)
                if m:
                    ch_num = m.group(1).zfill(2)
                else:
                    # 序/前言/附录 → 不编号表格
                    ch_num = "pre"
                tbl_counter[ch_num] = 0

            # ---- 正文段落首行缩进 ----
            if style in (None, "Normal"):
                if pPr is None:
                    pPr = OxmlElement("w:pPr"); elem.insert(0, pPr)
                indent = pPr.find(qn("w:ind"))
                if indent is None:
                    indent = OxmlElement("w:ind"); pPr.append(indent)
                indent.set(qn("w:firstLine"), str(int(first_line_indent_emu())))
                indent.set(qn("w:firstLineChars"), "200")

            # ---- 代码块格式 ----
            if style in ("CodeBlock", "Fragment", "Source Code"):
                if pPr is None: pPr = OxmlElement("w:pPr"); elem.insert(0, pPr)
                shd = pPr.find(qn("w:shd"))
                if shd is None:
                    shd = OxmlElement("w:shd"); pPr.append(shd)
                shd.set(qn("w:val"),   "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"),  "F2F2F2")
                for run in elem.iter(qn("w:r")):
                    rn = run.find(qn("w:rPr"))
                    if rn is None: rn = OxmlElement("w:rPr"); run.insert(0, rn)
                    rFonts = rn.find(qn("w:rFonts"))
                    if rFonts is None: rFonts = OxmlElement("w:rFonts"); rn.insert(0, rFonts)
                    rFonts.set(qn("w:ascii"),   MONO_FONT)
                    rFonts.set(qn("w:hAnsi"),   MONO_FONT)
                    rFonts.set(qn("w:eastAsia"), MONO_FONT)
                    sz = rn.find(qn("w:sz"))
                    if sz is None: sz = OxmlElement("w:sz"); rn.append(sz)
                    sz.set(qn("w:val"), "18")  # 9pt → 18 half-points

        elif tag == "tbl":
            # 找当前表格索引
            tbl_idx = sum(1 for j in range(i) if children[j].tag.endswith('}tbl'))
            table   = doc.tables[tbl_idx] if tbl_idx < len(doc.tables) else None

            if table is not None:
                # ★★★ 三线表格式
                _tbl_borders_three_line(table)

                # ★★★ 表格自动宽度（撑满页面）
                n_cols = len(table.columns)
                if n_cols > 0:
                    col_w = page_width / n_cols
                    for row in table.rows:
                        for cell in row.cells:
                            _set_cell_width(cell, col_w)
                            # ★★★ 表格内字体+行距+去缩进
                            _set_cell_font(cell)
                    _set_tbl_width(table, page_width)

                # ★★★ 表名注入（仅正文章节，不是 pre）
                if ch_num != "pre":
                    tbl_counter[ch_num] = tbl_counter.get(ch_num, 0) + 1
                    ch_display = str(int(ch_num))
                    tbl_num = f"表 {ch_display}-{tbl_counter[ch_num]:02d}"
                    _inject_tbl_caption(elem, tbl_num, doc)
                    children = list(body); i += 1

        i += 1

    log(f"  ✅ 后处理完成！{sum(tbl_counter.values())}个表格完成编号")
    return doc

# ================================================================
# 步骤5：公式编号注入（公式后右对齐编号，同一行 → 制表位实现）
# ================================================================
def step5_inject_eq_numbers(doc):
    """
    公式编号：
    - 跳过 序/前言/大纲/知识图谱/例题 → 只给正文章节正文公式编号
    - 例题识别：标题含"例题" → 跳过该区域直到下一个一级标题
    - 编号放在公式同一行 → 添加一个右对齐制表位，然后写入编号
    """
    print("🔢 步骤5：公式编号注入（仅正文，格式 (m-n)）...")

    SKIP_HEADINGS = {
        "序", "前言", "附录",
        "《土力学》创新教材 — 详细大纲",
        "土力学创新教材知识图谱"
    }
    ch_num   = "00"
    eq_no    = {}
    injected = 0
    skip     = True  # 默认跳过正文外
    in_example = False  # ★ 是否在例题区域（例题不编号）

    for para in doc.paragraphs:
        sname = para.style.name if para.style else ""

        # 检测 Heading 1 → 更新章节+例题状态
        if "Heading 1" in sname:
            text = para.text.strip()
            m = re.search(r'第(\d+)章', text)
            if m:
                ch_num = m.group(1).zfill(2)
                eq_no[ch_num] = 0
                skip = False
            elif text in SKIP_HEADINGS or "知识图谱" in text:
                skip = True; ch_num = "00"
            else:
                # 序/前言 → 跳过
                skip = True; ch_num = "00"
            # ★ 只要更新章节，就重置例题状态
            in_example = False

        # 例题检测：Heading 包含"例题" → 跳过编号
        if para.text.strip().startswith("例题") and any(h in sname for h in ["Heading","Heading 2","Heading 3"]):
            in_example = True

        # 跳过规则：
        # 1. skip = True → 全局跳过（序言/前言）
        # 2. in_example = True → 当前在例题中，跳过编号
        if skip or ch_num == "00" or in_example:
            continue

        # 检测段落是否有块级公式（oMathPara）
        has_eq = any("oMathPara" in e.tag for e in para._p.iter())
        if not has_eq:
            continue

        # ★★★ 公式编号同一行：添加右对齐制表位，然后插入编号 ★★★
        # 1. 设置制表位：右对齐到页宽（16cm → 16*360000 = 5760000 emu）
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr"); para._p.insert(0, pPr)
        tabstops = pPr.find(qn("w:tabs"))
        if tabstops is None:
            tabstops = OxmlElement("w:tabs"); pPr.append(tabstops)
        # 添加最右侧制表位（右对齐）
        tab = OxmlElement("w:tab"); tab.set(qn("w:pos"), str(int(16 * 360000)))
        tab.set(qn("w:val"), "right"); tabstops.append(tab)

        # 2. 在段落最后添加制表符 + 编号
        eq_no[ch_num] = eq_no.get(ch_num, 0) + 1
        eq_num_str = f"({int(ch_num)}-{eq_no[ch_num]:02d})"
        run = para.add_run("\t" + eq_num_str)
        run.font.name = EN_FONT; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(70,70,70)
        r = run._element; rPr = r.get_or_add_rPr()
        rF = rPr.find(qn("w:rFonts"))
        if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
        rF.set(qn("w:eastAsia"), ZH_FONT)
        rF.set(qn("w:ascii"), EN_FONT); rF.set(qn("w:hAnsi"), EN_FONT)

        injected += 1

    ch_stats = {int(k): v for k,v in sorted(eq_no.items()) if v>0}
    log(f"  ✅ 注入 {injected} 个公式编号（公式同一行右对齐）")
    log(f"     各章公式数：{ch_stats}")
    return doc

# ================================================================
# 主流程
# ================================================================
def run():
    print("=" * 60)
    print("📖 土力学创新教材 v4 — Markdown → DOCX")
    print("=" * 60)
    print(f"📂 源: {TEXTBOOK_DIR}")
    print(f"📄 输出: {OUTPUT_DOCX}")
    print(f"📚 出版章节: {' / '.join(f[:f.index('_')] if '_' in f else f for f in BOOK_ORDER[:5])}...")

    if not step1_merge():  return
    if not step2_pandoc(): return

    doc = Document(OUTPUT_DOCX)
    doc = step3_styles(doc)
    doc = step4_postprocess(doc)
    doc = step5_inject_eq_numbers(doc)
    doc.save(OUTPUT_DOCX)

    sz = os.path.getsize(OUTPUT_DOCX) / (1024*1024)
    print(f"\n{'='*60}")
    print(f"✅ 转换完成！")
    print(f"   📄 {OUTPUT_DOCX}")
    print(f"   📦 {sz:.2f} MB")
    print(f"   📅 {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")
    print(f"💡 打开：open \"{OUTPUT_DOCX}\"")
    print(f"💡 本地脚本：cd \"{TEXTBOOK_DIR}\" && python3 md2docx_v4.py")

if __name__ == "__main__":
    run()
