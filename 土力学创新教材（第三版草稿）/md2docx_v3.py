#!/usr/bin/env python3
"""
土力学创新教材 v3 — Markdown → DOCX 高质量转换
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
修复清单（来自用户反馈）：
  1. 二三级标题加上编号（## 1.1.1 保留；去序言里的重复章节标记）
  2. 表格加标题（"表 X-Y: 名称"）+ 自动调宽度（撑满页面）
  3. 表格加自动宽度（set_width 撑满）
  4. 公式只在正文（第1-12章）编号，跳过序/前言
  5. 公式编号改为 (m-n) 格式（章节号-序号）
  6. 章节顺序：序→前言→目录→第1-12章→附录；大纲/知识图谱等放最后（不出版）
  7. 前言和第一章之间不插入多余内容
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
TEMP_MD      = "/tmp/土力学_v3_整合.md"

# ================================================================
# ★★★ 章节顺序 v3 ★★★
#   出版部分：序→前言→（目录由pandoc自动生成）→第1-12章→附录
#   非出版部分（校内评审材料）：大纲、知识图谱 → 放最后
# ================================================================
BOOK_ORDER = [          # ★ 出版目录
    "序.md",
    "前言.md",
    # --- 第1-12章 ---
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
BODY_SIZE  = 10.5

# 公式编号：只在第1-12章生效，跳过序/前言
# 编号格式：(m-n)，m=章号，n=该章第n个公式
EQ_COUNTER = {}   # {ch_num: n}
EQ_TOTAL   = 0
_eq_active = False  # 是否进入正文第一章

# 表格计数器
TBL_COUNTER = {}   # {ch_num: n}

# 页面宽度（CM）用于表格自动撑满
PAGE_WIDTH_CM = 16.0   # 去掉左右边距2.5cm后的净宽

# ================================================================
# 工具函数
# ================================================================
def log(msg): print(f"  {msg}")

def first_line_indent_emu(chars=2):
    """2字符首行缩进"""
    return Emu(int(chars * 914400 * 10.5 / 10))  # ~30000 emu

# ================================================================
# 步骤1：合并文件 + 重写章节标题（加编号） + 公式编号（正文才编号）
# ================================================================
def step1_merge():
    print("\n📋 步骤1：合并章节、写入编号...")
    global EQ_COUNTER, EQ_TOTAL, _eq_active, TBL_COUNTER

    EQ_COUNTER.clear(); EQ_TOTAL = 0; _eq_active = False
    TBL_COUNTER.clear()

    merged_parts = []

    # ---- 处理出版部分 ----
    for fname in BOOK_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if not os.path.exists(fpath):
            log(f"⚠️  跳过不存在：{fname}"); continue

        text = Path(fpath).read_text(encoding="utf-8")
        ch_num = _chapter_num_from_filename(fname)
        is_chapter = fname.startswith("第") and "章" in fname

        if is_chapter:
            _eq_active = True
            EQ_COUNTER[ch_num] = 0
            TBL_COUNTER[ch_num] = 0

        # ★★★ 标题重写 ★★★
        text = _rewrite_headings(text, ch_num, is_chapter)
        # 公式编号（仅正文）
        text = _add_eq_numbers(text, ch_num)
        # 表格编号（所有章节都编）
        text = _add_tbl_numbers(text, ch_num)

        merged_parts.append(text)

    # ---- 处理非出版部分（大纲/知识图谱）----
    # 它们作为无编号附录跟在附录后面
    appendix_extra = []
    for fname in APPENDIX_ORDER:
        fpath = os.path.join(TEXTBOOK_DIR, fname)
        if os.path.exists(fpath):
            text = Path(fpath).read_text(encoding="utf-8")
            # 去掉 FILE: 前缀（不是章节）
            text = re.sub(r'^FILE:.*$\n?', '', text, flags=re.MULTILINE)
            # 保留一级标题作为附录标记（不加章节号）
            text = re.sub(r'^#\s+', '# ', text, flags=re.MULTILINE)
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
            # 单#章节 → 补上章节号（去前导零）
            ch_display = str(int(ch_num))  # 07 → 7
            result.append(f"# 第{ch_display}章 {title}")
            i += 1; continue

        # ★ 不是章标题行：往前找最近的"第X章"标题级别
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
                # 找到"第X章"就停（不再往更早回溯）

        if chapter_level is not None and level == chapter_level:
            # 同级 → 升一级（二级/三级标题）
            new_level = '#' * (level + 1)
            result.append(f"{new_level} {title}")
        else:
            result.append(raw)

        i += 1

    return '\n'.join(result)

def _add_eq_numbers(text, ch_num):
    """
    公式编号统一在 step5（docx XML层）处理。
    这里只打标记行，step5 扫描 oMathPara 时可跳过含 MARKER 的行。
    """
    # ★ 不做任何操作，所有编号在 docx post-process 层做
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
        "--table-of-contents",    # ★ 自动生成目录
        "--toc-depth=3",
        # ★ 不加 --number-sections（全手动编号）
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
# 步骤4：后处理（标题格式/三线表/表格自动宽度/段落首行缩进/代码块）
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
    """三线表：只保留顶线、底线、隔线"""
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

def _add_tbl_caption(table, caption, doc):
    """在表格上方插入表名标题段落"""
    tbl_elem = table._tbl
    tbl_pr   = tbl_elem.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr"); tbl_elem.insert(0, tbl_pr)

    # 在 tblPr 之后插入空行（pPr before table）
    # 实际上我们在 table element 之前插入一个段落
    body = doc.element.body
    tbl_idx = list(body).index(tbl_elem)
    para = doc.add_paragraph()
    para.clear()
    pf = para.paragraph_format
    pf.alignment    = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(6); pf.space_after = Pt(2)
    run = para.add_run(caption)
    run.font.name  = ZH_FONT
    run.font.size  = Pt(9)
    run.font.bold  = True
    r = run._element; rPr = r.get_or_add_rPr()
    rF = rPr.find(qn("w:rFonts"))
    if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
    rF.set(qn("w:eastAsia"), ZH_FONT)

    # 插入到表格之前
    body.insert(tbl_idx, para._element)
    return para

def step4_postprocess(doc):
    print("🛠️ 步骤4：后处理（标题/三线表/表名/自动宽度/代码块）...")

    tbl_counter  = {}   # {ch_num: n}
    ch_num       = "00"
    tbl_active   = False
    page_width   = PAGE_WIDTH_CM
    p_style_buf  = []   # 缓冲段落样式队列

    def _rFonts(run, zh, en):
        r = run._element; rPr = r.get_or_add_rPr()
        rF = rPr.find(qn("w:rFonts"))
        if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
        rF.set(qn("w:eastAsia"), zh); rF.set(qn("w:ascii"), en)
        rF.set(qn("w:hAnsi"), en); rF.set(qn("w:cs"), en)

    def _inject_tbl_caption(tbl_elem, tbl_num, doc):
        body = doc.element.body
        idx  = list(body).index(tbl_elem)
        p    = doc.add_paragraph()
        p.clear()
        p.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(2)
        run  = p.add_run(tbl_num)
        run.font.name = ZH_FONT; run.font.size = Pt(9); run.font.bold = True
        r = run._element; rPr = r.get_or_add_rPr()
        rF = rPr.find(qn("w:rFonts"))
        if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
        rF.set(qn("w:eastAsia"), ZH_FONT)
        body.insert(idx, p._element)
        return p

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

            # ---- 跟踪章节（更新 ch_num）----
            if style and "Heading 1" in style:
                m = re.search(r'第(\d+)章', text)
                if m:
                    ch_num = m.group(1).zfill(2)
                else:
                    # 序/前言/附录 → 禁用表格编号
                    ch_num = "pre"
                tbl_counter[ch_num] = 0; tbl_active = False

            # ---- 段落首行缩进（正文）----
            if style in (None, "Normal"):
                if pPr is None:
                    pPr = OxmlElement("w:pPr"); elem.insert(0, pPr)
                indent = pPr.find(qn("w:ind"))
                if indent is None:
                    indent = OxmlElement("w:ind"); pPr.append(indent)
                indent.set(qn("w:firstLine"), str(int(first_line_indent_emu())))
                indent.set(qn("w:firstLineChars"), "200")

            # ---- 代码块 ----
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
                    sz.set(qn("w:val"), "18")  # 9pt

            # ---- 表格标题（检测表格前紧邻的文本行）----
            # 策略：找形如 "表X-Y" 或 "表 X-Y" 的段落，后面紧跟表格
            if i + 1 < len(children):
                next_elem = children[i + 1]
                next_tag  = next_elem.tag.split('}')[-1]
                if next_tag == "tbl":
                    m_tbl = re.match(r'^(表\s*[\d\-]+[:：]?\s*)(.*)$', text.strip())
                    if m_tbl:
                        # 已有表名 → 格式化
                        tbl_counter[ch_num] = tbl_counter.get(ch_num, 0) + 1
                        new_caption = f"表 {ch_num}-{tbl_counter[ch_num]:02d}：{m_tbl.group(2) or '（续）'}"
                        next_elem2 = next_elem
                        # 替换这段落文字
                        for t in elem.iter(qn("w:t")):
                            t.text = ""
                        # 找第一个 run
                        runs = list(elem.iter(qn("w:r")))
                        if runs:
                            runs[0].find(qn("w:t")).text = new_caption
                        # 加粗
                        for run in elem.iter(qn("w:r")):
                            rn = run.find(qn("w:rPr"))
                            if rn is None: rn = OxmlElement("w:rPr"); run.insert(0, rn)
                            b = rn.find(qn("w:b"))
                            if b is None: b = OxmlElement("w:b"); rn.append(b)
                            b.set(qn("w:val"), "1")
                            rF = rn.find(qn("w:rFonts"))
                            if rF is None: rF = OxmlElement("w:rFonts"); rn.insert(0, rF)
                            rF.set(qn("w:eastAsia"), ZH_FONT)
                            rF.set(qn("w:ascii"),    EN_FONT)

        elif tag == "tbl":
            # 找当前是第几个表格
            tbl_idx = sum(1 for j in range(i) if children[j].tag.endswith('}tbl'))
            table   = doc.tables[tbl_idx] if tbl_idx < len(doc.tables) else None

            if table is not None:
                # ★★★ 三线表（所有表格都做）
                _tbl_borders_three_line(table)

                # ★★★ 表格自动宽度（撑满页面）
                n_cols = len(table.columns)
                if n_cols > 0:
                    col_w = page_width / n_cols
                    for row in table.rows:
                        for k, cell in enumerate(row.cells):
                            _set_cell_width(cell, col_w)
                    _set_tbl_width(table, page_width)

                # ★★★ 表名（仅正文）
                if ch_num != "pre":
                    tbl_counter[ch_num] = tbl_counter.get(ch_num, 0) + 1
                    ch_display = str(int(ch_num))   # 07→7
                    tbl_num = f"表 {ch_display}-{tbl_counter[ch_num]:02d}"
                    _inject_tbl_caption(elem, tbl_num, doc)
                    children = list(body); i += 1

        i += 1

    log(f"  ✅ 后处理完成！")
    return doc

# ================================================================
# 步骤5：公式编号注入（仅正文，用 (m-n) 格式）
# ================================================================
def _make_eq_number_para(eq_num, doc):
    para = doc.add_paragraph()
    para.clear()
    pf = para.paragraph_format
    pf.alignment    = WD_ALIGN_PARAGRAPH.RIGHT
    pf.space_before = Pt(0); pf.space_after = Pt(4)
    run = para.add_run(eq_num)
    run.font.name  = EN_FONT; run.font.size  = Pt(9)
    run.font.bold  = False;   run.font.color.rgb = RGBColor(70,70,70)
    r = run._element; rPr = r.get_or_add_rPr()
    rF = rPr.find(qn("w:rFonts"))
    if rF is None: rF = OxmlElement("w:rFonts"); rPr.insert(0, rF)
    rF.set(qn("w:eastAsia"), ZH_FONT)
    rF.set(qn("w:ascii"), EN_FONT); rF.set(qn("w:hAnsi"), EN_FONT)
    return para

def step5_inject_eq_numbers(doc):
    """
    用 python-docx paragraphs 迭代器注入公式编号（含表格内段落）。
    跳过 序/前言/大纲/知识图谱，只对第1-12章的公式编号。
    """
    print("🔢 步骤5：公式编号注入（正文，仅 (m-n) 格式）...")

    SKIP_HEADINGS = {
        "序", "前言", "附录",
        "《土力学》创新教材 — 详细大纲",
        "土力学创新教材知识图谱"
    }
    ch_num   = "00"
    eq_no    = {}
    injected = 0
    skip     = True
    pending  = []   # (paragraph_obj, eq_num_str)

    # doc.paragraphs 包含：顶层段落 + 表格/页眉/脚注中的段落
    for para in doc.paragraphs:
        sname = para.style.name if para.style else ""

        # 检测 Heading 1 → 更新章节状态
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
                # 序/前言 → 关闭编号
                skip = True; ch_num = "00"

        if skip or ch_num == "00":
            continue

        # 检测段落是否含 oMathPara（块级公式）
        has_eq = any("oMathPara" in e.tag for e in para._p.iter())
        if has_eq:
            eq_no[ch_num] = eq_no.get(ch_num, 0) + 1
            pending.append((para, f"({ch_num}-{eq_no[ch_num]:02d})"))
            injected += 1

    # ★ 从后往前插入（避免段落索引偏移）★
    for para, eq_num_str in reversed(pending):
        new_p = _make_eq_number_para(eq_num_str, doc)
        para._p.addnext(new_p._element)

    ch_stats = {k: v for k, v in sorted(eq_no.items()) if v > 0}
    log(f"  ✅ 注入 {injected} 个公式编号")
    log(f"     各章公式数：{ch_stats}")
    return doc

# ================================================================
# 主流程
# ================================================================
def run():
    print("=" * 60)
    print("📖 土力学创新教材 v3 — Markdown → DOCX")
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
    print(f"💡 本地脚本：cd \"{TEXTBOOK_DIR}\" && python3 md2docx_v2.py")

if __name__ == "__main__":
    run()
