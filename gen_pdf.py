# -*- coding: utf-8 -*-
import re, os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, ListFlowable, ListItem)
from reportlab.lib.styles import ParagraphStyle

SRC = '/Users/jizhiqiang/.qclaw/workspace-main/烟台南站流态固化土回填技术咨询_20260827.md'
OUT = '/Users/jizhiqiang/.qclaw/workspace-main/烟台南站流态固化土回填技术咨询_20260827.pdf'

# ---- 字体：Songti.ttc 简体中文 Regular/Bold/Black + Menlo(等宽, 流程图) ----
SONG = '/System/Library/Fonts/Supplemental/Songti.ttc'
MONO = '/System/Library/Fonts/Menlo.ttc'
pdfmetrics.registerFont(TTFont('Song', SONG, subfontIndex=6))      # Regular
pdfmetrics.registerFont(TTFont('SongB', SONG, subfontIndex=1))     # Bold
pdfmetrics.registerFont(TTFont('SongBlack', SONG, subfontIndex=0)) # Black
pdfmetrics.registerFont(TTFont('Mono', MONO, subfontIndex=0))      # Menlo Regular
registerFontFamily('Song', normal='Song', bold='SongB', italic='Song', boldItalic='SongB')

# ---- 样式 ----
bodyStyle  = ParagraphStyle('body', fontName='Song', fontSize=10.5, leading=15.5, spaceAfter=5)
titleStyle = ParagraphStyle('title', fontName='SongBlack', fontSize=17, leading=23, alignment=TA_CENTER, spaceAfter=14)
h1Style    = ParagraphStyle('h1', fontName='SongBlack', fontSize=14, leading=19, spaceBefore=12, spaceAfter=7)
h2Style    = ParagraphStyle('h2', fontName='SongBlack', fontSize=12, leading=16.5, spaceBefore=9, spaceAfter=5)
h3Style    = ParagraphStyle('h3', fontName='SongB', fontSize=11, leading=15, spaceBefore=7, spaceAfter=4)
quoteStyle = ParagraphStyle('quote', fontName='Song', fontSize=9, leading=13.5, leftIndent=4)
monoStyle  = ParagraphStyle('mono', fontName='Mono', fontSize=8.5, leading=12)
cellStyle  = ParagraphStyle('cell', fontName='Song', fontSize=8.5, leading=11.5)
cellHead   = ParagraphStyle('cellh', fontName='SongB', fontSize=8.5, leading=11.5)
listStyle  = ParagraphStyle('li', parent=bodyStyle, spaceAfter=3)

PAGE_W, PAGE_H = A4
MARGIN = 1.9 * cm
FRAME_W = PAGE_W - 2 * MARGIN
BOXCHARS = set('├│└┌┐┘─┬┴┼↓')

def md_inline(s):
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    out = ''
    for i, p in enumerate(s.split('**')):
        if p == '':
            continue
        out += ('<b>' + p + '</b>') if (i % 2 == 1) else p
    return out

def make_table(tbl_lines):
    rows = [[c.strip() for c in l.split('|')[1:-1]] for l in tbl_lines]
    ncol = len(rows[0]); colW = FRAME_W / ncol
    data = []
    for ri, r in enumerate(rows):
        if ri == 1 and all(re.match(r'^:?-+:?$', c) for c in r):
            continue
        st = cellHead if ri == 0 else cellStyle
        data.append([Paragraph(md_inline(c), st) for c in r])
    t = Table(data, colWidths=[colW] * ncol, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9E2F3')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#999999')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4), ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]))
    return t

def make_flow():
    lines = open(SRC, encoding='utf-8').read().split('\n')
    flow = []; i = 0
    while i < len(lines):
        line = lines[i]; t = line.strip()
        if t.startswith('|'):
            tbl = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                tbl.append(lines[i]); i += 1
            flow.append(make_table(tbl)); flow.append(Spacer(1, 6)); continue
        if line.startswith('# '):
            flow.append(Paragraph(t[2:].strip(), titleStyle)); i += 1; continue
        if line.startswith('## '):
            flow.append(Paragraph(md_inline(t[3:].strip()), h1Style)); i += 1; continue
        if line.startswith('### '):
            flow.append(Paragraph(md_inline(t[4:].strip()), h2Style)); i += 1; continue
        if line.startswith('#### '):
            flow.append(Paragraph(md_inline(t[5:].strip()), h3Style)); i += 1; continue
        if t.startswith('> '):
            txt = t[2:].strip()
            q = Table([[Paragraph(md_inline(txt), quoteStyle)]], colWidths=[FRAME_W])
            q.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F2F2F2')),
                ('LEFTPADDING', (0, 0), (-1, -1), 8), ('RIGHTPADDING', (0, 0), (-1, -1), 8),
                ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
            flow.append(q); i += 1; continue
        m = re.match(r'^(\s*)-\s+(.*)$', line)
        if m:
            lvl = min(2, len(m.group(1).replace('\t', '  ')) // 2)
            bt = '○' if lvl >= 1 else '•'
            flow.append(Paragraph(md_inline(m.group(2).strip()),
                                  ParagraphStyle('b%d' % lvl, parent=listStyle,
                                                 leftIndent=14 + lvl * 14, bulletIndent=4),
                                  bulletText=bt)); i += 1; continue
        m = re.match(r'^(\d+)\.\s+(.*)$', line)
        if m:
            grp = []
            while i < len(lines) and re.match(r'^(\d+)\.\s+(.*)$', lines[i]):
                grp.append(re.match(r'^(\d+)\.\s+(.*)$', lines[i]).group(2).strip()); i += 1
            items = [ListItem(Paragraph(md_inline(x), listStyle), value=j + 1) for j, x in enumerate(grp)]
            flow.append(ListFlowable(items, bulletType='1', leftIndent=18,
                                     bulletFontName='Song', bulletFontSize=10)); continue
        if t == '':
            i += 1; continue
        if any(c in BOXCHARS for c in line):
            flow.append(Paragraph(line.replace(' ', '&nbsp;'), monoStyle)); i += 1; continue
        flow.append(Paragraph(md_inline(line), bodyStyle)); i += 1
    return flow

# ---- 页脚：第一遍计数，第二遍绘制总页码 ----
TOTAL = [0]; COUNTING = [True]
def draw_footer(canvas, doc):
    if COUNTING[0]:
        TOTAL[0] += 1; return
    canvas.saveState(); canvas.setFont('Song', 8)
    canvas.setFillColor(colors.HexColor('#666666'))
    canvas.drawCentredString(PAGE_W / 2, 1.1 * cm, '第 %d 页 / 共 %d 页' % (doc.page, TOTAL[0]))
    canvas.restoreState()

frame = Frame(MARGIN, MARGIN, FRAME_W, PAGE_H - 2 * MARGIN, id='main')
doc = BaseDocTemplate(OUT, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                      topMargin=MARGIN, bottomMargin=MARGIN,
                      title='烟台南站站前路流态固化土回填技术咨询')
doc.addPageTemplates([PageTemplate(id='main', frames=[frame], onPage=draw_footer)])

doc.build(make_flow())      # 第一遍：计数
COUNTING[0] = False
doc.build(make_flow())      # 第二遍：绘制总页码

print('OK: %s  (%d bytes, %d pages)' % (OUT, os.path.getsize(OUT), TOTAL[0]))
