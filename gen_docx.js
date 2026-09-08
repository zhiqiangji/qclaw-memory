const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
        PageNumber, Footer, VerticalAlign } = require('docx');

const SRC = '/Users/jizhiqiang/.qclaw/workspace-main/烟台南站流态固化土回填技术咨询_20260827.md';
const OUT = '/Users/jizhiqiang/.qclaw/workspace-main/烟台南站流态固化土回填技术咨询_20260827.docx';

const CJK = '宋体';
const CJK_HEAD = '黑体';

const md = fs.readFileSync(SRC, 'utf8');
const lines = md.split('\n');

// 行内加粗解析：按 ** 切分，奇数段加粗
function inlineRuns(text, opts) {
  const runs = [];
  const parts = text.split('**');
  parts.forEach((part, i) => {
    if (part === '') return;
    runs.push(new TextRun(Object.assign({ text: part }, opts, { bold: opts.bold || (i % 2 === 1) })));
  });
  if (runs.length === 0) runs.push(new TextRun(Object.assign({ text: '' }, opts)));
  return runs;
}

const bodyOpts = { font: CJK, size: 21 };       // 10.5pt
const bodyBoldOpts = { font: CJK, size: 21, bold: true };

const PAGE_W = 11906, PAGE_H = 16838;
const MARGIN = 1440;
const CONTENT_W = PAGE_W - MARGIN * 2; // 9026

function titlePara(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({ text, font: CJK_HEAD, size: 32, bold: true })]
  });
}
function h1(text) {
  return new Paragraph({
    spacing: { before: 300, after: 160 },
    children: [new TextRun({ text, font: CJK_HEAD, size: 28, bold: true })]
  });
}
function h2(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: CJK_HEAD, size: 24, bold: true })]
  });
}
function h3(text) {
  return new Paragraph({
    spacing: { before: 160, after: 100 },
    children: [new TextRun({ text, font: CJK_HEAD, size: 22, bold: true })]
  });
}
function body(text) {
  return new Paragraph({
    spacing: { after: 80 },
    children: inlineRuns(text, bodyOpts)
  });
}
function blockquote(text) {
  return new Paragraph({
    spacing: { after: 100, before: 60 },
    indent: { left: 360, right: 360 },
    shading: { fill: 'F2F2F2', type: ShadingType.CLEAR },
    children: inlineRuns(text, { font: CJK, size: 20, italics: false })
  });
}
function bullet(text, level) {
  return new Paragraph({
    spacing: { after: 60 },
    numbering: { reference: 'bullets', level },
    children: inlineRuns(text, bodyOpts)
  });
}
function numbered(text) {
  return new Paragraph({
    spacing: { after: 60 },
    numbering: { reference: 'numbers', level: 0 },
    children: inlineRuns(text, bodyOpts)
  });
}

const border = { style: BorderStyle.SINGLE, size: 4, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border,
                  insideHorizontal: border, insideVertical: border };

function buildTable(tblLines) {
  const rows = tblLines.map(l => l.split('|').slice(1, -1).map(c => c.trim()));
  const ncol = rows[0].length;
  const colW = Math.floor(CONTENT_W / ncol);
  const tableRows = [];
  rows.forEach((r, ri) => {
    if (ri === 1 && r.every(c => /^:?-+:?$/.test(c))) return; // 分隔行
    const isHeader = ri === 0;
    const cells = r.map(c => new TableCell({
      borders,
      width: { size: colW, type: WidthType.DXA },
      shading: isHeader ? { fill: 'D9E2F3', type: ShadingType.CLEAR } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        spacing: { after: 0 },
        children: inlineRuns(c, { font: CJK, size: 19, bold: isHeader })
      })]
    }));
    tableRows.push(new TableRow({ children: cells, tableHeader: isHeader }));
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: rows[0].map(() => colW),
    borders,
    rows: tableRows
  });
}

const children = [];
let i = 0;
while (i < lines.length) {
  const line = lines[i];
  const trimmed = line.trim();

  if (trimmed.startsWith('|')) {
    const tbl = [];
    while (i < lines.length && lines[i].trim().startsWith('|')) { tbl.push(lines[i]); i++; }
    children.push(buildTable(tbl));
    children.push(new Paragraph({ spacing: { after: 80 }, children: [] })); // 表后间距
    continue;
  }
  if (line.startsWith('# ')) { children.push(titlePara(line.slice(2).trim())); i++; continue; }
  if (line.startsWith('## ')) { children.push(h1(line.slice(3).trim())); i++; continue; }
  if (line.startsWith('### ')) { children.push(h2(line.slice(4).trim())); i++; continue; }
  if (line.startsWith('#### ')) { children.push(h3(line.slice(5).trim())); i++; continue; }
  if (trimmed.startsWith('> ')) { children.push(blockquote(trimmed.slice(2).trim())); i++; continue; }

  const bulletMatch = line.match(/^(\s*)-\s+(.*)$/);
  if (bulletMatch) {
    const lead = bulletMatch[1].replace(/\t/g, '  ');
    const level = Math.min(2, Math.floor(lead.length / 2));
    children.push(bullet(bulletMatch[2].trim(), level));
    i++; continue;
  }
  const numMatch = line.match(/^(\d+)\.\s+(.*)$/);
  if (numMatch) {
    children.push(numbered(numMatch[2].trim()));
    i++; continue;
  }
  if (trimmed === '') { i++; continue; }
  children.push(body(line));
  i++;
}

const doc = new Document({
  styles: { default: { document: { run: { font: CJK, size: 21 } } } },
  numbering: {
    config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        { level: 2, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
      ]},
      { reference: 'numbers', levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ]},
    ]
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } },
    footers: { default: new Footer({ children: [ new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: '第 ', font: CJK, size: 18 }),
        new TextRun({ children: [PageNumber.CURRENT], font: CJK, size: 18 }),
        new TextRun({ text: ' 页 / 共 ', font: CJK, size: 18 }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: CJK, size: 18 }),
        new TextRun({ text: ' 页', font: CJK, size: 18 }),
      ]
    }) ] }) },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('OK: ' + OUT + '  (' + buf.length + ' bytes)');
}).catch(e => { console.error('ERR', e); process.exit(1); });
