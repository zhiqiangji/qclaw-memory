
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat, ShadingType, WidthType, BorderStyle, Table, TableRow, TableCell } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "宋体", size: 24 },
        paragraph: { spacing: { line: 360 } }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, font: "宋体" },
        paragraph: { spacing: { before: 240, after: 240 }, alignment: AlignmentType.CENTER, outlineLevel: 0 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "-",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("技术服务合同补充协议")]
      }),

      new Paragraph({ children: [new TextRun("甲方：山东大禹水务建设集团有限公司")] }),
      new Paragraph({ children: [new TextRun("乙方：烟台大学")] }),
      new Paragraph({ spacing: { after: 240 } }),

      new Paragraph({
        children: [
          new TextRun("甲方与乙方于"),
          new TextRun({ text: "二〇二二年十二月十日", bold: true }),
          new TextRun("签订了《渠道边坡衬砌季节性冻融破坏机理研究技术服务合同》（以下简称\"原合同\"）。经双方友好协商，就原合同相关事宜达成如下补充协议：")
        ]
      }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: "一、原合同延期", bold: true, size: 28 })] }),
      new Paragraph({
        children: [
          new TextRun("1. 原合同执行期限为"),
          new TextRun({ text: "3年", bold: true }),
          new TextRun("，自"),
          new TextRun({ text: "2022年12月1日至2025年11月30日", bold: true }),
          new TextRun("止。")
        ]
      }),
      new Paragraph({
        children: [
          new TextRun("2. 经双方协商一致，原合同有效期延长至"),
          new TextRun({ text: "2027年12月31日", bold: true }),
          new TextRun("。")
        ]
      }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: "二、研究内容调整", bold: true, size: 28 })] }),
      new Paragraph({
        children: [
          new TextRun("因项目研究需要深化，专利申请周期比以前有了大幅度的增长且更加严格，专著的撰写和出版周期也比较长，双方同意对原合同研究内容进行调整，具体如下：")
        ]
      }),
      new Paragraph({ children: [new TextRun("1. 乙方继续完成原合同约定的研究任务：")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("技术报告1套")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("发明专利2件，实用新型专利4件")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("北大核心论文2篇，科技核心论文4篇")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("署名甲方有关技术人员的学术专著1本")] }),
      new Paragraph({
        children: [
          new TextRun("2. 乙方应在"),
          new TextRun({ text: "2027年12月31日前", bold: true }),
          new TextRun("完成全部研究成果并提交甲方验收。")
        ]
      }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: "三、经费支付", bold: true, size: 28 })] }),
      new Paragraph({
        children: [
          new TextRun("1. 原合同总金额为人民币"),
          new TextRun({ text: "400,000.00元", bold: true }),
          new TextRun("（大写：肆拾万元整）。")
        ]
      }),
      new Paragraph({
        children: [
          new TextRun("2. 甲方已向乙方支付原合同总金额的"),
          new TextRun({ text: "50%", bold: true }),
          new TextRun("，即人民币"),
          new TextRun({ text: "200,000.00元", bold: true }),
          new TextRun("（大写：贰拾万元整）。")
        ]
      }),
      new Paragraph({ children: [new TextRun("3. 剩余款项按原合同约定的支付方式执行：")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("30%（120,000.00元）：完成合同中第三条3.1.2、3.1.3、3.1.4约定的成果后支付")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("20%（80,000.00元）：完成技术报告并验收后支付")] }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: "四、其他", bold: true, size: 28 })] }),
      new Paragraph({ children: [new TextRun("1. 本协议是原合同不可分割的组成部分，与原合同具有同等法律效力。")] }),
      new Paragraph({ children: [new TextRun("2. 除本协议明确约定的条款外，原合同的其他条款继续有效。")] }),
      new Paragraph({ children: [new TextRun("3. 本协议自双方签字盖章之日起生效。")] }),
      new Paragraph({
        children: [
          new TextRun("4. 本协议一式"),
          new TextRun({ text: "四份", bold: true }),
          new TextRun("，甲乙双方各执"),
          new TextRun({ text: "两份", bold: true }),
          new TextRun("，具有同等法律效力。")
        ]
      }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun("（以下无正文）")] }),

      new Paragraph({ spacing: { before: 480 } }),
      new Paragraph({ children: [new TextRun("甲方：山东大禹水务建设集团有限公司")] }),
      new Paragraph({ children: [new TextRun("（盖章）")] }),
      new Paragraph({ children: [new TextRun("法定代表人")] }),
      new Paragraph({ children: [new TextRun("（或委托代理人）：")] }),
      new Paragraph({ children: [new TextRun("单位地址：济南市历城区唐冶中路绿地城商办A区一栋")] }),
      new Paragraph({ children: [new TextRun("邮政编码：250102")] }),
      new Paragraph({ children: [new TextRun("电 话：0531-58256015")] }),
      new Paragraph({ children: [new TextRun("传 真：0531-58256016")] }),
      new Paragraph({ children: [new TextRun("签订日期：")] }),

      new Paragraph({ spacing: { before: 480 } }),
      new Paragraph({ children: [new TextRun("乙方：烟台大学")] }),
      new Paragraph({ children: [new TextRun("（盖章）")] }),
      new Paragraph({ children: [new TextRun("法定代表人")] }),
      new Paragraph({ children: [new TextRun("（或委托代理人）：")] }),
      new Paragraph({ children: [new TextRun("单位地址：烟台市莱山区清泉路30号")] }),
      new Paragraph({ children: [new TextRun("邮政编码：264005")] }),
      new Paragraph({ children: [new TextRun("电 话：0535-6902967")] }),
      new Paragraph({ children: [new TextRun("传 真：")] }),
      new Paragraph({ children: [new TextRun("签订日期：")] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\Users\\32891\\.qclaw\\workspace\\技术服务合同补充协议.docx", buffer);
  console.log("Word文档创建成功！");
});
