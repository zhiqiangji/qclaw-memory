# export_survey.py - 导出问卷数据为 Excel
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 数据
rows = [
    ["吉植强",   "硕士+博士", "2003", "2005", "烟台",       "烟台大学",                         "教授",             "副院长",             "13031616187"],
    ["李治国",   "硕士",     "2007", "",     "济南",       "国家电投山东电力工程咨询院有限公司",  "高级工程师",        "主任兼项目经理",      "18678650912"],
    ["汪恩良",   "博士",     "",     "2002", "哈尔滨",     "东北农业大学",                      "教授",             "",                  "15546182172"],
    ["孙雨洋",   "硕士",     "2010", "",     "哈尔滨",     "哈工大建筑设计研究院有限公司",       "副高级工程师",      "",                  "15134566904"],
    ["李孝臣",   "硕士",     "2006", "",     "深圳",       "刚离职",                           "中级工程师",        "",                  "18665826587"],
    ["马元顺",   "硕士",     "2008", "",     "北京",       "中国恩菲工程技术有限公司",           "正高级工程师",      "工管部副部长",        "15001232998"],
    ["李勇",     "硕士",     "2003", "",     "石家庄",     "河北北方绿野建筑设计有限公司",        "正高级工程师",      "总工",               "13483138708"],
    ["张辰熙",   "硕士+博士", "2006", "2023","哈尔滨",    "黑龙江省寒地建筑科学研究院",          "高级工程师",        "所长",               "15004688377"],
    ["洪文江",   "硕士",     "2013", "",     "哈尔滨",     "安博健康科技（黑龙江）有限公司",      "工程师",            "",                  "15124527886"],
    ["柴艳飞",   "硕士",     "2010", "",     "天津",       "中国建筑第六工程局有限公司",          "高级工程师",        "",                  "18920286350"],
    ["刘用海",   "硕士+博士", "2002", "2005", "宁波",      "宁波建乐工程集团有限公司",            "教授级高级工程师",  "总工程师",           "13777045313"],
    ["常俊德",   "硕士",     "2013", "",     "哈尔滨",     "黑龙江省水利科学研究院",              "高级工程师",        "规划中心副主任",      "13796688779"],
    ["齐云静",   "硕士",     "2016", "",     "雄安",       "雄商发展有限公司",                   "工程师",            "设计管理部业务经理",  "16633302301"],
    ["王超",     "硕士",     "2012", "",     "南京",       "南京华泽环保工程有限公司",            "副高级工程师",      "结构专业负责人",      "18994046762"],
    ["徐春华",   "硕士+博士", "2000", "2003", "哈尔滨",    "哈工大建筑设计院",                   "",                 "",                  ""],
    ["刘复孝",   "硕士",     "2004", "",     "青岛",       "青岛腾远设计事务所有限公司",          "高级工程师",        "所长",               "15853208545"],
    ["陈伟华",   "硕士",     "2009", "",     "珠海",       "世荣兆业",                          "工程师",            "部门副职",           "13680300067"],
    ["王立悦",   "硕士",     "2005", "",     "哈尔滨",     "黑龙江省建筑设计研究院",              "正高级工程师",      "",                  "15104608366"],
]

headers = ["序号", "姓名", "培养层次", "硕士入学年份", "博士入学年份",
           "现工作城市", "工作单位", "职称", "职务", "手机号"]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "徐学燕教授历届学生信息"

# 样式
header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", fgColor="2E75B6")
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
data_font = Font(name="微软雅黑", size=10)
data_align = Alignment(horizontal="center", vertical="center")
data_align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 写表头
for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = border

# 写数据
for row_idx, row_data in enumerate(rows, 2):
    # 序号
    ws.cell(row=row_idx, column=1, value=row_idx - 1).alignment = data_align
    ws.cell(row=row_idx, column=1).border = border
    ws.cell(row=row_idx, column=1).font = data_font
    # 其余列
    for col_idx, val in enumerate(row_data, 2):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = data_font
        cell.border = border
        cell.alignment = data_align_left if col_idx in (7, 8) else data_align

# 隔行底色
light_fill = PatternFill("solid", fgColor="DEEAF1")
for row_idx in range(2, len(rows) + 2):
    if row_idx % 2 == 0:
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=col_idx).fill = light_fill

# 列宽
col_widths = {1: 6, 2: 9, 3: 10, 4: 12, 5: 12,
              6: 10, 7: 36, 8: 16, 9: 20, 10: 16}
for col_idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# 行高
ws.row_dimensions[1].height = 32
for r in range(2, len(rows) + 2):
    ws.row_dimensions[r].height = 22

# 冻结首行
ws.freeze_panes = "A2"

# 保护手机号列（列10）
# 隐藏手机号：按需去掉注释
# for row_idx in range(2, len(rows) + 2):
#     ws.cell(row=row_idx, column=10).value = ""

output_path = r"C:\Users\32891\Desktop\徐学燕教授历届学生信息_2026.xlsx"
wb.save(output_path)
print(f"已保存至：{output_path}")
print(f"共 {len(rows)} 条记录")
