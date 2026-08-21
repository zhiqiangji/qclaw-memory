import xlrd, csv
wb = xlrd.open_workbook(r'C:\Users\32891\Desktop\成绩导入模板_已填分数.xls')
ws = wb.sheet_by_index(0)
print('=== 填充后模板前12行 ===')
for r in range(min(12, ws.nrows)):
    row = [ws.cell_value(r, c) for c in range(8)]
    print(row)
print()
print('=== CSV前15行 ===')
with open(r'C:\Users\32891\Desktop\实训成绩_百分制_合并结果.csv', encoding='utf-8-sig') as f:
    for i, line in enumerate(f):
        print(line.strip())
        if i >= 15: break
