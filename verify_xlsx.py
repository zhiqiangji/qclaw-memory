# verify_xlsx.py - 验证排序
import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\32891\Desktop\徐学燕教授历届学生信息_2026_25人_v2.xlsx')
ws = wb.active
prev = 0
ok = True
for row in ws.iter_rows(min_row=2, values_only=True):
    y = row[3] if row[3] else row[4]
    flag = ''
    if y and str(y).isdigit():
        if int(y) < prev:
            flag = '  <-- 乱序!'
            ok = False
        prev = int(y)
    print(row[0], row[1], '硕%s 博%s%s' % (row[3], row[4], flag))
print('排序OK' if ok else '存在乱序')
