import openpyxl, xlrd, xlwt, csv, statistics
from collections import defaultdict

# ======== 1. 读取模板 ========
wb2 = xlrd.open_workbook(r'C:\Users\32891\Desktop\成绩导入模板.xls')
ws2 = wb2.sheet_by_index(0)

# 保留原始说明和表头
note_row = [ws2.cell_value(0, c) for c in range(ws2.ncols)]
header_row = [ws2.cell_value(1, c) for c in range(ws2.ncols)]

template = []
for r in range(2, ws2.nrows):
    raw_sid = ws2.cell_value(r, 0)
    sid = str(int(raw_sid)) if isinstance(raw_sid, float) else str(raw_sid).strip()
    name = str(ws2.cell_value(r, 1)).strip()
    cls = str(ws2.cell_value(r, 2)).strip()
    if sid:
        template.append({'xl_row': r + 1, 'sid': sid, 'name': name, 'cls': cls})
print(f"模板学生: {len(template)}人")

# ======== 2. 读取AI通识（期末成绩） ========
wb3 = openpyxl.load_workbook(r'C:\Users\32891\Desktop\25-26-2人工智能通识考试成绩.xlsx')
ws3 = wb3['学生信息']
ai_scores = {}
for row in ws3.iter_rows(min_row=2, values_only=True):
    sid = str(row[0]).strip()
    pct = row[6]
    if sid and pct and isinstance(pct, (int, float)):
        ai_scores[sid] = {'name': str(row[1]).strip(), 'pct': int(pct), 'status': str(row[4]).strip()}
print(f"AI通识成绩: {len(ai_scores)}条")

# ======== 3. 读取实训（实验成绩） ========
wb1 = openpyxl.load_workbook(r'C:\Users\32891\Desktop\导出实训成绩-20260802-816690146199515136.xlsx')
ws1 = wb1['sheet1']
student_cases = defaultdict(list)
for row in ws1.iter_rows(min_row=2, values_only=True):
    name, sid, case_name, practice, realtime, status, score, grade = row
    if status == '已确认':
        try:
            student_cases[sid].append({'name': name, 'score': float(score), 'grade': grade})
        except (ValueError, TypeError):
            pass
print(f"有实训已确认记录的学生: {len(student_cases)}人，共{sum(len(v) for v in student_cases.values())}条")

# ======== 4. 计算实训综评分 ========
# 规则：取最多5个已确认案例均分，多于5个每多1个+1分，少于5个每少1个-5分，上限100下限60
def calc_train_score(cases):
    """
    cases: list of {'score': float, ...}
    取最多5个最高分算均分，多5个+1/个，少5个-5/个，clamp [60, 100]
    """
    if not cases:
        return None, 0, 0, 0, 0
    # 取最多5个最高分
    top5 = sorted([c['score'] for c in cases], reverse=True)[:5]
    n = len(top5)
    base_avg = statistics.mean(top5)
    bonus = (len(cases) - n) * 1.0   # 多于5个每多1个+1分
    raw = base_avg + bonus
    final = max(60.0, min(100.0, round(raw, 1)))
    return final, len(cases), n, round(base_avg, 1), round(bonus, 1)

train_scores = {}
for sid, cases in student_cases.items():
    score, total_cnt, used_cnt, base_avg, bonus = calc_train_score(cases)
    train_scores[sid] = {
        'name': cases[0]['name'],
        'score': score,
        'total_cnt': total_cnt,
        'used_cnt': used_cnt,
        'base_avg': base_avg,
        'bonus': bonus
    }
print(f"计算出实训综评的学生: {len(train_scores)}人")

# ======== 5. 合并并统计 ========
results = []
for s in template:
    sid = s['sid']
    ai = ai_scores.get(sid)
    tr = train_scores.get(sid)
    results.append({
        'xl_row': s['xl_row'], 'sid': sid, 'name': s['name'], 'cls': s['cls'],
        'ai': ai,
        'tr': tr,
    })

# 统计
ai_list = [s['ai']['pct'] for s in results if s['ai']]
tr_list = [s['tr']['score'] for s in results if s['tr'] and s['tr']['score'] is not None]
print(f"\n=== 成绩统计 ===")
print(f"期末成绩(AI通识): n={len(ai_list)}, min={min(ai_list)}, max={max(ai_list)}, avg={statistics.mean(ai_list):.1f}")
print(f"实验成绩(实训综评): n={len(tr_list)}, min={min(tr_list):.1f}, max={max(tr_list):.1f}, avg={statistics.mean(tr_list):.1f}")

# ======== 6. 用 xlwt 写入 .xls ========
wb_out = xlwt.Workbook(encoding='utf-8')
ws_out = wb_out.add_sheet('学生成绩信息统计表')

# 说明行（原样）
ws_out.write(0, 0, note_row[0])
for c_idx in range(1, ws2.ncols):
    ws_out.write(0, c_idx, '')

# 表头行
for c_idx in range(ws2.ncols):
    ws_out.write(1, c_idx, header_row[c_idx])

# 列映射: 0学号 1姓名 2班级 3平时成绩 4期末成绩(AI通识) 5实验成绩(实训综评) 6总成绩 7未通过原因
for s in results:
    xl_r = s['xl_row'] - 1  # xlrd 0-index
    ai = s['ai']
    tr = s['tr']

    ws_out.write(xl_r, 0, s['sid'])
    ws_out.write(xl_r, 1, s['name'])
    ws_out.write(xl_r, 2, s['cls'])

    # 期末成绩 = AI通识机考百分制
    if ai:
        ws_out.write(xl_r, 4, ai['pct'])

    # 实验成绩 = 实训综评
    if tr and tr['score'] is not None:
        ws_out.write(xl_r, 5, tr['score'])

    # 未通过原因标注
    note = ''
    if not ai and not tr:
        note = 'AI和实训均无数据'
    elif not ai:
        note = '无AI通识成绩(缺考?)'
    elif not tr:
        note = '无实训成绩'

    if note:
        ws_out.write(xl_r, 7, note)

out_xls = r'C:\Users\32891\Desktop\成绩导入模板_已填分数v2.xls'
wb_out.save(out_xls)
print(f"\n模板v2已输出: {out_xls}")

# ======== 7. 详细CSV ========
csv_path = r'C:\Users\32891\Desktop\实训成绩_百分制_合并结果_v2.csv'
with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.DictWriter(f, fieldnames=[
        'xl行号', '学号', '姓名', '班级',
        '期末成绩_AI通识', 'AI状态',
        '实验成绩_实训综评', '实训案例总数', '取用案例数', '基础均分', '加分',
        '建议期末录入', '建议实验录入', '备注'
    ])
    writer.writeheader()
    for s in results:
        ai = s['ai']
        tr = s['tr']
        writer.writerow({
            'xl行号': s['xl_row'],
            '学号': s['sid'],
            '姓名': s['name'],
            '班级': s['cls'],
            '期末成绩_AI通识': ai['pct'] if ai else '',
            'AI状态': ai['status'] if ai else '',
            '实验成绩_实训综评': tr['score'] if tr and tr['score'] is not None else '',
            '实训案例总数': tr['total_cnt'] if tr else '',
            '取用案例数': tr['used_cnt'] if tr else '',
            '基础均分': tr['base_avg'] if tr else '',
            '加分': tr['bonus'] if tr else '',
            '建议期末录入': ai['pct'] if ai else '',
            '建议实验录入': tr['score'] if tr and tr['score'] is not None else '',
            '备注': 'AI和实训均无数据' if (not ai and not tr) else ('无AI通识' if not ai else ('无实训' if not tr else '')),
        })
print(f"CSV v2: {csv_path}")

# ======== 8. 详细打印 ========
print(f"\n=== 全部74人成绩明细 ===")
print(f"{'学号':<15} {'姓名':<8} {'班级':<10} {'期末(AI)':>8} {'实训综评':>8} {'案例数':>5} {'备注'}")
print('-' * 80)
for s in results:
    ai = s['ai']
    tr = s['tr']
    ai_str = str(ai['pct']) if ai else '—'
    tr_score = f"{tr['score']:.1f}" if tr and tr['score'] is not None else '—'
    tr_cnt = f"{tr['total_cnt']}/5" if tr else '—'
    note = ''
    if not ai and not tr: note = '两者均无'
    elif not ai: note = '无AI'
    elif not tr: note = '无实训'
    print(f"{s['sid']:<15} {s['name']:<8} {s['cls']:<10} {ai_str:>8} {tr_score:>8} {tr_cnt:>5} {note}")
