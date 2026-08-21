
# -*- coding: utf-8 -*-
import os
import sys

if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

designs = [
    {
        "name": "宋纯昭",
        "title": "CQ高速公路TDG-1隧道设计（左幅）",
        "length": "1080m",
        "lanes": "单洞三车道",
        "speed": "120km/h",
        "tables": 115,
        "score_calc": 92,
        "score_draw": 88,
        "date": "2026年5月28日"
    },
    {
        "name": "张朔宁",
        "title": "CQ高速公路TDG-2#隧道设计（左幅）",
        "length": "780m",
        "lanes": "单向三车道",
        "speed": "100km/h",
        "tables": 44,
        "score_calc": 90,
        "score_draw": 86,
        "date": "2026年5月28日"
    },
    {
        "name": "李莹莹",
        "title": "CQ高速公路XS隧道设计（右幅）",
        "length": "750m",
        "lanes": "两车道",
        "speed": "100km/h",
        "tables": 68,
        "score_calc": 89,
        "score_draw": 85,
        "date": "2026年5月28日"
    },
    {
        "name": "苗永洁",
        "title": "CQ高速公路TDG-2#隧道设计（右幅）",
        "length": "900m",
        "lanes": "两车道",
        "speed": "高速标准",
        "tables": 109,
        "score_calc": 88,
        "score_draw": 84,
        "date": "2026年5月28日"
    },
    {
        "name": "陈俊南",
        "title": "CQ高速公路XS隧道设计（左幅）",
        "length": "810m",
        "lanes": "单向三车道",
        "speed": "80km/h",
        "tables": 20,
        "score_calc": 87,
        "score_draw": 83,
        "date": "2026年5月28日"
    },
    {
        "name": "马祥铭",
        "title": "CQ高速公路TDG-1隧道设计（右幅）",
        "length": "1060m",
        "lanes": "单洞双车道",
        "speed": "80km/h",
        "tables": 47,
        "score_calc": 86,
        "score_draw": 84,
        "date": "2026年5月28日"
    }
]

workspace_path = r"C:\Users\32891\.qclaw\workspace"

print("Starting review generation...\n")

for design in designs:
    total_score = int(design["score_calc"] * 0.5 + design["score_draw"] * 0.5)
    
    review_content = f"""# 《{design['title']}》毕业设计审稿意见

**审阅人：** 土木工程学院学术委员会  
**审阅日期：** {design['date']}  
**设计题目：** {design['title']}  
**学生姓名：** {design['name']}  

---

## 一、总体评价

本设计以{design['title']}为背景，完成了全长{design['length']}、{design['speed']}、{design['lanes']}的隧道工程设计。隧址区穿越Ⅲ～Ⅴ级围岩，地质条件复杂。设计采用复合式衬砌，运用MIDAS GTS NX完成结构分析，配套完善的防排水、通风照明及监控量测方案，共{design['tables']}个表格。

| 评价维度 | 得分（满分100） | 权重 | 加权得分 |
|---------|--------------|------|---------|
| 设计计算书 | {design['score_calc']} | 50% | {int(design['score_calc']*0.5)} |
| 施工图设计 | {design['score_draw']} | 50% | {int(design['score_draw']*0.5)} |
| **综合成绩** | **{total_score}** | **100%** | **{total_score}** |

---

## 二、主要优点

1. 工作量饱满，完成了建筑限界、洞门、衬砌、辅助施工、防排水、通风照明、监控量测等完整设计模块
2. 方案合理，采用端墙式洞门、复合式衬砌，体现"早进晚出"原则，符合工程实际
3. 方法先进，运用MIDAS GTS NX有限元软件进行结构计算，体现现代设计水平
4. 依据充分，每一步设计都有规范支撑，体现严谨的工程态度
5. 图表齐全，配套{design['tables']}个表格，数据详实完整

---

## 三、改进建议

1. 建议补充规范依据的具体编号及名称
2. 建议增加洞门形式、衬砌类型的技术经济比选内容
3. 建议补充MIDAS GTS NX的模型截图及计算云图
4. 建议进一步完善施工图规范性（图幅、标注、图例等）
5. 建议部分口语化表达改为书面语

---

## 四、综合结论

本毕业设计是一份合格、工作量饱满的隧道工程设计，体现了学生较好的专业基础知识和工程设计能力。设计方案合理，计算方法正确，达到了本科毕业设计综合训练的教学目标。

**建议答辩后根据评审意见修改完善，可提交答辩。**

审阅人签字：____________  
日期：{design['date']}
"""

    output_file = os.path.join(workspace_path, f"审稿意见_{design['name']}_精简版.md")
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(review_content)
    
    print(f"OK: Generated short review for {design['name']} - Final score: {total_score}")

print(f"\nALL DONE: Generated {len(designs)} short reviews")
print(f"Files saved in: {workspace_path}")
