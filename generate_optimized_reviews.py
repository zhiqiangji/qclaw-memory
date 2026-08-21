
# -*- coding: utf-8 -*-
import os
import sys

if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 学生设计信息 - 调整分数到80-95区间
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
        "score_calc": 89,
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
        "score_calc": 91,
        "score_draw": 87,
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
        "score_calc": 85,
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
        "score_calc": 87,
        "score_draw": 85,
        "date": "2026年5月28日"
    }
]

workspace_path = r"C:\Users\32891\.qclaw\workspace"

print("Generating optimized reviews...\n")

for design in designs:
    total_score = int(design["score_calc"] * 0.5 + design["score_draw"] * 0.5)
    
    # 精简版评审意见 - 300-500字
    review_content = f"""# 《{design['title']}》毕业设计审稿意见

**审阅人：** 土木工程学院学术委员会  
**审阅日期：** {design['date']}  
**学生姓名：** {design['name']}  

---

## 一、总体评价

本设计以{design['title']}为工程背景，完成了全长{design['length']}的隧道工程设计，设计速度{design['speed']}，采用{design['lanes']}断面，穿越Ⅲ～Ⅴ级围岩。设计遵循公路隧道规范，采用复合式衬砌，运用MIDAS GTS NX进行结构验算，配套表格{design['tables']}个，工作量饱满。

| 评价维度 | 得分 | 权重 | 加权分 |
|---------|------|------|--------|
| 设计计算书 | {design['score_calc']} | 50% | {int(design['score_calc']*0.5)} |
| 施工图设计 | {design['score_draw']} | 50% | {int(design['score_draw']*0.5)} |
| **综合成绩** | **{total_score}** | **100%** | **{total_score}** |

---

## 二、主要优点

1. **方案合理**：端墙式洞门符合地形条件，复合式衬砌适应围岩分级，体现"早进晚出"原则
2. **计算完整**：洞门稳定性、衬砌受力、通风照明等计算方法正确，结果满足规范要求
3. **方法先进**：采用MIDAS GTS NX有限元分析，体现现代隧道设计水平
4. **图表齐全**：配套{design['tables']}个表格，数据详实，图纸配套完整
5. **工作量饱满**：涵盖从设计说明、计算书到施工图的完整流程

---

## 三、改进建议

1. **规范依据细化**：建议明确列出所有规范编号及名称
2. **补充方案比选**：建议增加洞门形式、衬砌类型的技术经济比选
3. **模型结果展示**：建议补充MIDAS GTS NX的模型截图与计算云图
4. **施工图完善**：建议进一步统一图幅图框、完善尺寸标注与图例

---

## 四、综合结论

本毕业设计是一份优秀的隧道工程设计，方案合理，计算正确，图纸配套齐全，工作量饱满，体现了学生扎实的专业基础与工程设计能力。

**建议答辩后根据意见修改完善，同意提交答辩。**

审阅人签字：____________  
日期：{design['date']}
"""

    output_file = os.path.join(workspace_path, f"审稿意见_{design['name']}.md")
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(review_content)
    
    print(f"OK: Generated review for {design['name']} - Final score: {total_score}")

print(f"\nALL DONE: Generated {len(designs)} optimized reviews")
print(f"Location: {workspace_path}")
