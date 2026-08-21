# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

doc = Document()

# 设置页边距
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

# 标题
title = doc.add_paragraph()
run = title.add_run('土木工程学院转专业工作实施方案')
run.font.size = Pt(22)
run.font.bold = True
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.space_after = Pt(20)

# 开头段落
intro = doc.add_paragraph()
intro.add_run('根据《烟台大学本科生转专业管理办法（修订）》（烟大校发〔2026〕19号）规定，结合土木工程学院学科专业建设、师资、实验室条件、在校生情况等，确定土木工程学院各专业接收计划和转专业工作实施方案。')
intro.paragraph_format.first_line_indent = Cm(0.74)
intro.paragraph_format.line_spacing = 1.5

# 一、成立院转专业工作领导小组
h1 = doc.add_paragraph()
h1.add_run('一、成立院转专业工作领导小组')
h1.paragraph_format.first_line_indent = Cm(0.74)
h1.paragraph_format.space_before = Pt(12)
h1.runs[0].font.bold = True

group = doc.add_paragraph()
group.add_run('组  长：书记，院长')
group.paragraph_format.first_line_indent = Cm(0.74)
group.paragraph_format.line_spacing = 1.5

group2 = doc.add_paragraph()
group2.add_run('副组长：教学副院长，副书记')
group2.paragraph_format.first_line_indent = Cm(0.74)

group3 = doc.add_paragraph()
group3.add_run('成  员：专业负责人、系主任、辅导员、骨干教师')
group3.paragraph_format.first_line_indent = Cm(0.74)

group4 = doc.add_paragraph()
group4.add_run('秘  书：教务员')
group4.paragraph_format.first_line_indent = Cm(0.74)

# 二、接收专业简介
h2 = doc.add_paragraph()
h2.add_run('二、接收专业简介')
h2.paragraph_format.first_line_indent = Cm(0.74)
h2.paragraph_format.space_before = Pt(12)
h2.runs[0].font.bold = True

intro2 = doc.add_paragraph()
intro2.add_run('土木工程学院现设有4个本科专业，各专业及方向介绍如下。')
intro2.paragraph_format.first_line_indent = Cm(0.74)

# （一）土木工程专业
p1 = doc.add_paragraph()
p1.add_run('（一）土木工程专业')
p1.paragraph_format.first_line_indent = Cm(0.74)
p1.runs[0].font.bold = True

p1_code = doc.add_paragraph()
p1_code.add_run('专业代码：081001')
p1_code.paragraph_format.first_line_indent = Cm(0.74)

p1_degree = doc.add_paragraph()
p1_degree.add_run('授予学位：工学学士学位')
p1_degree.paragraph_format.first_line_indent = Cm(0.74)

p1_intro = doc.add_paragraph()
p1_intro.add_run('专业介绍：土木工程专业下设建筑工程、道路和桥梁工程、岩土与地下工程三个专业方向，是山东省特色专业、省卓越工程师培养计划项目、省名校工程建设项目重点建设专业、省高水平应用型立项建设核心专业，2014年通过专业评估，2017年通过工程教育认证，2020年获批国家一流本科专业建设点，2024年通过工程教育认证复评。')
p1_intro.paragraph_format.first_line_indent = Cm(0.74)

p1_course = doc.add_paragraph()
p1_course.add_run('核心课程：结构力学、土力学、混凝土结构设计原理、钢结构设计原理、基础工程、土木工程施工等')
p1_course.paragraph_format.first_line_indent = Cm(0.74)

# （二）工程管理专业
p2 = doc.add_paragraph()
p2.add_run('（二）工程管理专业')
p2.paragraph_format.first_line_indent = Cm(0.74)
p2.runs[0].font.bold = True

p2_code = doc.add_paragraph()
p2_code.add_run('专业代码：120103')
p2_code.paragraph_format.first_line_indent = Cm(0.74)

p2_degree = doc.add_paragraph()
p2_degree.add_run('授予学位：管理学学士学位')
p2_degree.paragraph_format.first_line_indent = Cm(0.74)

p2_intro = doc.add_paragraph()
p2_intro.add_run('专业介绍：2002年土木工程专业下设的施工管理教研室发展为工程管理专业，2011年开始招收硕士研究生，2013年成为山东省首批"名校工程"建设专业，2016年获批山东省高水平应用型立项建设专业（群）专业，2022年获批山东省一流专业建设点。工程管理专业下设智能建造管理、工程造价管理两个专业方向。')
p2_intro.paragraph_format.first_line_indent = Cm(0.74)

p2_course = doc.add_paragraph()
p2_course.add_run('核心课程：工程经济学、工程项目管理、工程估价、工程合同管理、运筹学、工程力学等')
p2_course.paragraph_format.first_line_indent = Cm(0.74)

# （三）给排水科学与工程专业
p3 = doc.add_paragraph()
p3.add_run('（三）给排水科学与工程专业')
p3.paragraph_format.first_line_indent = Cm(0.74)
p3.runs[0].font.bold = True

p3_code = doc.add_paragraph()
p3_code.add_run('专业代码：081003')
p3_code.paragraph_format.first_line_indent = Cm(0.74)

p3_degree = doc.add_paragraph()
p3_degree.add_run('授予学位：工学学士学位')
p3_degree.paragraph_format.first_line_indent = Cm(0.74)

p3_intro = doc.add_paragraph()
p3_intro.add_run('专业介绍：给排水科学与工程专业下设建筑给排水、市政给排水两个方向，培养具备给排水科学与工程领域基本理论和专业知识，具有较强工程实践能力和创新意识，能在建筑给排水、市政给排水等领域从事规划、设计、施工、管理等工作的高素质应用型人才。')
p3_intro.paragraph_format.first_line_indent = Cm(0.74)

p3_course = doc.add_paragraph()
p3_course.add_run('核心课程：水力学、水分析化学、水泵与水泵站、给水排水管网系统、建筑给水排水工程、水质工程学等')
p3_course.paragraph_format.first_line_indent = Cm(0.74)

# （四）智慧水利专业
p4 = doc.add_paragraph()
p4.add_run('（四）智慧水利专业')
p4.paragraph_format.first_line_indent = Cm(0.74)
p4.runs[0].font.bold = True

p4_code = doc.add_paragraph()
p4_code.add_run('专业代码：081101T')
p4_code.paragraph_format.first_line_indent = Cm(0.74)

p4_degree = doc.add_paragraph()
p4_degree.add_run('授予学位：工学学士学位')
p4_degree.paragraph_format.first_line_indent = Cm(0.74)

p4_intro = doc.add_paragraph()
p4_intro.add_run('专业介绍：智慧水利专业下设智能感知与数字孪生、智慧建造与运维管理两个方向，培养具备水利工程、信息技术、人工智能等跨学科知识，能在水利工程建设、运维管理等领域从事智慧化设计、施工、管理等工作的高素质复合型人才。')
p4_intro.paragraph_format.first_line_indent = Cm(0.74)

p4_course = doc.add_paragraph()
p4_course.add_run('核心课程：水利工程概论、水力学、水文学、传感器与检测技术、数据结构与算法、智慧水利导论等')
p4_course.paragraph_format.first_line_indent = Cm(0.74)

# 学位点说明
degree_intro = doc.add_paragraph()
degree_intro.add_run('学院于2003年获结构工程专业硕士学位授予权，2011年获批土木工程一级硕士点，2010年获批建筑与土木工程专业（现为土木水利专业学位类别）硕士专业学位点，形成了从本科到研究生比较完善的人才培养体系。')
degree_intro.paragraph_format.first_line_indent = Cm(0.74)

# 三、接收计划
h3 = doc.add_paragraph()
h3.add_run('三、接收计划')
h3.paragraph_format.first_line_indent = Cm(0.74)
h3.paragraph_format.space_before = Pt(12)
h3.runs[0].font.bold = True

plan1 = doc.add_paragraph()
plan1.add_run('各专业分两批接收转入专业的计划总数为该专业同级原有人数的15%（四舍五入取整）。')
plan1.paragraph_format.first_line_indent = Cm(0.74)

plan2 = doc.add_paragraph()
plan2.add_run('第一批：第二学期初，接收该专业同级原有人数的8%。')
plan2.paragraph_format.first_line_indent = Cm(0.74)

plan3 = doc.add_paragraph()
plan3.add_run('第二批：第三学期初，接收该专业同级原有人数的7%。')
plan3.paragraph_format.first_line_indent = Cm(0.74)

# 四、申请条件
h4 = doc.add_paragraph()
h4.add_run('四、申请条件')
h4.paragraph_format.first_line_indent = Cm(0.74)
h4.paragraph_format.space_before = Pt(12)
h4.runs[0].font.bold = True

cond_intro = doc.add_paragraph()
cond_intro.add_run('申请转入土木工程学院的学生，除须满足《烟台大学本科生转专业管理办法（修订）》（烟大校发〔2026〕19号）第四条、第五条、第六条规定的基本条件外，还应满足以下条件：')
cond_intro.paragraph_format.first_line_indent = Cm(0.74)

cond1 = doc.add_paragraph()
cond1.add_run('1. 各专业身体健康状况要求符合最新教育部、原卫生部、中国残疾人联合会颁布的《普通高等学校招生体检工作指导意见》。')
cond1.paragraph_format.first_line_indent = Cm(0.74)

cond2 = doc.add_paragraph()
cond2.add_run('2. 高考综合改革省份的学生，高考选考科目需为"物理"或"化学"科目；非高考综合改革省份的学生，申报专业不能跨文理大类。')
cond2.paragraph_format.first_line_indent = Cm(0.74)

cond3 = doc.add_paragraph()
cond3.add_run('3. 转入土木类专业的学生，应具备基本的物理基础，并对工程实践有兴趣。')
cond3.paragraph_format.first_line_indent = Cm(0.74)

# 五、考核名单确定方式
h5 = doc.add_paragraph()
h5.add_run('五、考核名单确定方式')
h5.paragraph_format.first_line_indent = Cm(0.74)
h5.paragraph_format.space_before = Pt(12)
h5.runs[0].font.bold = True

m1 = doc.add_paragraph()
m1.add_run('1. 资格初审：学院对报名学生提交的材料进行审核，主要核查是否满足《办法》第四条、第五条的基本条件，以及高考选考科目是否符合转入专业要求。')
m1.paragraph_format.first_line_indent = Cm(0.74)

m2 = doc.add_paragraph()
m2.add_run('2. 名单确定：通过资格初审的学生，均获得考核资格，进入录取排序环节。')
m2.paragraph_format.first_line_indent = Cm(0.74)

# 六、考核方式
h6 = doc.add_paragraph()
h6.add_run('六、考核方式')
h6.paragraph_format.first_line_indent = Cm(0.74)
h6.paragraph_format.space_before = Pt(12)
h6.runs[0].font.bold = True

method = doc.add_paragraph()
method.add_run('本学院不另设专门的考核（如笔试、面试等环节），直接以学生第一学期或第一学年所学课程平均学分绩点（GPA）作为考核依据。')
method.paragraph_format.first_line_indent = Cm(0.74)

# 七、录取规则
h7 = doc.add_paragraph()
h7.add_run('七、录取规则')
h7.paragraph_format.first_line_indent = Cm(0.74)
h7.paragraph_format.space_before = Pt(12)
h7.runs[0].font.bold = True

r1 = doc.add_paragraph()
r1.add_run('1. 录取依据：按照学生第一学期或第一学年所学课程平均学分绩点（GPA）从高到低排序。')
r1.paragraph_format.first_line_indent = Cm(0.74)

r2_intro = doc.add_paragraph()
r2_intro.add_run('2. 录取原则：')
r2_intro.paragraph_format.first_line_indent = Cm(0.74)

r2_1 = doc.add_paragraph()
r2_1.add_run('（1）优先录取第一志愿：分专业按第一志愿学生GPA从高到低排序，依次录取，直至该专业接收计划满额。')
r2_1.paragraph_format.first_line_indent = Cm(0.74)

r2_2 = doc.add_paragraph()
r2_2.add_run('（2）第二志愿递补：若有专业第一志愿录取后仍有空缺名额，则从未被第一志愿录取、且第二志愿填报该专业的学生中，按GPA从高到低排序，依次递补录取。')
r2_2.paragraph_format.first_line_indent = Cm(0.74)

r2_3 = doc.add_paragraph()
r2_3.add_run('（3）GPA相同的情况下，依次比较高等数学成绩、大学英语成绩，分数高者优先录取。')
r2_3.paragraph_format.first_line_indent = Cm(0.74)

r3_intro = doc.add_paragraph()
r3_intro.add_run('3. 特殊政策：对入伍退役后复学的学生，按《办法》第十二条执行，GPA达2.0（含）以上即可报名，录取时在原GPA基础上增加0.5个绩点后参与排名。')
r3_intro.paragraph_format.first_line_indent = Cm(0.74)

# 八、工作流程及时间安排
h8 = doc.add_paragraph()
h8.add_run('八、工作流程及时间安排')
h8.paragraph_format.first_line_indent = Cm(0.74)
h8.paragraph_format.space_before = Pt(12)
h8.runs[0].font.bold = True

w1 = doc.add_paragraph()
w1.add_run('按照学校统一部署，具体时间节点如下：')
w1.paragraph_format.first_line_indent = Cm(0.74)

w2 = doc.add_paragraph()
w2.add_run('1. 报名阶段：本院学生根据学校通知，在规定时间内提交转专业申请，每人可填报2个志愿。')
w2.paragraph_format.first_line_indent = Cm(0.74)

w3 = doc.add_paragraph()
w3.add_run('2. 资格审核及公示：学院对申请学生进行报名资格初审，将符合条件的学生名单在学院内公示3个工作日。公示无异议后，将汇总表报送教务处。')
w3.paragraph_format.first_line_indent = Cm(0.74)

w4 = doc.add_paragraph()
w4.add_run('3. 组织考核及录取排序：学院收到其他学院学生的转专业申请后，按照申请学生第一学期或第一学年所学课程的平均学分绩点（GPA）从高到低进行排序，确定拟录取名单。')
w4.paragraph_format.first_line_indent = Cm(0.74)

w5 = doc.add_paragraph()
w5.add_run('4. 录取公示：将拟录取名单报送教务处审核，审核通过后在教务处网站公示3个工作日。')
w5.paragraph_format.first_line_indent = Cm(0.74)

w6 = doc.add_paragraph()
w6.add_run('5. 报到确认：公示无异议后，学生凭转专业通知单到土木工程学院教务办公室办理报到手续。逾期未确认者，视为自动放弃。')
w6.paragraph_format.first_line_indent = Cm(0.74)

# 九、学籍管理与教学安排
h9 = doc.add_paragraph()
h9.add_run('九、学籍管理与教学安排')
h9.paragraph_format.first_line_indent = Cm(0.74)
h9.paragraph_format.space_before = Pt(12)
h9.runs[0].font.bold = True

s1 = doc.add_paragraph()
s1.add_run('1. 学籍异动：学生转专业后，由教务处统一办理学籍异动。学生可根据自身情况申请转入同年级或下一年级学习。')
s1.paragraph_format.first_line_indent = Cm(0.74)

s2 = doc.add_paragraph()
s2.add_run('2. 学分认定与补修：在原专业已修读且考核合格的课程，若与转入专业人才培养方案中的课程相同或相近（由学院认定），可申请课程替代。不符合要求或者未修读的课程，应进行补修。')
s2.paragraph_format.first_line_indent = Cm(0.74)

s3 = doc.add_paragraph()
s3.add_run('3. 学业指导：学院将为每位转入学生安排学业导师，指导其制定个性化补修计划，确保顺利完成学业。')
s3.paragraph_format.first_line_indent = Cm(0.74)

# 十、工作纪律
h10 = doc.add_paragraph()
h10.add_run('十、工作纪律')
h10.paragraph_format.first_line_indent = Cm(0.74)
h10.paragraph_format.space_before = Pt(12)
h10.runs[0].font.bold = True

d1 = doc.add_paragraph()
d1.add_run('学院成立转专业工作领导小组，由院长、党委书记任组长，教学副院长、党委副书记任副组长，成员包括专业负责人、教务员、辅导员等。领导小组全程负责监督，确保工作公平、公正、公开。')
d1.paragraph_format.first_line_indent = Cm(0.74)

d2 = doc.add_paragraph()
d2.add_run('申请学生对提交材料的真实性负责，一经查实有弄虚作假或违纪行为，取消其转专业资格，并按学校相关规定处理。')
d2.paragraph_format.first_line_indent = Cm(0.74)

# 十一、附则
h11 = doc.add_paragraph()
h11.add_run('十一、附则')
h11.paragraph_format.first_line_indent = Cm(0.74)
h11.paragraph_format.space_before = Pt(12)
h11.runs[0].font.bold = True

final = doc.add_paragraph()
final.add_run('本实施方案自发布之日起施行，由土木工程学院负责解释。未尽事宜，按照《烟台大学本科生转专业管理办法（修订）》（烟大校发〔2026〕19号）执行。')
final.paragraph_format.first_line_indent = Cm(0.74)

# 空行和落款
doc.add_paragraph()
doc.add_paragraph()

sign1 = doc.add_paragraph()
sign1.add_run('                            土木工程学院')
sign1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

sign2 = doc.add_paragraph()
sign2.add_run('                                 2026年5月14日')
sign2.alignment = WD_ALIGN_PARAGRAPH.RIGHT

# 保存
doc.save(r'C:\Users\32891\Desktop\12_土木工程学院转专业工作实施方案_修订稿.docx')
print('文档已生成')
