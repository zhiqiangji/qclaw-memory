// create_survey.js
const fs = require('fs');
const { execSync } = require('child_process');

const text = `徐学燕教授八十华诞庆贺——历届学生信息收集

为了筹备徐学燕教授八十华诞庆祝活动，现收集历届学生的基本信息，请各位同学填写。以下培养层次问题请逐项回答，不涉及的直接跳过即可。

1. 您的姓名[单行文本题][必答]

2. 请选择您的培养层次[多选题][必答]（涉及哪项填哪项，不涉及的跳过）
是硕士
是博士

3. 以下硕士相关信息，请读了硕士的同学填写；未读硕士请直接跳过。[段落说明]

4. 您的硕士入学年份[单行文本题]（请填写年份，如：2005）

5. 以下博士相关信息，请读了博士的同学填写；未读博士请直接跳过。[段落说明]

6. 您的博士入学年份[单行文本题]（请填写年份，如：2008）

7. 您现工作所在城市[单行文本题]（请填写城市名，如：哈尔滨）

8. 您现工作单位[单行文本题]（请填写单位全称）

9. 您的职称[单行文本题]（如：教授、副教授、讲师、工程师等，不方便填写可填不便透露）

10. 您的职务[单行文本题]（如：系副主任、项目经理等，不方便填写可填无或不便透露）

11. 您的手机号[单行文本题]（仅用于活动联络，绝不公开，联系人会妥善保管）

感谢您的填写！祝老师八十华诞快乐！[段落说明]`;

const payload = JSON.stringify({
  scene: 1,
  text: text
});

// Write payload to temp file
const tmpFile = 'C:\\Users\\32891\\.qclaw\\workspace\\survey_payload.json';
fs.writeFileSync(tmpFile, payload, 'utf8');

const result = execSync(
  `mcporter call tencent-survey.create_survey --args-file "${tmpFile}"`,
  { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 }
);

console.log(result);
