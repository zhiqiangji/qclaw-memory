// create_survey_direct.js - 直接通过 HTTP 调用腾讯问卷 MCP
const https = require('https');

const TOKEN = 'wjpt_bc282cfe07fce574588ba9c6ee410a675432ce7064dd8bfde1e7a3adb98a95ce';

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

const body = JSON.stringify({
  jsonrpc: '2.0',
  method: 'tools/call',
  params: {
    name: 'create_survey',
    arguments: { scene: 1, text: text }
  },
  id: 1
});

const options = {
  hostname: 'wj.qq.com',
  path: '/api/v2/mcp',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${TOKEN}`,
    'Content-Length': Buffer.byteLength(body)
  }
};

const req = https.request(options, (res) => {
  let data = '';
  res.on('data', chunk => data += chunk);
  res.on('end', () => {
    console.log('Status:', res.statusCode);
    console.log('Response:', data);
  });
});

req.on('error', (e) => console.error('Error:', e.message));
req.write(body);
req.end();
