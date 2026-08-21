// list_answers.js - 翻页拉取问卷全部回答，输出去重后的学生清单
const https = require('https');

const TOKEN = 'wjpt_bc282cfe07fce574588ba9c6ee410a675432ce7064dd8bfde1e7a3adb98a95ce';
const SURVEY_ID = 27514431;

function call(last_answer_id) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      jsonrpc: '2.0', method: 'tools/call',
      params: { name: 'list_answers', arguments: { survey_id: SURVEY_ID, per_page: 100, last_answer_id } },
      id: 1
    });
    const options = {
      hostname: 'wj.qq.com', path: '/api/v2/mcp', method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${TOKEN}`, 'Content-Length': Buffer.byteLength(body) }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          const text = parsed.result.content[0].text;
          resolve(JSON.parse(text));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body); req.end();
  });
}

function extract(answerObj) {
  // answerObj.answer[].questions[] -> map by question id suffix
  const q = {};
  for (const page of answerObj.answer) {
    for (const question of page.questions) {
      const m = question.id.match(/^q-(\d+)/); // e.g. "q-1-4d55" -> "1"
      if (m) q[m[1]] = question;
    }
  }
  const name = q['1'] ? q['1'].text : '';
  const levelOpts = q['2'] ? q['2'].options.filter(o => o.checked).map(o => o.text) : [];
  const master = q['4'] ? q['4'].text : '';
  const phd = q['6'] ? q['6'].text : '';
  const city = q['7'] ? q['7'].text : '';
  const unit = q['8'] ? q['8'].text : '';
  const title = q['9'] ? q['9'].text : '';
  const position = q['10'] ? q['10'].text : '';
  const phone = q['11'] ? q['11'].text : '';
  return { name, level: levelOpts.join('/'), master, phd, city, unit, title, position, phone };
}

(async () => {
  let all = [];
  let last = 0;
  while (true) {
    const r = await call(last);
    all = all.concat(r.list);
    if (r.list.length < 100) break;
    last = r.last_answer_id;
  }
  console.log(`总数 total=${all.length}（含重复）`);
  // 去重（按姓名）
  const seen = new Map();
  for (const a of all) {
    const e = extract(a);
    if (!e.name) continue;
    if (!seen.has(e.name)) seen.set(e.name, e);
  }
  const unique = [...seen.values()];
  console.log(`去重后 ${unique.length} 人：\n`);
  const order = {};
  unique.forEach((e, i) => {
    const my = e.master && /^[\d]{4}$/.test(e.master) ? parseInt(e.master) : (e.phd && /^[\d]{4}$/.test(e.phd) ? parseInt(e.phd) : 9999);
    order[i] = my;
  });
  const sorted = unique.map((e, i) => ({ e, k: order[i] })).sort((a, b) => a.k - b.k);
  sorted.forEach(({ e }, idx) => {
    console.log(`${String(idx + 1).padStart(2, '0')} | ${e.name} | ${e.level} | 硕${e.master} 博${e.phd} | ${e.city} | ${e.unit} | ${e.title} | ${e.position} | ${e.phone}`);
  });
})();
