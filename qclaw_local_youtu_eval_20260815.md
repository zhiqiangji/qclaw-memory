# Youtu-FT-V1-35B-A3B 本地模型评测报告

评测日期：2026-08-15
评测环境：AMD Ryzen 7 8845H (8核), 27.8GB RAM, AMD Radeon 780M 核显(4GB VRAM, 不可用)
推理引擎：Ollama (localhost:11434), 纯 CPU+内存
模型来源：QClaw 本地模型库 (package-cdn.qclaw.qq.com)

---

## 一、模型规格

| 项目 | 值 |
|------|-----|
| 全称 | Youtu-FT-V1-35B-A3B（腾讯优图实验室） |
| 底模架构 | qwen35moe（Qwen3.5 MoE，34.7B 参数，A3B≈3B 激活） |
| 量化 | Q4_K_M |
| 磁盘/内存占用 | 21GB / 常驻 22GB |
| 上下文长度 | 262144（256K，超长） |
| 能力 | completion / thinking（推理）/ tools（工具调用） |
| 默认参数 | temperature=1, top_p=0.95, top_k=20, presence_penalty=1.5 |
| License | Apache 2.0 |

结论：这是腾讯优图基于 Qwen3.5-MoE 微调的**推理型大模型**，底子强、带思维链、且支持工具调用。

---

## 二、速度表现（CPU 推理）

| 指标 | 实测 |
|------|------|
| 首层加载 | ~16s |
| 生成速度 | **6~7 tok/s**（多次实测 6.37 / 7.14 / 7.23 / 7.35） |
| 预填充/首 token（TTFT） | 慢且不稳定：think:false 的代码题 125 token 用了 201s（疑似内部 prefill 极重） |
| 思维链长度 | 极长：简单任务也能想满 600~1500 token 才开始答 |

注：机器无独显（AMD 780M 仅 4GB VRAM），完全靠 CPU+内存，GPU 帮不上忙。生成速度约云端模型的 1/30~1/50。

---

## 三、质量评测（四项全部通过，均正确/准确）

### 1. 中文数学推理（水池进/出水管）—— EXCELLENT
> 问：单开进水管6小时注满，单开出水管8小时放空，同时打开几小时注满？

答：净效率 1/6 - 1/8 = 1/24，时间 = 24 小时。推理清晰、带 LaTeX 公式，结果正确。

### 2. 英文代码（第二大元素）—— EXCELLENT
> 问：写 Python 函数返回列表第二大元素，正确处理空列表和重复。

答：用 `set` 去重后排序，`sorted_unique[1]`，空列表/无第二大返回 None。完全正确、规范、含注释。

### 3. 英文常识（摩尔定律）—— 正确
准确描述了摩尔定律（晶体管密度约每两年翻番）及提出者 Gordon Moore。

### 4. 中文知识（区块链，三句话）—— EXCELLENT
准确通俗地讲清：分布式公开账本、密码学链式结构、共识机制（PoW）、不可篡改。概念全部正确。

**质量小结**：中文/英文、推理/代码/知识均达到很强水平，作为免费本地模型质量惊喜。注意四项均是在 `think:false` 或足够 token 预算下测得（详见"测试坑"）。

---

## 四、资源占用与稳定性（关键风险）

- 模型常驻 **22GB** 内存，机器总内存 **27.8GB** → 空闲仅 **600~860MB**。
- **OOM 风险**：第 2 轮 `[knowledge]` 测试（think:false 之前）曾被 **SIGKILL 直接杀掉**，结合内存红线，高度疑似内存耗尽（OOM）。
- 无独显可用，纯内存推理已顶到天花板；若同时开浏览器/其他程序，极易触发 OOM 或系统靠页面文件硬撑（严重降速）。
- 不适合并发或多任务场景；更适合"关掉其他程序、专心跑本地模型"的轻量/离线/隐私场景。

---

## 五、测试踩坑记录（供复现参考）

1. **中文乱码（GBK）**：PowerShell 5.1 默认 GBK 解析 .ps1，脚本里含中文会连 ASCII 一起解析崩。解决办法：中文一律放 UTF-8 文件，用 `[System.IO.File]::ReadAllText(path, UTF8)` 读取；HTTP 请求体用 `[System.Text.Encoding]::UTF8.GetBytes()` 转字节数组 + `Invoke-RestMethod -ContentType 'application/json; charset=utf-8'`。
2. **WebClient 超时 100s**：`System.Net.WebClient` 在本机无可读写 `Timeout` 属性，长推理必超时。改用 `Invoke-RestMethod -TimeoutSec 600`。
3. **HttpClient 不可用**：`System.Net.Http.HttpClient` 未加载，直接 new 报错。改用上面方案。
4. **思维链占满 token 预算**：该模型 thinking 极长，num_predict=600/1500 都曾全被思考吃掉、答案字段为空。解决：① 调大 num_predict（≥1500）；② 或请求加 `"think": false` 直接出答案（推荐，又快又稳）。

---

## 六、结论与建议

### 能否替代云端模型（省钱）？
- **质量**：能。四项评测全过，作为免费本地模型能力很强，中文尤其好。
- **速度**：慢。生成 ~7 tok/s，云端几乎瞬时。日常对话会明显卡顿。
- **内存**：本机 27.8GB 跑 22GB 模型属"顶红线"，有 OOM 风险，不宜边跑边开别的东西。

### 建议
1. **当前 QClaw 对话不要切换成 Youtu**：否则交互降到 ~7 tok/s 且可能 OOM 卡死。
2. **适用场景**：离线/隐私/批量任务、或在不常用机器时跑。
3. **硬件升级**：要想舒服地本地日用，建议内存升到 **32GB+**，或加一块 **≥24GB VRAM** 的独显（目前 4GB 核显完全带不动）。
4. **使用技巧**：不需要深度推理时，开 `think:false`（或降低 num_predict）可显著提速、避免 token 预算被思考吃光。
5. **省钱本质**：能否真省取决于 QClaw 收费模式——若是"按调用量"则本地免费；若是"软件订阅费"则本地仍免调用费但订阅费照付。需向 QClaw 确认收费类型。

---

## 七、复现用文件
- 评测结果：`C:\Users\32891\AppData\Local\Temp\youtu_results.txt`
- 中文 prompt 文件：`C:\Users\32891\AppData\Local\Temp\youtu_prompts.json` / `youtu_prompts2.json`
- 评测脚本：`C:\Users\32891\AppData\Local\Temp\youtu_eval8.ps1`（byte[] UTF-8 + TimeoutSec 600）
