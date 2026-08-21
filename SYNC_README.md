# 记忆同步 - Git 方案

## 两台电脑共用同一个仓库

### 本机（Windows）→ 初始化
```bash
cd C:\Users\32891\.qclaw\workspace
git init
git add AGENTS.md HEARTBEAT.md IDENTITY.md MEMORY.md SOUL.md USER.md TOOLS.md memory/
git commit -m "Initial commit: memory scaffold"
git remote add origin https://github.com/YOUR_USERNAME/qclaw-memory.git
git push -u origin master
```

### Mac → 克隆并使用
```bash
git clone https://github.com/YOUR_USERNAME/qclaw-memory.git ~/.qclaw/workspace
cd ~/.qclaw/workspace
# 之后每天：
git pull   # 拉取更新
# QClaw 产生的记忆自动写入本地文件
git add memory/ MEMORY.md USER.md
git commit -m "$(date '+%Y-%m-%d')"
git push   # 推送回 GitHub
```

Windows 每次开机后：
```bash
cd C:\Users\32891\.qclaw\workspace
git pull
# 用完 QClaw 后
git add memory/ MEMORY.md USER.md
git commit -m "$(date '+%Y-%m-%d')"
git push
```

### 同步哪些文件
**必须同步（记忆核心）：**
- `MEMORY.md` — 长期记忆
- `USER.md` — 用户档案
- `memory/*.md` — 日记

**可选同步：**
- `AGENTS.md` / `IDENTITY.md` / `SOUL.md` / `TOOLS.md` / `HEARTBEAT.md` — 行为配置，Mac 上可单独维护

**不要同步：**
- `openclaw.json` — 含机器特定配置（端口/Token/模型路径/workspace路径）
- `sessions/` — 对话历史，机器本地
- `skills/` — 技能目录，按需单独同步
- `lossless/lcm.db` — 压缩数据库，在线重建

### ⚠️ 注意事项
1. 两台电脑**不要同时开 QClaw**，可能导致 `MEMORY.md` 冲突
2. 每次使用后**及时 push/pull**，避免积累大量冲突
3. 如果冲突，用文本编辑器手动合并（保留两台的内容）
4. 首次 Mac 同步时，记得把补写的 21 个日记文件先放进去再 push
