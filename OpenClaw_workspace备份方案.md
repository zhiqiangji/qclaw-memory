# OpenClaw Workspace 备份方案

## 📋 当前需要备份的内容清单

### 核心配置与记忆（必须备份）
```
~/.qclaw/workspace/
├── AGENTS.md              # 工作区规则
├── HEARTBEAT.md           # 心跳检查清单
├── IDENTITY.md            # 我的身份定义
├── MEMORY.md              # 长期核心记忆 ⭐️ 最重要
├── SOUL.md                # 性格与行为准则
├── TOOLS.md               # 本地工具配置笔记
├── USER.md                # 你的信息
└── memory/               # 日常日志
    ├── YYYY-MM-DD.md     # 每日记录
```

### 工作成果（必须备份）
```
├── *.md                   # 各种报告、审稿结果、笔记
├── *.docx                 # 文档文件
├── 土力学创新教材（第三版草稿）/  # 教材编写项目
└── skills/                # 自定义技能
```

### 不需要备份的内容
- `.git/` (git本身是版本控制，备份裸库即可)
- `.openclaw/` (运行时缓存，可重新生成)
- 大型二进制文件如 `openclaw_shouce.pdf`（如果原始文件 elsewhere 可不用重复备份）

---

## 💾 推荐备份方案：三层备份策略

### 第一层：Git 版本控制（增量备份，随时可用）

**当前已经有 `.git` 了，很好！只需要配置正确的 `.gitignore` 并提交即可。**

#### 1. 创建合理的 `.gitignore`
```gitignore
# 运行时缓存
.openclaw/
*.log
*.tmp
.DS_Store

# 大型二进制文件（可选，如果原始文件别处已有）
# *.pdf
# *.docx

# 本地环境配置（如果有）
.env
*.env.local
```

#### 2. 提交所有内容到本地 Git
```bash
cd ~/.qclaw/workspace
git add .
git commit -m "Initial commit: 完整备份工作区 - $(date +%Y-%m-%d)"
```

#### 3. 推送到远程私有仓库（GitHub/GitLab）⭐️
```bash
# 比如推送到 GitHub 私有仓库
git remote add origin https://github.com/你的用户名/openclaw-workspace.git
git push -u origin main
```

**优势**：
- ✅ 增量备份，只存变化内容，省空间
- ✅ 保留完整修改历史，可回滚到任意版本
- ✅ 异地备份，即使本地硬盘挂了也不怕
- ✅ 方便在不同设备之间同步

---

### 第二层：定时全量压缩备份（自动执行，防误删）

创建一个自动备份脚本，每周日凌晨自动打包整个工作区，保存到你的云盘。

#### 创建备份脚本 `scripts/backup_workspace.sh`

```bash
#!/bin/bash
# OpenClaw Workspace 自动备份脚本
# 保存到 ~/.qclaw/workspace/scripts/backup_workspace.sh

WORKSPACE_DIR="$HOME/.qclaw/workspace"
BACKUP_DIR="$HOME/Documents/OpenClaw_Backups"
DATE=$(date +%Y%m%d)
BACKUP_FILE="openclaw_workspace_$DATE.tar.gz"

# 创建备份目录
mkdir -p "$BACKUP_DIR"

# 压缩备份（排除git和缓存）
echo "开始备份 OpenClaw 工作区..."
tar -czf "$BACKUP_DIR/$BACKUP_FILE" \
    -C "$WORKSPACE_DIR" \
    --exclude='.git' \
    --exclude='.openclaw' \
    --exclude='*.DS_Store' \
    .

# 只保留最近 10 个备份，避免占太多空间
cd "$BACKUP_DIR" && ls -t *.tar.gz | tail -n +11 | xargs rm -f

echo "备份完成: $BACKUP_DIR/$BACKUP_FILE"
```

#### 添加执行权限
```bash
chmod +x ~/.qclaw/workspace/scripts/backup_workspace.sh
```

#### 通过 cron 每周日凌晨 2:00 自动执行
```bash
crontab -e
# 添加一行：
0 2 * * 0 ~/.qclaw/workspace/scripts/backup_workspace.sh >> ~/.qclaw/workspace/logs/backup.log 2>&1
```

**优势**：
- ✅ 全量快照，即使误删文件也能找回
- ✅ 自动清理旧备份，不会占太多空间
- ✅ 定时自动执行，不用手动管

---

### 第三层：云盘实时同步（重要文件额外保护）

把整个 `~/.qclaw/workspace` 目录加到你的百度网盘/阿里云盘/OneDrive 同步文件夹中。

**设置建议**：
- 如果使用 iCloud（你在 macOS 上）：可以直接把 `~/.qclaw/workspace` 放在 `~/Documents` 下，iCloud 会自动同步
- 如果使用第三方云盘：设置为实时同步，忽略 `.git` 目录节省空间

---

## 🚀 一步一步马上落实

### 第 1 步：初始化 Git 并提交（现在就做）

我已经看到这里已有 `.git`，直接提交即可：
```bash
cd ~/.qclaw/workspace
git status
git add .
git commit -m "完整备份工作区 - 2026-04-13"
```

### 第 2 步：创建备份脚本

我来帮你创建备份脚本：</think_never_used_51bce0c785ca2f68081bfa7d91973934><seed:tool_call><function name="exec"><parameter name="command" string="true">mkdir -p /Users/jizhiqiang/.qclaw/workspace/scripts /Users/jizhiqiang/.qclaw/workspace/logs