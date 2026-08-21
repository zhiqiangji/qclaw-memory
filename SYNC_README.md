# 记忆同步 - Git 方案

> 仓库地址：https://github.com/zhiqiangji/qclaw-memory
> 分支：main

## 已完成的设置

- **Windows**：HTTPS 方式（git clone/pull/push）
- **Mac**：SSH 方式（git clone/pull/push），公钥已加入 GitHub

## 每日同步流程

在任意一端操作即可（Mac 或 Windows）：

```bash
cd ~/.qclaw/workspace        # Mac
# 或 Windows:
# cd C:\Users\32891\.qclaw\workspace

git pull origin main
git add memory/ MEMORY.md USER.md
git commit -m "sync YYYY-MM-DD"
git push origin main
```

**建议**：每次用完 QClaw 后顺手推一下，避免两台设备内容越拉越远。

## 同步的文件

**核心（每次必推）：**
- `memory/*.md` — 日记
- `MEMORY.md` — 长期记忆
- `USER.md` — 用户档案

**可选（按需推）：**
- `AGENTS.md` / `SOUL.md` / `IDENTITY.md` / `TOOLS.md` / `HEARTBEAT.md`

**不要推：**
- `openclaw.json` — 含机器特定配置（端口/Token/模型/workspace 路径）
- `sessions/` — 对话历史，机器本地
- `skills/` — 技能目录，按需单独同步
- `lossless/lcm.db` — 压缩数据库，可在线重建

## ⚠️ 注意事项

1. **不要同时开两台电脑的 QClaw** ——可能导致 MEMORY.md 冲突
2. **冲突处理**：如遇冲突，保留两台的内容，手动合并后再 push
3. **首次新设备**：clone 后直接用，无需再 init

## Windows PowerShell 注意事项

PowerShell 不支持 `&&` 串联命令，用分号分隔：
```bash
cd C:\Users\32891\.qclaw\workspace; git pull origin main
```

如报错"标记 '&&' 不是有效语句分隔符"，改用：
```bash
cd C:\Users\32891\.qclaw\workspace; git pull; git add memory/ MEMORY.md USER.md; git commit -m "sync"; git push
```
