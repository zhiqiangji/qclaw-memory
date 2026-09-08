# Obsidian 每日 Git 自动备份 — 2026-08-31

## 执行结果：commit 成功，push 失败

### 已执行命令
1. `git add -A` ✓
2. `git commit -m "auto-backup 2026-08-31" --allow-empty` ✓ → `[master 5f873c2] auto-backup 2026-08-31`
3. `git push origin master` ✗ 失败

### Push 失败原因（已确认非偶发）
```
ERROR: Repository not found.
fatal: Could not read from remote repository.
```

- 本地 SSH 认证通过：`ssh -T git@github.com` 返回 `Hi zhiqiangji! You've successfully authenticated...`
- 但 `origin` 地址为 `git@github.com:jizhiqiang/obsidian-vault.git`
- 认证用户 `zhiqiangji` 与仓库所属 `jizhiqiang` 不一致 → 该用户对该仓库无访问权限，故返回 "Repository not found"
- 重试 push 一次，同样失败，确认非网络抖动

### 修复建议（需用户处理）
认证账户与仓库归属不匹配，任选其一：
1. 将 origin 改为认证用户 `zhiqiangji` 名下的仓库：`git remote set-url origin git@github.com:zhiqiangji/obsidian-vault.git`
2. 在 GitHub 上将 `zhiqiangji` 添加为 `jizhiqiang/obsidian-vault` 的 collaborator
3. 或换用属于 `jizhiqiang` 账户的 SSH key

### 本地状态
- 提交已落盘，本地 master 领先远端 1 个 commit（5f873c2），数据未丢失。
- 下次 cron 触发时若未修复，仍会 commit 成功但 push 失败（因 --allow-empty，无变更时正常跳过）。
