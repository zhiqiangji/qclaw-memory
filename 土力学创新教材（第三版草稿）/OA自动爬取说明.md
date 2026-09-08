# OA办公系统通知自动爬取工具

## 📋 功能说明

每天自动访问烟台大学OA系统，爬取指定通知内容，提取关键信息并保存到本地，自动生成汇总清单。

## 🚀 第一步：配置登录信息

### 获取 Cookies（必须）

因为OA需要登录认证，所以需要你从浏览器获取登录后的`Cookies`：

1. 打开浏览器，登录 [OA系统](http://oa.ytu.edu.cn/)
2. 登录成功后，在新标签页打开你要监控的通知：
   ```
   http://oa.ytu.edu.cn/seeyon/bulData.do?method=bulView&bulId=9131753115002882447&from=list
   ```
3. 按 `F12` 打开开发者工具
4. 切换到 **网络(Network)** 标签
5. 刷新页面
6. 在请求列表中点击第一个请求（就是通知页面本身）
7. 在右侧**请求头(Request Headers)**中找到 `Cookie:` 这一行
8. 复制整段Cookie内容

### 填入脚本

打开 `oa_notification_monitor.py`，找到 `COOKIES` 变量，按格式填入：

```python
COOKIES = {
    # 复制进来的cookie是形如 "key1=value1; key2=value2"，拆分后：
    'JSESSIONID': '这里填JSESSIONID对应的值',
    'SEAWAID': '这里填SEAWAID对应的值',
    # ... 其他key-value依次添加
}
```

> 💡 Cookie过期后需要重新获取，一般能保持几个月

## 🧪 第二步：测试运行

```bash
cd "/Users/jizhiqiang/Desktop/土力学创新教材（第三版草稿）"
python3 oa_notification_monitor.py
```

如果配置正确，会输出：
```
🚀 开始爬取OA通知...
📄 提取成功：通知标题
✅ 已保存到：./oa_notifications/xxx.txt
📋 汇总文件：./oa_notifications/latest_summary.md
✅ 完成！
```

## ⏰ 第三步：设置每天自动运行

### 方法一：使用 macOS launchd（推荐，Mac系统自带）

1. 创建 plist 文件：`~/Library/LaunchAgents/com.ytu.oa.monitor.plist`

内容如下（修改路径）：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.ytu.oa.monitor</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/Users/jizhiqiang/Desktop/土力学创新教材（第三版草稿）/oa_notification_monitor.py</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>/Users/jizhiqiang/Desktop/土力学创新教材（第三版草稿）/oa_notifications/launchd.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/jizhiqiang/Desktop/土力学创新教材（第三版草稿）/oa_notifications/launchd.error.log</string>
</dict>
</plist>
```

这个配置是**每天上午9:00自动运行**，可以修改 `Hour` 和 `Minute` 改时间。

2. 加载定时任务：

```bash
launchctl load ~/Library/LaunchAgents/com.ytu.oa.monitor.plist
```

3. 查看任务状态：

```bash
launchctl list | grep oa
```

### 方法二：使用 crontab（更简单，适合高级用户）

```bash
crontab -e
```

添加一行（每天9点运行）：

```
0 9 * * * /usr/bin/python3 /Users/jizhiqiang/Desktop/土力学创新教材\（第三版草稿）/oa_notification_monitor.py >> /Users/jizhiqiang/Desktop/土力学创新教材\（第三版草稿）/oa_notifications/crontab.log 2>&1
```

保存退出即可。

## 📂 输出文件说明

爬取结果保存在 `./oa_notifications/` 目录：

| 文件 | 说明 |
|------|------|
| `notification_YYYYMMDD_HHMMSS.json` | 结构化数据，包含完整信息 |
| `notification_YYYYMMDD_HHMMSS.txt` | 纯文本格式，方便直接阅读 |
| `latest_summary.md` | 最近20条通知汇总，快速浏览 |
| `launchd.log` / `crontab.log` | 运行日志，排错用 |

## ✨ 梳理功能

脚本会自动提取：

- 通知标题
- 通知完整内容
- 发布时间
- 爬取时间

你可以：

1. 每天打开 `latest_summary.md` 看最新通知汇总
2. 点击打开对应的txt文件看全文
3. 如果通知有更新，会保存新版本，历史版本都保留

## 🔧 常见问题

**Q: 提示"未提取到有效内容"怎么办？**
A: OA页面结构可能变化，或者Cookie过期了。重新获取Cookie，如果还是不行，脚本会保存原始HTML到debug文件，可以发给我帮你调整解析规则。

**Q: Cookie多久过期？**
A: 一般学校OA保持登录几个月没问题，如果长期不访问可能过期，过期了重新获取即可。

**Q: 可以监控多个通知吗？**
A: 可以复制脚本改一下URL和保存目录，每个通知对应一个脚本，在定时任务中添加即可。

**Q: 通知内容变化了能提醒我吗？**
A: 脚本会保存每次爬取的版本，你可以对比历史版本看哪里变了。需要邮件/微信提醒的话，可以加功能。
