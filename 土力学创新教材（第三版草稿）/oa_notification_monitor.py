#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OA办公系统通知自动爬取与梳理工具
功能：每天自动访问OA，获取**通知列表**中所有新通知，提取关键信息并保存到本地
"""

import requests
from bs4 import BeautifulSoup
import json
import time
from datetime import datetime
import os
import sys
import argparse
import re

class OANotificationMonitor:
    def __init__(self, oa_list_url, save_dir="./oa_notifications"):
        self.oa_list_url = oa_list_url
        self.save_dir = save_dir
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Referer': oa_list_url
        }
        # 创建保存目录
        os.makedirs(save_dir, exist_ok=True)
        # 加载已爬取记录
        self.crawled_file = os.path.join(save_dir, 'crawled_history.json')
        self.crawled_history = self.load_crawled_history()

    def load_crawled_history(self):
        """加载已爬取历史"""
        if os.path.exists(self.crawled_file):
            try:
                with open(self.crawled_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_crawled_history(self):
        """保存已爬取历史"""
        with open(self.crawled_file, 'w', encoding='utf-8') as f:
            json.dump(self.crawled_history, f, ensure_ascii=False, indent=2)

    def set_cookies(self, cookie_dict):
        """设置登录cookies，需要从浏览器获取"""
        self.cookies = cookie_dict

    def fetch_notification_list(self):
        """获取通知列表 - 处理AJAX动态加载"""
        try:
            # OA使用AJAX动态加载通知列表，需要调用API获取
            # 根据浏览器抓包得到正确API路径和参数
            import random
            rnd = random.random()
            api_url = "http://oa.ytu.edu.cn/seeyon/ajax.do?method=ajaxAction&managerName=bulDataManager&rnd={rnd}"
            api_url = api_url.format(rnd=rnd)

            # 正确参数格式来自浏览器抓包
            # managerMethod=findBulDatas 调用方法
            # arguments 是JSON编码的参数数组
            from datetime import datetime
            start_date = "2020-01-01"
            end_date = datetime.now().strftime('%Y-%m-%d')
            arguments = [
                {
                    "pageSize": "20",
                    "pageNo": 1,
                    "listType": "1",
                    "spaceType": "",
                    "spaceId": "",
                    "typeId": "",
                    "search": "firstPage",
                    "condition": "publishDate",
                    "textfield1": start_date,
                    "textfield2": end_date,
                    "myBul": "all"
                }
            ]
            import json
            data = {
                "managerMethod": "findBulDatas",
                "arguments": json.dumps(arguments)
            }
            # 添加正确请求头
            headers = self.headers.copy()
            headers['RequestType'] = 'AJAX'
            headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8'

            response = requests.post(
                api_url,
                headers=headers,
                cookies=self.cookies,
                data=data,
                timeout=30
            )
            response.encoding = 'utf-8'
            print(f"ℹ️  响应状态码：{response.status_code}")
            if response.status_code == 200:
                # 尝试解析JSON
                try:
                    json_data = response.json()
                    print(f"✅ JSON解析成功，获取到数据")
                    # 泛微OA返回格式：result字段包含结果
                    if isinstance(json_data, dict) and 'result' in json_data:
                        import json
                        result_str = json_data['result']
                        # result本身又是JSON字符串，需要再次解析
                        try:
                            actual_data = json.loads(result_str)
                            return self.parse_notification_list_json(actual_data)
                        except:
                            return self.parse_notification_list_json(json_data)
                    else:
                        return self.parse_notification_list_json(json_data)
                except Exception as e:
                    print(f"⚠️  JSON解析失败：{str(e)}，尝试HTML解析")
                    if len(response.text) < 500:
                        print(f"ℹ️  响应内容：{response.text[:500]}...")
                    return self.parse_notification_list_html(response.text)
            else:
                print(f"❌ 请求列表失败，状态码：{response.status_code}")
                print(f"ℹ️  响应内容：{response.text[:500]}...")
                return None
        except Exception as e:
            print(f"❌ 请求列表异常：{str(e)}")
            return None

    def parse_notification_list_json(self, json_data):
        """解析JSON格式的通知列表"""
        notifications = []
        # 不同OA返回结构可能不同，尝试常见结构
        if isinstance(json_data, list):
            items = json_data
        elif isinstance(json_data, dict):
            # 尝试不同key
            for key in ['list', 'data', 'rows', 'result', 'obj']:
                if key in json_data and isinstance(json_data[key], list):
                    items = json_data[key]
                    break
            else:
                items = []
        else:
            items = []

        for item in items:
            if not isinstance(item, dict):
                continue
            # 提取字段 - 尝试多种可能的字段名
            bul_id = str(item.get('bulId', item.get('id', item.get('bulid', ''))))
            if not bul_id:
                continue
            # 标题尝试多个字段名
            title = item.get('subject', item.get('title', item.get('bulTitle', item.get('contentTitle', ''))))
            # 发布时间尝试多个字段名
            publish_time = item.get('publishTime', item.get('createTime', item.get('pubTime', '')))
            # 部门尝试多个字段名
            department = item.get('deptName', item.get('department', item.get('pubDeptName', item.get('departmentName', ''))))

            full_url = f"http://oa.ytu.edu.cn/seeyon/bulData.do?method=bulView&bulId={bul_id}&from=list"

            notifications.append({
                'bul_id': bul_id,
                'title': title,
                'publish_time': publish_time,
                'department': department,
                'url': full_url,
                'crawl_time': None
            })

        print(f"📋 获取到 {len(notifications)} 条通知")
        return notifications

    def parse_notification_list_html(self, html):
        """回退方案：解析HTML格式的通知列表"""
        soup = BeautifulSoup(html, 'html.parser')
        notifications = []

        # 烟台大学OA公告列表解析
        # 寻找通知链接
        links = soup.select('a[href*="bulId"]')
        for link in links:
            href = link.get('href', '')
            # 提取bulId
            match = re.search(r'bulId=([\d]+)', href)
            if not match:
                continue
            bul_id = match.group(1)

            # 提取标题
            title = link.get_text(strip=True)
            if not title:
                title = link.get('title', '').strip()

            # 尝试找发布时间和部门
            # 这里简化处理，后续爬取详情页会再提取
            full_url = f"http://oa.ytu.edu.cn/seeyon/bulData.do?method=bulView&bulId={bul_id}&from=list"

            notifications.append({
                'bul_id': bul_id,
                'title': title,
                'publish_time': '',
                'department': '',
                'url': full_url,
                'crawl_time': None
            })

        # 去重
        seen = set()
        unique_notifications = []
        for n in notifications:
            if n['bul_id'] not in seen:
                seen.add(n['bul_id'])
                unique_notifications.append(n)

        print(f"📋 获取到 {len(unique_notifications)} 条通知")
        return unique_notifications

    def fetch_single_notification(self, url):
        """获取单个通知详细内容"""
        try:
            response = requests.get(
                url,
                headers=self.headers,
                cookies=self.cookies,
                timeout=30
            )
            response.encoding = 'utf-8'
            if response.status_code == 200:
                return self.parse_single_notification(response.text, url)
            else:
                print(f"❌ 请求通知失败，状态码：{response.status_code}")
                return None
        except Exception as e:
            print(f"❌ 请求通知异常：{str(e)}")
            return None

    def parse_single_notification(self, html, url):
        """解析单个通知内容"""
        soup = BeautifulSoup(html, 'html.parser')

        # 提取基本信息
        title = ""
        content = ""
        publisher = ""
        publish_time = ""

        # 获取全部文本，逐行提取
        full_text = soup.get_text(separator='\n', strip=True)
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]

        # 烟台大学OA格式 - 在详情页，标题已经在页面中
        # 先查找是否有缓存链接在当前页面中（有时候直接在这里）
        cache_link = None
        scripts = soup.find_all('script')
        for script in scripts:
            script_text = str(script)
            match = re.search(r'var officeTransSrc = "(.+?)";', script_text)
            if match:
                cache_link = match.group(1)
                break

        # 如果没找到，查找iframe（尝试所有iframes）
        if not cache_link:
            iframes = soup.select('iframe[src*="officeTrans"]')
            for iframe in iframes:
                src = iframe.get('src', '')
                if src:
                    # 处理相对路径
                    if src.startswith('/'):
                        src = 'http://oa.ytu.edu.cn' + src
                    elif not src.startswith('http'):
                        src = 'http://oa.ytu.edu.cn/' + src
                    # 获取iframe内容，里面会有cache链接
                    try:
                        full_resp = requests.get(src, headers=self.headers, cookies=self.cookies, timeout=30)
                        full_resp.encoding = 'utf-8'
                        iframe_soup = BeautifulSoup(full_resp.text, 'html.parser')
                        iframe_scripts = iframe_soup.find_all('script')
                        for iscript in iframe_scripts:
                            match = re.search(r'var officeTransSrc = "(.+?)";', str(iscript))
                            if match:
                                cache_link = match.group(1)
                                break
                        if cache_link:
                            break
                    except Exception as e:
                        print(f"⚠️  获取iframe内容失败：{str(e)}")
                        continue

        # 获取缓存的正文内容
        if cache_link:
            # 处理相对路径
            if cache_link.startswith('/'):
                cache_link = 'http://oa.ytu.edu.cn' + cache_link
            elif not cache_link.startswith('http'):
                cache_link = 'http://oa.ytu.edu.cn/' + cache_link
            try:
                cache_resp = requests.get(cache_link, headers=self.headers, cookies=self.cookies, timeout=30)
                cache_resp.encoding = 'utf-8'
                if cache_resp.status_code == 200:
                    cache_soup = BeautifulSoup(cache_resp.text, 'html.parser')
                    content = cache_soup.get_text(' ', strip=True)
                    # 如果内容还没提取到标题，尝试从缓存页面提取完整标题
                    if (not title or len(title) < 5) and content:
                        cache_text = cache_soup.get_text(separator='\n', strip=True)
                        cache_lines = [line.strip() for line in cache_text.split('\n') if line.strip()]
                        for line in cache_lines:
                            if len(line) > 10:
                                title = line
                                break
                else:
                    print(f"⚠️  缓存链接返回{cache_resp.status_code}，使用全文提取")
                    content = full_text
            except Exception as e:
                print(f"⚠️  获取缓存内容失败：{str(e)}，使用全文提取")
                content = full_text

        # 如果还是没获取到内容，直接使用页面全文
        if (not content or len(content) < 10) and full_text:
            content = full_text

        # 从全文提取标题、发布部门、发布时间
        if not title or len(title) < 5:
            found_title = False
            for line in lines:
                if not found_title and len(line) > 5:
                    # 跳过导航和脚本相关的行
                    if '公告' not in line or len(line) > 8:
                        title = line
                        found_title = True
                        break

        # 提取发布信息 - 搜索全文
        all_text = '\n'.join(lines)
        dept_match = re.search(r'发布部门\s*[:：]\s*([^\n]+)', all_text)
        if dept_match:
            publisher = dept_match.group(1).strip()
        time_match = re.search(r'发布时间\s*[:：]\s*([^\n]+)', all_text)
        if time_match:
            publish_time = time_match.group(1).strip()

        # 如果还是没内容，使用全文
        if not content or len(content) < 10:
            content = full_text

        # 清理内容中重复的头部信息
        # 去掉重复的标题和发布信息
        if title:
            content = content.replace(title, '', 1)
        content_clean = []
        for line in content.split('\n'):
            line = line.strip()
            if line and not any(skip in line for skip in ['发布部门', '发布时间', '公告版块', '阅读量', '发布范围', '打印', '收藏', '查看原文档']):
                content_clean.append(line)
        content = '\n'.join(content_clean).strip()

        return {
            'title': title.strip() if title else '无标题',
            'content': content,
            'publisher': publisher.strip() if publisher else '',
            'publish_time': publish_time.strip() if publish_time else '',
            'crawl_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'url': url
        }

    def save_single_notification(self, data):
        """保存单个通知"""
        bul_id = data.get('bul_id', datetime.now().strftime('%Y%m%d%H%M%S'))
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        # 保存JSON（结构化数据）
        json_file = os.path.join(self.save_dir, f"notification_{bul_id}_{timestamp}.json")
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # 保存纯文本（方便阅读）
        txt_file = os.path.join(self.save_dir, f"notification_{bul_id}_{timestamp}.txt")
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write(f"标题：{data['title']}\n")
            f.write(f"发布部门：{data['publisher']}\n")
            f.write(f"发布时间：{data['publish_time']}\n")
            f.write(f"爬取时间：{data['crawl_time']}\n")
            f.write(f"来源：{data['url']}\n")
            f.write("\n" + "="*50 + "\n\n")
            f.write(data['content'])

        print(f"✅ 已保存：{data['title']} → {txt_file}")

        # 更新历史
        self.crawled_history[bul_id] = {
            'title': data['title'],
            'last_crawl_time': data['crawl_time']
        }
        self.save_crawled_history()

        return txt_file

    def run(self):
        """运行一次完整爬取：获取列表 → 检查新通知 → 爬取新通知 → 保存"""
        print(f"🚀 开始爬取OA通知列表：{self.oa_list_url}")
        print(f"⏰ 时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        notifications = self.fetch_notification_list()
        if notifications is None:
            print("❌ 获取通知列表失败，请检查网络连接和登录状态")
            return False

        new_count = 0
        for notification in notifications:
            bul_id = notification['bul_id']
            if bul_id not in self.crawled_history:
                print(f"🔍 发现新通知：{notification['title']}")
                data = self.fetch_single_notification(notification['url'])
                if data:
                    data['bul_id'] = bul_id
                    data['department'] = notification['department']
                    self.save_single_notification(data)
                    new_count += 1
                    time.sleep(2)  # 礼貌延时
            else:
                print(f"✓ 已爬取过：{notification['title']}")

        print(f"\n🏁 爬取完成！发现 {new_count} 条新通知")
        self.generate_summary()
        return True

    def generate_summary(self):
        """生成所有通知汇总"""
        all_files = sorted([
            f for f in os.listdir(self.save_dir)
            if f.endswith('.json') and not f in ['crawled_history.json']
        ], reverse=True)

        summary_file = os.path.join(self.save_dir, "latest_summary.md")

        summary = "# OA通知汇总\n\n"
        summary += f"最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        summary += f"共爬取 {len(all_files)} 条通知\n\n"
        summary += "## 最近通知列表\n\n"

        for i, filename in enumerate(all_files[:30], 1):
            filepath = os.path.join(self.save_dir, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                title = data.get('title', '无标题')
                publish_time = data.get('publish_time', '')
                crawl_time = data.get('crawl_time', '')
                department = data.get('department', '')

                summary += f"{i}. **{title}**  \n"
                summary += f"   🏢 {department} | 📅 {publish_time} | ⏰ {crawl_time}  \n\n"
            except:
                pass

        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(summary)

        print(f"📋 汇总文件：{summary_file}")


def main():
    parser = argparse.ArgumentParser(description='OA通知自动爬取工具 - 爬取通知列表中所有新通知')
    parser.add_argument('--url',
        default='http://oa.ytu.edu.cn/seeyon/bulData.do?method=bulIndex',
        help='OA通知列表URL')
    parser.add_argument('--save-dir', default='./oa_notifications',
        help='保存目录')
    args = parser.parse_args()

    # 创建监控器
    monitor = OANotificationMonitor(args.url, args.save_dir)

    # ===== 重要：需要从浏览器获取登录后的cookies =====
    # 在浏览器登录OA后，按F12打开开发者工具，复制请求头中的cookie
    # 格式示例：
    # COOKIES = {
    #     'JSESSIONID': 'xxxxxxxxxxxxxxx',
    #     'SEAWAID': 'yyyyyyyyyyyyyyy',
    # }
    COOKIES = {
        'hostname': '202.194.116.76:80',
        'JSESSIONID': 'F54640428A45E37D6A602D8838416FCE',
        'login_locale': 'zh_CN',
        'avatarImageUrl': '-3476203101945380435',
        'loginPageURL': '',
    }

    if not COOKIES or not 'JSESSIONID' in COOKIES:
        print("""
⚠️  需要配置登录Cookies才能使用！

配置方法：
1. 打开浏览器，登录烟台大学OA系统 http://oa.ytu.edu.cn/
2. 登录后，打开通知公告列表页面
3. 按F12 → 网络(Network) → 刷新页面 → 点击第一个请求 → 复制请求头中的 Cookie
4. 打开本脚本，找到 COOKIES 变量，按键值对格式填入
5. 保存后再次运行即可
""")
        sys.exit(1)

    monitor.set_cookies(COOKIES)
    monitor.run()


if __name__ == '__main__':
    main()
