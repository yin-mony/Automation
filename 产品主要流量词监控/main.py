"""
产品主要流量词监控

通过易得客浏览器打开 TikTok 店铺 → Amazon 商品页 → 西柚「查看所有流量词」导出 Excel，
筛选自然排名大于 10 的关键词并推送企业微信预警。
"""

import math
import time
import threading
from DrissionPage import ChromiumPage,Chromium
import socket
import psutil
from YidekeLogin import Specification
import re
import os
import sys
import cv2
import pandas as pd
import requests
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


class Comment:
    """易得客 + 西柚流量词采集、Excel 预警与企业微信通知。"""

    def __init__(self, config):
        """从 config 读取账号、店铺 IP/端口、ASIN 列表与下载目录。"""
        self.username = config["username"]
        self.password = config["password"]
        self.ip = config["ip"]
        self.port = config["port"]
        self.experts = config["asin"]
        self.file_path = config["file_path"]
        self.number = config["number"]
        self.wechat_webhook = config.get(
            "wechat_webhook",
            "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=c4e9f21f-c365-496f-8c05-c56dd358926e",
        )

    def stop_program(self):
        """强制结束程序"""
        import psutil
        # 进程名称
        process_name = "chrome.exe"

        # 遍历所有进程
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # 检查进程名称是否匹配
                if proc.info['name'] == process_name:
                    # 终止进程
                    proc.kill()
                    print(f"已终止进程: {process_name} (PID: {proc.info['pid']})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        os._exit(0)  # 立即终止程序

    def kill_edecker(self, exclude_pid):
        """结束除指定 PID 外的所有 edecker 进程（避免多实例冲突）。"""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                pid = proc.info['pid']
                name = proc.info['name']
                if name and name.lower() == 'edecker.exe':
                    if pid != exclude_pid:
                        proc.kill()
            except:
                pass

    def kill_edecker_on_port(self, port):
        """启动店铺浏览器前，结束占用该调试端口的 edecker 进程"""
        flag = f'--remote-debugging-port={port}'
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                name = proc.info['name']
                if name and name.lower() == 'edecker.exe':
                    cmdline = proc.info.get('cmdline') or []
                    if any(flag in str(arg) for arg in cmdline):
                        proc.kill()
            except:
                pass

    def wait_for_port(self, port, timeout=60):
        """等待调试端口就绪"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(('127.0.0.1', port), timeout=2):
                    return
            except OSError:
                time.sleep(1)
        raise RuntimeError(f'等待 127.0.0.1:{port} 超时 ({timeout}s)')

    def wait_for_seller_page(self, page, timeout=90):
        """等待页面进入 TikTok 卖家后台"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            url = (page.url or '').lower()
            if 'seller' in url or 'tiktok' in url:
                return
            time.sleep(2)
        raise RuntimeError(f'店铺浏览器未进入 TikTok 后台，当前 URL: {page.url}')

    def snapshot_download_files(self, folder):
        """记录下载目录中文件的修改时间和大小，用于识别本次导出的新文件。"""
        folder = Path(folder)
        snapshot = {}
        if not folder.exists():
            return snapshot
        for file_path in folder.iterdir():
            if not file_path.is_file():
                continue
            try:
                stat = file_path.stat()
            except OSError:
                continue
            snapshot[file_path.name] = (stat.st_mtime, stat.st_size)
        return snapshot

    def wait_for_export_file(self, folder, before_files, timeout=300, stable_seconds=2):
        """等待页面异步导出的 Excel/CSV 文件出现在下载目录并写入稳定。"""
        folder = Path(folder)
        deadline = time.time() + timeout
        target_suffixes = {'.xlsx', '.xls', '.csv'}
        temp_suffixes = {'.crdownload', '.tmp', '.download', '.part'}
        seen_sizes = {}

        while time.time() < deadline:
            candidates = []
            has_temp_file = False

            for file_path in folder.iterdir():
                if not file_path.is_file():
                    continue

                suffix = file_path.suffix.lower()
                if suffix in temp_suffixes or file_path.name.endswith('.crdownload'):
                    try:
                        stat = file_path.stat()
                    except OSError:
                        continue
                    old = before_files.get(file_path.name)
                    if old is None or stat.st_mtime > old[0] or stat.st_size != old[1]:
                        has_temp_file = True
                    continue
                if suffix not in target_suffixes:
                    continue

                try:
                    stat = file_path.stat()
                except OSError:
                    continue

                old = before_files.get(file_path.name)
                is_new_or_changed = old is None or stat.st_mtime > old[0] or stat.st_size != old[1]
                if is_new_or_changed:
                    candidates.append((stat.st_mtime, file_path, stat.st_size))

            if candidates:
                candidates.sort(key=lambda item: item[0], reverse=True)
                _, newest_file, newest_size = candidates[0]
                last_size, stable_since = seen_sizes.get(newest_file.name, (None, time.time()))

                if newest_size != last_size:
                    seen_sizes[newest_file.name] = (newest_size, time.time())
                elif time.time() - stable_since >= stable_seconds and not has_temp_file:
                    return str(newest_file)

            time.sleep(1)

        return None

    def visit_shop(self, ip, port=9222):
        """
        点击指定店铺访问
        :param ip: 店铺IP
        :param port: 易得客管理浏览器端口
        """
        browser = Chromium(port)
        tab = browser.latest_tab
        tab.ele(f'x://div[text()="{ip}"]//following-sibling::button').click()
        time.sleep(3)

    def run_edecker_automation(self, ips,port=9222):
        """
        全部启动店铺
        :param port:
        :return:
        """
        browser = Chromium(port)
        tab = browser.latest_tab
        # buttons = tab.eles("t:button@@text()=访问")
        # for btn in buttons:
        #     btn.click()
        #     time.sleep(3)
        time.sleep(2)
        for ip in ips:
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="美国"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30
            ).click()
            time.sleep(3)
        self.kill_edecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def start_edecker(self, ip: str, port: int):
        """按店铺 IP 匹配 eDecker profile 目录，以指定调试端口启动浏览器。"""
        import subprocess
        from pathlib import Path

        base = Path.home() / "AppData/Local/eDecker6"
        exe_path = base / "Application/edecker.exe"
        profiles_path = base / "Profiles"

        print("EXE:", exe_path, exe_path.exists())
        print("Profiles dir exists:", profiles_path.exists())

        if not exe_path.exists():
            raise FileNotFoundError(f"找不到 exe: {exe_path}")

        if not profiles_path.exists():
            raise FileNotFoundError(f"找不到 profiles 目录: {profiles_path}")

        ip_dot = ip
        ip_underline = ip.replace('.', '_')

        all_profiles = list(profiles_path.iterdir())
        print("所有 profile:")
        for p in all_profiles:
            print(" -", p.name)

        # 目录名含 IP（点号或下划线形式）即视为该店铺 profile
        candidates = [
            p for p in all_profiles
            if p.is_dir() and (ip_dot in p.name or ip_underline in p.name)
        ]

        if not candidates:
            raise Exception(f"未找到 IP={ip} 的 profile")

        latest = max(candidates, key=lambda p: p.stat().st_mtime)

        print("使用 profile:", latest)

        cmd = [
            str(exe_path),
            f'--user-data-dir={latest}',
            '--no-sandbox',
            f'--remote-debugging-port={port}'  # DrissionPage 接管用
        ]
        print("启动命令:")
        print(" ".join(cmd))

        try:
            subprocess.Popen(cmd, cwd=str(base))
            print("启动成功（已发起进程）")
        except Exception as e:
            print("启动失败:", e)
            raise

    def find_seller_tab(self, port, timeout=90):
        """在指定调试端口的浏览器中查找已进入 TikTok 卖家后台的标签页。"""
        browser = Chromium(port)
        deadline = time.time() + timeout

        while time.time() < deadline:
            for tab_id in browser.tab_ids:
                tab = browser.get_tab(tab_id)
                url = (tab.url or '').lower()
                print("检测 tab URL:", url)

                if url.startswith("chrome-extension://"):
                    continue

                if any(k in url for k in ("seller", "tiktok", "tiktokglobalshop")):
                    return tab

            time.sleep(1)

        raise RuntimeError(f"未找到 TikTok seller 后台标签页，端口={port}")

    # 自动化流程：按 ASIN 搜索 Amazon → 西柚导出流量词
    def main(self, page):
        """在已接管的易得客浏览器中，逐个 ASIN 下载西柚关键词反查表。"""
        # 新建标签页进入
        # page.new_tab("https://www.amazon.com")
        browser = page.browser
        for asin in self.experts:
            print(f"\n正在搜索商品: {asin}")
            # 每个 ASIN 都重新打开 Amazon 搜索页，避免上一轮流量词详情页影响下一轮
            page = browser.new_tab("https://www.amazon.com")
            time.sleep(3)
            page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()
            page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=30).input(f'{asin}\n', clear=True)
            time.sleep(3)
            page.ele(f'x://div[@data-asin="{asin}"]//a', timeout=30).click()
            time.sleep(3)
            # 进入商品详情页（可能在当前标签或新标签，按 URL 找含 ASIN 的页）
            for tab_id in browser.tab_ids:
                tab = browser.get_tab(tab_id)
                url = tab.url or ''
                if asin in url or '/dp/' in url:
                    page = tab
                    break
            else:
                page = browser.latest_tab
            time.sleep(3)
            print("当前商品详情页:", page.url)
            # 记录当前商品详情页
            old_tab = page
            # 等待西柚插件加载，滑动到【查看所有流量词】，等待2秒后点击
            traffic_btn = None
            for _ in range(30):
                traffic_btn = page.ele('x://div[text()="查看所有流量词"]', timeout=3)
                if traffic_btn:
                    print("已定位到【查看所有流量词】")
                    break
                print("等待西柚插件加载...")
                time.sleep(2)
            else:
                print(f"未找到【查看所有流量词】，当前页面：{page.url}")
                continue
            traffic_btn.scroll.to_see()
            time.sleep(2)
            traffic_btn.click()
            print("已点击【查看所有流量词】")
            time.sleep(3)
            # 点击【查看所有流量词】后获取新打开的标签页
            new_tab = page.browser.latest_tab
            if new_tab != old_tab:
                page = new_tab
            time.sleep(3)
            print("当前流量词详情页:", page.url)
            # 等待西柚扫码登录完成
            while True:
                print("未检测到登录弹窗，等待西柚页面加载...")
                time.sleep(3)

                login = page.ele('x://div[contains(@class,"login_modal")]', timeout=2)
                if login:
                    print("当前西柚账号未登录，请扫码登录，等待中...")
                    time.sleep(3)
                    continue

                # download_btn = page.ele('x://div[@class="operate"]//span[@class="download"]', timeout=3)
                else:
                    print("当前西柚账号已登录，流量词页面已加载")
                    break

                

            # 开始下载当前 ASIN 的关键词列表
            download_dir = Path(self.file_path)
            download_dir.mkdir(parents=True, exist_ok=True)
            page.browser.set.download_path(download_dir)
            page.set.download_path(download_dir)
            before_files = self.snapshot_download_files(download_dir)

            download = page.ele('x://div[@class="operate"]//span[@class="download"]', timeout=30)
            if not download:
                print(f"当前商品 {asin} 未定位到下载按钮，当前页面：{page.url}")
                continue
                
            print('等待判断是否有广告弹窗，请手动关闭....')
            time.sleep(3)

            download.click()
            print(f"当前商品 {asin} 已点击下载，等待页面导出文件...")
            file_path = self.wait_for_export_file(download_dir, before_files, timeout=300)
            if file_path:
                print(f"当前商品 {asin} 下载完成，文件路径：{file_path}")
            else:
                print(f"当前商品 {asin} 下载失败，300 秒内未在 {download_dir} 检测到新的 Excel/CSV 文件")
                continue
            # 下载完成后关闭流量词详情页,当前商品详情页
            page.close()
            time.sleep(1)
            old_tab.close()
            time.sleep(2)
        print("所有 ASIN 商品全部下载完成，准备关闭当前易得可浏览器")
        browser.quit(timeout=5, force=True)

    # 处理下载文件：筛选自然排名 > 10 并汇总保存
    def excel(self, path=None):
        """读取下载目录中的西柚导出文件，提取自然排名大于 10 的关键词。

        跳过已生成的预警汇总文件；从文件名解析 ASIN；输出完整预警 xlsx。
        """
        path = Path(path or self.file_path)
        if not path.exists():
            print(f"路径不存在: {path}")
            return []
        if path.is_dir():
            files = []
            for suffix in ("*.xlsx", "*.xls", "*.csv"):
                files.extend(path.glob(suffix))
        else:
            files = [path]
        if not files:
            print(f"当前路径下未找到下载文件: {path}")
            return []
        data = []
        for file_path in files:
            if not file_path.exists():
                print(f"文件不存在: {file_path}")
                continue
            if "自然排名提取结果" in file_path.stem or "自然排名预警完整数据" in file_path.stem:
                continue
            print(f"正在处理下载文件: {file_path}")
            # 从文件名提取 ASIN（B + 9 位字母数字）
            asin_match = re.search(r"(?<![A-Z0-9])B[A-Z0-9]{9}(?![A-Z0-9])", file_path.stem.upper())
            asin = asin_match.group() if asin_match else ""
            if file_path.suffix.lower() == ".csv":
                df = pd.read_csv(file_path)
            else:
                try:
                    df = pd.read_excel(file_path, sheet_name="关键词反查结果")
                except Exception:
                    df = pd.read_excel(file_path)

            columns = list(df.columns)
            keyword_index = None
            nature_index = None
            sp_index = None
            time_index = None
            # 按列名定位：关键词、自然排名、SP 广告排名（右边界）、抓取时间
            for index, col in enumerate(columns):
                col = str(col).strip()
                if keyword_index is None and "关键词" in col:
                    keyword_index = index
                if nature_index is None and col == "自然排名":
                    nature_index = index
                if sp_index is None and "SP广告排名" in col:
                    sp_index = index
            if keyword_index is None:
                print(f"未找到关键词列: {file_path}")
                continue
            if nature_index is None:
                print(f"未找到自然排名列: {file_path}")
                continue
            end_index = sp_index if sp_index is not None else len(columns)
            for index in range(nature_index + 1, end_index):
                col = str(columns[index]).strip()
                if "抓取时间" in col:
                    time_index = index
                    break
            if time_index is None:
                print(f"未找到自然排名对应的抓取时间列: {file_path}")
                continue
            result = df.iloc[:, [keyword_index, nature_index, time_index]].copy()
            result.columns = ["关键词", "自然排名", "抓取时间"]
            result = result.dropna(subset=["关键词", "自然排名"])
            result["自然排名"] = pd.to_numeric(result["自然排名"], errors="coerce")
            result = result.dropna(subset=["自然排名"])
            result["自然排名"] = result["自然排名"].astype(int)
            result = result[result["自然排名"] > 10]
            result["asin"] = asin
            data.extend(result[["asin", "关键词", "自然排名", "抓取时间"]].to_dict("records"))
        print(data)
        if not data:
            print("没有自然排名大于10的数据")
            return []

        save_dir = path if path.is_dir() else path.parent
        save_path = save_dir / f"自然排名预警完整数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        pd.DataFrame(data).to_excel(save_path, index=False)
        self.warning_file_path = str(save_path)
        print(f"完整预警数据已保存: {save_path}")
        return data

    def message_send(self, data=None):
        """按 ASIN 分组推送企业微信文本预警（每 ASIN 最多展示前 10 条）。"""
        if data is None:
            data = self.excel()
        if not data:
            print("没有需要发送的企业微信预警数据")
            return

        webhook = self.wechat_webhook
        session = requests.Session()
        session.trust_env = False  # 避免系统代理干扰 webhook 请求
        asin_map = {}
        for item in data:
            asin = item.get("asin", "") or "未识别ASIN"
            asin_map.setdefault(asin, []).append(item)

        for asin, items in asin_map.items():
            content = (
                f"产品关键词监控预警\n"
                f"ASIN：{asin}\n"
                f"自然排名大于10的总计：{len(items)} 条\n"
                f"仅展示前10条，完整数据已保存到本地：{getattr(self, 'warning_file_path', '')}\n"
            )

            for index, item in enumerate(items[:10], 1):
                content += (
                    f"\n{index}. 关键词：{item.get('关键词', '')}\n"
                    f"自然排名：{item.get('自然排名', '')}\n"
                    f"抓取时间：{item.get('抓取时间', '')}\n"
                )

            payload = {
                "msgtype": "text",
                "text": {
                    "content": content,
                    "mentioned_mobile_list": [f"{self.number}"]
                }
            }
            for attempt in range(1, 4):
                try:
                    res = session.post(
                        webhook,
                        json=payload,
                        headers={"Connection": "close"},
                        timeout=(5, 20)
                    )
                    result = res.json()
                    print(f"{asin} 企业微信发送结果: {result}")
                    if result.get("errcode") == 0:
                        break
                    if result.get("errcode") == 45009:
                        # 频率限制，递增等待后重试
                        time.sleep(10 * attempt)
                        continue
                    break
                except requests.exceptions.RequestException as e:
                    print(f"{asin} 企业微信第 {attempt} 次发送失败: {e}")
                    if attempt < 3:
                        time.sleep(3 * attempt)
            time.sleep(3)







    # 启动：登录易得客 → 批量访问店铺 → 逐店铺下载 → 汇总预警 → 发微信
    def run(self):
        """完整流程：易得客登录、启动店铺浏览器、采集、excel 预警、message_send。"""
        sp = Specification(self.username, self.password)  # 易得客登录
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        # self.visit_shop(self.ip, self.port)
        self.run_edecker_automation(self.ip)  # 访问全部店铺
        time.sleep(4)
        for index, ip in enumerate(self.ip):
            self.kill_edecker_on_port(self.port[index])  # 启动前清理占用端口的 edecker
            time.sleep(1)
            self.start_edecker(self.ip[index], self.port[index])  # 启动指定易得客浏览器
            # time.sleep(2)
            self.wait_for_port(self.port[index])
            page = ChromiumPage("127.0.0.1:" + str(self.port[index]))  # 接管浏览器
            # time.sleep(2)
            try:
                page.set.window.max()
            except RuntimeError:
                pass  # 易得客不支持，不影响下载
            self.main(page)
        data = self.excel(self.file_path)
        self.message_send(data)



if __name__ == '__main__':
    # CLI 入口：config 仅在此处定义
    config = {        # 易得客账号
        "username": "13778451825",
        # 易得客密码
        "password": "wjh12345.",
        # 企业微信使用的手机号
        "number":'18280194086',
        # 易得客中的店铺ip(由于数据量过大，建议只填写单个店铺ip单次执行)
        "ip": ["35.84.243.7"],  # 多家店铺依次填写，与 port 一一对应
        # 端口号,4位数字(一般以8和9开头,例如：8200,9527,9999,8888)
        "port": [9999],
        # 目标商品对应的ASIN
        "asin": ["B0963P4V3B","B09YVFYTGX"],
        # 文件存放路径
        "file_path": r"C:\Users\admin\Desktop\产品流量词监控"
    }
    comment = Comment(config)
    comment.run()
