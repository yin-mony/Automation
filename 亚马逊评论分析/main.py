"""
亚马逊评论分析 — 评论下载核心逻辑

通过易得客启动店铺浏览器，在 Amazon 搜索 ASIN，
按 1～5 星展开评论并加载全部页，导出为「亚马逊评论.xlsx」。
"""

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
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


class Comment:
    """易得客 + Amazon 评论抓取与 Excel 导出。"""

    def __init__(self, config):
        """从 config 读取账号、店铺 IP/端口、ASIN 列表与保存目录。"""
        self.username = config["username"]
        self.password = config["password"]
        # 店铺IP
        ips = config["ip"]
        self.ip = ips if isinstance(ips, list) else [ips]
        # 端口
        ports = config["port"]
        self.port = ports if isinstance(ports, list) else [ports]
        self.experts = config["experts"]
        self.file_path = config["file_path"]

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
        """结束除指定 PID 外的所有 edecker 进程。"""
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
            # 在易得客店铺列表中按 IP 匹配「美国」店铺并点击「访问」
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="美国"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30
            ).click()
            time.sleep(3)
        self.kill_edecker(browser.process_id)  # 关闭易得客管理窗口，仅保留店铺浏览器
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def start_edecker(self, ip: str, port: int):
        """按店铺 IP 匹配 eDecker profile，以指定调试端口启动浏览器。"""
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
        ip_underline = ip.replace('.', '_')  # profile 目录名可能用下划线表示 IP

        all_profiles = list(profiles_path.iterdir())
        print("所有 profile:")
        for p in all_profiles:
            print(" -", p.name)

        candidates = [
            p for p in all_profiles
            if p.is_dir() and (ip_dot in p.name or ip_underline in p.name)
        ]

        if not candidates:
            raise Exception(f"未找到 IP={ip} 的 profile")

        latest = max(candidates, key=lambda p: p.stat().st_mtime)  # 取最近使用的 profile

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
        """在指定调试端口的浏览器中查找卖家后台标签页。"""
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

    # 自动化流程：登录易得客 → 打开 Amazon → 按 ASIN/星级抓评论
    def main(self):
        """逐个 ASIN 抓取 1～5 星评论，汇总后调用 excel_files 导出。"""
        sp = Specification(self.username, self.password)  # 易得客登录
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        # self.visit_shop(self.ip, self.port)
        self.run_edecker_automation(self.ip)  # 访问全部店铺
        time.sleep(4)
        # 按店铺 IP 逐个启动浏览器并采集评论
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
            # https://www.amazon.com
            # page = Chromium()
            # 新建标签页进入
            # 定义星级列表（Amazon 评论页下拉选项文案）
            stars = ["5 star only", "4 star only", "3 star only", "2 star only", "1 star only"]
            star_names = ["5星", "4星", "3星", "2星", "1星"]
            # 存储所有评论
            all_comments = {}

            page = page.new_tab("https://www.amazon.com")
            time.sleep(3)
            page.ele('x://div/input[@placeholder="Search Amazon"]',timeout=25).click()
            for experts in self.experts:
                print(f"\n正在处理商品: {experts}")
                all_comments[experts] = {}
                page.ele('x://div/input[@placeholder="Search Amazon"]',timeout=30).input(f'{experts}\n',clear=True)
                time.sleep(3)
                # 进入商品详情 → See more reviews → All stars 筛选
                page.ele(f'x://div[@data-asin="{experts}"]//a',timeout=30).click()
                time.sleep(3)
                # 滚动页面至中间指定评论元素
                ele = page.ele('x://div[text()="See more reviews"]')
                page.scroll.to_see(ele)
                time.sleep(3)
                # 定位点击“评论”一级页面
                page.ele('x://div[text()="See more reviews"]').click()
                time.sleep(3)
                # 具体评论页面
                page.ele('x://span[text() = "All stars"]').click()
                time.sleep(3)
                # 指定星级评论
                # page.ele('x://ul[@role="listbox"]/li/a[text()="5 star only"]').click()
                for star, star_name in zip(stars, star_names):
                    print(f"\n正在处理 {star_name}...")
                    page.ele(f'x://ul[@role="listbox"]/li/a[text()="{star}"]').click()
                    time.sleep(3)
                    # 循环点击「Show 10 more reviews」直到无更多评论
                    while True:
                        # 每次循环重新定位按钮（避免元素失效）
                        comment_button = page.ele('x://span/a[text()="Show 10 more reviews"]',timeout=0)
                        # comment_button.scroll.to_see()
                        if not comment_button:
                            print("没有更多评论了")
                            break
                        else:
                            comment_button.click()
                            time.sleep(2)  # 等待新评论加载
                            print(f"点击加载更多{star_name}评论...")

                    reviews = []
                    # 提取当前星级下页面全部评论正文
                    for elem in page.eles('x://span[@data-hook="review-body"]/span'):
                        text = elem.text.strip()
                        if text:
                            reviews.append(text)
                    all_comments[experts][star_name] = reviews

                    # 重新打开星级下拉，准备切换下一档
                    page.ele(f'x://span[text() = "{star}"]').click()
                    time.sleep(3)

                    # print(f"共提取 {len(reviews)} 条评论")
                    # print(f"评论: {reviews}")
                    print(f"{star_name} 共提取 {len(reviews)} 条评论")
                    print("-" * 50)  # 分隔线
                # 重新定位回到搜索框并点击
                page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()

                # 打印所有评论
                for star_name, reviews in all_comments[experts].items():
                    print(f"\n{'=' * 50}")
                    print(f"{star_name} 评论 (共 {len(reviews)} 条)")
                    print(f"{'=' * 50}")
                    for i, review in enumerate(reviews, 1):
                        print(f"{i}. {review}")
            # 导出为 Excel
            self.excel_files(all_comments)
            return all_comments

    def excel_files(slef, all_comments):
        """将 all_comments 展平为行，写入 {file_path}/亚马逊评论.xlsx。"""
        # 转换数据格式
        data = []
        for asin, star_comments in all_comments.items():  # asin 就是商品ID
            for star, reviews in star_comments.items():
                for review in reviews:
                    data.append({
                        'ASIN': asin,  # 修正1：添加 ASIN 值
                        '星级': star,
                        '评论内容': review,
                        '评论长度': len(review)
                    })

        # 创建DataFrame
        df = pd.DataFrame(data)

        # 保存到Excel
        output_dir = Path(slef.file_path)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "亚马逊评论.xlsx"
        df.to_excel(output_path, index=False)

        print(f"已保存 {len(data)} 条评论到 {output_path}")
        return output_path



    # 启动
    def run(self):
        """入口：执行 main()。"""
        self.main()


if __name__ == '__main__':
    # CLI 入口：config 仅在此处定义
    config = {
        "username": "19944318805",
        "password": "DY0924DY0924",
        "ip": ["35.82.248.104"],  # 多家店铺依次填写，与 port 一一对应
        "port": [8945],
        "experts": ["B0963P4V3B","B09YVFYTGX"],
        "file_path": r"C:\RPA流程\亚马逊评论分析\flie"
    }
    comment = Comment(config)
    comment.run()
