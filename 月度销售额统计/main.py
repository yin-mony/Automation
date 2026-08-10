
import time
import threading
import socket
from DrissionPage import ChromiumPage,Chromium

import psutil
from YidekeLogin import Specification
from TikTokSellerLogin import TikTokSellerLogin, _agent_dbg
import re
import os
import sys
import cv2
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

class Automation():
    def __init__(self, config):
        self.username = config["username"]
        self.password = config["password"]
        self.ip = config["ip"]
        self.port = config["port"]
        self.experts = config["experts"]
        self.file_path = config["file_path"]
        self.tiktok_email = config.get("tiktok_email", "")
        self.tiktok_password = config.get("tiktok_password", "")
        self.on_captcha = config.get("on_captcha")

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
        resolver = TikTokSellerLogin(page)
        deadline = time.time() + timeout
        loop = 0
        while time.time() < deadline:
            page = resolver._pick_seller_tab(page)
            url = (page.url or '').lower()
            has_analytics = bool(page.ele('x://div[text()="Analytics"]', timeout=2))
            in_seller_backend = resolver.is_seller_backend(page)
            if loop == 0 or loop % 5 == 0:
                # #region agent log
                _agent_dbg("D", "main.wait_for_seller_page", "poll", {
                    "runId": "post-fix",
                    "loop": loop, "url": url, "has_analytics": has_analytics,
                    "in_seller_backend": in_seller_backend,
                })
                # #endregion
            loop += 1
            if '/account/login' in url:
                time.sleep(2)
                continue
            if has_analytics or in_seller_backend:
                # #region agent log
                _agent_dbg("A", "main.wait_for_seller_page", "success", {
                    "runId": "post-fix", "url": url, "has_analytics": has_analytics,
                })
                # #endregion
                return page
            time.sleep(2)
        # #region agent log
        _agent_dbg("E", "main.wait_for_seller_page", "timeout", {
            "runId": "post-fix",
            "final_url": (page.url or "").lower(), "loops": loop,
        })
        # #endregion
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
            f'--remote-debugging-port={port}'
        ]

        print("启动命令:")
        print(" ".join(cmd))

        try:
            subprocess.Popen(cmd, cwd=str(base))
            print("启动成功（已发起进程）")
        except Exception as e:
            print("启动失败:", e)
            raise

    def Start(self):
        sp = Specification(self.username, self.password)  # 其他易得客
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        self.run_edecker_automation(self.ip)  # 访问全部店铺
        time.sleep(4)
        for index ,ip in enumerate(self.ip):
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
            # page.set.window.max()
            month_map = {
                "1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr",
                "5": "May", "6": "Jun", "7": "Jul", "8": "Aug",
                "9": "Sept", "10": "Oct", "11": "Nov", "12": "Dec"
            }
            month = str((datetime.now() - relativedelta(months=1)).month)
            month_cn = month_map[month]
            # time.sleep(3)
            if self.tiktok_email and self.tiktok_password:
                tiktok_login = TikTokSellerLogin(page, on_captcha=self.on_captcha)
                tiktok_login.login(self.tiktok_email, self.tiktok_password)
                page = tiktok_login.get_active_page()
                # #region agent log
                _agent_dbg("A", "main.Start", "after tiktok login", {
                    "runId": "post-fix",
                    "main_page_url": (page.url or "").lower(),
                    "tiktok_page_url": (tiktok_login.page.url or "").lower(),
                    "same_object": page is tiktok_login.page,
                })
                # #endregion
            page = self.wait_for_seller_page(page)
            page.ele('x://div[text()="Analytics"]',timeout=30).click()
            time.sleep(2)
            page.ele('x://span[text()="Shop analytics"]', timeout=25)
            time.sleep(4)
            # page.ele('x://span[text()="LIVE & video analytics"]').click(by_js=True)
            # 如果不存在该标签，则点击Content analytics
            if not page.ele('x://span[text()="LIVE & video analytics"]'):
                page.ele('x://span[text()="Content analytics"]', timeout=60).click(by_js=True)
                time.sleep(5)
                page.ele('x://div[text()="Video & photo overview"]', timeout=60).click(by_js=True)
                time.sleep(4)
            else:
                page.ele('x://span[text()="LIVE & video analytics"]').click(by_js=True)
                time.sleep(4)
                page.ele('x://div[text()="Performance"]', timeout=60).click(by_js=True)
                time.sleep(2)
            # time.sleep(4)
            # page.ele('x://div[text()="Performance"]', timeout=25).click(by_js=True)
            # time.sleep(4)
            page.ele('x://button[text()="View Video Details"]', timeout=25).click()
            time.sleep(4)
            page.ele('x://div[text()="(GMT-08:00)"]//following-sibling::div', timeout=25).click()
            time.sleep(1.5)
            page.ele('x://div[text()="Month"]').click()
            time.sleep(1.5)
            page.ele('x://div[text()="' + month_cn + '"]').click()
            time.sleep(10)
            page.ele('x://span[contains(text(), "Filter")]').click()
            time.sleep(1.5)
            page.ele('x://span[text()="All linked accounts"]', timeout=60).click(by_js=True)
            for _ in range(25):
                page.actions.key_down('BACKSPACE')
                time.sleep(0.1)
            time.sleep(1.5)
            for expert in self.experts:
                # self._wait_page_loaded(page)
                page.ele(f'x://div[contains(text(), "{expert}")]', timeout=60).click(by_js=True)
                time.sleep(0.78)
                page.ele('x://span[text()="Confirm"]', timeout=60).click(by_js=True)
                time.sleep(7.5)
                page.ele('x://span[contains(., "Linked account video-attributed") and contains(., "GMV")]',
                         timeout=60).click(by_js=True)
                time.sleep(7)
                page.ele('x://span[contains(text(), "Linked account video-attributed") and contains(., "GMV")]',
                         timeout=60).click(by_js=True)
                time.sleep(7)
                download = page.ele('x://span[text()="Export"]').click.to_download(save_path=self.file_path,
                                                                                   suffix='xlsx', timeout=20)
                download.wait()
                time.sleep(4)
                # 下载完成后：再次点击 Filter，聚焦账号输入框并退格清空当前达人；最后一位无需清空，循环结束
                if expert != self.experts[-1]:
                    # self._wait_page_loaded(page)
                    page.ele('x://span[contains(text(), "Filter")]', timeout=60).click(by_js=True)
                    time.sleep(1.5)
                    page.ele('x://div[contains(text(), "Account type")]/following-sibling::div//input',
                             timeout=60).click(by_js=True)
                    time.sleep(0.5)
                    for _ in range(25):
                        page.actions.key_down('BACKSPACE')
                        time.sleep(0.1)
                    time.sleep(1.5)
            # return page
            # self.dlData(page)

    def dlData(self,page):

        month_map = {
            "1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr",
            "5": "May", "6": "Jun", "7": "Jul", "8": "Aug",
            "9": "Sept", "10": "Oct", "11": "Nov", "12": "Dec"
        }
        month = str((datetime.now() - relativedelta(months=1)).month)
        month_cn = month_map[month]

        page.ele('x://div[text()="Analytics"]').click()
        time.sleep(2)
        page.ele('x://span[text()="Shop analytics"]',timeout=25)
        time.sleep(4)
        page.ele('x://span[text()="LIVE & video analytics"]').click(by_js=True)
        time.sleep(4)
        page.ele('x://div[text()="Performance"]',timeout=25).click(by_js=True)
        time.sleep(4)
        page.ele('x://button[text()="View Video Details"]',timeout=25).click()
        time.sleep(4)
        page.ele('x://div[text()="(GMT-08:00)"]//following-sibling::div',timeout=25).click()
        time.sleep(1.5)
        page.ele('x://div[text()="Month"]').click()
        time.sleep(1.5)
        page.ele('x://div[text()="' + month_cn + '"]').click()
        time.sleep(10)
        page.ele('x://span[contains(text(), "Filter")]').click()
        time.sleep(1.5)
        page.ele('x://span[text()="All linked accounts"]').click()
        for _ in range(25):
            page.actions.key_down('BACKSPACE')
            time.sleep(0.1)
        time.sleep(1.5)
        for expert in self.experts:
            page.ele(f'x://div[contains(text(), "{expert}")]').click(by_js=True)
            time.sleep(0.78)
            page.ele('x://span[text()="Confirm"]').click()
            time.sleep(7.5)
            page.ele('x://span[contains(., "Linked account video-attributed") and contains(., "GMV")]').click()
            time.sleep(7)
            page.ele('x://span[contains(text(), "Linked account video-attributed") and contains(., "GMV")]').click()
            time.sleep(7)
            download =  page.ele('x://span[text()="Export"]').click.to_download(save_path=self.file_path, suffix='xlsx',timeout=20)
            download.wait()
if __name__ == '__main__':
    yideke_ips = [item.strip() for item in os.getenv("MONTHLY_YIDEKE_IPS", "").split(",") if item.strip()]
    yideke_ports = [int(item.strip()) for item in os.getenv("MONTHLY_YIDEKE_PORTS", "").split(",") if item.strip()]
    config = {
        "username": os.getenv("YIDEKE_USERNAME", ""),
        "password": os.getenv("YIDEKE_PASSWORD", ""),
        "ip": yideke_ips,  # 多家店铺依次填写，与 port 一一对应
        "port": yideke_ports,
        "experts": ["lydia_homegoods","carhack_ryan","k8paz0xqw4","chicpicksbylydia","c7crfmav15","dailyfindsbylydia","detailing_dave_","furfreeliving_","haley1110","lydiashomefinds","homewithcamila","kerryshares","cleanwithlydia18","pltejkffq9","shopwithlydia_","sneakerheadmax_","spicypotato571","gppzoa2o03","cleaningwithemma91","cleaningwithsofia_","hrmb03eak0"],
        "file_path": r"C:\RPA流程\月度销售额统计\flie",
        "tiktok_email": os.getenv("MONTHLY_TIKTOK_EMAIL", ""),
        "tiktok_password": os.getenv("MONTHLY_TIKTOK_PASSWORD", ""),
    }
    automation = Automation(config)
    automation.Start()
