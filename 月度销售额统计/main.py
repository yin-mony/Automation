
import time
import threading
from DrissionPage import ChromiumPage,Chromium

import psutil
from YidekeLogin import Specification
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

    def _wait_page_loaded(self, page, timeout=180):
        """等待当前标签页文档加载完成后再继续（境外网络慢时减少元素未找到）。不改变业务步骤。"""
        try:
            page.set.timeouts(page_load=timeout)
        except Exception:
            pass
        page.wait.doc_loaded(timeout=timeout, raise_err=False)
        deadline = time.perf_counter() + timeout
        while time.perf_counter() < deadline:
            try:
                if getattr(page.states, "ready_state", None) == "complete":
                    break
            except Exception:
                pass
            time.sleep(0.2)
        time.sleep(0.3)

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

    def run_edecker_automation(self, ips,port=9222):
        """
        全部启动店铺
        :param port:
        :return:
        """
        browser = Chromium(port)
        tab = browser.latest_tab
        buttons = tab.eles("t:button@@text()=访问")
        for btn in buttons:
            btn.click()
            time.sleep(3)
        time.sleep(2)
        for ip in ips:
            tab.ele(f'x://div[text()="{ip}"]//following-sibling::button').click()
            time.sleep(3)
        self.kill_edecker(browser.process_id)
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
            self.start_edecker(self.ip[index], self.port[index])  # 启动指定易得客浏览器
            time.sleep(2)
            page = ChromiumPage("127.0.0.1:" + str(self.port[index]))  # 接管浏览器
            time.sleep(2)
            page.set.window.max()
            self._wait_page_loaded(page)
            self.dlData(page)

    def dlData(self,page):

        self._wait_page_loaded(page)
        month_map = {
            "1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr",
            "5": "May", "6": "Jun", "7": "Jul", "8": "Aug",
            "9": "Sept", "10": "Oct", "11": "Nov", "12": "Dec"
        }
        month = str((datetime.now() - relativedelta(months=1)).month)
        month_cn = month_map[month]

        page.ele('x://div[text()="Analytics"]',timeout=25).click(by_js=True)
        time.sleep(2)
        page.ele('x://span[text()="Shop analytics"]',timeout=25)
        time.sleep(2)
        self._wait_page_loaded(page)
        # 为考虑后续网站可能退回原始标签界面，所以需要判断是否存在LIVE & video analytics
        if not page.ele('x://span[text()="LIVE & video analytics"]'):
            # 如果不存在LIVE & video analytics，则点击Content analytics
            page.ele('x://span[text()="Content analytics"]',timeout=60).click(by_js=True)
            time.sleep(5)
            page.ele('x://div[text()="Video & photo overview"]',timeout=60).click(by_js=True)
            time.sleep(4)
        else:
            page.ele('x://span[text()="LIVE & video analytics"]').click(by_js=True)
            time.sleep(4)
            page.ele('x://div[text()="Performance"]',timeout=60).click(by_js=True)
            time.sleep(2)
        self._wait_page_loaded(page)
        page.ele('x://span[text()="Video details"]',timeout=60).click(by_js=True)
        self._wait_page_loaded(page)
        time.sleep(4)
        page.ele('x://div[text()="(GMT-08:00)"]//following-sibling::div',timeout=60).click(by_js=True)
        self._wait_page_loaded(page)
        # time.sleep(4)
        page.ele('x://div[text()="Month"]',timeout=60).click()
        self._wait_page_loaded(page)
        # time.sleep(4)
        page.ele('x://div[text()="' + month_cn + '"]',timeout=60).click()
        # time.sleep(10)
        self._wait_page_loaded(page)
        page.ele('x://span[contains(text(), "Filter")]',timeout=60).click(by_js=True)
        self._wait_page_loaded(page)
        # time.sleep(5)
        page.ele('x://span[text()="All linked accounts"]',timeout=60).click(by_js=True)
        self._wait_page_loaded(page)
        for _ in range(25):
            page.actions.key_down('BACKSPACE')
            time.sleep(0.1)
        time.sleep(1.5)
        for expert in self.experts:
            self._wait_page_loaded(page)
            page.ele(f'x://div[contains(text(), "{expert}")]',timeout=60).click(by_js=True)
            time.sleep(0.78)
            page.ele('x://span[text()="Confirm"]',timeout=60).click(by_js=True)
            time.sleep(7.5)
            page.ele('x://span[contains(., "Linked account video-attributed") and contains(., "GMV")]',timeout=60).click(by_js=True)
            time.sleep(7)
            page.ele('x://span[contains(text(), "Linked account video-attributed") and contains(., "GMV")]',timeout=60).click(by_js=True)
            time.sleep(7)
            download = page.ele('x://span[text()="Export"]').click.to_download(save_path=self.file_path, suffix='xlsx', timeout=20)
            download.wait()
            time.sleep(4)
            # 下载完成后：再次点击 Filter，聚焦账号输入框并退格清空当前达人；最后一位无需清空，循环结束
            if expert != self.experts[-1]:
                self._wait_page_loaded(page)
                page.ele('x://span[contains(text(), "Filter")]', timeout=60).click(by_js=True)
                time.sleep(1.5)
                page.ele('x://div[contains(text(), "Account type")]/following-sibling::div//input', timeout=60).click(by_js=True)
                time.sleep(0.5)
                for _ in range(25):
                    page.actions.key_down('BACKSPACE')
                    time.sleep(0.1)
                time.sleep(1.5)
                    
if __name__ == '__main__':
    # 示例占位：请按环境填写后再运行；GUI 入口请使用 run.py
    config = {
        "username": "",
        "password": "",
        "ip": ["127.0.0.1"],
        "port": [9222],
        "experts": [],
        "file_path": ""
    }
    automation = Automation(config)
    automation.Start()