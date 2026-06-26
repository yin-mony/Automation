"""
下载全站点汇总报告

通过易得客登录 Amazon 卖家后台，按站点（美/加/墨/巴）在 Reports Repository
申请当月 SELLER_SUMMARY_DATE_RANGE（Summary）汇总报告。
"""

import time
import socket
import psutil
import os
import subprocess
from pathlib import Path
from DrissionPage import ChromiumPage,Chromium
from YidekeLogin import Specification
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# YidekeLogin 启动易得客时固定使用的调试端口
# EDECKE_DEBUG_PORT = 9222


class TestPage:
    """易得客 + Amazon 卖家后台：按站点申请 Summary 月度汇总报告。"""

    def __init__(self, config):
        """从 config 读取账号、店铺 IP/端口；站点列表默认四国。"""
        self.username = config["username"]
        self.password = config["password"]
        # 店铺IP
        ips = config["ip"]
        self.ip = ips if isinstance(ips, list) else [ips]
        # 端口
        ports = config["port"]
        self.port = ports if isinstance(ports, list) else [ports]
        # 站点
        datas = ["United States", "Canada", "Mexico", "Brazil"]
        self.data = datas if isinstance(datas, list) else [datas]


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
        """在指定调试端口的浏览器中查找 Amazon 卖家后台标签页。"""
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


    def main(self):
        """完整自动化：登录易得客 → 启动店铺 → 逐站点申请 Summary 报告。"""
        sp = Specification(self.username, self.password)  # 易得客登录
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        self.run_edecker_automation(self.ip)  # 访问全部店铺
        time.sleep(4)
        for index, ip in enumerate(self.ip):
            self.kill_edecker_on_port(self.port[index])  # 启动前清理占用端口的 edecker
            time.sleep(1)
            self.start_edecker(self.ip[index], self.port[index])  # 启动指定易得客浏览器
            # time.sleep(2)
            self.wait_for_port(self.port[index])
            # page = ChromiumPage("127.0.0.1:" + str(self.port[index]))  # 接管浏览器
            page = self.find_seller_tab(self.port[index])
            # time.sleep(2)
            try:
                page.set.window.max()
            except RuntimeError:
                pass  # 易得客不支持，不影响下载

            print("当前 tab URL:", page.url)
            # 获取当前年月份
            now = datetime.now()
            # year = str(now.year)  # 2026当前年份
            month = str(now.month) # 月份
            # 定义月份
            month_map = {
                "1": "January", "2": "February", "3": "March", "4": "April",
                "5": "May", "6": "June", "7": "July", "8": "August",
                "9": "September", "10": "October", "11": "November", "12": "December"
            }

            month_cn = month_map[month]
            # 切换后台显示语言为 English
            page.ele('x://div[@aria-label="语言"] | //div[@aria-label="Language"]').click()
            time.sleep(1.5)
            # //div[text()="English"]
            page.ele('x://div[text()="English"]').click()
            time.sleep(5)

            # 站点切换：四国分别申请 Summary 报告（各站点表单字段数不同）
            for data in self.data:

                page.ele('x://span[text()="KORCCI LLC"]').click()
                time.sleep(0.78)
                page.ele('x://*[text()="See all"][1]').click()
                time.sleep(0.78)
                page.ele(f'x://span[contains(text(), "{data}")]',timeout=20).click()
                time.sleep(0.78)
                page.ele('x://kat-button[@label="Select account"]',timeout=20).click()
                # 汉堡菜单（shadow DOM）进入 Payments → Reports Repository
                menu_host = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=30)
                menu = menu_host.shadow_root
                # print(11111)
                menu.ele('x://div/img', timeout=30).click()
                time.sleep(1)
                menu.ele('x://div/span[text()="Payments"]',timeout=20).click()
                time.sleep(3)
                page.wait(1)
                menu.ele('x://div/span[text()="Reports Repository"]',timeout=20).click()
                time.sleep(3)

                page.wait.load_start()
                # 美国站：三列下拉（店铺/账户/报告类型）+ 选月
                if data == 'United States':
                    dropdown_stores = page.ele('x://form/div[1]/kat-dropdown')
                    dropdown_stores.click()
                    time.sleep(3)
                    dropdown_stores.shadow_root('x://kat-option[@value="ALL_STORES"]',timeout=20).click()
                    time.sleep(3)
                    dropdown_account = page.ele('x://form/div[2]/kat-dropdown')
                    dropdown_account.click()
                    time.sleep(1)
                    dropdown_account.shadow_root('x://kat-option[@value="ALL"]',timeout=20).click()
                    time.sleep(3)
                    dropdown_report = page.ele('x://form/div[3]/kat-dropdown')
                    dropdown_report.click()
                    time.sleep(1)
                    dropdown_report.shadow_root('x://kat-option[@value="SELLER_SUMMARY_DATE_RANGE"]',timeout=20).click()
                    time.sleep(3)
                    page.ele('x://div/kat-radiobutton[@label="Month"]', timeout=20).click(by_js=True)
                    time.sleep(3)
                    # 年月份选择//div[@class="date-range-item"]/kat-dropdown
                    # 月份
                    dropdown_month = page.ele('x://div[@class="date-range-item"][1]/kat-dropdown')
                    dropdown_month.click()
                    dropdown_month.shadow_root(f'x://kat-option//div[text()="{month_cn}"]',timeout=20).click()
                    time.sleep(3)

                # 巴西站：两列下拉（账户/报告类型）+ 选月
                elif data == 'Brazil':
                    dropdown_account = page.ele('x://form/div[1]/kat-dropdown')
                    dropdown_account.click()
                    time.sleep(1)
                    dropdown_account.shadow_root('x://kat-option[@value="ALL"]', timeout=20).click()
                    time.sleep(3)
                    dropdown_report = page.ele('x://form/div[2]/kat-dropdown')
                    dropdown_report.click()
                    time.sleep(3)
                    dropdown_report.shadow_root('x://kat-option[@value="SELLER_SUMMARY_DATE_RANGE"]', timeout=20).click()
                    time.sleep(3)
                    page.ele('x://div/kat-radiobutton[@label="Month"]', timeout=20).click(by_js=True)
                    time.sleep(3)

                    # 年月份选择//div[@class="date-range-item"]/kat-dropdown
                    # 月份
                    dropdown_month = page.ele('x://div[@class="date-range-item"][1]/kat-dropdown')
                    dropdown_month.click()
                    time.sleep(1)
                    dropdown_month.shadow_root(f'x://kat-option//div[text()="{month_cn}"]', timeout=20).click()
                    time.sleep(3)

                # 墨西哥站：报告类型 + 选月（表单列数与加拿大类似）
                elif data == 'Mexico':
                    dropdown_report = page.ele('x://form/div[2]/kat-dropdown')
                    dropdown_report.click()
                    time.sleep(3)
                    dropdown_report.shadow_root('x://kat-option[@value="SELLER_SUMMARY_DATE_RANGE"]', timeout=20).click()
                    time.sleep(3)
                    page.ele('x://div/kat-radiobutton[@label="Month"]', timeout=20).click(by_js=True)
                    time.sleep(3)

                    # 年月份选择//div[@class="date-range-item"]/kat-dropdown
                    # 月份
                    dropdown_month = page.ele('x://div[@class="date-range-item"][1]/kat-dropdown')
                    dropdown_month.click()
                    time.sleep(3)
                    dropdown_month.shadow_root(f'x://kat-option//div[text()="{month_cn}"]', timeout=20).click()
                    time.sleep(3)

                # 加拿大站：报告类型 + 选月
                elif data == 'Canada':
                    dropdown_report = page.ele('x://form/div[2]/kat-dropdown')
                    dropdown_report.click()
                    time.sleep(3)
                    dropdown_report.shadow_root('x://kat-option[@value="SELLER_SUMMARY_DATE_RANGE"]',timeout=20).click()
                    time.sleep(3)
                    page.ele('x://div/kat-radiobutton[@label="Month"]', timeout=20).click(by_js=True)
                    time.sleep(3)
                    # 年月份选择//div[@class="date-range-item"]/kat-dropdown
                    # 月份
                    dropdown_month = page.ele('x://div[@class="date-range-item"][1]/kat-dropdown')
                    dropdown_month.click()
                    time.sleep(3)
                    dropdown_month.shadow_root(f'x://kat-option//div[text()="{month_cn}"]', timeout=20).click()
                    time.sleep(3)

                page.ele('x://kat-button[translate(@label, "ABCDEFGHIJKLMNOPQRSTUVWXYZ",'
                         ' "abcdefghijklmnopqrstuvwxyz")="request report"]',timeout=20).click()
                time.sleep(3)

    def run(self):
        """入口：执行 main()。"""
        self.main()


if __name__ == '__main__':
    # CLI 入口：config 仅在此处定义
    config = {        "username": "13281439638",
        "password": "13281439638@MM",
        "ip": ["54.70.92.80"],  # 多家店铺依次填写
        "port": [9999],
        # "data": ["United States","Canada","Mexico","Brazil"],
        # "file_path": r"F:\RPA流程"
    }
    dev = TestPage(config)
    dev.run()
