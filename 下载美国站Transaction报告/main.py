"""
下载 Transaction 报告

通过易得客登录 Amazon 卖家后台，在 Reports Repository 申请
上一个自然月指定站点 SELLER_TRANSACTION_DATE_RANGE（Transaction）报告。
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
from email_util import PROJECT_NAME, deliver_outputs

# YidekeLogin 启动易得客时固定使用的调试端口
# EDECKE_DEBUG_PORT = 9222


class TestPage:
    """易得客 + Amazon 卖家后台：申请所选站点上一个自然月的 Transaction 报告。"""

    def __init__(self, config=None):
        """设置默认初始值，并使用外部 config 覆盖对应配置。"""
        config = config or {}

        # 易得客登录账号，用于进入易得客工作台；未传配置时读取环境变量。
        self.username = (config.get("username") or os.getenv("YIDEKE_USERNAME") or "").strip()

        # 易得客登录密码，与上方账号配套使用；未传配置时读取环境变量。
        self.password = config.get("password") or os.getenv("YIDEKE_PASSWORD") or ""

        # 需要运行的易得客店铺 IP 列表；多个店铺时按顺序继续添加。
        ips = config.get("ip", ["54.70.92.80"])
        self.ip = ips if isinstance(ips, list) else [ips]

        # 每个店铺浏览器使用的远程调试端口，必须与 ip 列表按下标一一对应。
        ports = config.get("port", [9527])
        self.port = ports if isinstance(ports, list) else [ports]

        # Amazon 后台站点：中文用于 GUI，英文用于 Seller Central 账号切换
        self.siteEnglishMap = {
            "美国": "United States",
            "加拿大": "Canada",
            "墨西哥": "Mexico",
            "巴西": "Brazil",
            "英国": "United Kingdom",
            "法国": "France",
            "德国": "Germany",
            "意大利": "Italy",
            "西班牙": "Spain",
            "荷兰": "Netherlands",
            "瑞典": "Sweden",
            "波兰": "Poland",
            "比利时": "Belgium",
            "爱尔兰": "Ireland",
            "日本": "Japan",
            "新加坡": "Singapore",
            "澳大利亚": "Australia",
            "印度": "India",
            "阿联酋": "United Arab Emirates",
            "沙特阿拉伯": "Saudi Arabia",
            "土耳其": "Turkey",
            "埃及": "Egypt",
            "南非": "South Africa",
        }

        # 易得客工作台中的店铺站点，用于在对应站点区域查找店铺访问按钮。
        rawAutoSiteName = config.get("autoSiteName") or config.get("shopSiteName") or "美国"
        if isinstance(rawAutoSiteName, list):
            rawAutoSiteName = rawAutoSiteName[0] if rawAutoSiteName else "美国"
        self.autoSiteName = str(rawAutoSiteName).strip() or "美国"

        # Amazon Seller Central 后台目标站点列表，按配置顺序逐个切换并请求报告。
        rawSiteNames = (
            config.get("amazonSiteNames")
            or config.get("amazonSiteName")
            or config.get("siteName")
            or config.get("data")
            or ["美国"]
        )
        if not isinstance(rawSiteNames, (list, tuple)):
            rawSiteNames = [rawSiteNames]
        self.siteNames = []
        for rawSiteName in rawSiteNames:
            siteName = str(rawSiteName).strip()
            if siteName and siteName not in self.siteNames:
                self.siteNames.append(siteName)
        if not self.siteNames:
            self.siteNames = ["美国"]

        # 首个站点用于 Amazon 登录后的账户选择，后续站点由主流程依次切换。
        self.siteName = self.siteNames[0]
        self.data = self.siteEnglishMap.get(self.siteName, self.siteName)

        # Amazon 登录邮箱；留空时优先使用易得客浏览器中已保存的登录状态。
        self.amazonEmail = (config.get("amazonEmail") or config.get("amazon_email") or "").strip()

        # Amazon 登录密码；留空时不主动填写密码，使用浏览器已有登录状态。
        self.amazonPassword = config.get("amazonPassword") or config.get("amazon_password") or ""

        # 运行环境标识：True 为线上，False 为线下；仅影响 GUI 和日志显示。
        # 两种模式都会先登录易得客并访问对应店铺，不改变主要自动化流程。
        self.isOnline = bool(config.get("isOnline", False))

        # 是否在流程完成后发送邮件：True 发送，False 不发送。
        self.sendEmail = bool(config.get("sendEmail", False))

        # 报告接收邮箱；sendEmail 为 True 时必须填写。
        self.email = (config.get("email") or "").strip()

        # SMTP 发件邮箱，用于发送 Transaction 报告邮件；未传配置时读取环境变量。
        self.senderEmail = (config.get("sender_email") or os.getenv("SMTP_SENDER") or "").strip()

        # SMTP 发件邮箱授权码，不是邮箱登录密码；未传配置时读取环境变量。
        self.smtpAuthCode = (config.get("smtp_auth_code") or os.getenv("SMTP_AUTH_CODE") or "").strip()

        # 报告文件查找目录，默认使用当前 Windows 用户的桌面目录。
        self.filePath = Path(config.get("file_path") or (Path.home() / "Desktop"))

    def collectOutputFiles(self):
        """在报告目录中查找文件名包含项目名的文件。"""
        folder = self.filePath
        if not folder.exists():
            print(f"报告目录不存在: {folder}")
            return []

        files = [
            p for p in folder.iterdir()
            if p.is_file() and PROJECT_NAME in p.name
        ]
        files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        if files:
            print(f"在 {folder} 找到 {len(files)} 个匹配文件")
        else:
            print(f"在 {folder} 未找到文件名含「{PROJECT_NAME}」的文件")
        return [str(p) for p in files]

    def stopProgram(self):
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

    def killEdecker(self, excludePid):
        """结束除指定 PID 外的所有 edecker 进程。"""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                pid = proc.info['pid']
                name = proc.info['name']
                if name and name.lower() == 'edecker.exe':
                    if pid != excludePid:
                        proc.kill()
            except:
                pass

    def killEdeckerOnPort(self, port):
        """启动店铺浏览器前，结束占用该调试端口的旧进程"""
        flag = f'--remote-debugging-port={port}'
        for proc in psutil.process_iter(['pid', 'cmdline']):
            try:
                cmdline = proc.info.get('cmdline') or []
                if any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except:
                pass

    def waitForPort(self, port, timeout=60):
        """等待调试端口就绪"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(('127.0.0.1', port), timeout=2):
                    return
            except OSError:
                time.sleep(1)
        raise RuntimeError(f'等待 127.0.0.1:{port} 超时 ({timeout}s)')

    def waitForSellerPage(self, page, timeout=90):
        """等待页面进入 TikTok 卖家后台"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            url = (page.url or '').lower()
            if 'seller' in url or 'tiktok' in url:
                return
            time.sleep(2)
        raise RuntimeError(f'店铺浏览器未进入 TikTok 后台，当前 URL: {page.url}')

    def visitShop(self, ip, port=9222):
        """
        点击指定店铺访问
        :param ip: 店铺IP
        :param port: 易得客管理浏览器端口
        """
        browser = Chromium(port)
        tab = browser.latest_tab
        tab.ele(f'x://div[text()="{ip}"]//following-sibling::button').click()
        time.sleep(3)

    def runEdeckerAutomation(self, ips, port=9222):
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
        print(f"按易得客店铺站点访问: {self.autoSiteName}")
        for ip in ips:
            # 按配置站点区域和店铺 IP 找到访问按钮
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="{self.autoSiteName}"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30
            ).click()
            time.sleep(3)
        self.killEdecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def startEdecker(self, ip, port):
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

    def findSellerTab(self, port, timeout=90):
        """在指定调试端口的浏览器中查找 Amazon 卖家后台标签页。"""
        deadline = time.time() + timeout
        lastError = ""

        while time.time() < deadline:
            try:
                # 端口开始监听后 CDP 仍可能处于启动阶段，接管失败时继续等待
                page = ChromiumPage(f"127.0.0.1:{port}")
                browser = page.browser
                for tab_id in browser.tab_ids:
                    tab = browser.get_tab(tab_id)
                    url = (tab.url or '').lower()
                    print("检测 tab URL:", url)

                    if url.startswith("chrome-extension://"):
                        continue

                    if "sellercentral.amazon" in url or "/ap/signin" in url:
                        return tab
            except Exception as exc:
                lastError = str(exc)
                print(f"Amazon 浏览器尚未就绪，继续等待: {exc}")

            time.sleep(1)

        raise RuntimeError(
            f"未找到 Amazon Seller Central 后台标签页，端口={port}，最后错误: {lastError}"
        )



    def main(self):
        """完整自动化：登录易得客 → 启动店铺 → 逐站申请 Transaction 报告。"""
        env = "线上" if self.isOnline else "线下"
        print(f"{env}模式：登录易得客管理端并访问店铺")
        sp = Specification(self.username, self.password)
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        self.runEdeckerAutomation(self.ip)
        time.sleep(4)

        for index, ip in enumerate(self.ip):
            self.killEdeckerOnPort(self.port[index])  # 启动前清理占用端口的 edecker
            time.sleep(1)
            self.startEdecker(self.ip[index], self.port[index])  # 启动指定易得客浏览器
            # time.sleep(2)
            self.waitForPort(self.port[index])
            # page = ChromiumPage("127.0.0.1:" + str(self.port[index]))  # 接管浏览器
            page = self.findSellerTab(self.port[index])
            # 检查 Amazon 登录页、密码页、账户选择页和二步验证码
            page = Specification.AmazonSeller(
                page,
                store_password=self.amazonPassword or None,
                siteName=self.data,
                siteChineseName=self.siteName,
            ).login(
                self.amazonEmail or None,
                self.amazonPassword or None,
            )
            print(f"Amazon 登录检查完成，当前 URL: {page.url}")
            # time.sleep(2)
            try:
                page.set.window.max()
            except RuntimeError:
                pass  # 易得客不支持，不影响下载

            print("当前 tab URL:", page.url)
            # 以系统时间减一个月，跨年时自动选择上一年的 12 月
            reportMonth = datetime.now() - relativedelta(months=1)
            year = str(reportMonth.year)
            month = str(reportMonth.month)
            monthValue = str(reportMonth.month - 1)
            # 定义月份
            monthEnglishMap = {
                "1": "January", "2": "February", "3": "March", "4": "April",
                "5": "May", "6": "June", "7": "July", "8": "August",
                "9": "September", "10": "October", "11": "November", "12": "December"
            }
            monthChineseMap = {
                "1": "一月", "2": "二月", "3": "三月", "4": "四月",
                "5": "五月", "6": "六月", "7": "七月", "8": "八月",
                "9": "九月", "10": "十月", "11": "十一月", "12": "十二月"
            }

            monthEnglishName = monthEnglishMap[month]
            monthChineseName = monthChineseMap[month]
            print(f"本次请求报告月份: {year}年{month}月")
            # 先确认 Amazon 页面语言，已是中文时不重复切换
            chineseLang = page.ele('x://div[@aria-label="语言"] | //*[normalize-space()="ZH"]', timeout=3)
            chineseText = page.ele('x://*[contains(text(),"管理库存") or contains(text(),"帮助")]', timeout=2)
            if chineseLang or chineseText:
                print("Amazon 后台已是中文简体，跳过语言切换")
            else:
                page.ele('x://div[@aria-label="Language"] | //div[@aria-label="语言"]', timeout=15).click()
                time.sleep(1.5)
                page.ele('x://div[text()="中文(简体)"] | //*[normalize-space()="中文(简体)"]', timeout=10).click()
                time.sleep(5)
                print("已切换为中文简体")

            # 语言确认后按配置顺序切换 Amazon 后台站点并逐站请求报告
            allSiteNames = list(self.siteEnglishMap.keys()) + list(self.siteEnglishMap.values())
            siteButtonCondition = " or ".join([f'contains(normalize-space(), "{siteName}")' for siteName in allSiteNames])
            for siteIndex, amazonSiteName in enumerate(self.siteNames, 1):
                amazonSiteEnglishName = self.siteEnglishMap.get(amazonSiteName, amazonSiteName)
                print(
                    f"开始处理 Amazon 后台站点 "
                    f"({siteIndex}/{len(self.siteNames)}): {amazonSiteName}"
                )

                switchEntry = None
                switchSelectors = [
                    f'x://button[{siteButtonCondition}]',
                    f'x://span[{siteButtonCondition}]/ancestor::button[1]',
                    f'x://span[{siteButtonCondition}]/ancestor::div[contains(@class,"picker") or contains(@class,"switch") or contains(@class,"selector")][1]',
                    f'x://span[{siteButtonCondition}]/ancestor::div[1]',
                    'x://span[contains(text(),"Amazon(")]',
                    'x://span[contains(text(),"Amazon（")]',
                    'x://span[text()="KORCCI LLC"]',
                ]
                for selector in switchSelectors:
                    try:
                        switchEntry = page.ele(selector, timeout=3)
                    except Exception:
                        switchEntry = None
                    if switchEntry:
                        break
                if not switchEntry:
                    raise Exception("没有找到 Amazon 后台账号站点切换入口")
                switchText = " ".join(str(switchEntry.text or "").split())
                if amazonSiteName in switchText or amazonSiteEnglishName in switchText:
                    print(f"Amazon 后台已是目标站点: {amazonSiteName}")
                else:
                    switchEntry.click()
                    time.sleep(1)
                    seeAll = page.ele(
                        'x://*[normalize-space()="查看所有"]'
                        ' | //*[normalize-space()="See all"]',
                        timeout=3,
                    )
                    if not seeAll:
                        switchEntry.click(by_js=True)
                        time.sleep(1)
                        seeAll = page.ele(
                            'x://*[normalize-space()="查看所有"]'
                            ' | //*[normalize-space()="See all"]',
                            timeout=20,
                        )
                    seeAll.click()
                    time.sleep(1)
                    page.ele(
                        f'x://*[normalize-space()="{amazonSiteName}" '
                        f'or contains(normalize-space(), "{amazonSiteName}（") '
                        f'or contains(normalize-space(), "{amazonSiteName} (") '
                        f'or normalize-space()="{amazonSiteEnglishName}" '
                        f'or contains(normalize-space(), "{amazonSiteEnglishName} (")]',
                        timeout=20,
                    ).click()
                    time.sleep(1)
                    page.ele(
                        'x://kat-button[@label="选择账户"]'
                        ' | //kat-button[@label="Select account"]'
                        ' | //button[normalize-space()="选择账户"]'
                        ' | //button[normalize-space()="Select account"]'
                        ' | //span[normalize-space()="选择账户"]/ancestor::button[1]'
                        ' | //span[normalize-space()="Select account"]/ancestor::button[1]',
                        timeout=20,
                    ).click()
                    time.sleep(5)
                    print(f"Amazon 后台已切换到目标站点: {amazonSiteName}")

                # 汉堡菜单（shadow DOM）进入 Payments → Reports Repository
                menu_host = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=30)
                menu = menu_host.shadow_root
                menu.ele('x://div/img', timeout=30).click()
                time.sleep(1)
                menu.ele(
                    'x://div/span[normalize-space()="付款" '
                    'or normalize-space()="支付" '
                    'or normalize-space()="Payments"]',
                    timeout=20,
                ).click()
                time.sleep(3)
                page.wait(1)
                menu.ele(
                    'x://div/span[normalize-space()="报告库" '
                    'or normalize-space()="报告存储库" '
                    'or normalize-space()="Reports Repository"]',
                    timeout=20,
                ).click()
                time.sleep(3)

                page.wait.load_start()
                # 美国站有商城、账户类型、报告类型 3 个筛选，其他站点通常只有后两个。
                filterDropdowns = [
                    dropdown for dropdown in page.eles(
                        'x://form//div[contains(@class,"selection-filter")]/kat-dropdown'
                    )
                    if dropdown.states.is_displayed
                ]
                filterCount = len(filterDropdowns)
                expectedFilterCount = 3 if amazonSiteEnglishName == "United States" else 2
                if filterCount != expectedFilterCount:
                    raise RuntimeError(
                        f"{amazonSiteName} 报告筛选数量异常，"
                        f"预期 {expectedFilterCount} 个，实际 {filterCount} 个"
                    )
                print(f"{amazonSiteName} 报告筛选数量: {filterCount}")

                # 只有存在商城筛选时才选择所有商城，避免非美国站把账户类型错当成商城。
                dropdownStores = page.ele(
                    'x://form//div[contains(@class,"store-selection")]/kat-dropdown',
                    timeout=2,
                )
                if filterCount == 3 and not dropdownStores:
                    dropdownStores = filterDropdowns[0]
                if dropdownStores:
                    storeValue = dropdownStores.attr("value") or ""
                    storeClickable = dropdownStores.states.is_clickable
                    print(
                        f"商城筛选: 当前值={storeValue or '空'}，"
                        f"可点击={storeClickable}"
                    )
                    if storeValue != "ALL_STORES":
                        if not storeClickable:
                            raise RuntimeError(
                                f"{amazonSiteName} 商城筛选不可点击，"
                                f"当前值为 {storeValue or '空'}"
                            )
                        storeOption = dropdownStores.shadow_root(
                            'x://kat-option[@value="ALL_STORES"]'
                            ' | //kat-option[normalize-space()="所有商城" '
                            'or normalize-space()="All stores" '
                            'or normalize-space()="All Stores"]'
                            ' | //kat-option[.//*[normalize-space()="所有商城" '
                            'or normalize-space()="All stores" '
                            'or normalize-space()="All Stores"]]',
                            timeout=3,
                        )
                        if not storeOption:
                            raise RuntimeError(
                                f"{amazonSiteName} 商城筛选中没有“所有商城”选项"
                            )
                        dropdownStores.click()
                        time.sleep(1)
                        storeOption.click()
                        time.sleep(1)
                        if dropdownStores.attr("value") != "ALL_STORES":
                            raise RuntimeError(
                                f"{amazonSiteName} 商城筛选未切换到所有商城"
                            )
                else:
                    print(f"{amazonSiteName} 没有商城筛选，按 2 个筛选项处理")

                # 账户类型优先选择全部；站点没有该选项时保留当前有效类型。
                dropdownAccount = page.ele(
                    'x://form//div[contains(@class,"account-type-selection")]/kat-dropdown',
                    timeout=2,
                )
                if not dropdownAccount:
                    dropdownAccount = filterDropdowns[-2]
                accountValue = dropdownAccount.attr("value") or ""
                accountClickable = dropdownAccount.states.is_clickable
                accountAllOption = dropdownAccount.shadow_root(
                    'x://kat-option[@value="ALL"]'
                    ' | //kat-option[starts-with(normalize-space(),"全部") '
                    'or starts-with(normalize-space(),"All")]'
                    ' | //kat-option[.//*[starts-with(normalize-space(),"全部") '
                    'or starts-with(normalize-space(),"All")]]',
                    timeout=2,
                )
                print(
                    f"账户类型筛选: 当前值={accountValue or '空'}，"
                    f"可点击={accountClickable}，"
                    f"包含全部选项={bool(accountAllOption)}"
                )
                if accountValue != "ALL" and accountAllOption:
                    if not accountClickable:
                        raise RuntimeError(
                            f"{amazonSiteName} 账户类型筛选不可点击，"
                            f"当前值为 {accountValue or '空'}"
                        )
                    dropdownAccount.click()
                    time.sleep(1)
                    accountAllOption.click()
                    time.sleep(1)
                    if dropdownAccount.attr("value") != "ALL":
                        raise RuntimeError(
                            f"{amazonSiteName} 账户类型未切换到全部"
                        )
                elif accountValue != "ALL":
                    if not accountValue:
                        raise RuntimeError(
                            f"{amazonSiteName} 账户类型没有当前值，也没有全部选项"
                        )
                    print(
                        f"{amazonSiteName} 不提供全部账户类型，"
                        f"保留当前值: {accountValue}"
                    )

                # 报告类型必须为交易；当前已正确时不重复点击。
                dropdownReport = page.ele(
                    'x://form//div[contains(@class,"report-type-selection")]/kat-dropdown',
                    timeout=2,
                )
                if not dropdownReport:
                    dropdownReport = filterDropdowns[-1]
                reportValue = dropdownReport.attr("value") or ""
                reportClickable = dropdownReport.states.is_clickable
                print(
                    f"报告类型筛选: 当前值={reportValue or '空'}，"
                    f"可点击={reportClickable}"
                )
                if reportValue != "SELLER_TRANSACTION_DATE_RANGE":
                    if not reportClickable:
                        raise RuntimeError(
                            f"{amazonSiteName} 报告类型筛选不可点击，"
                            f"当前值为 {reportValue or '空'}"
                        )
                    reportOption = dropdownReport.shadow_root(
                        'x://kat-option[@value="SELLER_TRANSACTION_DATE_RANGE"]'
                        ' | //kat-option[normalize-space()="交易" '
                        'or normalize-space()="Transaction"]'
                        ' | //kat-option[.//*[normalize-space()="交易" '
                        'or normalize-space()="Transaction"]]',
                        timeout=3,
                    )
                    if not reportOption:
                        raise RuntimeError(
                            f"{amazonSiteName} 报告类型中没有“交易”选项"
                        )
                    dropdownReport.click()
                    time.sleep(1)
                    reportOption.click()
                    time.sleep(1)
                    if dropdownReport.attr("value") != "SELLER_TRANSACTION_DATE_RANGE":
                        raise RuntimeError(
                            f"{amazonSiteName} 报告类型未切换到交易"
                        )

                page.ele(
                    'x://div/kat-radiobutton[@label="月" or @label="月份" or @label="Month"]',
                    timeout=20,
                ).click(by_js=True)
                time.sleep(3)

                # Amazon 月份 value 从 0 开始，目标月份需要减 1。
                dropdownMonth = page.ele('x://div[@class="date-range-item"][1]/kat-dropdown')
                dropdownMonth.click()
                dropdownMonth.shadow_root(
                    f'x://kat-option[@value="{monthValue}"]'
                    f' | //kat-option//div[normalize-space()="{monthEnglishName}" '
                    f'or normalize-space()="{monthChineseName}" '
                    f'or normalize-space()="{month}月"]',
                    timeout=20,
                ).click()
                time.sleep(1)
                if dropdownMonth.attr("value") != monthValue:
                    raise RuntimeError(
                        f"报告月份选择失败，目标月份为 {year}年{month}月，"
                        f"页面月份值为 {dropdownMonth.attr('value')}"
                    )
                print(f"报告月份已确认: {year}年{month}月")
                time.sleep(2)

                # 同步选择目标年份，确保 1 月运行时能正确选择上一年 12 月
                dropdownYear = page.ele(
                    'x://div[@class="date-range-item"][2]/kat-dropdown',
                    timeout=20,
                )
                dropdownYear.click()
                dropdownYear.shadow_root(
                    f'x://kat-option[@value="{year}"]'
                    f' | //kat-option//div[normalize-space()="{year}"]',
                    timeout=20,
                ).click()
                time.sleep(3)

                page.ele(
                    'x://kat-button[@label="请求报告" or @label="申请报告" or '
                    'translate(@label, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", '
                    '"abcdefghijklmnopqrstuvwxyz")="request report"]',
                    timeout=20,
                ).click()
                time.sleep(3)
                print(f"Amazon 后台站点 {amazonSiteName} 的 Transaction 报告请求完成")



    def run(self):
        """入口：执行 main()；若启用 sendEmail 则发送报告邮件。"""
        env = "线上" if self.isOnline else "线下"
        print(f"运行环境：{env}")
        self.main()
        if not self.sendEmail:
            print("未启用邮件发送，流程结束")
            return

        outputFiles = self.collectOutputFiles()
        deliver_outputs(
            {
                "sendEmail": True,
                "email": self.email,
                "sender_email": self.senderEmail,
                "smtp_auth_code": self.smtpAuthCode,
            },
            outputFiles,
        )


if __name__ == '__main__':
    # 使用 __init__ 中定义的默认初始值创建自动化任务实例。
    dev = TestPage()

    # 启动完整的 Transaction 报告请求流程。
    dev.run()
