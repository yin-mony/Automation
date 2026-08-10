"""亚马逊后台站点报告请求主流程。"""

import socket
import subprocess
import time
from datetime import datetime
from pathlib import Path

import psutil
from DrissionPage import Chromium

from YidekeLogin import YidekeLogin


siteMap = {
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

siteRegionMap = {
    "美国": "US",
    "加拿大": "CA",
    "墨西哥": "MX",
    "巴西": "BR",
    "英国": "UK",
    "法国": "FR",
    "德国": "DE",
    "意大利": "IT",
    "西班牙": "ES",
    "荷兰": "NL",
    "瑞典": "SE",
    "波兰": "PL",
    "比利时": "BE",
    "爱尔兰": "IE",
    "日本": "JP",
    "新加坡": "SG",
    "澳大利亚": "AU",
    "印度": "IN",
    "阿联酋": "AE",
    "沙特阿拉伯": "SA",
    "土耳其": "TR",
    "埃及": "EG",
    "南非": "ZA",
}

reportTypes = {
    "summary": {
        "label": "汇总报告",
        "optionValue": "SELLER_SUMMARY_DATE_RANGE",
    },
    "transaction": {
        "label": "Transaction 报告",
        "optionValue": "SELLER_TRANSACTION_DATE_RANGE",
    },
}

reportLabelToKey = {
    value["label"]: key
    for key, value in reportTypes.items()
}

siteModes = {
    "single": "单站点自主切换模式",
    "all": "全站点自动切换模式",
}

siteModeLabelToKey = {
    label: key
    for key, label in siteModes.items()
}

monthEnglish = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
}

monthChinese = {
    1: "一月",
    2: "二月",
    3: "三月",
    4: "四月",
    5: "五月",
    6: "六月",
    7: "七月",
    8: "八月",
    9: "九月",
    10: "十月",
    11: "十一月",
    12: "十二月",
}


class AmazonReport:
    """通过易得客店铺环境请求指定亚马逊后台站点报告。"""

    def __init__(self, config):
        """读取运行配置，报告请求以点击 Request Report 为完成边界。"""
        self.config = config
        self.username = str(config.get("yidekeUsername") or config.get("username") or "").strip()
        self.password = str(config.get("yidekePassword") or config.get("password") or "")
        self.autoSiteName = str(config.get("autoSiteName") or config.get("siteName") or "美国").strip()
        self.shopIps = self.ensureList(config.get("ip") or config.get("shopIp"))
        self.shopPorts = self.ensureList(config.get("port") or config.get("shopPort"))
        self.amazonEmail = str(config.get("amazonEmail") or "").strip()
        self.amazonPassword = str(config.get("amazonPassword") or "")
        self.isOnline = bool(config.get("isOnline", False))
        self.controlPort = int(config.get("controlPort") or 9222)
        self.reportType = self.resolveReportType(config.get("reportType") or "summary")
        self.reportOptionValue = reportTypes[self.reportType]["optionValue"]
        self.siteMode = self.resolveSiteMode(config.get("siteMode") or "single")
        self.amazonSiteNames = self.resolveSiteNames(config)
        self.currentSiteName = ""
        self.currentSiteEnglishName = ""

    def ensureList(self, value):
        """把单值、逗号文本或列表统一整理为列表。"""
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        text = str(value or "").replace("\n", ",")
        return [item.strip() for item in text.split(",") if item.strip()]

    def resolveReportType(self, value):
        """把 GUI 文案或内部 key 转成报告类型 key。"""
        text = str(value or "").strip()
        if text in reportTypes:
            return text
        if text in reportLabelToKey:
            return reportLabelToKey[text]
        raise ValueError(f"暂不支持该报告类型: {text}")

    def resolveSiteEnglishName(self, siteName):
        """把中文站点名映射为后台账号切换所需英文站点名。"""
        if siteName in siteMap:
            return siteMap[siteName]
        if siteName in siteMap.values():
            return siteName
        raise ValueError(f"暂不支持该亚马逊后台站点: {siteName}")

    def resolveSiteMode(self, value):
        """把 GUI 文案或内部 key 转成站点切换模式 key。"""
        text = str(value or "").strip()
        if text in siteModes:
            return text
        if text in siteModeLabelToKey:
            return siteModeLabelToKey[text]
        raise ValueError(f"暂不支持该站点切换模式: {text}")

    def resolveSiteNames(self, config):
        """根据站点切换模式整理要请求的后台站点列表。"""
        if self.siteMode == "all":
            siteNames = self.ensureList(config.get("amazonSiteNames") or config.get("amazonSiteName"))
            if not siteNames:
                siteNames = list(siteMap.keys())
        else:
            siteNames = [str(config.get("amazonSiteName") or "美国").strip()]

        result = []
        for siteName in siteNames:
            self.resolveSiteEnglishName(siteName)
            chineseName = siteName
            if siteName in siteMap.values():
                for key, value in siteMap.items():
                    if value == siteName:
                        chineseName = key
                        break
            if chineseName not in result:
                result.append(chineseName)
        return result

    def validate(self):
        """校验运行所需的基础配置。"""
        if not self.username:
            raise ValueError("请填写易得客账号")
        if not self.password:
            raise ValueError("请填写易得客密码")
        if not self.autoSiteName:
            raise ValueError("请选择店铺站点")
        self.resolveSiteEnglishName(self.autoSiteName)
        if not self.shopIps:
            raise ValueError("请填写店铺 IP")
        if not self.shopPorts:
            raise ValueError("请填写店铺端口")
        if not self.amazonSiteNames:
            raise ValueError("请至少选择一个亚马逊后台站点")
        try:
            self.shopPorts = [int(port) for port in self.shopPorts]
        except ValueError as exc:
            raise ValueError("店铺端口只能填写数字，多个端口用逗号或换行分隔") from exc
        if len(self.shopPorts) == 1 and len(self.shopIps) > 1:
            self.shopPorts = self.shopPorts * len(self.shopIps)
        if len(self.shopPorts) != len(self.shopIps):
            raise ValueError("端口数量需要和 IP 数量一致，或只填写一个端口")

    def killEdeckerOnPort(self, port):
        """启动店铺浏览器前，结束占用该调试端口的易得客进程。"""
        flag = f"--remote-debugging-port={port}"
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = proc.info.get("name") or ""
                cmdline = proc.info.get("cmdline") or []
                if name.lower() == "edecker.exe" and any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                pass

    def killEdecker(self, excludePid):
        """关闭除指定 PID 外的易得客进程。"""
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                name = proc.info.get("name") or ""
                pid = proc.info.get("pid")
                if name.lower() == "edecker.exe" and pid != excludePid:
                    proc.kill()
            except Exception:
                pass

    def waitForPort(self, port, timeout=60):
        """等待本机调试端口可用。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=2):
                    return
            except OSError:
                time.sleep(1)
        raise RuntimeError(f"等待 127.0.0.1:{port} 超时 ({timeout}s)")

    def runEdeckerAutomation(self, ips, port=9222):
        """在易得客管理端按店铺 IP 点击访问。"""
        browser = Chromium(port)
        tab = browser.latest_tab
        time.sleep(2)
        for ip in ips:
            print(f"易得客管理端访问店铺: {self.autoSiteName} / {ip}")
            self.clickFirst(
                tab,
                [
                    f'x://div[contains(@class,"platform-region")]//span[normalize-space()="{self.autoSiteName}"]'
                    f'/ancestor::div[contains(@class,"shop-item")]'
                    f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                    f'//button[normalize-space()="访问"]',
                    f'x://div[contains(@class,"platform-region")]//span[normalize-space()="{self.autoSiteName}"]'
                    f'/ancestor::div[contains(@class,"shop-item")]'
                    f'[.//*[normalize-space()="{ip}"]]'
                    f'//button[contains(normalize-space(),"访问")]',
                    f'x://div[contains(@class,"shop-item")][.//div[contains(@class,"text") and normalize-space()="{ip}"]]//button[contains(normalize-space(),"访问")]',
                    f'x://div[text()="{ip}"]//following-sibling::button',
                ],
                timeout=30,
                errorMessage=f"未找到店铺 IP={ip} 的访问按钮",
            )
            time.sleep(3)
        self.killEdecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def startEdecker(self, ip, port):
        """按店铺 IP 匹配易得客 profile，并用指定端口启动。"""
        base = Path.home() / "AppData/Local/eDecker6"
        exePath = base / "Application/edecker.exe"
        profilesPath = base / "Profiles"
        print("易得客 EXE:", exePath, exePath.exists())
        print("Profiles dir exists:", profilesPath.exists())
        if not exePath.exists():
            raise FileNotFoundError(f"找不到 exe: {exePath}")
        if not profilesPath.exists():
            raise FileNotFoundError(f"找不到 profiles 目录: {profilesPath}")

        ipDot = ip
        ipUnderline = ip.replace(".", "_")
        profiles = list(profilesPath.iterdir())
        candidates = [
            profile for profile in profiles
            if profile.is_dir() and (ipDot in profile.name or ipUnderline in profile.name)
        ]
        if not candidates:
            raise RuntimeError(f"未找到 IP={ip} 的易得客 profile")
        latest = max(candidates, key=lambda profile: profile.stat().st_mtime)
        profileParts = latest.name.split("_")
        shopId = profileParts[-1] if profileParts else ""
        region = siteRegionMap.get(self.autoSiteName, "US")
        redirectUrl = "https://sellercentral.amazon.com/gp/homepage.html"
        customAction = (
            "chrome-extension://cbakaehmphdgknbdbpfgejbiiehadcfe/bundle/loading/loading.html"
            f"?shopId={shopId}&region={region}&redirectUrl={redirectUrl}"
        )
        cmd = [
            str(exePath),
            f"--user-data-dir={latest}",
            "--no-sandbox",
            f"--scope={latest.name}",
            f"--custom_action={customAction}",
            f"--remote-debugging-port={port}",
        ]
        print("启动店铺浏览器:", " ".join(cmd))
        subprocess.Popen(cmd, cwd=str(base))

    def findSellerTab(self, port, timeout=90):
        """在指定调试端口查找卖家后台标签页。"""
        browser = Chromium(port)
        deadline = time.time() + timeout
        loadingError = ""
        while time.time() < deadline:
            for tabId in browser.tab_ids:
                tab = browser.get_tab(tabId)
                url = (tab.url or "").lower()
                print("检测 tab URL:", url)
                if url.startswith("chrome-extension://"):
                    if "loading/loading.html" in url:
                        try:
                            loadingText = tab.run_js('return document.body ? document.body.innerText : ""') or ""
                        except Exception:
                            loadingText = ""
                        if "检测失败" in loadingText or "当前环境存在异常" in loadingText:
                            loadingError = " ".join(loadingText.split())
                            raise RuntimeError(f"易得客店铺环境检测失败：{loadingError}")
                        try:
                            visitButton = tab.ele('x://button[.//span[normalize-space()="访问账号"]]', timeout=1)
                            buttonClass = visitButton.attr("class") or ""
                            if visitButton and "is-disabled" not in buttonClass:
                                visitButton.click()
                                time.sleep(5)
                        except Exception:
                            pass
                    continue
                if any(key in url for key in ("seller", "amazon", "tiktok", "tiktokglobalshop")):
                    return tab
            time.sleep(1)
        if loadingError:
            raise RuntimeError(f"易得客店铺环境检测失败：{loadingError}")
        raise RuntimeError(f"未找到卖家后台标签页，端口 {port}")

    def clickFirst(self, page, selectors, timeout=5, errorMessage=None):
        """按顺序尝试多个 XPath/CSS 选择器并点击第一个可用元素。"""
        lastError = None
        for selector in selectors:
            try:
                ele = page.ele(selector, timeout=timeout)
                if ele:
                    ele.click()
                    return ele
            except Exception as exc:
                lastError = exc
        if errorMessage:
            raise RuntimeError(errorMessage) from lastError
        return None

    def switchLanguageChinese(self, page):
        """优先切换后台页面语言为中文。"""
        print("优先切换后台页面语言为中文")
        try:
            languageButton = self.clickFirst(
                page,
                [
                    'x://div[@aria-label="语言"]',
                    'x://div[@aria-label="Language"]',
                    'x://button[contains(@aria-label,"语言")]',
                    'x://button[contains(@aria-label,"Language")]',
                ],
                timeout=5,
            )
            if not languageButton:
                print("未找到语言切换入口，继续后续流程")
                return
            time.sleep(1)
            option = self.clickFirst(
                page,
                [
                    'x://div[normalize-space()="简体中文"]',
                    'x://span[normalize-space()="简体中文"]',
                    'x://div[normalize-space()="中文(简体)"]',
                    'x://span[normalize-space()="中文(简体)"]',
                    'x://div[normalize-space()="中文"]',
                    'x://span[normalize-space()="中文"]',
                    'x://div[contains(normalize-space(),"Chinese")]',
                ],
                timeout=5,
            )
            if option:
                time.sleep(5)
                print("已切换或确认中文页面语言")
            else:
                print("未找到中文语言选项，继续后续流程")
        except Exception as exc:
            print(f"切换中文语言失败，继续后续流程: {exc}")

    def loginAmazonBackend(self, page):
        """接入 FBA 通用 Amazon Seller Central 登录逻辑。"""
        loginSiteName = self.amazonSiteNames[0] if self.amazonSiteNames else self.autoSiteName
        loginSiteEnglishName = self.resolveSiteEnglishName(loginSiteName)
        amazonEmail = self.amazonEmail or None
        amazonPassword = self.amazonPassword or None
        if not amazonEmail and not amazonPassword:
            print("未填写 Amazon 邮箱/密码，将尝试使用浏览器已保存登录态或人工确认")
        print(f"准备执行 Amazon 后台登录/确认: {loginSiteName} / {loginSiteEnglishName}")
        page = YidekeLogin.AmazonSeller(
            page,
            storePassword=amazonPassword,
            siteName=loginSiteEnglishName,
        ).login(
            amazonEmail,
            amazonPassword,
        )
        print(f"Amazon 登录后 URL: {page.url}")
        return page

    def enterReportsRepository(self, page):
        """通过汉堡菜单进入 Reports Repository。"""
        print("进入 Payments / Reports Repository")
        menuHost = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=30)
        menu = menuHost.shadow_root
        menu.ele('x://div/img', timeout=30).click()
        time.sleep(1)

        for labels in (["Payments", "付款", "款项", "支付"], ["Reports Repository", "报告存储库", "报告库", "报告"]):
            selectors = []
            for label in labels:
                selectors.append(f'x://div/span[normalize-space()="{label}"]')
                selectors.append(f'x://span[contains(normalize-space(),"{label}")]')
                selectors.append(f'x://div[contains(normalize-space(),"{label}")]')
            lastError = None
            clicked = None
            for selector in selectors:
                try:
                    clicked = menu.ele(selector, timeout=20)
                    if clicked:
                        clicked.click()
                        break
                except Exception as exc:
                    lastError = exc
            if not clicked:
                raise RuntimeError(f"未找到菜单项: {'/'.join(labels)}") from lastError
            time.sleep(3)
            page.wait(1)
        page.wait.load_start()

    def selectKatOption(self, page, optionValue=None, optionTexts=None, required=False):
        """在页面 kat-dropdown 中选择指定 value 或文本的选项。"""
        optionTexts = optionTexts or []
        dropdowns = page.eles('x://form//kat-dropdown', timeout=20)
        if not dropdowns:
            dropdowns = page.eles('x://kat-dropdown', timeout=10)
        lastError = None
        for dropdown in dropdowns:
            try:
                dropdown.click()
                time.sleep(1)
                shadow = dropdown.shadow_root
                if optionValue:
                    try:
                        option = shadow(f'x://kat-option[@value="{optionValue}"]', timeout=2)
                        if option:
                            option.click()
                            time.sleep(1.5)
                            return True
                    except Exception as exc:
                        lastError = exc
                for text in optionTexts:
                    try:
                        option = shadow(
                            f'x://kat-option//*[contains(normalize-space(),"{text}")]/ancestor::kat-option[1]'
                            f' | //kat-option[contains(normalize-space(),"{text}")]',
                            timeout=2,
                        )
                        if option:
                            option.click()
                            time.sleep(1.5)
                            return True
                    except Exception as exc:
                        lastError = exc
            except Exception as exc:
                lastError = exc
        if required:
            target = optionValue or "/".join(optionTexts)
            raise RuntimeError(f"未找到下拉选项: {target}") from lastError
        return False

    def processSite(self, page, siteName):
        """按顺序处理当前浏览器中的单个后台站点报告请求。"""
        self.currentSiteName = siteName
        self.currentSiteEnglishName = self.resolveSiteEnglishName(siteName)
        print(f"准备请求站点: {self.currentSiteName} / {self.currentSiteEnglishName}")

        # 第一步：打开顶部账号/站点切换区域并切换目标站点
        print(f"切换 Amazon 后台站点: {self.currentSiteName} / {self.currentSiteEnglishName}")
        aliases = [self.currentSiteEnglishName]
        if self.currentSiteName not in aliases:
            aliases.append(self.currentSiteName)

        allSiteNames = []
        for chineseName, englishName in siteMap.items():
            if chineseName not in allSiteNames:
                allSiteNames.append(chineseName)
            if englishName not in allSiteNames:
                allSiteNames.append(englishName)
        siteButtonCondition = " or ".join([
            f'contains(normalize-space(), "{name}")'
            for name in allSiteNames
        ])
        switchEntry = None
        for selector in [
            f'x://button[{siteButtonCondition}]',
            f'x://span[{siteButtonCondition}]/ancestor::button[1]',
            f'x://span[{siteButtonCondition}]/ancestor::div[contains(@class,"picker") or contains(@class,"switch") or contains(@class,"selector")][1]',
            f'x://span[{siteButtonCondition}]/ancestor::div[1]',
            'x://span[contains(text(),"Amazon(")]',
            'x://span[contains(text(),"Amazon（")]',
        ]:
            try:
                switchEntry = page.ele(selector, timeout=3)
            except Exception:
                switchEntry = None
            if switchEntry:
                break
        if not switchEntry:
            currentUrl = (page.url or "").lower()
            if self.currentSiteEnglishName == "United States" and "sellercentral.amazon.com" in currentUrl:
                print("当前页面已是美国站 Seller Central，未找到切换入口时跳过站点切换")
            else:
                raise RuntimeError("未找到 Amazon 后台账号/站点切换入口")

        if switchEntry:
            switchText = " ".join(str(switchEntry.text or "").split())
            if any(alias in switchText for alias in aliases):
                print(f"Amazon 后台已是目标站点: {self.currentSiteName}")
            else:
                switchEntry.click()
                time.sleep(1)
                seeAll = self.clickFirst(
                    page,
                    [
                        'x://*[normalize-space()="See all"]',
                        'x://*[contains(normalize-space(),"See all")]',
                        'x://*[normalize-space()="查看所有"]',
                        'x://*[contains(normalize-space(),"查看所有")]',
                        'x://*[normalize-space()="查看全部"]',
                        'x://*[contains(normalize-space(),"全部")]',
                    ],
                    timeout=3,
                )
                if not seeAll:
                    try:
                        switchEntry.click(by_js=True)
                    except Exception:
                        switchEntry.click()
                    time.sleep(1)
                    self.clickFirst(
                        page,
                        [
                            'x://*[normalize-space()="See all"]',
                            'x://*[contains(normalize-space(),"See all")]',
                            'x://*[normalize-space()="查看所有"]',
                            'x://*[contains(normalize-space(),"查看所有")]',
                            'x://*[normalize-space()="查看全部"]',
                            'x://*[contains(normalize-space(),"全部")]',
                        ],
                        timeout=20,
                        errorMessage="未找到 See all/查看全部 入口",
                    )
                time.sleep(1)

                siteSelectors = []
                for alias in aliases:
                    siteSelectors.append(f'x://*[normalize-space()="{alias}"]')
                    siteSelectors.append(f'x://*[contains(normalize-space(),"{alias}（")]')
                    siteSelectors.append(f'x://*[contains(normalize-space(),"{alias} (")]')
                    siteSelectors.append(f'x://span[contains(normalize-space(),"{alias}")]')
                    siteSelectors.append(f'x://*[contains(normalize-space(),"{alias}")]')
                self.clickFirst(
                    page,
                    siteSelectors,
                    timeout=20,
                    errorMessage=f"未找到目标亚马逊后台站点: {self.currentSiteEnglishName}",
                )
                time.sleep(1)
                self.clickFirst(
                    page,
                    [
                        'x://kat-button[@label="Select account"]',
                        'x://kat-button[contains(@label,"选择")]',
                        'x://button[contains(normalize-space(),"Select account")]',
                        'x://button[contains(normalize-space(),"选择账户")]',
                        'x://button[contains(normalize-space(),"选择账号")]',
                    ],
                    timeout=20,
                    errorMessage="未找到 Select account/选择账户 按钮",
                )
                time.sleep(5)

        # 第二步：进入 Reports Repository
        self.enterReportsRepository(page)

        # 第三步：选择报告类型、当前月份并点击 Request Report
        print(f"选择报告类型: {reportTypes[self.reportType]['label']}")
        self.selectKatOption(page, optionValue="ALL_STORES", required=False)
        self.selectKatOption(page, optionValue="ALL", required=False)
        self.selectKatOption(page, optionValue=self.reportOptionValue, required=True)
        self.clickFirst(
            page,
            [
                'x://kat-radiobutton[@label="Month"]',
                'x://kat-radiobutton[contains(@label,"月")]',
                'x://div/kat-radiobutton[@label="Month"]',
            ],
            timeout=20,
            errorMessage="未找到按月选项",
        )
        time.sleep(1.5)

        month = datetime.now().month
        self.selectKatOption(
            page,
            optionTexts=[
                monthEnglish[month],
                monthChinese[month],
                f"{month}月",
            ],
            required=True,
        )
        self.clickFirst(
            page,
            [
                'x://kat-button[translate(@label, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")="request report"]',
                'x://kat-button[contains(@label,"请求") or contains(@label,"申请")]',
                'x://button[contains(normalize-space(),"Request Report")]',
                'x://button[contains(normalize-space(),"请求报告")]',
                'x://button[contains(normalize-space(),"申请报告")]',
            ],
            timeout=20,
            errorMessage="未找到 Request Report/请求报告 按钮",
        )
        time.sleep(3)
        print("已点击 Request Report，请求报告流程完成")

    def processShop(self, ip, port):
        """处理单个店铺浏览器中的全部目标站点报告请求。"""
        self.killEdeckerOnPort(port)
        time.sleep(1)
        self.startEdecker(ip, port)
        self.waitForPort(port)
        page = self.findSellerTab(port)
        try:
            page.set.window.max()
        except RuntimeError:
            pass
        print("当前 tab URL:", page.url)
        page = self.loginAmazonBackend(page)
        self.switchLanguageChinese(page)
        for siteName in self.amazonSiteNames:
            self.processSite(page, siteName)

    def main(self):
        """完整自动化：登录易得客、打开店铺、切中文和站点、请求报告。"""
        self.validate()
        if self.isOnline:
            print("线上模式：登录易得客管理端并批量访问店铺")
            yideke = YidekeLogin(self.username, self.password)
            time.sleep(3)
            yideke.login()
            time.sleep(3)
            self.runEdeckerAutomation(self.shopIps, self.controlPort)
            time.sleep(4)
        else:
            print("线下模式：跳过易得客管理端登录，直接启动本机店铺 profile")
        for index, ip in enumerate(self.shopIps):
            self.processShop(ip, self.shopPorts[index])

    def run(self):
        """流程入口。"""
        print(f"报告类型：{reportTypes[self.reportType]['label']}")
        print(f"站点切换模式：{siteModes[self.siteMode]}")
        print(f"店铺站点：{self.autoSiteName}")
        print(f"目标站点：{', '.join(self.amazonSiteNames)}")
        self.main()


if __name__ == "__main__":
    config = {
        "username": "",
        "password": "",
        "autoSiteName": "美国",
        "ip": [""],
        "port": [8888],
        "amazonEmail": "",
        "amazonPassword": "",
        "siteMode": "single",
        "amazonSiteName": "美国",
        "reportType": "summary",
        "isOnline": False,
    }
    AmazonReport(config).run()
