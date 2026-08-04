"""易得客与 Amazon 后台货件索赔自动化。"""

import json
import re
import socket
import subprocess
import time
from datetime import datetime
from pathlib import Path

import psutil
from DrissionPage import Chromium, ChromiumPage

from YidekeLogin import Specification
from email_util import deliverCase
from export import PopExport
from wechat import Wechat


class Auto:
    """易得客进店、Amazon Seller Central 货件详情处理与凭证上传"""

    def __init__(self, config):
        # 运行配置
        self.config = config
        # 易得客登录信息
        self.yidekeUsername = config.get("yidekeUsername") or config.get("yideke_username") or ""
        self.yidekePassword = config.get("yidekePassword") or config.get("yideke_password") or ""
        # 店铺站点只用于易得客区域进店
        self.siteName = str(config.get("autoSiteName") or config.get("siteName") or "美国").strip()
        # Amazon 后台站点只用于 Seller Central 账号选择和站点切换
        self.amazonSiteName = str(config.get("amazonSiteName") or self.siteName).strip()
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
        self.siteEnglishName = self.siteEnglishMap.get(self.amazonSiteName, self.amazonSiteName)
        self.shipmentJsonName = "shipment_ids.json"
        # 店铺 IP 与调试端口
        shopIp = config.get("shopIp") or config.get("shop_ip") or config.get("ip")
        self.shopIp = shopIp if isinstance(shopIp, list) else [shopIp]
        self.shopIp = [str(item).strip() for item in self.shopIp if str(item or "").strip()]
        shopPort = config.get("shopPort") or config.get("shop_port") or config.get("port") or 9222
        if isinstance(shopPort, list):
            self.shopPort = [int(item) for item in shopPort]
        else:
            self.shopPort = [int(shopPort)]
        # 易得客主程序调试端口
        self.controlPort = int(config.get("controlPort") or 9222)
        # Amazon 登录信息
        self.amazonEmail = config.get("amazonEmail") or config.get("amazon_email") or ""
        self.amazonPassword = config.get("amazonPassword") or config.get("amazon_password") or ""
        # 项目资源与固定库存所有权证明文件
        self.baseDir = Path(config.get("baseDir") or PopExport.getBaseDir())
        resourceDir = PopExport.getResourceDir()
        self.podAwd = resourceDir / "AWD_POD.pdf"
        self.podFba = resourceDir / "FBA_POD.pdf"
        # 已完成导出的 POP PDF 存放目录
        popDir = config.get("popDir") or config.get("pop_dir") or ""
        self.popDir = Path(str(popDir)) if str(popDir or "").strip() else None
        # 企业微信汇总通知配置
        self.sendWechat = str(config.get("sendWechat") or "").strip().lower() in {"1", "true", "yes", "y", "是"}
        self.wechatWebhook = str(config.get("wechatWebhook") or "").strip()
        self.wechatMobile = str(config.get("wechatMobile") or "").strip()
        # 邮件汇总通知配置，易得客流程用于发送 CASE 结果与 POP 附件
        self.sendEmail = str(config.get("sendEmail") or "").strip().lower() in {"1", "true", "yes", "y", "是"}
        self.email = str(config.get("email") or "").strip()
        # 当前接管页面
        self.page = config.get("page")

    def killEdecker(self, excludePid):
        """关闭除指定进程外的易得客进程"""
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                pid = proc.info["pid"]
                name = proc.info["name"]
                # 只处理易得客进程
                if name and name.lower() == "edecker.exe" and pid != excludePid:
                    proc.kill()
                    # 已关闭非当前控制的易得客进程
            except Exception:
                # 进程已退出或权限不足时跳过
                pass

    def killEdeckerPort(self, port):
        """按远程调试端口关闭易得客进程"""
        flag = f"--remote-debugging-port={port}"
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = proc.info["name"]
                cmdline = proc.info.get("cmdline") or []
                # 匹配目标调试端口后关闭进程
                if name and name.lower() == "edecker.exe" and any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                # 进程状态不可读时跳过
                pass

    def waitPort(self, port, timeout=60):
        """等待本机调试端口可连接"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=2):
                    # 端口已可连接
                    return
            except OSError:
                # 端口未就绪，继续等待
                time.sleep(1)
        raise RuntimeError(f"等待 127.0.0.1:{port} 超时 ({timeout}s)")

    def visitShop(self, ips, port=9222):
        """在易得客主窗口中按站点访问指定店铺"""
        browser = Chromium(port)
        tab = browser.latest_tab
        time.sleep(2)
        for ip in ips:
            # 按配置站点区域和店铺 IP 找到访问按钮
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="{self.siteName}"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30,
            ).click()
            time.sleep(3)
        # 保留当前易得客主控进程，关闭其他同名进程
        self.killEdecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def startEdecker(self, ip, port):
        """按店铺 profile 启动易得客浏览器"""
        base = Path.home() / "AppData/Local/eDecker6"
        exePath = base / "Application/edecker.exe"
        profilesPath = base / "Profiles"

        # 校验易得客安装路径
        if not exePath.exists():
            raise FileNotFoundError(f"找不到 exe: {exePath}")
        # 校验店铺 profile 目录
        if not profilesPath.exists():
            raise FileNotFoundError(f"找不到 profiles 目录: {profilesPath}")

        ipDot = ip
        ipUnderline = ip.replace(".", "_")
        candidates = [
            item for item in profilesPath.iterdir()
            if item.is_dir() and (ipDot in item.name or ipUnderline in item.name)
        ]
        # 未找到 profile 时停止流程
        if not candidates:
            raise RuntimeError(f"未找到 IP={ip} 的 profile")

        latest = max(candidates, key=lambda item: item.stat().st_mtime)
        print(f"使用 profile: {latest}", flush=True)

        cmd = [
            str(exePath),
            f"--user-data-dir={latest}",
            "--no-sandbox",
            f"--remote-debugging-port={port}",
        ]
        subprocess.Popen(cmd, cwd=str(base))
        # 已按目标 profile 启动易得客浏览器

    def openShop(self):
        """登录易得客、访问店铺并接管 Amazon 卖家后台"""
        sp = Specification(self.yidekeUsername, self.yidekePassword)
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        # 在易得客中访问配置店铺
        self.visitShop(self.shopIp, self.controlPort)
        time.sleep(4)

        ip = self.shopIp[0]
        port = self.shopPort[0]
        # 关闭同端口旧浏览器后重新启动目标店铺 profile
        self.killEdeckerPort(port)
        time.sleep(1)
        self.startEdecker(ip, port)
        self.waitPort(port)
        self.page = ChromiumPage(f"127.0.0.1:{port}")
        try:
            self.page.set.window.max()
            # 已最大化店铺浏览器窗口
        except RuntimeError:
            # 最大化失败不影响登录流程
            pass
        print(f"已接管店铺浏览器，当前 URL: {self.page.url}", flush=True)

        amazonEmail = self.amazonEmail or None
        amazonPassword = self.amazonPassword or None
        # 登录 Amazon 卖家后台
        self.page = sp.amazonSellerLogin(
            self.page,
            amazonEmail,
            amazonPassword,
            siteEnglishName=self.siteEnglishName,
        )
        print(f"Amazon 登录后 URL: {self.page.url}", flush=True)
        return self.page

    def getSearchInput(self, page, searchSelectors, timeout=3):
        """按多个选择器查找 Amazon 货件编号搜索框"""
        for selector in searchSelectors:
            try:
                searchInput = page.ele(selector, timeout=timeout)
            except Exception:
                searchInput = None
            if searchInput:
                return searchInput
        return None

    def enterShipmentPage(self, page, searchSelectors):
        """通过 Amazon 菜单重新进入库存-货件页面，并返回货件编号搜索框"""
        lastError = ""
        for attempt in range(1, 3):
            try:
                # 打开 Amazon 左上角汉堡菜单
                menuHost = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=30)
                menu = menuHost.shadow_root
                menu.ele('x://div/img', timeout=30).click()
                time.sleep(1)

                # 先点击库存菜单，再点击货件入口
                inventoryBtn = None
                for text in ["库存", "Inventory"]:
                    try:
                        inventoryBtn = menu.ele(f'x://div/span[normalize-space()="{text}"]', timeout=5)
                    except Exception:
                        inventoryBtn = None
                    if inventoryBtn:
                        break
                if not inventoryBtn:
                    raise Exception("未找到库存菜单")
                inventoryBtn.click()
                time.sleep(3)
                page.wait(1)

                shipmentBtn = None
                for text in ["货件", "Shipments"]:
                    try:
                        shipmentBtn = menu.ele(f'x://div/span[normalize-space()="{text}"]', timeout=5)
                    except Exception:
                        shipmentBtn = None
                    if shipmentBtn:
                        break
                if not shipmentBtn:
                    raise Exception("未找到货件菜单")
                shipmentBtn.click()
                time.sleep(5)

                # 进入货件页后必须能找到搜索框
                searchInput = self.getSearchInput(page, searchSelectors, timeout=5)
                if searchInput:
                    print("已通过库存-货件菜单进入货件列表页", flush=True)
                    return searchInput
                raise Exception("进入货件页面后未找到货件编号搜索框")
            except Exception as exc:
                lastError = str(exc)
                print(f"第 {attempt} 次进入库存-货件失败：{lastError}", flush=True)
                time.sleep(2)
        raise Exception(f"重新进入库存-货件失败：{lastError}")

    def main(self):
        """在 Amazon 货件页面逐个处理货件并上传凭证"""
        page = self.page
        if not page:
            raise RuntimeError("未接管浏览器页面，无法执行易得客流程")

        # 先建立 POP PDF 文件映射，后续上传交货证明时精确匹配货件编号
        shipmentIds = []
        popFileMap = {}
        seen = set()
        if self.popDir and self.popDir.is_dir():
            for pdfPath in sorted(self.popDir.iterdir()):
                # 只读取最终导出的 PDF 文件
                if not pdfPath.is_file() or pdfPath.suffix.lower() != ".pdf":
                    continue
                match = re.search(r"(FBA[A-Z0-9]{6,})", pdfPath.stem, re.IGNORECASE)
                if not match:
                    continue
                shipmentId = match.group(1).upper()
                if shipmentId not in popFileMap:
                    popFileMap[shipmentId] = pdfPath

        # 优先读取赛狐本轮覆盖生成的 shipment_ids.json，避免旧 POP 文件混入处理数量
        jsonPath = self.popDir / self.shipmentJsonName if self.popDir else None
        if jsonPath and jsonPath.is_file():
            data = json.loads(jsonPath.read_text(encoding="utf-8"))
            if isinstance(data, list):
                rawIds = data
            elif isinstance(data, dict):
                if "shipmentIds" in data:
                    rawIds = data.get("shipmentIds") or []
                elif "allShipmentIds" in data:
                    rawIds = data.get("allShipmentIds") or []
                else:
                    rawIds = data.get("ids") or []
            else:
                rawIds = []
            for item in rawIds:
                shipmentId = str(item or "").strip().upper()
                if not re.fullmatch(r"FBA[A-Z0-9]{6,}", shipmentId):
                    continue
                if shipmentId in seen:
                    continue
                seen.add(shipmentId)
                shipmentIds.append(shipmentId)
            print(f"从本轮 shipment_ids.json 提取到 {len(shipmentIds)} 个货件编号: {jsonPath}", flush=True)
            if not shipmentIds:
                raise ValueError(f"本轮 shipment_ids.json 中没有可处理的货件编号: {jsonPath}")
        else:
            # 没有本轮 JSON 时才兜底使用 POP PDF 文件名
            for shipmentId in popFileMap:
                if shipmentId in seen:
                    continue
                seen.add(shipmentId)
                shipmentIds.append(shipmentId)
            if shipmentIds:
                print(f"未找到 shipment_ids.json，兜底从 POP 目录提取到 {len(shipmentIds)} 个货件编号", flush=True)
            else:
                print(f"POP 目录中未提取到货件编号: {self.popDir}", flush=True)
                raise ValueError("未找到可处理的货件编号")

        # 本轮易得客提交结果，流程结束后统一推送企业微信
        caseResults = []
        failResults = []
        skipResults = []

        # 先确认 Amazon 页面语言，已是中文时不重复切换
        chineseLang = page.ele('x://div[@aria-label="语言"] | //*[normalize-space()="ZH"]', timeout=3)
        chineseText = page.ele('x://*[contains(text(),"管理库存") or contains(text(),"帮助")]', timeout=2)
        if chineseLang or chineseText:
            print("Amazon 后台已是中文简体，跳过语言切换", flush=True)
        else:
            page.ele('x://div[@aria-label="Language"] | //div[@aria-label="语言"]', timeout=15).click()
            time.sleep(1.5)
            page.ele('x://div[text()="中文(简体)"] | //*[normalize-space()="中文(简体)"]', timeout=10).click()
            time.sleep(5)
            print("已切换为中文简体", flush=True)

        # 语言确认后再切换 Amazon 后台站点，已是目标站点时不重复切换
        amazonSiteName = self.amazonSiteName
        allSiteNames = list(self.siteEnglishMap.keys())
        siteButtonCondition = " or ".join([f'contains(normalize-space(), "{siteName}")' for siteName in allSiteNames])
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
        if amazonSiteName in switchText:
            print(f"Amazon 后台已是目标站点: {amazonSiteName}", flush=True)
        else:
            switchEntry.click()
            time.sleep(1)
            seeAll = page.ele('x://*[normalize-space()="查看所有"] | //*[normalize-space()="See all"]', timeout=3)
            if not seeAll:
                switchEntry.click(by_js=True)
                time.sleep(1)
                seeAll = page.ele('x://*[normalize-space()="查看所有"] | //*[normalize-space()="See all"]', timeout=20)
            seeAll.click()
            time.sleep(1)
            page.ele(
                f'x://*[normalize-space()="{amazonSiteName}" '
                f'or contains(normalize-space(), "{amazonSiteName}（") '
                f'or contains(normalize-space(), "{amazonSiteName} (")]',
                timeout=20,
            ).click()
            time.sleep(1)
            page.ele(
                'x://kat-button[@label="选择账户"]'
                ' | //kat-button[@label="Select account"]'
                ' | //button[normalize-space()="选择账户"]'
                ' | //span[normalize-space()="选择账户"]/ancestor::button[1]',
                timeout=20,
            ).click()
            time.sleep(5)
            print(f"Amazon 后台已切换到目标站点: {amazonSiteName}", flush=True)

        # 通过汉堡菜单进入库存货件页面
        searchSelectors = [
            'x://input[@placeholder="按货件编号搜索"]',
            'x://input[contains(@placeholder,"货件编号")]',
            'x://input[contains(@placeholder,"Shipment") or contains(@placeholder,"shipment")]',
            'x://input[contains(@aria-label,"货件编号") or contains(@aria-label,"Shipment")]',
        ]
        searchInput = self.enterShipmentPage(page, searchSelectors)

        # 清空上一次货件筛选条件
        try:
            clearBtn = page.ele('x://span[text()="清除筛选条件"]', timeout=3)
        except Exception:
            clearBtn = None
        if clearBtn:
            clearBtn.click()
            time.sleep(1.5)

        firstFailReasons = {}
        retryShipmentIds = []
        pendingShipmentIds = list(shipmentIds)
        retryMode = False
        index = 0
        while pendingShipmentIds or (not retryMode and retryShipmentIds):
            # 第一轮全部结束后，再统一处理失败货件
            if not pendingShipmentIds and not retryMode and retryShipmentIds:
                pendingShipmentIds = list(retryShipmentIds)
                retryMode = True
                index = 0
                print(f"开始重新处理第一轮失败货件，共 {len(pendingShipmentIds)} 个", flush=True)
            shipmentId = pendingShipmentIds.pop(0)
            index += 1
            currentTotal = len(retryShipmentIds) if retryMode else len(shipmentIds)
            detailPage = None
            shouldCloseDetail = False
            detailOpenedInCurrentTab = False
            try:
                roundName = "重试" if retryMode else "处理"
                print(f"开始{roundName}货件 {index}/{currentTotal}：{shipmentId}", flush=True)

                # 搜索当前货件编号
                searchInput = self.getSearchInput(page, searchSelectors, timeout=5)
                if not searchInput:
                    print("未找到货件编号搜索框，重新走库存-货件菜单后重试", flush=True)
                    searchInput = self.enterShipmentPage(page, searchSelectors)
                if not searchInput:
                    raise Exception(f"没有找到货件编号搜索框，当前 URL: {page.url}")
                searchInput.input(f"{shipmentId}\n", clear=True)
                time.sleep(3)
                print(f"已搜索货件编号：{shipmentId}", flush=True)

                # 进入货件详情页
                shipmentRow = page.ele(f'x://tr[contains(., "{shipmentId}")]', timeout=15)
                print("已找到对应行", flush=True)
                beforeTabIds = set(page.browser.tab_ids)
                shipmentRow.ele('x:.//div[contains(@class, "awsui_t-left")]//a').click()
                print("已进入对应的货件编号详情页", flush=True)
                page.wait(3)
                afterTabIds = set(page.browser.tab_ids)
                newTabIds = [tabId for tabId in afterTabIds if tabId not in beforeTabIds]
                if newTabIds:
                    for tabId in newTabIds:
                        tab = page.browser.get_tab(tabId)
                        if shipmentId in (tab.url or ""):
                            detailPage = tab
                            break
                    if not detailPage:
                        detailPage = page.browser.get_tab(newTabIds[-1])
                    shouldCloseDetail = True
                else:
                    latestTab = page.browser.latest_tab
                    if latestTab and shipmentId in (latestTab.url or ""):
                        detailPage = latestTab
                        detailOpenedInCurrentTab = True
                    else:
                        detailPage = page
                        detailOpenedInCurrentTab = True
                detailPage.wait(3)
                print("已切换到新详情页", flush=True)

                print("当前详情页:", detailPage.url, flush=True)
                detailPage.wait(3)

                # 已提交过问题的货件直接记录已有 CASE，不重复上传与提交
                existingCaseId = ""
                caseText = ""
                caseEle = detailPage.ele('x://*[contains(text(),"问题编号")]', timeout=2)
                if caseEle:
                    caseText = caseEle.text
                if not caseText:
                    # 页面文本无法直接读取时，从 HTML 中兜底提取
                    caseText = re.sub(r"<[^>]+>", " ", detailPage.html)
                caseMatch = re.search(r"问题编号\s*[:：]?\s*(\d+)", caseText)
                if caseMatch:
                    existingCaseId = caseMatch.group(1)
                if existingCaseId:
                    popFile = popFileMap.get(shipmentId)
                    caseResults.append({
                        "shipmentId": shipmentId,
                        "caseId": existingCaseId,
                        "popFile": popFile.name if popFile else "",
                        "popPath": str(popFile) if popFile else "",
                        "status": "已有问题编号",
                    })
                    print(f"货件已存在 CASE 问题编号，跳过提交：{shipmentId}，{existingCaseId}", flush=True)
                    continue

                # 获取所需上传交货证明的分类
                selectSort = detailPage.ele('x://kat-box[contains(@class, "shipment-info-box")]', timeout=10)
                sortLabel = selectSort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-label', timeout=5)
                if sortLabel:
                    sort = sortLabel.attr("text")
                else:
                    sortLabel = selectSort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-link', timeout=5)
                    sort = sortLabel.attr("label")
                print(f"获取到的值: {sort}", flush=True)

                # 查看详情页的货件差值
                print(detailPage.html.find("shipmentContentsIframe"), flush=True)
                frame = detailPage.get_frame("@data-testid=shipmentContentsIframe", timeout=10)
                headers = frame.eles('x://kat-table-head//kat-table-cell[@role="columnheader"]')
                colIndex = None
                for headerIndex, header in enumerate(headers, start=1):
                    # 确认页面中存在差值列，真实行取值以后续商品行结构为准
                    headerText = header.text
                    if "差值" in headerText:
                        colIndex = headerIndex
                if colIndex is None:
                    raise Exception("没有找到差值列")
                print(f"差值所在列:{colIndex}", flush=True)

                # 读取商品明细总行数
                pageCountSpan = detailPage.ele('x://span[@id="page-item-count"]', timeout=5)
                totalText = pageCountSpan.text.strip()
                match = re.search(r"显示\s+\d+\s+到\s+(\d+)", totalText)
                if match:
                    totalRows = int(match.group(1))
                else:
                    totalRows = 0
                print(f"总共 {totalRows} 行商品", flush=True)

                hasDiff = False
                selectedRows = []
                productRows = frame.eles('x://kat-table-body//kat-table-row[.//kat-dropdown[@id="action-required"]]')
                if not productRows:
                    raise Exception("没有识别到可操作商品行")
                print(f"识别到 {len(productRows)} 行可操作商品", flush=True)

                # 已提交过问题的商品行不重复上传与提交，避免同一货件重复开 CASE
                submittedRowTexts = []
                for productRow in productRows:
                    productRowText = " ".join(productRow.text.split())
                    if "已提交问题" in productRowText:
                        submittedRowTexts.append(productRowText)
                if submittedRowTexts:
                    skipResults.append({
                        "shipmentId": shipmentId,
                        "reason": "页面商品行已提交问题",
                    })
                    print(f"货件商品行已提交问题，跳过重复操作：{shipmentId}", flush=True)
                    continue

                for rowIndex, productRow in enumerate(productRows, start=1):
                    # 商品行中会混入展开子行，因此以操作下拉框定位当前行差值
                    cells = productRow.eles('x:./kat-table-cell')
                    actionDropdown = productRow.ele('x:.//kat-dropdown[@id="action-required"]', timeout=5)
                    actionCellIndex = None
                    for cellIndex, cell in enumerate(cells):
                        if cell.ele('x:.//kat-dropdown[@id="action-required"]', timeout=0):
                            actionCellIndex = cellIndex
                            break
                    if actionCellIndex is None or actionCellIndex <= 0:
                        raise Exception(f"第 {rowIndex} 行没有识别到差值所在单元格")

                    valueText = " ".join(cells[actionCellIndex - 1].text.split()).replace("+", "")
                    valueInt = int(valueText)
                    print(f"第 {rowIndex} 行差值: {valueInt}", flush=True)

                    # 差值为 0 时跳过当前行，不选择调查类型
                    if valueInt == 0:
                        print(f"第 {rowIndex} 行差值为0，不操作", flush=True)
                        continue

                    # 按差值正负选择调查类型
                    hasDiff = True
                    mskuText = ""
                    if cells:
                        mskuText = " ".join(cells[0].text.split())
                    actionDropdown.scroll.to_see()
                    time.sleep(0.5)
                    actionDropdown.click(by_js=True)
                    time.sleep(1)
                    shadow = actionDropdown.shadow_root
                    if valueInt < 0:
                        actionText = "调查缺失商品"
                        shadow.ele('x://*[(contains(text(),"调查缺失商品"))]', timeout=5).click(by_js=True)
                        print(f"第 {rowIndex} 行选择调查缺失商品", flush=True)
                    elif valueInt > 0:
                        actionText = "调查额外商品"
                        shadow.ele('x://*[(contains(text(),"调查额外商品"))]', timeout=5).click(by_js=True)
                        print(f"第 {rowIndex} 行选择调查额外商品", flush=True)
                    selectedRows.append({
                        "rowIndex": rowIndex,
                        "msku": mskuText,
                        "diff": valueInt,
                        "action": actionText,
                    })

                # 全部商品差值为 0 时跳过后续上传文件操作
                if not hasDiff:
                    skipResults.append({
                        "shipmentId": shipmentId,
                        "reason": "全部商品差值为0",
                    })
                    print(f"货件全部差值为0，已跳过上传：{shipmentId}", flush=True)
                    continue

                # 查找当前货件对应的 POP PDF，交货证明必须上传当前货件编号的 POP
                popFile = popFileMap.get(shipmentId)
                if not popFile and self.popDir and self.popDir.is_dir():
                    for pdfPath in sorted(self.popDir.iterdir()):
                        if not pdfPath.is_file() or pdfPath.suffix.lower() != ".pdf":
                            continue
                        if shipmentId in pdfPath.stem.upper():
                            popFile = pdfPath
                            popFileMap[shipmentId] = pdfPath
                            break
                if not popFile or not popFile.is_file():
                    raise FileNotFoundError(f"未找到货件 {shipmentId} 对应的 POP PDF 文件")
                if shipmentId not in popFile.stem.upper():
                    raise ValueError(f"POP 文件与当前货件编号不匹配: {shipmentId} -> {popFile.name}")

                # 根据货件创建来源选择库存所有权证明使用的内置 POD 文件
                if sort == "亚马逊分销":
                    podFile = self.podAwd
                else:
                    podFile = self.podFba
                if not podFile.is_file():
                    raise FileNotFoundError(f"库存所有权证明 POD 文件不存在: {podFile}")
                if "_POD" not in podFile.stem.upper():
                    raise ValueError(f"库存所有权证明文件命名异常: {podFile.name}")

                # 上传交货证明文件：当前货件编号对应的 POP PDF
                proofChooseBtn = detailPage.ele(
                    'x://div[contains(@class,"document-proof")]'
                    '[.//div[contains(text(),"交货证明")]]'
                    '//kat-button[@label="选择文件"]',
                    timeout=10,
                )
                if not proofChooseBtn:
                    raise Exception("没有找到交货证明选择文件按钮框")
                proofChooseBtn.click.to_upload(str(popFile))
                print(f"交货证明 POP 文件选择完成: {popFile.name}", flush=True)

                proofUploadBtn = detailPage.ele(
                    'x://div[contains(@class,"document-proof")]'
                    '[.//div[contains(text(),"交货证明")]]'
                    '//kat-button[@label="上传文档"]',
                    timeout=10,
                )
                if not proofUploadBtn:
                    raise Exception("没有找到交货证明上传文档按钮框")
                proofUploadBtn.click()
                print("交货证明文件已点击上传", flush=True)
                for waitIndex in range(30):
                    time.sleep(1)
                    disabled = proofUploadBtn.attr("disabled")
                    ariaDisabled = proofUploadBtn.attr("aria-disabled")
                    loading = proofUploadBtn.attr("loading")
                    className = str(proofUploadBtn.attr("class") or "").lower()
                    if disabled is not None or ariaDisabled == "true" or loading == "true" or "disabled" in className:
                        print("交货证明上传按钮已不可点击，视为上传完成", flush=True)
                        break
                else:
                    raise TimeoutError("交货证明上传后按钮未变为不可点击状态")

                # 上传库存所有权证明文件：内置 POD 文件
                ownershipChooseBtn = detailPage.ele(
                    'x://div[contains(@class,"document-proof")]'
                    '[.//div[contains(text(),"库存所有权证明")]]'
                    '//kat-button[@label="选择文件"]',
                    timeout=10,
                )
                if not ownershipChooseBtn:
                    raise Exception("没有找到库存所有权证明选择文件按钮框")
                ownershipChooseBtn.click.to_upload(str(podFile))
                print(f"库存所有权证明 POD 文件选择完成: {podFile.name}", flush=True)

                ownershipUploadBtn = detailPage.ele(
                    'x://div[contains(@class,"document-proof")]'
                    '[.//div[contains(text(),"库存所有权证明")]]'
                    '//kat-button[@label="上传文档"]',
                    timeout=10,
                )
                if not ownershipUploadBtn:
                    raise Exception("没有找到库存所有权证明上传文档按钮框")
                ownershipUploadBtn.click()
                print("库存所有权证明文件已点击上传", flush=True)
                for waitIndex in range(30):
                    time.sleep(1)
                    disabled = ownershipUploadBtn.attr("disabled")
                    ariaDisabled = ownershipUploadBtn.attr("aria-disabled")
                    loading = ownershipUploadBtn.attr("loading")
                    className = str(ownershipUploadBtn.attr("class") or "").lower()
                    if disabled is not None or ariaDisabled == "true" or loading == "true" or "disabled" in className:
                        print("库存所有权证明上传按钮已不可点击，视为上传完成", flush=True)
                        break
                else:
                    raise TimeoutError("库存所有权证明上传后按钮未变为不可点击状态")
                print(
                    f"提交前文件校验通过：交货证明={popFile.name}，库存所有权证明={podFile.name}",
                    flush=True,
                )

                # 进入预览请求
                detailPage.ele('x://kat-button[@label="预览您的请求"]', timeout=10).click()
                print(f"货件已进入预览请求：{shipmentId}", flush=True)

                # 预览内容在 iframe 中展示，读取表格内容确认与已选行一致
                previewFrame = detailPage.get_frame("@id=iframe-preview-reconcile-request", timeout=10)
                previewText = ""
                previewTexts = []
                for pageIndex in range(max(1, len(selectedRows))):
                    currentText = ""
                    for waitIndex in range(10):
                        time.sleep(1)
                        previewBody = previewFrame.ele("tag:body", timeout=2)
                        if previewBody:
                            currentText = " ".join(previewBody.text.split())
                        if not currentText:
                            currentText = " ".join(re.sub(r"<[^>]+>", " ", previewFrame.html).split())
                        if "MSKU" in currentText and "差值" in currentText and "需要操作" in currentText:
                            break
                    if currentText and currentText not in previewTexts:
                        previewTexts.append(currentText)

                    pagination = previewFrame.ele('x://kat-pagination[@id="preview-table-pagination"]', timeout=1)
                    if not pagination:
                        break
                    pageNum = int(pagination.attr("page") or 1)
                    pageSize = int(pagination.attr("items-per-page") or len(selectedRows) or 1)
                    totalItems = int(pagination.attr("total-items") or len(selectedRows) or 0)
                    if pageNum * pageSize >= totalItems:
                        break
                    nextBtn = pagination.shadow_root.ele('x://span[contains(@part,"pagination-nav-right")]', timeout=2)
                    if not nextBtn:
                        break
                    className = str(nextBtn.attr("class") or "").lower()
                    if "end" in className:
                        break
                    nextBtn.click(by_js=True)
                    for waitIndex in range(10):
                        time.sleep(0.5)
                        newPageNum = int(pagination.attr("page") or pageNum)
                        if newPageNum != pageNum:
                            break
                previewText = " ".join(previewTexts)
                if not previewText:
                    raise Exception("未读取到预览表格内容")

                # 校验预览表格中每个非 0 差值行的 MSKU、差值与操作类型
                for selectedRow in selectedRows:
                    mskuText = str(selectedRow.get("msku") or "").strip()
                    diffText = str(selectedRow.get("diff"))
                    actionText = str(selectedRow.get("action") or "").strip()
                    if mskuText and mskuText not in previewText:
                        raise Exception(f"预览内容缺少 MSKU：{mskuText}")
                    if diffText not in previewText:
                        raise Exception(f"预览内容缺少差值：{diffText}")
                    if actionText and actionText not in previewText:
                        raise Exception(f"预览内容缺少操作类型：{actionText}")
                print(f"预览内容校验通过：{shipmentId}", flush=True)

                # 点击预览弹窗内的提交按钮完成索赔提交
                submitBtn = previewFrame.ele(
                    'x://kat-button[@label="提交"]'
                    ' | //button[normalize-space()="提交"]'
                    ' | //span[normalize-space()="提交"]/ancestor::button[1]',
                    timeout=10,
                )
                if not submitBtn:
                    raise Exception("没有找到预览弹窗内提交按钮")
                submitBtn.click()
                print(f"货件已点击提交：{shipmentId}", flush=True)

                # 提交后读取页面提示中的 CASE 问题编号
                caseId = ""
                for waitIndex in range(15):
                    time.sleep(2)
                    caseText = ""
                    caseEle = detailPage.ele('x://*[contains(text(),"问题编号")]', timeout=2)
                    if caseEle:
                        caseText = caseEle.text
                    if not caseText:
                        # 页面文本无法直接读取时，从 HTML 中兜底提取
                        caseText = re.sub(r"<[^>]+>", " ", detailPage.html)
                    caseMatch = re.search(r"问题编号\s*[:：]?\s*(\d+)", caseText)
                    if caseMatch:
                        caseId = caseMatch.group(1)
                        break
                if not caseId:
                    raise Exception("提交后未读取到 CASE 问题编号")
                caseResults.append({
                    "shipmentId": shipmentId,
                    "caseId": caseId,
                    "popFile": popFile.name if popFile else "",
                    "popPath": str(popFile) if popFile else "",
                    "status": "提交成功",
                })
                print(f"货件提交成功，货件编号：{shipmentId}，CASE 问题编号：{caseId}", flush=True)
            except Exception as exc:
                # 第一轮失败只放入最后重试队列，重试失败才写入最终失败结果
                if retryMode:
                    failResults.append({
                        "shipmentId": shipmentId,
                        "reason": str(exc),
                        "firstReason": firstFailReasons.get(shipmentId, ""),
                    })
                    print(f"货件重试后仍失败，需要人工处理：{shipmentId}，{exc}", flush=True)
                else:
                    firstFailReasons[shipmentId] = str(exc)
                    if shipmentId not in retryShipmentIds:
                        retryShipmentIds.append(shipmentId)
                    print(f"货件处理失败，已加入最后重试队列：{shipmentId}，{exc}", flush=True)
                continue
            finally:
                if shouldCloseDetail and detailPage:
                    try:
                        detailPage.close()
                        print(f"已关闭货件详情页：{shipmentId}", flush=True)
                    except Exception as exc:
                        print(f"关闭货件详情页失败：{shipmentId}，{exc}", flush=True)
                    time.sleep(1)
                elif detailOpenedInCurrentTab and detailPage:
                    try:
                        page = detailPage
                        self.page = page
                        self.enterShipmentPage(page, searchSelectors)
                        print(f"已通过库存-货件菜单返回货件列表页：{shipmentId}", flush=True)
                    except Exception as exc:
                        print(f"返回货件列表页失败：{shipmentId}，{exc}", flush=True)

        saveDir = self.popDir if self.popDir and self.popDir.is_dir() else self.baseDir
        # 使用时间戳避免覆盖历史运行结果
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        failedPath = ""
        if failResults:
            failedFile = saveDir / f"failed_shipments_{timestamp}.json"
            failedData = {
                "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "message": "以下货件自动重试后仍失败，需要人工手动处理。",
                "failList": failResults,
            }
            failedFile.write_text(json.dumps(failedData, ensure_ascii=False, indent=2), encoding="utf-8")
            failedPath = str(failedFile)
            print(f"失败货件编号文件已保存，请人工处理: {failedFile}", flush=True)
        caseResultPath = saveDir / f"case_result_{timestamp}.json"
        caseData = {
            "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "resultList": caseResults,
            "failList": failResults,
            "skipList": skipResults,
            "failedPath": failedPath,
        }
        caseResultPath.write_text(json.dumps(caseData, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"CASE 汇总结果已保存: {caseResultPath}", flush=True)

        if self.sendEmail:
            try:
                deliverCase(
                    self.config,
                    caseResults,
                    failResults,
                    skipResults,
                    str(caseResultPath),
                )
                print("CASE 汇总邮件已发送", flush=True)
            except Exception as exc:
                # 邮件发送失败不影响本地 CASE 结果保存
                print(f"CASE 汇总邮件发送失败：{exc}", flush=True)

        if self.sendWechat:
            try:
                Wechat(self.config).sendCase({
                    "resultList": caseResults,
                    "failList": failResults,
                    "skipList": skipResults,
                    "email": self.email,
                    "wechatMobile": self.wechatMobile,
                    "caseResultPath": str(caseResultPath),
                })
                print("企业微信 CASE 汇总消息已发送", flush=True)
            except Exception as exc:
                # 企业微信发送失败不影响主流程结果
                print(f"企业微信 CASE 汇总消息发送失败：{exc}", flush=True)

    def run(self, detailResults=None):
        """按正式模式执行易得客流程"""
        print("正式模式，执行登录流程", flush=True)
        self.openShop()
        self.main()
        if detailResults:
            print(f"待处理货件数: {len(detailResults)}", flush=True)
        return self.page


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "yidekeUsername": "19167561839",
        "yidekePassword": "yxh643208yang",
        "autoSiteName": "美国",
        "amazonSiteName": "美国",
        "shopIp": "54.201.27.19",
        "shopPort": 8888,
        "amazonEmail": "happymike9@outlook.com",
        "amazonPassword": "Happylife989.",
        "baseDir": str(PopExport.getBaseDir()),
        "popDir": str(PopExport.getBaseDir() / "output"),
        "sendWechat": False,
        "wechatWebhook": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=b0049d97-c114-4b16-9434-ca6534a7e1f2",
        "wechatMobile": "",
        "email": "",
    }

    service = Auto(config)
    service.run()
