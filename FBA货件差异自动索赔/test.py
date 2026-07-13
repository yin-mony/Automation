from SaihuERPLogin import SaiHuERPLogin
from DrissionPage import ChromiumPage, Chromium
from pywinauto import Desktop
from YidekeLogin import Specification
from email_util import deliver_outputs
from export import PopExport
import socket
import subprocess
import psutil
import time
import re
import json
from pathlib import Path


class Test:
    def __init__(self,config):
        self.config = config
        self.page = config["page"]
        self.username = config["username"]
        self.password = config["password"]
        self.isOnline = bool(config.get("isOnline", False))
        self.baseDir = Path(config.get("baseDir") or PopExport.getBaseDir())
        # FBA 货件列表接口地址
        self.listApiUrl = "https://www.sellfox.com/api/inbound/shipmentCommodity/page.json"
        # FBA 货件详情接口地址
        self.detailApiUrl = "https://www.sellfox.com/api/inbound/shipment/detail.json"
        
        # 固定筛选条件字段
        # 接口行数据里的货件 ID 字段名
        self.shipmentIdKey = "amazonShipmentId"
        # 申收差异字段名
        self.diffKey = "quaRecDifference"
        # SKU 明细里的申收差异字段名
        self.itemDiffKey = "quantityRecDiff"
        # 默认筛选站点，从调试配置读取
        self.siteName = str(config.get("siteName") or "美国").strip()
        # 无下拉大区的站点，站点面板中直接点击
        self.directSiteNames = {
            "日本", "新加坡", "澳大利亚", "印度", "阿联酋",
            "沙特阿拉伯", "土耳其", "埃及", "南非",
        }
        # 有下拉大区的站点映射
        self.areaSiteMap = {
            "美国": "北美区",
            "加拿大": "北美区",
            "墨西哥": "北美区",
            "巴西": "北美区",
            "英国": "欧洲区",
            "法国": "欧洲区",
            "德国": "欧洲区",
            "意大利": "欧洲区",
            "西班牙": "欧洲区",
            "荷兰": "欧洲区",
            "瑞典": "欧洲区",
            "波兰": "欧洲区",
            "比利时": "欧洲区",
            "爱尔兰": "欧洲区",
        }
        # POP 文档导出
        self.popExport = PopExport(self.baseDir)
        exportDir = config.get("exportDir")
        if exportDir:
            self.popExport.exportDir = Path(str(exportDir)).resolve()
        self.generatedFiles = []

    def getShipmentIds(self, totalPages):
        """监听列表接口并翻页，提取申收差异大于 0 的 amazonShipmentId"""
        page = self.page
        allShipmentIds = []
        seenShipmentIds = set()

        for pageNum in range(1, totalPages + 1):
            # 第 2 页起先点页码触发请求
            if pageNum > 1:
                page.ele(f'x://ul[contains(@class,"el-pager")]//li[text()="{pageNum}"]', timeout=8).click()
                time.sleep(1)

            # 等待当前页接口响应
            packet = page.listen.wait(timeout=20, fit_count=False)
            if not packet:
                if pageNum == 1:
                    print("未监听到 FBA 货件列表接口")
                else:
                    print(f"第 {pageNum} 页未监听到接口，停止翻页")
                break

            # 取出响应体
            body = packet.response.body
            # 响应体为字符串时解析为 JSON
            if isinstance(body, str):
                body = json.loads(body)

            # 取 data 节点
            data = body.get("data", {})
            # 兼容 records / list / rows 三种列表字段
            rows = (
                data.get("records")
                or data.get("list")
                or data.get("rows")
                or []
            )

            # 只提取 quantityRecDiff 大于 0 的父级货件 ID，同一货件只保留一次
            pageIds = []
            currentShipmentId = ""
            for row in rows:
                if not isinstance(row, dict):
                    continue

                # 顶层行有货件编号时，记录为当前多 SKU 货件的父级编号
                rowShipmentId = str(row.get(self.shipmentIdKey) or "").strip()
                if rowShipmentId:
                    currentShipmentId = rowShipmentId

                # 子行可能没有货件编号，沿用最近的父级货件编号
                shipmentId = rowShipmentId or currentShipmentId
                checkRows = [row]
                childKeys = ["children", "childList", "items", "records", "list", "rows"]
                for childKey in childKeys:
                    childRows = row.get(childKey)
                    if not isinstance(childRows, list):
                        continue
                    for childRow in childRows:
                        if isinstance(childRow, dict):
                            checkRows.append(childRow)

                # 父行或任一 SKU 子行存在正向申收差异，则加入父级货件编号
                for checkRow in checkRows:
                    checkShipmentId = str(checkRow.get(self.shipmentIdKey) or shipmentId or "").strip()
                    if not checkShipmentId:
                        continue
                    # 父级行和 SKU 明细行的申收差异字段不同，两个字段都要兼容
                    diff = checkRow.get(self.diffKey)
                    if diff is None:
                        diff = checkRow.get(self.itemDiffKey)
                    if diff is None:
                        continue
                    diffText = str(diff).strip().replace(",", "")
                    try:
                        if float(diffText) <= 0:
                            continue
                    except (TypeError, ValueError):
                        continue
                    if checkShipmentId in seenShipmentIds:
                        continue
                    seenShipmentIds.add(checkShipmentId)
                    pageIds.append(checkShipmentId)
                    print(f"{checkShipmentId} quantityRecDiff: {diff}", flush=True)
                    break

            # 旧逻辑：提取全部货件 ID，不筛申收差异
            # pageIds = []
            # for row in rows:
            #     shipmentId = row.get(self.shipmentIdKey)
            #     if shipmentId:
            #         pageIds.append(shipmentId)

            allShipmentIds.extend(pageIds)
            print(f"第 {pageNum} 页申收差异>0 共 {len(pageIds)} 个，累计 {len(allShipmentIds)} 个", flush=True)
            # print(f"第 {pageNum} 页提取到 {len(pageIds)} 个，累计 {len(allShipmentIds)} 个", flush=True)

        return allShipmentIds

    def getShipmentDetails(self, shipmentIds):
        """逐个新标签页访问详情接口，解析页面 JSON 并提取货件详情"""
        page = self.page
        detailResults = []

        # 列表阶段监听已结束，停止监听避免干扰
        if page.listen.listening:
            page.listen.stop()

        for shipmentId in shipmentIds:
            # 拼接带货件 ID 的详情地址
            detailUrl = f"{self.detailApiUrl}?amazonShipmentId={shipmentId}"
            # 新标签页打开详情，不离开列表页
            tab = page.new_tab(url=detailUrl)
            time.sleep(1)

            # 从页面读取 JSON 正文
            pre = tab.ele('tag:pre', timeout=3)
            if pre:
                rawText = pre.text
            else:
                rawText = tab.html.strip()

            try:
                body = json.loads(rawText)
            except json.JSONDecodeError:
                print(f"{shipmentId} 详情 JSON 解析失败，跳过", flush=True)
                tab.close()
                continue

            # 取 data 节点
            data = body.get("data")
            if not data:
                print(f"{shipmentId} 详情无数据: {body.get('msg')}", flush=True)
                tab.close()
                continue

            # 货件级字段
            shopName = data.get("shopName")
            fulfillmentCenterId = data.get("fulfillmentCenterId")
            shipmentName = data.get("name")
            rawCreateTime = data.get("createTime")
            createTime = ""
            if rawCreateTime:
                text = str(rawCreateTime).strip()
                if len(text) >= 10:
                    createTime = text[:10]

            # source：拼接对象里所有非空字段
            source = data.get("source") or {}
            sourceText = ""
            if isinstance(source, dict):
                parts = []
                for val in source.values():
                    if val is None:
                        continue
                    text = str(val).strip()
                    if text:
                        parts.append(text)
                sourceText = " ".join(parts)
            elif source:
                sourceText = str(source).strip()

            # items：提取 fnSku、quantity、msku
            itemList = []
            for item in data.get("items") or []:
                if not isinstance(item, dict):
                    continue
                itemList.append({
                    "fnSku": item.get("fnSku"),
                    "quantity": item.get("quantity"),
                    "msku": item.get("msku"),
                })

            detailInfo = {
                "shipmentId": shipmentId,
                "shopName": shopName,
                "fulfillmentCenterId": fulfillmentCenterId,
                "name": shipmentName,
                "createTime": createTime,
                "source": sourceText,
                "items": itemList,
            }
            detailResults.append(detailInfo)
            print(detailInfo, flush=True)

            # 旧逻辑：只提取 msku 列表
            # if isinstance(data, list):
            #     rows = data
            # else:
            #     rows = (
            #         data.get("items")
            #         or data.get("list")
            #         or data.get("records")
            #         or data.get("commodityList")
            #         or [data]
            #     )
            # mskuList = []
            # for row in rows:
            #     if not isinstance(row, dict):
            #         continue
            #     msku = row.get("msku")
            #     if msku:
            #         mskuList.append(msku)
            # detailResults.append({
            #     "shipmentId": shipmentId,
            #     "mskuList": mskuList
            # })
            # print(f"{shipmentId} msku: {mskuList}", flush=True)
            tab.close()

        return detailResults

    def main(self):
        """执行赛狐 FBA 差异货件筛选与 POP 生成调试流程"""
        page = self.page
        env = "线上" if self.isOnline else "线下"
        print(f"运行环境：{env}", flush=True)
        page_size = 200
        login = SaiHuERPLogin(
            page=page,
            username=self.username,
            password=self.password,
            img_path=self.baseDir
        )
        login.login()
        print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)
        # 进入 FBA 货件产品维度页面
        page.ele('x://div/ul/li/span[text()="FBA"]', timeout=8).click()
        time.sleep(3)
        page.ele('x://a[text()="FBA货件"]', timeout=8).click()
        time.sleep(3)
        page.ele('x://span[text()="产品"]', timeout=8).click()
        time.sleep(3)

        # 选择当前调试配置中的站点
        page.ele('x://div/input[@placeholder="全部站点"]').click()
        time.sleep(3)

        siteName = self.siteName
        if siteName in self.directSiteNames:
            print(f"检测到该国家【{siteName}】, 不属于北美区与欧洲区, 直接点击...")
            # 直接找到页面上该国家的 span 并点击
            page.ele(f'x://span[contains(text(), "{siteName}")]').click()
            # return
        elif siteName in self.areaSiteMap:
            site = self.areaSiteMap[siteName]
            print(f"检测到该国家【{siteName}】, 属于【{site}】, 直接点击...")
            page.ele(f'x://span[text()="{site}"]', timeout=8).hover()
            # page.write(1.5)
            time.sleep(1.5)
            page.ele(f'x://span[normalize-space()="{siteName}"]', timeout=8).click()
            time.sleep(1.5)
        else:
            raise ValueError(f"未配置站点区域映射: {siteName}")
        page.ele('x://div[contains(@class, "cascader_footer")]//span[text()="确定"]', timeout=8).click()
        time.sleep(3)

        # 筛选 CLOSED 已完成货件
        page.ele('x://div/input[@placeholder="所有状态"]', timeout=15).click()
        time.sleep(3)
        page.ele('x://span[text()="CLOSED(已完成)"]', timeout=15).click()
        time.sleep(3)
        page.ele('x://div[@class="sf_select__footer"]//span[text()="确定"]', timeout=8).click()
        time.sleep(1)

        # 时间字段切换为更新时间
        page.ele('x://div[@class="picker_group"]//input[@readonly="readonly"]', timeout=8).click()
        time.sleep(3)
        page.ele('x://div[@class="el-scrollbar"]//span[text()="更新时间"]', timeout=8).click()
        time.sleep(3)

        # 启动列表接口监听后选择上月日期范围
        page.listen.start('https://www.sellfox.com/api/inbound/shipmentCommodity/page.json')

        page.ele('x://div[@class="picker_group"]//input[@class="el-range-input"]', timeout=8).click()
        time.sleep(3)
        page.ele('x://div/button[text()="上月"]', timeout=8).click()
        time.sleep(3)

         # 每页条数：200条/页（最大）
        page.ele('x://span[contains(@class,"el-pagination__sizes")]//input', timeout=8).click()
        time.sleep(0.5)
        page.ele('x://span[text()="200条/页"]', timeout=8).click()
        time.sleep(1)
        
        # 读取总条数并计算总页数
        total_text = page.ele('x://span[contains(@class,"total_style")]', timeout=8).text
        total_text = str(total_text).strip().replace(",", "")
        total_match = re.search(r"\d+", total_text)
        if not total_match:
            raise ValueError(f"未能解析总条数: {total_text}")
        total = int(total_match.group())
        total_pages = (total + page_size - 1) // page_size
        print(f"共 {total} 条，每页 {page_size} 条，共 {total_pages} 页", flush=True)
        
        allShipmentIds = self.getShipmentIds(total_pages)

        # print(f"共提取到 {len(allShipmentIds)} 个 amazonShipmentId")
        # for shipmentId in allShipmentIds:
        #     print(shipmentId)
        print(f"共提取到 {len(allShipmentIds)} 个申收差异大于 0 的 amazonShipmentId", flush=True)

        # 拉取货件详情并生成 POP 文件
        detailResults = self.getShipmentDetails(allShipmentIds)
        print(f"共获取 {len(detailResults)} 条货件详情 msku")
        for item in detailResults:
            print(item)
            savePath = self.popExport.build(item)
            if savePath:
                item["popPath"] = savePath
                self.generatedFiles.append(savePath)

        if self.config.get("sendEmail"):
            print(f"准备发送邮件，共 {len(self.generatedFiles)} 个 POP 文件", flush=True)
            deliver_outputs(self.config, self.generatedFiles)

        print("FBA货件页面加载完成")

    def run(self):
        """运行调试流程入口"""
        self.main()


class EdeckerClaim:
    """易得客：登录、进入店铺后台、索赔上传（第3步起，调试骨架）"""

    def __init__(self, config):
        self.debug = config.get("debug", False)
        self.yideke_username = config["yideke_username"]
        self.yideke_password = config["yideke_password"]
        ips = config.get("shop_ip") or config.get("ip")
        self.shop_ip = ips if isinstance(ips, list) else [ips]
        ports = config.get("shop_port") or config.get("port")
        if isinstance(ports, list):
            self.shop_port = [int(p) for p in ports]
        else:
            self.shop_port = [int(ports)]
        self.base_dir = Path(config.get("baseDir") or PopExport.getBaseDir())
        self.pod_awd = self.base_dir / "AWD亚马逊分销POD.pdf"
        self.pod_fba = self.base_dir / "FBA直发POD.pdf"
        self.amazon_email = config.get("amazon_email") or config.get("amazonEmail") or ""
        self.amazon_password = config.get("amazon_password") or config.get("amazonPassword") or ""
        self.page = None

    def kill_edecker(self, exclude_pid):
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                pid = proc.info['pid']
                name = proc.info['name']
                if name and name.lower() == 'edecker.exe':
                    if pid != exclude_pid:
                        proc.kill()
            except Exception:
                pass

    def kill_edecker_on_port(self, port):
        flag = f'--remote-debugging-port={port}'
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                name = proc.info['name']
                if name and name.lower() == 'edecker.exe':
                    cmdline = proc.info.get('cmdline') or []
                    if any(flag in str(arg) for arg in cmdline):
                        proc.kill()
            except Exception:
                pass

    def wait_for_port(self, port, timeout=60):
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(('127.0.0.1', port), timeout=2):
                    return
            except OSError:
                time.sleep(1)
        raise RuntimeError(f'等待 127.0.0.1:{port} 超时 ({timeout}s)')

    def run_edecker_automation(self, ips, port=9222):
        browser = Chromium(port)
        tab = browser.latest_tab
        time.sleep(2)
        for ip in ips:
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="美国"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30,
            ).click()
            time.sleep(3)
        self.kill_edecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def start_edecker(self, ip: str, port: int):
        base = Path.home() / "AppData/Local/eDecker6"
        exe_path = base / "Application/edecker.exe"
        profiles_path = base / "Profiles"

        if not exe_path.exists():
            raise FileNotFoundError(f"找不到 exe: {exe_path}")
        if not profiles_path.exists():
            raise FileNotFoundError(f"找不到 profiles 目录: {profiles_path}")

        ip_dot = ip
        ip_underline = ip.replace('.', '_')
        candidates = [
            p for p in profiles_path.iterdir()
            if p.is_dir() and (ip_dot in p.name or ip_underline in p.name)
        ]
        if not candidates:
            raise RuntimeError(f"未找到 IP={ip} 的 profile")

        latest = max(candidates, key=lambda p: p.stat().st_mtime)
        print(f"使用 profile: {latest}", flush=True)

        cmd = [
            str(exe_path),
            f'--user-data-dir={latest}',
            '--no-sandbox',
            f'--remote-debugging-port={port}',
        ]
        subprocess.Popen(cmd, cwd=str(base))

    def attach_existing_browser(self):
        """
        调试模式：
        直接接管已经打开的eDecker浏览器
        """
        port = self.shop_port[0]
        self.wait_for_port(port)
        self.page = ChromiumPage(f"127.0.0.1:{port}")

        try:
            self.page.set.window.max()
        except RuntimeError:
            pass
        print(f"✅ 已接管当前浏览器: {self.page.url}",flush=True)

        return self.page

    def login_and_open_shop(self):
        """第3步：易得客登录 → 访问店铺 → 启动 profile 浏览器"""
        sp = Specification(self.yideke_username, self.yideke_password)
        time.sleep(5)
        sp.YidekeLogin()
        time.sleep(3)
        self.run_edecker_automation(self.shop_ip)
        time.sleep(4)

        ip = self.shop_ip[0]
        port = self.shop_port[0]
        self.kill_edecker_on_port(port)
        time.sleep(1)
        self.start_edecker(ip, port)
        self.wait_for_port(port)
        self.page = ChromiumPage(f"127.0.0.1:{port}")
        try:
            self.page.set.window.max()
        except RuntimeError:
            pass
        print(f"已接管店铺浏览器，当前 URL: {self.page.url}", flush=True)
        amazon_email = self.amazon_email or None
        amazon_password = self.amazon_password or None
        # sp.amazonSellerLogin(self.page,amazon_email,amazon_password)
        self.page = sp.amazonSellerLogin(self.page, amazon_email, amazon_password)
        # self.page = self.page.latest_tab
        print(f"Amazon 登录后 URL: {self.page.url}", flush=True)
        return self.page

    def main(self):
        """第4步起：Amazon 卖家后台索赔页面操作"""
        page  = self.page

        # 调试/正式 模式切换
        if not self.debug:
            # 第4步：切换语言为中文简体（复用下载美国站Transaction报告）
            page.ele('x://div[@aria-label="语言"] | //div[@aria-label="Language"]', timeout=15).click()
            time.sleep(1.5)
            page.ele('x://div[text()="中文(简体)"]', timeout=10).click()
            time.sleep(5)
            print("已切换为中文简体", flush=True)
            # 第5步：汉堡菜单进入 库存 → 货件（复用下载美国站Transaction报告 shadow DOM 写法）
            menu_host = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=30)
            menu = menu_host.shadow_root
            menu.ele('x://div/img', timeout=30).click()
            time.sleep(1)
            menu.ele('x://div/span[text()="库存"]', timeout=20).click()
            time.sleep(3)
            page.wait(1)
            menu.ele('x://div/span[text()="货件"]', timeout=20).click()
            time.sleep(3)
            print("已进入货件界面", flush=True)

            # 查找判断是否有“清除筛选条件”按钮
            # 存在：点击-重新输入  不存在则直接进行输入
            if page.ele('x://span[text()="清除筛选条件"]'):
                page.ele('x://span[text()="清除筛选条件"]').click()
                time.sleep(1.5)
                # 第6步：搜索货件
                ship_id = "FBA19FG2L923"  # 测试货件编号
                page.ele('x://input[@placeholder="按货件编号搜索"]', timeout=15).input(f"{ship_id}\n", clear=True)
                time.sleep(3)
            else:
                # 第6步：搜索货件
                ship_id = "FBA199DZSQQQ"    #测试货件编号
                page.ele('x://input[@placeholder="按货件编号搜索"]', timeout=15).input(f"{ship_id}\n", clear=True)
                time.sleep(3)
            print(f"已搜索货件编号：{ship_id}", flush=True)

            # 第7步：进入对应的货件编号详情页
            ship_row = page.ele(f'x://tr[contains(., "{ship_id}")]', timeout=15)
            print("已找到对应行")
            # 查看当前行的HTML 结构
            # print(ship_row.html)
            # 在当前行内，找 sku表头下的值进入具体详情页

            ship_row.ele('x:.//div[contains(@class, "awsui_t-left")]//a').click()
            print("已进入对应的货件编号详情页")
            # 切换进入新tab，进行page操作
            page.wait(3)
            browser = page.browser
            detail_page = browser.latest_tab
            print("✅ 已切换到新详情页")
            detail_page.wait(3)


            # 第8步：在当前货件编号详情页，查看具体所需内容
            # 获取所需上传交货证明的分类
            select_sort = detail_page.ele('x://kat-box[contains(@class, "shipment-info-box")]', timeout=10)
            # print(select_sort.html)
            sort_label = select_sort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-label', timeout=5)
            if sort_label:
                sort = sort_label.attr('text')
            else:
                sort_label = select_sort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-link', timeout=5)
                # 交货证明的分类
                sort = sort_label.attr('label')
            print(f"获取到的值: {sort}")

            # 查看详情页的货件差值
            print(detail_page.html.find("shipmentContentsIframe"))
            frame = detail_page.get_frame('@data-testid=shipmentContentsIframe', timeout=10)
            headers = frame.eles('x://kat-table-head//kat-table-cell[@role="columnheader"]')
            col_index = None
            for i, item in enumerate(headers, start=1):
                if "差值" in item.text:
                    col_index = i
                    break
            if col_index is None:
                raise Exception("没有找到差值列")
            print(f"差值所在列:{col_index}")

            # 获取商品总量
            page_count_span = detail_page.ele('x://span[@id="page-item-count"]', timeout=5)
            total_text = page_count_span.text.strip()  # 例如 "显示 1 到 7 件商品"
            match = re.search(r'显示\s+\d+\s+到\s+(\d+)', total_text)
            if match:
                total_rows = int(match.group(1))
            else:
                total_rows = 0
            print(f"总共 {total_rows} 行商品")
            for row_index in range(1, total_rows + 1):
                # 当前行
                row = frame.ele(f'x://kat-table-body//kat-table-row[{row_index}]/kat-table-cell[{col_index}]',
                                timeout=5)
                # 提取文本并转换数字
                value_text = row.text.strip()
                value_int = int(value_text)
                print(f"第 {row_index} 行差值: {value_int}")

                # 第9步：定位需要操作的状态, 并依次选择对应状态
                # 定位需要操作的状态
                action_dropdown = frame.ele(
                    f'x://kat-table-body//kat-table-row[{row_index}]//kat-dropdown[@id="action-required"]', timeout=5)
                action_dropdown.click()
                time.sleep(1)
                shadow = action_dropdown.shadow_root
                if value_int < 0:
                    shadow.ele('x://*[(contains(text(),"调查缺失商品"))]', timeout=5).click()
                    print(f"第 {row_index} 行选择调查缺失商品")
                elif value_int > 0:
                    shadow.ele('x://*[(contains(text(),"调查额外商品"))]', timeout=5).click()
                    print(f"第 {row_index} 行选择调查额外商品")
                else:
                    print(f"第 {row_index} 行差值为0，不操作")


            # 第10步：提交上传POP/POD文件
            # 定位上传交货证明文件的按钮
            choose_btn = detail_page.ele('x://div[contains(@class,"document-proof")]'
                                         '[.//div[contains(text(),"交货证明")]]//kat-button[@label="选择文件"]', timeout=10)
            if not choose_btn:
                raise Exception("没有找到交货证明上传框")
                # 判断sort获取到的值
                # sort = '亚马逊分销' 上传 AWD亚马逊分销POD.pdf 文件
                # sort = 'Send to Amazon（视图）' 上传 FBA直发POD.pdf 文件
            if sort == '亚马逊分销':
                # 获取到的sort为亚马逊分销时，上传 AWD亚马逊分销POD.pdf 文件
                choose_btn.click.to_upload(str(self.pod_awd))
            else:
                # 否则 sort = 'Send to Amazon（视图）' 上传 FBA直发POD.pdf 文件
                choose_btn.click.to_upload(str(self.pod_fba))
            print("交货证明文件选择完成")

            # 以下还需继续调整补充
            # 定位上传库存所有权证明文件（赛狐流程完成导出的发票文件）的按钮
            invoice_btn = detail_page.ele('x://div[contains(@class,"document-proof")]'
                                          '[.//div[contains(text(),"库存所有权证明")]]'
                                          '//kat-button[@label="选择文件"]', timeout=10)
            if not invoice_btn:
                raise Exception("没有找到库存所有权证明上传框")
            invoice_btn.click.to_upload(str(self.pod_fba))
            # TODO: 第6-11步：搜索货件、上传 POP/POD、提交、记录 CASE 编号
        else:
            print("DEBUG模式，直接使用当前页面",flush=True)
            detail_page = page
        print("当前详情页:",detail_page.url)
        detail_page.wait(3)
        # 获取所需上传交货证明的分类
        select_sort = detail_page.ele('x://kat-box[contains(@class, "shipment-info-box")]', timeout=10)
        # print(select_sort.html)
        sort_label = select_sort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-label', timeout=5)
        if sort_label:
            sort = sort_label.attr('text')
        else:
            sort_label = select_sort.ele('x://dd[.//div[contains(text(),"创建自")]]//kat-link', timeout=5)
            # 交货证明的分类
            sort = sort_label.attr('label')
        print(f"获取到的值: {sort}")

        # 查看详情页的货件差值
        # 不要用 @src* 了，直接用这个
        print(detail_page.html.find("shipmentContentsIframe"))
        frame = detail_page.get_frame('@data-testid=shipmentContentsIframe', timeout=10)
        # print(frame.html[:3000])
        html = frame.html
        idx = html.find("差值")
        # print(html[idx - 1000:idx + 1000])

        headers = frame.eles('x://kat-table-head//kat-table-cell[@role="columnheader"]')
        col_index = None
        for i, item in enumerate(headers, start=1):
            if "差值" in item.text:
                col_index = i
                break
        if col_index is None:
            raise Exception("没有找到差值列")
        print(f"差值所在列:{col_index}")

        # 获取商品总量
        page_count_span = detail_page.ele('x://span[@id="page-item-count"]', timeout=5)
        total_text = page_count_span.text.strip()  # 例如 "显示 1 到 7 件商品"
        match = re.search(r'显示\s+\d+\s+到\s+(\d+)', total_text)
        if match:
            total_rows = int(match.group(1))
        else:
            total_rows = 0
        print(f"总共 {total_rows} 行商品")
        for row_index in range(1, total_rows + 1):
            # 当前行
            row = frame.ele(f'x://kat-table-body//kat-table-row[{row_index}]/kat-table-cell[{col_index}]', timeout=5)
            # 提取文本并转换数字
            value_text = row.text.strip()
            value_int = int(value_text)
            print(f"第 {row_index} 行差值: {value_int}")
            # 定位需要操作的状态
            action_dropdown = frame.ele(f'x://kat-table-body//kat-table-row[{row_index}]//kat-dropdown[@id="action-required"]', timeout=5)
            action_dropdown.click()
            time.sleep(1)
            shadow = action_dropdown.shadow_root
            if value_int < 0:
                shadow.ele('x://*[(contains(text(),"调查缺失商品"))]', timeout=5).click()
                print(f"第 {row_index} 行选择调查缺失商品")
            elif value_int > 0:
                shadow.ele('x://*[(contains(text(),"调查额外商品"))]', timeout=5).click()
                print(f"第 {row_index} 行选择调查额外商品")
            else:
                print(f"第 {row_index} 行差值为0，不操作")


        # 单行
        # value = frame.ele(f'x://kat-table-body//kat-table-row[1]/kat-table-cell[{col_index}]',timeout=10).text.strip()
        # print(f"获取到的差值为:{value}")
        # if int(value) < 0:
        #     # 定位需要操作的状态
        #     action_dropdown = frame.ele('x://kat-table-body//kat-table-row[1]//kat-dropdown[@id="action-required"]',timeout=10).click()
        #     time.sleep(1.5)
        #     shadow = action_dropdown.shadow_root
        #     shadow.ele('x://*[(contains(text(),"调查缺失商品"))]',timeout=10).click()
        #     print('当前差值为负数，已选择 "调查缺失商品" ')
        # elif int(value) > 0:
        #     # 定位需要操作的状态
        #     action_dropdown = frame.ele('x://kat-table-body//kat-table-row[1]//kat-dropdown[@id="action-required"]',
        #                                 timeout=10).click()
        #     time.sleep(1.5)
        #     shadow = action_dropdown.shadow_root
        #     shadow.ele('x://*[(contains(text(),"调查额外商品"))]', timeout=10).click()
        #     print('当前差值为正数，已选择 "调查额外商品" ')
        # else:
        #     print('无差异不操作')

        # 定位上传交货证明文件的按钮
        choose_btn = detail_page.ele('x://div[contains(@class,"document-proof")]'
                                     '[.//div[contains(text(),"交货证明")]]'
                                     '//kat-button[@label="选择文件"]', timeout=10)
        if not choose_btn:
            raise Exception("没有找到交货证明上传框")
        # 判断sort获取到的值
        # sort = '亚马逊分销' 上传 AWD亚马逊分销POD.pdf 文件
        # sort = 'Send to Amazon（视图）' 上传 FBA直发POD.pdf 文件
        if sort == '亚马逊分销':
            # 获取到的sort为亚马逊分销时，上传 AWD亚马逊分销POD.pdf 文件
            choose_btn.click.to_upload(str(self.pod_awd))
        else:
            # 否则 sort = 'Send to Amazon（视图）' 上传 FBA直发POD.pdf 文件
            choose_btn.click.to_upload(str(self.pod_fba))
        print("交货证明文件选择完成")

        # 定位上传库存所有权证明文件（赛狐流程完成导出的发票文件）的按钮
        invoice_btn = detail_page.ele('x://div[contains(@class,"document-proof")]'
                                     '[.//div[contains(text(),"库存所有权证明")]]'
                                     '//kat-button[@label="选择文件"]', timeout=10)
        # if not invoice_btn:
        #     raise Exception("没有找到库存所有权证明上传框")
        # invoice_btn.click.to_upload(str(self.pod_fba))





    def run(self, detail_results=None):
            if self.debug:
                print("进入调试模式，直接接管当前浏览器")
                self.attach_existing_browser()
            else:
                print("正常模式，执行登录流程")
                self.login_and_open_shop()
            self.main()
            if detail_results:
                print(f"待处理货件数: {len(detail_results)}",flush=True)
            return self.page


if __name__ == "__main__":
    # page = ChromiumPage()
    # config = {
    #     "page": page,
    #     "username": "",
    #     "password": "",
    #     "file_path": r"C:\Users\admin\Desktop",
    # }
    # test = Test(config)
    # test.main()

    # 易得客调试（取消注释后单独运行）
    edecker_config = {
        "yideke_username": "",
        "yideke_password": "",
        "shop_ip": "",
        "shop_port": 8888,
        "amazon_email": "",
        "amazon_password": "",
        "baseDir": str(PopExport.getBaseDir()),
        "debug":  True
    }
    EdeckerClaim(edecker_config).run()
