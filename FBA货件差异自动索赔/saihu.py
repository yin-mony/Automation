"""FBA 货件差异查询与 POP 导出核心编排。"""

import json
import os
import re
import ctypes
import time
from datetime import date, timedelta
from pathlib import Path

from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin
from email_util import deliver_outputs
from export import PopExport


class Saihu:
    """赛狐 FBA 货件申收差异筛选、详情拉取与 POP 文档生成"""

    def __init__(self, config):
        # 运行配置（GUI 或 main 块传入）
        self.config = config
        self.page = config["page"]
        self.username = config["username"]
        self.password = config["password"]
        # self.isOnline = bool(config.get("isOnline", False))
        self.baseDir = Path(config.get("baseDir") or PopExport.getBaseDir())
        # FBA 货件列表接口地址
        self.listApiUrl = "https://www.sellfox.com/api/inbound/shipmentCommodity/page.json"
        # FBA 货件详情接口地址
        self.detailApiUrl = "https://www.sellfox.com/api/inbound/shipment/detail.json"
        # 接口行数据里的货件 ID 字段名
        self.shipmentIdKey = "amazonShipmentId"
        # 申收差异字段名
        self.diffKey = "quaRecDifference"
        # SKU 明细里的申收差异字段名
        self.itemDiffKey = "quantityRecDiff"
        # 列表每页条数（赛狐最大 200）
        self.pageSize = 200
        # 默认筛选站点，从配置读取，未配置时使用美国
        self.siteName = str(config.get("siteName") or "美国").strip()
        # 默认筛选店铺，从配置读取，未配置时使用常用店铺
        self.shopName = str(config.get("shopName") or "Lydia deal-US").strip()
        # 筛选更新时间范围，从 GUI 读取，未配置时默认上月
        firstDayThisMonth = date.today().replace(day=1)
        lastDayLastMonth = firstDayThisMonth - timedelta(days=1)
        firstDayLastMonth = lastDayLastMonth.replace(day=1)
        self.startDate = str(config.get("startDate") or firstDayLastMonth.strftime("%Y-%m-%d")).strip()
        self.endDate = str(config.get("endDate") or lastDayLastMonth.strftime("%Y-%m-%d")).strip()
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
        templatePath = str(config.get("templatePath") or "").strip()
        if templatePath:
            self.popExport.customTemplatePath = Path(templatePath).resolve()
        exportDir = config.get("exportDir")
        if exportDir:
            self.popExport.exportDir = Path(str(exportDir)).resolve()
        # 本次运行生成的 POP 文件路径
        self.generatedFiles = []

    def getShipmentIds(self, totalPages):
        """监听列表接口并翻页，提取申收差异大于 0 的 amazonShipmentId"""
        page = self.page
        allShipmentIds = []
        seenShipmentIds = set()

        for pageNum in range(1, totalPages + 1):
            # 第 2 页起先点页码触发请求
            if pageNum > 1:
                page.listen.clear()
                print(f"准备切换到第 {pageNum} 页", flush=True)
                try:
                    page.ele(f'x://ul[contains(@class,"el-pager")]//li[text()="{pageNum}"]', timeout=8).click()
                except Exception:
                    pageClicked = page.run_js(
                        """
                        const target = String(arguments[0]);
                        const nodes = Array.from(document.querySelectorAll('.el-pager li'));
                        const node = nodes.find(item => {
                            const rect = item.getBoundingClientRect();
                            const text = (item.innerText || item.textContent || '').trim();
                            return rect.width > 0 && rect.height > 0 && text === target;
                        });
                        if (!node) return false;
                        node.click();
                        return true;
                        """,
                        pageNum,
                    )
                    if not pageClicked:
                        nextClicked = page.run_js(
                            """
                            const node = Array.from(document.querySelectorAll('button.btn-next, .el-pagination .btn-next')).find(item => {
                                const rect = item.getBoundingClientRect();
                                return rect.width > 0 && rect.height > 0 && !item.disabled;
                            });
                            if (!node) return false;
                            node.click();
                            return true;
                            """
                        )
                        if not nextClicked:
                            raise
                time.sleep(1)

            # 第一页取筛选/条数切换阶段产生的最新接口包，后续页只等待本次翻页请求
            if pageNum == 1:
                packet = page.listen.wait(count=999, timeout=3, fit_count=False)
            else:
                packet = page.listen.wait(timeout=20, fit_count=False)
            if isinstance(packet, list):
                packet = packet[-1] if packet else None
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
            responsePage = (
                data.get("current")
                or data.get("pageNum")
                or data.get("pageNo")
                or data.get("page")
                or data.get("currentPage")
                or data.get("pageIndex")
                or "未知"
            )
            print(f"第 {pageNum} 页接口响应页码: {responsePage}，返回行数: {len(rows)}", flush=True)

            # 只提取申收差异大于 0 的父级货件 ID，同一货件只保留一次
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

            allShipmentIds.extend(pageIds)
            print(f"第 {pageNum} 页申收差异>0 共 {len(pageIds)} 个，累计 {len(allShipmentIds)} 个", flush=True)

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
            tab.close()

        return detailResults

    def waitReset(self):
        """提示用户手动重置筛选状态并等待确认"""
        msg = (
            "未找到赛狐页面的“重置”按钮。\n\n"
            "请在当前赛狐 FBA货件产品页面手动点击“重置”，"
            "确认筛选状态已清空后，再点击本提示框的“确定”继续。"
        )
        ctypes.windll.user32.MessageBoxW(0, msg, "请手动重置筛选状态", 0x40 | 0x40000)

    def run(self):
        """登录赛狐、筛选货件、拉详情并生成 POP 文档"""
        page = self.page
        env = "线上" if self.isOnline else "线下"
        print(f"运行环境：{env}", flush=True)
        self.popExport.exportDir.mkdir(parents=True, exist_ok=True)
        shipmentJsonPath = self.popExport.exportDir / "shipment_ids.json"
        if shipmentJsonPath.exists():
            shipmentJsonPath.unlink()
            print(f"已删除上次货件编号 JSON: {shipmentJsonPath}", flush=True)
        runningData = {
            "createdAt": time.strftime("%Y-%m-%d %H:%M:%S"),
            "siteName": self.siteName,
            "shopName": self.shopName,
            "startDate": self.startDate,
            "endDate": self.endDate,
            "status": "running",
            "count": 0,
            "shipmentIds": [],
            "allShipmentIds": [],
            "failedShipmentIds": [],
            "popErrors": [],
        }
        shipmentJsonPath.write_text(json.dumps(runningData, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"本轮货件编号 JSON 已初始化: {shipmentJsonPath}", flush=True)

        # 赛狐登录（含验证码 OCR）
        login = SaiHuERPLogin(
            page=page,
            username=self.username,
            password=self.password,
            img_path=self.baseDir,
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
        login.closeNotice()
        time.sleep(1)

        # 优先重置上一次运行残留的筛选条件，找不到时等待人工处理
        try:
            page.ele('x://button[.//span[text()="重置"] or text()="重置"]', timeout=8).click()
            time.sleep(2)
        except Exception:
            print("未找到重置按钮，等待用户手动重置筛选状态。", flush=True)
            self.waitReset()
            time.sleep(1)
        login.closeNotice()
        time.sleep(1)

        # 选择当前配置中的站点
        siteName = self.siteName
        try:
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
        except Exception:
            filterText = ""
        selectedSiteTags = page.run_js(
            """
            const cascader = document.querySelector('.filter_container .el-cascader');
            if (!cascader) return [];
            return Array.from(cascader.querySelectorAll('.el-tag')).map(item => {
                return (item.innerText || item.textContent || '').replace(/\\s+/g, ' ').trim();
            }).filter(Boolean);
            """
        )
        if selectedSiteTags == [siteName]:
            print(f"当前筛选站点已是【{siteName}】，跳过重复选择。", flush=True)
        else:
            page.run_js(
                """
                const cascader = document.querySelector('.filter_container .el-cascader');
                const component = cascader && cascader.__vue__;
                if (component && typeof component.handleClear === 'function') {
                    component.handleClear();
                    return true;
                }
                return false;
                """
            )
            time.sleep(1)
            page.run_js("document.querySelector('.filter_container .el-cascader__tags').click();")
            time.sleep(4)
            if siteName in self.directSiteNames:
                print(f"检测到该国家【{siteName}】, 不属于北美区与欧洲区, 直接点击...", flush=True)
            elif siteName in self.areaSiteMap:
                site = self.areaSiteMap[siteName]
                print(f"检测到该国家【{siteName}】, 属于【{site}】, 直接点击...", flush=True)
                areaPoint = page.run_js(
                    """
                    const target = arguments[0];
                    const node = Array.from(document.querySelectorAll('.el-cascader-node')).find(item => {
                        const rect = item.getBoundingClientRect();
                        const text = (item.innerText || item.textContent || '').replace(/\\s+/g, ' ').trim();
                        return rect.width > 0 && rect.height > 0 && text.includes(target);
                    });
                    if (!node) return null;
                    const rect = node.getBoundingClientRect();
                    return {x: rect.left + rect.width / 2, y: rect.top + rect.height / 2};
                    """,
                    site,
                )
                if not areaPoint:
                    raise Exception(f"未找到可见站点区域: {site}")
                page.actions.move_to((areaPoint["x"], areaPoint["y"]), duration=.2)
                time.sleep(1.5)
            else:
                raise ValueError(f"未配置站点区域映射: {siteName}")
            siteClicked = page.run_js(
                """
                const target = arguments[0];
                const node = Array.from(document.querySelectorAll('.el-cascader-node')).find(item => {
                    const rect = item.getBoundingClientRect();
                    const text = (item.innerText || item.textContent || '').replace(/\\s+/g, ' ').trim();
                    return rect.width > 0 && rect.height > 0 && text.includes(target);
                });
                if (!node) return false;
                const checkbox = node.querySelector('.el-checkbox__inner') || node.querySelector('label') || node;
                checkbox.click();
                return true;
                """,
                siteName,
            )
            if not siteClicked:
                raise Exception(f"未找到可见站点选项: {siteName}")
            time.sleep(1.5)
            confirmClicked = page.run_js(
                """
                const node = Array.from(document.querySelectorAll('.el-cascader__dropdown .cascader_footer span, .el-cascader__dropdown .cascader_footer button')).find(item => {
                    const rect = item.getBoundingClientRect();
                    const text = (item.innerText || item.textContent || '').replace(/\\s+/g, ' ').trim();
                    return rect.width > 0 && rect.height > 0 && text === '确定';
                });
                if (!node) return false;
                node.click();
                return true;
                """
            )
            if not confirmClicked:
                page.ele('x://div[contains(@class, "cascader_footer")]//span[text()="确定"]', timeout=8).click()
            time.sleep(3)
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
            if siteName not in str(filterText):
                raise Exception(f"站点筛选未生效: {siteName}")

        # 选择当前配置中的店铺
        shopName = self.shopName
        try:
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
        except Exception:
            filterText = ""
        if shopName in str(filterText):
            print(f"当前筛选店铺已是【{shopName}】，跳过重复选择。", flush=True)
        else:
            shopChanged = page.run_js(
                """
                const target = arguments[0];
                const node = Array.from(document.querySelectorAll('*')).find(item => {
                    const component = item.__vue__;
                    return component && component.currentFilter && component.$refs && component.$refs.multiShopSelectorNew
                        && typeof component.filterChange === 'function';
                });
                if (!node) return false;
                const component = node.__vue__;
                const shopRef = component.$refs.multiShopSelectorNew;
                const options = (shopRef.shopSelectOptions && shopRef.shopSelectOptions.options) || [];
                const option = options.find(item => item.label === target);
                if (!option) return false;
                if (typeof shopRef.setData === 'function') shopRef.setData([option.value]);
                shopRef.innerValue = [option.value];
                if (shopRef.shopSelectOptions) shopRef.shopSelectOptions.value = [option.value];
                component.currentFilter.shopIds = [option.value];
                component.filterChange([option.value], 'shopIds');
                return true;
                """,
                shopName,
            )
            if not shopChanged:
                raise Exception(f"店铺筛选设置失败: {shopName}")
            time.sleep(2)
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
            if shopName not in str(filterText):
                raise Exception(f"店铺筛选未生效: {shopName}")

        # 筛选 CLOSED 已完成货件
        try:
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
        except Exception:
            filterText = ""
        if "CLOSED" in str(filterText) or "已完成" in str(filterText):
            print("当前筛选状态已是 CLOSED(已完成)，跳过重复选择。", flush=True)
        else:
            statusChanged = page.run_js(
                """
                const node = Array.from(document.querySelectorAll('*')).find(item => {
                    const component = item.__vue__;
                    return component && component.currentFilter && component.statusSelectOptions && typeof component.filterChange === 'function';
                });
                if (!node) return false;
                const component = node.__vue__;
                component.statusSelectOptions.value = ['CLOSED'];
                component.currentFilter.status = ['CLOSED'];
                component.filterChange(['CLOSED'], 'status');
                return true;
                """
            )
            if not statusChanged:
                raise Exception("状态筛选设置失败: CLOSED(已完成)")
            time.sleep(2)
            filterText = page.ele('x://div[contains(@class,"filter_container")]', timeout=3).text
            if "CLOSED" not in str(filterText) and "已完成" not in str(filterText):
                raise Exception("状态筛选未生效: CLOSED(已完成)")

        # 时间字段切换为更新时间
        try:
            timeValue = page.ele('x://div[@class="picker_group"]//input[@readonly="readonly"]', timeout=3).attr("value")
        except Exception:
            timeValue = ""
        if "更新时间" in str(timeValue):
            print("当前时间字段已是更新时间，跳过重复选择。", flush=True)
        else:
            timeChanged = page.run_js(
                """
                const input = Array.from(document.querySelectorAll('.picker_group input[readonly]')).find(item => {
                    return item.value === '创建时间' || item.value === '更新时间' || item.value === '货件签收时间' || item.value === '预计送达时间';
                });
                const component = input && input.closest('.picker_group_selector') && input.closest('.picker_group_selector').__vue__;
                if (!component) return false;
                component.singleInnerValue = 'updateTime';
                if (typeof component.confirm === 'function') component.confirm();
                return true;
                """
            )
            if not timeChanged:
                raise Exception("未找到时间字段选项: 更新时间")
            time.sleep(3)

        # 监听列表接口，按 GUI 传入的开始时间和结束时间筛选
        page.listen.start(self.listApiUrl)
        dateRange = page.run_js(
            """
            const range = [arguments[0], arguments[1]];
            const component = document.querySelector('.picker_group') && document.querySelector('.picker_group').__vue__;
            if (!component) return null;
            component.pickervalue = range;
            if (typeof component.dateChange === 'function') component.dateChange(range);
            if (typeof component.forceChange === 'function') component.forceChange();
            return range;
            """,
            self.startDate,
            self.endDate,
        )
        if not dateRange:
            raise Exception("未能设置 GUI 传入的时间范围")
        print(f"赛狐日期范围已切换为: {dateRange[0]} ~ {dateRange[1]}", flush=True)
        time.sleep(3)

        # 每页条数设为 200（最大），保留当前监听队列供第一页读取最新接口包
        pageSizeInput = page.ele('x://span[contains(@class,"el-pagination__sizes")]//input', timeout=8)
        if pageSizeInput.value == "200条/页":
            print("当前每页条数已是 200条/页，跳过重复选择。", flush=True)
        else:
            pageSizeInput.click()
            time.sleep(0.5)
            page.ele(
                'x://div[contains(@class,"el-select-dropdown") and not(contains(@style,"display: none"))]'
                '//span[text()="200条/页"]',
                timeout=8,
            ).click()
            time.sleep(1)

        # 读取总条数并计算总页数
        totalText = page.ele('x://span[contains(@class,"total_style")]', timeout=8).text
        totalText = str(totalText).strip().replace(",", "")
        totalMatch = re.search(r"\d+", totalText)
        if not totalMatch:
            raise ValueError(f"未能解析总条数: {totalText}")
        total = int(totalMatch.group())
        totalPages = (total + self.pageSize - 1) // self.pageSize
        print(f"共 {total} 条，每页 {self.pageSize} 条，共 {totalPages} 页", flush=True)

        # 翻页收集申收差异大于 0 的货件 ID
        allShipmentIds = self.getShipmentIds(totalPages)
        print(f"共提取到 {len(allShipmentIds)} 个申收差异大于 0 的 amazonShipmentId", flush=True)

        # 逐个拉取货件详情并生成 POP
        detailResults = self.getShipmentDetails(allShipmentIds)
        print(f"共获取 {len(detailResults)} 条货件详情 msku")
        detailShipmentIds = {str(item.get("shipmentId") or "").strip().upper() for item in detailResults}
        failedShipmentIds = []
        popErrors = []
        for shipmentId in allShipmentIds:
            if shipmentId not in detailShipmentIds:
                failedShipmentIds.append(shipmentId)
                popErrors.append({
                    "shipmentId": shipmentId,
                    "reason": "详情获取失败或无可导出明细",
                })
        successShipmentIds = []
        for item in detailResults:
            print(item)
            shipmentId = str(item.get("shipmentId") or "").strip().upper()
            try:
                savePath = self.popExport.build(item)
            except Exception as exc:
                # 单个 POP 生成失败时记录错误并继续处理后续货件
                item["popError"] = str(exc)
                failedShipmentIds.append(shipmentId)
                popErrors.append({
                    "shipmentId": shipmentId,
                    "reason": str(exc),
                })
                print(f"{item.get('shipmentId')} POP 生成失败: {exc}", flush=True)
                continue
            if savePath:
                item["popPath"] = savePath
                self.generatedFiles.append(savePath)
                if shipmentId and shipmentId not in successShipmentIds:
                    successShipmentIds.append(shipmentId)

        shipmentData = {
            "createdAt": time.strftime("%Y-%m-%d %H:%M:%S"),
            "siteName": self.siteName,
            "shopName": self.shopName,
            "startDate": self.startDate,
            "endDate": self.endDate,
            "status": "done",
            "sourceCount": len(allShipmentIds),
            "detailCount": len(detailResults),
            "count": len(successShipmentIds),
            "shipmentIds": successShipmentIds,
            "allShipmentIds": allShipmentIds,
            "failedShipmentIds": failedShipmentIds,
            "popErrors": popErrors,
        }
        shipmentJsonPath.write_text(json.dumps(shipmentData, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            f"本轮成功生成 POP 的货件编号 JSON 已保存: {shipmentJsonPath}，共 {len(successShipmentIds)} 个",
            flush=True,
        )

        # 可选：邮件发送本次生成的 POP 文件
        if self.config.get("sendEmail"):
            print(f"准备发送邮件，共 {len(self.generatedFiles)} 个 POP 文件", flush=True)
            deliver_outputs(self.config, self.generatedFiles)

        print("FBA货件页面加载完成")


if __name__ == "__main__":
    # 本文件独立调试配置
    defaultBaseDir = PopExport.getBaseDir()
    config = {
        "page": ChromiumPage(),
        "username": os.getenv("SAIHU_USERNAME", ""),
        "password": os.getenv("SAIHU_PASSWORD", ""),
        "exportDir": str(defaultBaseDir / "file"),
        "baseDir": str(defaultBaseDir),
        # "isOnline": False,
        "siteName": "美国",
        "sendEmail": False,
        "email": "",
        "templatePath": str(defaultBaseDir / "服务商模板.docx"),
    }

    service = Saihu(config)
    service.run()
