"""FBA 货件差异查询与 POP 导出核心编排。"""

import json
import re
import ctypes
import time
from pathlib import Path

from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin
from email_util import deliver_outputs
from export import PopExport


class FbaClaim:
    """赛狐 FBA 货件申收差异筛选、详情拉取与 POP 文档生成"""

    def __init__(self, config):
        # 运行配置（GUI 或 main 块传入）
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
        self.popExport.signatureName = str(config.get("signatureName") or "Xiaoyu Wang").strip() or "Xiaoyu Wang"
        self.popExport.signatureImagePath = str(config.get("signatureImagePath") or "").strip()
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
        ctypes.windll.user32.MessageBoxW(0, msg, "请手动重置筛选状态", 0x40)

    def run(self):
        """登录赛狐、筛选货件、拉详情并生成 POP 文档"""
        page = self.page
        env = "线上" if self.isOnline else "线下"
        print(f"运行环境：{env}", flush=True)

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

        # 优先重置上一次运行残留的筛选条件，找不到时等待人工处理
        try:
            page.ele('x://button[.//span[text()="重置"] or text()="重置"]', timeout=8).click()
            time.sleep(2)
        except Exception:
            print("未找到重置按钮，等待用户手动重置筛选状态。", flush=True)
            self.waitReset()
            time.sleep(1)

        # 选择当前配置中的站点
        page.ele('x://div/input[@placeholder="全部站点"]', timeout=8).click()
        time.sleep(3)
        siteName = self.siteName
        if siteName in self.directSiteNames:
            print(f"检测到该国家【{siteName}】, 不属于北美区与欧洲区, 直接点击...", flush=True)
            page.ele(f'x://span[contains(text(), "{siteName}")]', timeout=8).click()
        elif siteName in self.areaSiteMap:
            siteArea = self.areaSiteMap[siteName]
            print(f"检测到该国家【{siteName}】, 属于【{siteArea}】, 直接点击...", flush=True)
            page.ele(f'x://span[text()="{siteArea}"]', timeout=8).hover()
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

        # 监听列表接口，选择上月时间范围
        page.listen.start(self.listApiUrl)
        page.ele('x://div[@class="picker_group"]//input[@class="el-range-input"]', timeout=8).click()
        time.sleep(3)
        page.ele('x://div/button[text()="上月"]', timeout=8).click()
        time.sleep(3)

        # 每页条数设为 200（最大）
        page.ele('x://span[contains(@class,"el-pagination__sizes")]//input', timeout=8).click()
        time.sleep(0.5)
        page.ele('x://span[text()="200条/页"]', timeout=8).click()
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
        self.popExport.exportDir.mkdir(parents=True, exist_ok=True)
        shipmentJsonPath = self.popExport.exportDir / "shipment_ids.json"
        shipmentData = {
            "createdAt": time.strftime("%Y-%m-%d %H:%M:%S"),
            "siteName": self.siteName,
            "count": len(allShipmentIds),
            "shipmentIds": allShipmentIds,
        }
        shipmentJsonPath.write_text(json.dumps(shipmentData, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"货件编号兜底 JSON 已保存: {shipmentJsonPath}", flush=True)

        # 逐个拉取货件详情并生成 POP
        detailResults = self.getShipmentDetails(allShipmentIds)
        print(f"共获取 {len(detailResults)} 条货件详情 msku")
        for item in detailResults:
            print(item)
            try:
                savePath = self.popExport.build(item)
            except Exception as exc:
                # 单个 POP 生成失败时记录错误并继续处理后续货件
                item["popError"] = str(exc)
                print(f"{item.get('shipmentId')} POP 生成失败: {exc}", flush=True)
                continue
            if savePath:
                item["popPath"] = savePath
                self.generatedFiles.append(savePath)

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
        "username": "",
        "password": "",
        "exportDir": str(defaultBaseDir / "output"),
        "baseDir": str(defaultBaseDir),
        "isOnline": False,
        "siteName": "美国",
        "sendEmail": False,
        "email": "",
        "signatureName": "Xiaoyu Wang",
        "signatureImagePath": "",
    }

    claim = FbaClaim(config)
    claim.run()
