import csv
import json
import random
import re
import time
from pathlib import Path
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from DrissionPage import ChromiumPage, ChromiumOptions
from openpyxl import Workbook


class Test:
    """TREC 公司合作推广测试流程。

    本测试文件保留旧流程完整逻辑，并按根目录规范整理为更易读的顺序结构：
    1. 采集 TREC 网站列表页 license 基础数据。
    2. 融合 Texas Open Data 中的 broker/company 关系数据。
    3. 导出中间 CSV 文件。
    4. 根据公司和 broker 名称搜索邮箱、电话。
    5. 导出最终 CSV 和 Excel 文件。
    """

    def __init__(self, config=None):
        """初始化默认配置、固定地址、缓存数据和过滤规则。"""
        if config is None:
            config = {}

        # 基础运行配置
        self.outputDir = Path(config.get("outputDir", "output"))
        self.maxPage = int(config.get("maxPage", 5))
        self.maxBrokers = config.get("maxBrokers", None)
        self.searchOnly = bool(config.get("searchOnly", False))
        self.perPage = int(config.get("perPage", 100))
        self.waitSecond = float(config.get("waitSecond", 6))
        self.retryCount = int(config.get("retryCount", 5))

        # 固定接口地址和浏览器请求头
        self.trecHomeUrl = "https://www.trec.texas.gov/"
        self.trecDetailUrl = "https://www.trec.texas.gov/license-search/?detail_id={detailId}"
        self.openDataUrl = "https://data.texas.gov/resource/s7ft-44qi.json"
        self.googleSearchUrl = "https://www.google.com/search"
        self.userAgent = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
        )

        # 开放数据读取配置；按分页读取，避免只读取前 50000 条导致融合缺失
        self.openDataLimit = int(config.get("openDataLimit", 50000))
        self.openDataOffset = int(config.get("openDataOffset", 0))
        self.maxOpenDataPage = int(config.get("maxOpenDataPage", 100))

        # 运行过程缓存
        self.rows = []
        self.seen = set()

        # 邮箱过滤规则，避免明显无效或不适合推广的邮箱
        self.genericEmailPrefixes = {
            "admin", "info", "support", "contact", "sales", "office",
            "webmaster", "postmaster", "help", "noreply", "no-reply",
            "donotreply", "do-not-reply", "abuse", "billing", "hr",
            "payroll", "marketing", "legal", "privacy", "security",
            "feedback", "team", "hello", "service", "customerservice",
            "customer-service", "enquiries", "enquiry", "inquiries",
            "inquiry", "jobs", "careers", "recruitment", "press",
            "media", "spam", "root", "hostmaster", "techsupport",
            "it", "operations", "accounts", "purchasing", "orders",
            "returns", "shipping", "logistics", "reception",
        }
        self.blockedEmailTlds = {"gov", "mil", "edu"}
        self.blockedEmailDomains = {
            "mailinator.com", "guerrillamail.com", "guerrillamail.de",
            "tempmail.com", "throwaway.email", "yopmail.com",
            "trashmail.com", "trashmail.net", "sharklasers.com",
            "guerrillamailblock.com", "grr.la", "dispostable.com",
            "maildrop.cc", "temp-mail.org", "fakeinbox.com",
            "10minutemail.com", "guerrillamail.info", "mohmal.com",
        }

    def initiate(self):
        """打开 TREC 网站、选择 Active 状态，并监听真实列表接口。"""
        browserOptions = ChromiumOptions()
        browserOptions.set_argument("--incognito")
        browserOptions.set_user_agent(self.userAgent)

        page = ChromiumPage(browserOptions)
        page.get(self.trecHomeUrl)

        # 进入 License Holder Search 页面
        page.ele('x://button[@name="license_search"]').click()
        time.sleep(3)

        # 选择 Active 状态
        page.ele('x://label[text()=" License Status"]')("t:select").select("Active").click()
        time.sleep(2)

        # 开始监听列表接口，然后点击搜索触发请求
        page.listen.start("collections/licenses/documents/search", method="GET")
        page.ele('x://input[@name="license_search"]').click()
        time.sleep(5)

        # 固定次数等待接口，避免无限循环；只接收 q=* 的列表接口
        for listenIndex in range(1, 31):
            try:
                packet = page.listen.wait(timeout=10)
                params = dict(packet.request.params)
                print("监听到接口:", packet.request.url)

                if params.get("q") != "*":
                    continue

                apiUrl = packet.request.url.split("?")[0]
                headers = {}
                for key, value in dict(packet.request.headers).items():
                    lowerKey = key.lower()
                    if key.startswith(":"):
                        continue
                    if lowerKey in ["host", "content-length", "accept-encoding"]:
                        continue
                    headers[key] = value
                headers["Accept-Encoding"] = "identity"

                page.quit()
                return apiUrl, params, headers
            except Exception as error:
                print("第", listenIndex, "次监听失败:", error)
                time.sleep(self.waitSecond)

        page.quit()
        return None

    def requestJson(self, url, params=None, headers=None, timeout=30, retries=None):
        """发送 GET 请求并返回 JSON，失败时按固定次数重试。"""
        if retries is None:
            retries = self.retryCount

        queryUrl = url
        if params:
            queryUrl = url + "?" + urlencode(params, doseq=True)

        requestHeaders = {
            "User-Agent": self.userAgent,
            "Accept": "application/json",
        }
        if headers:
            requestHeaders.update(headers)

        lastError = None
        for attempt in range(1, retries + 1):
            try:
                request = Request(queryUrl, headers=requestHeaders)
                response = urlopen(request, timeout=timeout).read().decode("utf-8")
                return json.loads(response)
            except Exception as error:
                lastError = error
                sleepSecond = min(60, (2 ** (attempt - 1)) + random.random())
                print("请求失败，准备重试:", attempt, "/", retries, error)
                time.sleep(sleepSecond)

        print("请求最终失败:", queryUrl, lastError)
        return None

    def collectList(self):
        """采集 TREC 列表页数据，maxPage 为 0 时按接口 found 全量采集。"""
        result = self.initiate()
        if result is None:
            print("列表接口监听失败")
            return

        apiUrl, baseParams, headers = result

        # 先请求第一页，用 found 计算全量条数和页数
        firstParams = dict(baseParams)
        firstParams["page"] = "1"
        firstParams["per_page"] = str(self.perPage)
        firstData = self.requestJson(apiUrl, firstParams, headers, timeout=30, retries=3)
        if not firstData:
            print("第一页列表数据请求失败")
            return

        totalCount = int(firstData.get("found") or 0)
        realPerPage = int(firstData.get("request_params", {}).get("per_page") or self.perPage)
        totalPage = (totalCount + realPerPage - 1) // realPerPage

        runPage = totalPage
        if self.maxPage > 0 and self.maxPage < totalPage:
            runPage = self.maxPage

        print("列表接口:", apiUrl)
        print("全量数据条数:", totalCount)
        print("每页展示数量:", realPerPage)
        print("全量页数:", totalPage)
        print("本次采集页数:", runPage)

        for pageNo in range(1, runPage + 1):
            print("开始采集第", pageNo, "页")

            if pageNo == 1:
                data = firstData
            else:
                pageParams = dict(baseParams)
                pageParams["page"] = str(pageNo)
                pageParams["per_page"] = str(realPerPage)
                data = self.requestJson(apiUrl, pageParams, headers, timeout=30, retries=3)

            if not data:
                print("第", pageNo, "页无数据，停止采集")
                break

            groupedHits = data.get("grouped_hits") or []
            if not groupedHits:
                print("第", pageNo, "页为空，停止采集")
                break

            for rowNo, group in enumerate(groupedHits, 1):
                hits = group.get("hits") or []
                if not hits:
                    continue

                doc = hits[0].get("document") or {}
                uid = str(doc.get("detailId", "") or "")
                if not uid or uid in self.seen:
                    continue
                self.seen.add(uid)

                name = str(doc.get("fullName") or "") or (
                    f"{doc.get('lastName', '')} "
                    f"{doc.get('firstName', '')} "
                    f"{doc.get('middleName', '')}"
                ).strip()
                code = str(doc.get("customId", "") or "")
                url = self.trecDetailUrl.format(detailId=uid)

                self.rows.append({
                    "uid": uid,
                    "code": code,
                    "name": name,
                    "url": url,
                    "pg": str(pageNo),
                })

                print("第", pageNo, "页，第", rowNo, "条:", name, uid, url)

            print("第", pageNo, "页完成，当前数量:", len(self.rows))
            time.sleep(self.waitSecond)

    def mergeData(self):
        """分页读取 Texas Open Data，并融合 broker/company 关系字段。"""
        if not self.rows:
            print("没有列表数据，跳过关系融合")
            return

        # 列表 code 带有 -SA / -B 等后缀，先统一取数字前缀用于匹配
        matcher = re.compile(r"^(\d+)")
        codeMap = {}
        for row in self.rows:
            code = row.get("code", "").strip()
            match = matcher.match(code)
            if not match:
                continue
            key = match.group(1)
            codeMap.setdefault(key, []).append(row)

        if not codeMap:
            print("列表数据中没有可用于融合的 license code")
            return

        # Open Data 需要分页读取；只读前 50000 条时，很多 license 关系会缺失
        trecData = []
        for openPage in range(1, self.maxOpenDataPage + 1):
            offset = self.openDataOffset + (openPage - 1) * self.openDataLimit
            openParams = {
                "$limit": self.openDataLimit,
                "$offset": offset,
            }
            print("开始读取开放数据第", openPage, "页:", openParams)
            pageData = self.requestJson(self.openDataUrl, params=openParams, timeout=60, retries=3)
            if not pageData:
                print("开放数据读取为空，停止读取")
                break

            trecData.extend(pageData)
            print("开放数据累计数量:", len(trecData))

            if len(pageData) < self.openDataLimit:
                break

        if not trecData:
            print("开放数据为空，跳过关系融合")
            return

        mergeCount = 0
        for item in trecData:
            # 列表中的 code 对应 Open Data 的 license_number；related_license_number 是其关联 broker/company
            licenseNo = str(item.get("license_number", "") or "").strip()
            match = matcher.match(licenseNo)
            if not match:
                continue

            key = match.group(1)
            rowList = codeMap.get(key, [])
            if not rowList:
                continue

            fullName = str(item.get("full_name", "") or "").strip()
            relatedLicenseType = str(item.get("related_license_type", "") or "").strip()
            relatedNo = str(item.get("related_license_number", "") or "").strip()
            relatedFullName = str(item.get("related_license_full_name", "") or "").strip()
            recordType = str(item.get("license_type", "") or "").strip()

            extra = {
                "broker_full_name": fullName,
                "broker_license_type": relatedLicenseType,
                "broker_license_number": relatedNo,
                "broker_company_name": relatedFullName,
                "record_type": recordType,
            }

            if recordType == "Sales Agent":
                extra["relation_name"] = relatedFullName
            elif recordType == "Broker Company":
                extra["relation_name"] = fullName

            for row in rowList:
                row.update(extra)
                mergeCount += 1

        print("关系数据融合完成，命中数量:", mergeCount, "当前列表数量:", len(self.rows))

    def exportRows(self, outputPath):
        """导出列表采集和融合后的中间结果 CSV。"""
        if not self.rows:
            print("没有列表数据可导出")
            return

        fieldnames = [
            "uid", "code", "name", "url", "pg",
            "broker_full_name", "broker_license_type",
            "broker_license_number", "broker_company_name",
            "record_type", "relation_name",
        ]
        path = Path(outputPath)
        path.parent.mkdir(parents=True, exist_ok=True)

        with path.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(self.rows)

        print("中间结果已导出:", path, "数量:", len(self.rows))

    def extractEmailsAndPhones(self, text):
        """从搜索结果页面文本中提取邮箱和电话。"""
        emailsRaw = re.findall(
            r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text
        )
        phonesRaw = re.findall(
            r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text
        )

        emails = []
        companyEmails = []
        brokerEmails = []

        for email in emailsRaw:
            emailLower = email.lower()
            domain = emailLower.split("@")[-1] if "@" in emailLower else ""

            if domain in self.blockedEmailDomains:
                continue
            if domain.split(".")[-1] in self.blockedEmailTlds:
                continue
            if emailLower in emails:
                continue

            emails.append(emailLower)
            prefix = emailLower.split("@")[0]
            if prefix in self.genericEmailPrefixes or self.isLikelyPersonalEmail(emailLower):
                companyEmails.append(emailLower)
            else:
                brokerEmails.append(emailLower)

        phones = []
        for phone in phonesRaw:
            cleaned = re.sub(r"[^\d]", "", phone)
            if len(cleaned) == 10 and cleaned not in phones:
                phones.append(cleaned)

        return emails, phones, companyEmails, brokerEmails

    @staticmethod
    def isLikelyPersonalEmail(email):
        """判断邮箱是否更像个人邮箱或泛用邮箱。"""
        local = email.split("@")[0].lower() if "@" in email else ""
        if re.match(r"^(info|contact|hello|office|support|admin)", local):
            return True
        if re.match(r"^\d{3,}", local):
            return True
        return False

    def searchGoogle(self, query, page):
        """在 Google 搜索指定关键词，并从页面文本提取邮箱电话。"""
        url = self.googleSearchUrl + "?q=" + quote(query) + "&hl=en"
        page.get(url)
        time.sleep(random.uniform(4, 7))

        try:
            page.ele("x://div[@role='progressbar']", timeout=3)
            time.sleep(3)
        except Exception:
            pass

        body = page("x:/html/body")
        text = body.text if body else ""

        emails, phones, companyEmails, brokerEmails = self.extractEmailsAndPhones(text)
        return emails, phones, companyEmails, brokerEmails, text

    def brokerSearch(self, brokerName, companyName, page):
        """组合 broker/company 关键词搜索联系方式，失败时降级为公司关键词。"""
        if brokerName and companyName:
            query = f'"{brokerName}" "{companyName}" email phone'
        elif brokerName:
            query = f'"{brokerName}" real estate broker email phone Texas'
        else:
            return [], [], [], []

        emails, phones, companyEmails, brokerEmails, text = self.searchGoogle(query, page)

        if not emails and companyName:
            fallbackQuery = f'"{companyName}" "office email"'
            print("个人搜索无邮箱，改用公司关键词:", fallbackQuery)
            emails2, phones2, companyEmails2, brokerEmails2, text2 = self.searchGoogle(
                fallbackQuery, page
            )
            emails = emails or emails2
            phones = phones or phones2
            companyEmails = companyEmails or companyEmails2
            brokerEmails = brokerEmails or brokerEmails2

        return emails, phones, companyEmails, brokerEmails

    @staticmethod
    def formatResult(code, companyName, brokerName, emails, phones, companyEmails, brokerEmails):
        """格式化一条搜索结果，方便导出 CSV 和 Excel。"""
        return {
            "code": code,
            "company_name": companyName,
            "broker_name": brokerName,
            "emails": "; ".join(emails),
            "phones": "; ".join(phones),
            "company_email": "; ".join(companyEmails),
            "broker_email": "; ".join(brokerEmails),
        }

    def searchFlow(self, csvPath):
        """读取中间 CSV，逐个搜索 broker/company 联系方式。"""
        companies = []
        path = Path(csvPath)
        if not path.exists():
            print("中间 CSV 不存在:", path)
            return companies

        # 读取待搜索公司和 broker 信息
        with path.open("r", newline="", encoding="utf-8-sig") as file:
            for row in csv.DictReader(file):
                recordType = (row.get("record_type") or "").strip()
                relationName = (row.get("relation_name") or "").strip()
                name = (row.get("name") or "").strip()
                code = (row.get("code") or "").strip()

                if recordType == "Broker Company" and relationName:
                    companies.append({
                        "company_name": name,
                        "broker_name": relationName,
                        "code": code,
                    })
                elif recordType == "Sales Agent" and relationName:
                    companies.append({
                        "company_name": relationName,
                        "broker_name": name,
                        "code": code,
                    })

        if not companies:
            print("没有待搜索公司")
            return []

        print("读取待搜索公司数量:", len(companies))

        browserOptions = ChromiumOptions()
        browserOptions.set_argument("--incognito")
        browserOptions.set_user_agent(self.userAgent)

        page = ChromiumPage(browserOptions)
        page.get("https://www.google.com")
        time.sleep(2)

        results = []
        total = len(companies)
        if self.maxBrokers:
            total = min(total, int(self.maxBrokers))

        for index, company in enumerate(companies[:total]):
            code = company["code"]
            companyName = company["company_name"]
            brokerName = company["broker_name"]

            print("开始搜索:", index + 1, "/", total, companyName, brokerName)
            emails, phones, companyEmails, brokerEmails = self.brokerSearch(
                brokerName, companyName, page
            )

            results.append(self.formatResult(
                code, companyName, brokerName,
                emails, phones, companyEmails, brokerEmails,
            ))

            print(
                "搜索完成:", index + 1, "/", total,
                "邮箱", len(emails), "电话", len(phones)
            )
            if index < total - 1:
                time.sleep(random.uniform(3, 6))

        page.quit()
        return results

    def saveResults(self, results, pathStr):
        """保存搜索结果到 CSV。"""
        if not results:
            print("没有搜索结果可保存")
            return

        fieldnames = [
            "code", "company_name", "broker_name",
            "emails", "phones", "company_email", "broker_email",
        ]
        path = Path(pathStr)
        path.parent.mkdir(parents=True, exist_ok=True)

        with path.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)

        print("搜索结果 CSV 已保存:", path, "数量:", len(results))

    def exportExcel(self, results, pathStr):
        """保存搜索结果到 Excel。"""
        if not results:
            print("没有搜索结果可导出 Excel")
            return

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Broker Contacts"

        headers = [
            "code", "company_name", "broker_name",
            "emails", "phones", "company_email", "broker_email",
        ]
        worksheet.append(headers)

        for row in results:
            worksheet.append([row.get(header, "") for header in headers])

        worksheet.auto_filter.ref = worksheet.dimensions

        path = Path(pathStr)
        path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(str(path))
        print("搜索结果 Excel 已保存:", path, "数量:", len(results))

    def main(self):
        """按旧流程顺序执行：列表采集、关系融合、导出、搜索、最终导出。"""
        self.outputDir.mkdir(parents=True, exist_ok=True)
        listCsvPath = self.outputDir / "trec_output.csv"
        resultCsvPath = self.outputDir / "broker_contacts.csv"
        resultExcelPath = self.outputDir / "broker_contacts.xlsx"

        if not self.searchOnly:
            self.collectList()
            self.exportRows(listCsvPath)
            self.mergeData()
            self.exportRows(listCsvPath)

        results = self.searchFlow(str(listCsvPath))
        self.saveResults(results, resultCsvPath)
        self.exportExcel(results, resultExcelPath)

        print(json.dumps({
            "rows": len(self.rows),
            "results": len(results),
        }, ensure_ascii=False))

    def run(self):
        """测试入口，只负责调用 main。"""
        self.main()


if __name__ == "__main__":
    # 单文件测试配置：
    # maxPage=5 表示先采集前 5 页；如需全量采集，改为 0。
    # maxBrokers=None 表示搜索全部 broker；调试时建议改成较小数字。
    config = {
        "outputDir": "output",
        "maxPage": 5,
        "maxBrokers": None,
        "searchOnly": False,
        "perPage": 100,
        "waitSecond": 6,
        "retryCount": 5,
        "openDataLimit": 50000,
        "openDataOffset": 0,
        "maxOpenDataPage": 100,
    }
    Test(config).run()

