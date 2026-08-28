"""提取网页联系方式：直连抓取公开页面，并核对公司名称。"""

import html as html_lib
import json
import re
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen


class Browser:
    """使用直连 HTTP 抓取普通网页，Facebook 只保存公开链接。"""

    def __init__(self, baseDir, proxy, config, checkpoint, log=None):
        """初始化页面规则、联系方式规则和运行回调。"""
        self.baseDir = Path(baseDir)
        self.proxy = proxy
        self.config = config
        self.checkpoint = checkpoint
        self.log = log or (lambda message: None)
        self.profileDir = self.baseDir / "browser_profiles"
        self.profileDir.mkdir(parents=True, exist_ok=True)
        self.pages = {}
        self.verifiedProfiles = set()
        self.pageWait = max(1, int(config.get("pageWait", 4)))
        self.useBrowser = False
        self.useDirectFallback = True
        self.emailPattern = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
        self.phonePattern = re.compile(r"\+?1?[\s.-]?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}")
        self.publicEmailNames = {
            "admin", "administrator", "info", "information", "contact", "contactus", "support",
            "help", "hello", "office", "mail", "email", "team", "sales", "marketing",
            "service", "services", "customerservice", "customer.service", "webmaster",
            "postmaster", "abuse", "privacy", "legal", "compliance", "noreply", "no-reply",
            "donotreply", "do-not-reply", "reception", "receptionist", "billing", "accounting",
            "hr", "careers", "jobs", "press", "media",
        }
        self.blockedEmailWords = {
            "army", "navy", "airforce", "marines", "military", "defense", "dod", "usaf",
            "usarmy", "uscg", "police", "sheriff", "school", "k12", "isd",
        }
        self.socialHosts = {
            "facebook.com", "instagram.com", "linkedin.com", "youtube.com", "x.com", "twitter.com",
        }
        self.directoryHosts = {
            "brokercheck.finra.org", "adviserinfo.sec.gov", "sipc.org", "finra.org",
            "manta.com", "zoominfo.com", "yelp.com", "yellowpages.com", "bbb.org",
            "opencorporates.com", "buzzfile.com", "dandb.com", "dnb.com",
        }
        self.blockedExtensions = {
            ".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".zip", ".mp4", ".mp3",
        }
        self.challengeMarkers = (
            "verify you are human", "checking your browser", "security check", "unusual traffic",
            "complete the security check", "captcha", "cloudflare ray id", "human verification",
        )
        self.corporateSuffixes = {
            "inc", "incorporated", "llc", "ltd", "limited", "corp", "corporation",
            "co", "company", "lp", "llp", "pllc", "pc", "dba", "lllp", "l.l.c", "l.l.p",
        }
        self.stopwords = {"the", "a", "an", "of", "for", "and", "to", "at", "in", "on", "by"}

    # ---------- 名称核对 ----------

    def nameKey(self, value):
        """把名称转为小写字母数字串。"""
        text = str(value or "").lower()
        text = re.sub(r"[^a-z0-9 ]", " ", text)
        return " ".join(text.split())

    def stripSuffixes(self, key):
        """去掉公司名结尾的法人后缀。"""
        tokens = key.split()
        while tokens and tokens[-1] in self.corporateSuffixes:
            tokens.pop()
        return " ".join(tokens)

    def distinctiveTokens(self, name):
        """返回公司名中用于核对的显著词。"""
        key = self.stripSuffixes(self.nameKey(name))
        return [token for token in key.split() if token not in self.stopwords and len(token) >= 2]

    def nameMatch(self, name, text):
        """判断页面或摘要文本是否与该公司的名称一致。"""
        key = self.stripSuffixes(self.nameKey(name))
        textKey = self.nameKey(text)
        if not key:
            return False
        if key in textKey:
            return True
        tokens = self.distinctiveTokens(name)
        if not tokens:
            return False
        hits = [token for token in tokens if token in textKey]
        if len(tokens) == 1:
            return len(hits) == 1
        if len(tokens) == 2:
            return len(hits) == 2
        return len(hits) >= max(2, round(len(tokens) * 0.6))

    # ---------- 联系方式清洗 ----------

    def unique(self, values):
        """按出现顺序去重非空文本。"""
        output = []
        seen = set()
        for value in values:
            text = str(value or "").strip()
            if text and text not in seen:
                output.append(text)
                seen.add(text)
        return output

    def cleanEmails(self, text):
        """清理无效、公共职能、政府教育和军事域名邮箱。"""
        output = []
        for email in self.emailPattern.findall(str(text or "")):
            email = email.lower().strip("._-")
            if "@" not in email or ".." in email:
                continue
            name, domain = email.split("@", 1)
            if name in self.publicEmailNames:
                continue
            if domain.endswith((".gov", ".mil", ".edu")):
                continue
            if any(word in domain for word in self.blockedEmailWords):
                continue
            if email not in output:
                output.append(email)
        return output

    def cleanPhones(self, text):
        """统一美国十位电话格式并排除明显测试号码。"""
        output = []
        for phone in self.phonePattern.findall(str(text or "")):
            digits = re.sub(r"\D", "", phone)
            if len(digits) == 11 and digits.startswith("1"):
                digits = digits[1:]
            if len(digits) != 10:
                continue
            if digits[0] in "01":
                continue
            if len(set(digits)) == 1 or digits[3:6] == "555":
                continue
            if digits in {"1234567890", "0123456789", "5555555555"}:
                continue
            formatted = f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
            if formatted not in output:
                output.append(formatted)
        return output

    # ---------- 链接判断 ----------

    def isFacebook(self, url):
        host = (urlparse(str(url or "")).hostname or "").lower()
        return host == "facebook.com" or host.endswith(".facebook.com")

    def usableFacebook(self, url):
        if not self.isFacebook(url):
            return False
        path = urlparse(url).path.lower()
        blocked = ("/login", "/search", "/share", "/groups/", "/events/", "/posts/", "/photos/", "/videos/")
        return bool(path and path != "/" and not any(item in path for item in blocked))

    def isOrdinary(self, url):
        parsed = urlparse(str(url or ""))
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            return False
        host = parsed.hostname.lower()
        if host == "google.com" or host.endswith(".google.com"):
            return False
        if any(host == item or host.endswith("." + item) for item in self.socialHosts):
            return False
        if any(host == item or host.endswith("." + item) for item in self.directoryHosts):
            return False
        return not any(parsed.path.lower().endswith(item) for item in self.blockedExtensions)

    # ---------- 旧浏览器入口（当前流程固定直连） ----------

    def dependencyStatus(self):
        return True, "直连模式，无需浏览器代理"

    def ensurePage(self, profile):
        raise RuntimeError("当前版本固定使用直连抓取，未启用可见浏览器")

    def verifyExit(self, page, profile):
        if profile in self.verifiedProfiles:
            return
        page.get("https://api.ipify.org?format=json")
        time.sleep(1)
        body = page.ele("tag:body", timeout=5)
        bodyText = str(body.text if body else getattr(page, "html", "") or "")
        try:
            exitIp = str(json.loads(bodyText).get("ip") or "").strip()
        except Exception:
            match = re.search(r"(?:\d{1,3}\.){3}\d{1,3}", bodyText)
            exitIp = match.group(0) if match else ""
        if not exitIp:
            raise RuntimeError("浏览器未返回公网出口 IP")
        self.verifiedProfiles.add(profile)

    def pageText(self, page):
        body = page.ele("tag:body", timeout=5)
        return str(body.text if body else getattr(page, "html", "") or "")

    def pageLinks(self, page, currentUrl):
        links = []
        try:
            for anchor in page.eles("tag:a"):
                href = str(anchor.attr("href") or "").strip()
                if href:
                    links.append(urljoin(currentUrl, href))
        except Exception:
            pass
        return self.unique(links)

    def openPage(self, url):
        self.checkpoint()
        page = self.ensurePage("web")
        page.get(url)
        time.sleep(self.pageWait)
        self.checkpoint()
        text = self.pageText(page)
        currentUrl = str(getattr(page, "url", "") or url)
        challenged = any(marker in text.lower() for marker in self.challengeMarkers)
        if challenged:
            self.log(f"检测到人机验证，跳过该页面：{currentUrl}")
            return {"url": currentUrl, "text": "", "links": [], "challenged": True}
        return {"url": currentUrl, "text": text, "links": self.pageLinks(page, currentUrl), "challenged": challenged}

    # ---------- 直连抓取模式 ----------

    def htmlToText(self, htmlText):
        text = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", htmlText)
        text = re.sub(r"(?s)<[^>]+>", " ", text)
        text = html_lib.unescape(text)
        return re.sub(r"\s+", " ", text)

    def extractLinks(self, htmlText, baseUrl):
        links = []
        for match in re.finditer(r'''href\s*=\s*["']([^"']+)["']''', htmlText, re.IGNORECASE):
            href = match.group(1).strip()
            if href:
                links.append(urljoin(baseUrl, href))
        return self.unique(links)

    def fetchDirect(self, url, timeout=20):
        self.checkpoint()
        request = Request(url, headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/151.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        })
        response = urlopen(request, timeout=timeout)
        raw = response.read()
        charset = "utf-8"
        match = re.search(r"charset=([\w-]+)", response.headers.get("Content-Type", ""))
        if match:
            charset = match.group(1)
        try:
            htmlText = raw.decode(charset, errors="replace")
        except Exception:
            htmlText = raw.decode("utf-8", errors="replace")
        finalUrl = response.geturl()
        links = self.extractLinks(htmlText, finalUrl)
        text = self.htmlToText(htmlText)
        mailtos = [m.group(1) for m in re.finditer(r'''mailto:([^\s"'<>]+)''', htmlText, re.IGNORECASE)]
        tels = re.findall(r"tel:([+\d][\d\-.()\s]{5,})", htmlText, re.IGNORECASE)
        text += " " + " ".join(mailtos) + " " + " ".join(tels)
        challenged = any(marker in text.lower() for marker in self.challengeMarkers)
        return {"url": finalUrl, "text": text, "links": links, "challenged": challenged}

    # ---------- 汇总 ----------

    def collect(self, payload, candidate):
        """从搜索摘要和普通页面汇总联系方式，并核对公司名称。"""
        name = str(candidate.get("name") or "")
        organic = payload.get("organic_results") or []
        searchText = []
        sourceUrls = []
        facebookUrls = []
        matchedUrls = []
        otherUrls = []
        for item in organic:
            title = str(item.get("title") or "")
            link = str(item.get("link") or "").strip()
            snippet = str(item.get("snippet") or "")
            rich = json.dumps(item.get("rich_snippet") or "", ensure_ascii=False)
            searchText.extend((title, snippet, rich))
            if not link:
                continue
            sourceUrls.append(link)
            if self.usableFacebook(link):
                facebookUrls.append(link)
                continue
            if not self.isOrdinary(link):
                continue
            if self.nameMatch(name, title + " " + snippet):
                matchedUrls.append(link)
            else:
                otherUrls.append(link)

        emails = self.cleanEmails(" ".join(searchText))
        phones = self.cleanPhones(" ".join(searchText))
        detailUrls = []
        verifiedUrls = []
        pageErrors = []
        maximumPages = max(0, int(self.config.get("maxDetailPages", 6)))
        urlsToFetch = self.unique(matchedUrls + otherUrls)[:maximumPages]

        fetchKind = None
        if self.useBrowser:
            browserOk, reason = self.dependencyStatus()
            if browserOk:
                fetchKind = "browser"
            else:
                pageErrors.append("浏览器不可用：" + reason[:160])
                if self.useDirectFallback:
                    fetchKind = "direct"
        if fetchKind is None and self.useDirectFallback:
            fetchKind = "direct"
        if fetchKind is None:
            fetchKind = "none"

        if fetchKind != "none":
            self.log(f"二级页面抓取方式：{fetchKind}")
        for link in urlsToFetch:
            self.checkpoint()
            page = {"url": link, "text": "", "links": [], "challenged": False}
            try:
                if fetchKind == "browser":
                    page = self.openPage(link)
                elif fetchKind == "direct":
                    page = self.fetchDirect(link)
            except Exception as error:
                pageErrors.append(f"{link}: {str(error)[:120]}")
                if fetchKind == "browser":
                    self.log(f"浏览器抓取失败，切换到直连：{str(error)[:120]}")
                    fetchKind = "direct" if self.useDirectFallback else "none"
                    if fetchKind == "direct":
                        try:
                            page = self.fetchDirect(link)
                        except Exception as error2:
                            pageErrors.append(f"{link}: {str(error2)[:120]}")
                            continue
                    else:
                        continue
                else:
                    continue
            if page["text"]:
                detailUrls.append(page["url"])
                if self.nameMatch(name, page["text"]):
                    verifiedUrls.append(page["url"])
                    emails.extend(self.cleanEmails(page["text"]))
                    phones.extend(self.cleanPhones(page["text"]))
                facebookUrls.extend(item for item in page["links"] if self.usableFacebook(item))

        emails = self.unique(emails)
        phones = self.unique(phones)
        return {
            "emails": emails,
            "phones": phones,
            "verified": bool(verifiedUrls),
            "verifiedUrls": self.unique(verifiedUrls),
            "facebookUrls": self.unique(facebookUrls),
            "sourceUrls": self.unique(sourceUrls),
            "detailUrls": self.unique(detailUrls),
            "matchedResultCount": len(self.unique(matchedUrls)),
            "resultCount": len(organic),
            "contactStatus": "已找到联系方式" if emails or phones else "未找到联系方式",
            "pageErrors": pageErrors,
            "fetchMode": fetchKind,
            "searchId": str((payload.get("search_metadata") or {}).get("id") or ""),
        }

    def close(self):
        for page in list(self.pages.values()):
            try:
                page.quit()
            except Exception:
                pass
        self.pages.clear()
        self.verifiedProfiles.clear()
