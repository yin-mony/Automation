"""通过 DrissionPage 提取普通网页，并记录 Facebook 主页链接。"""

import json
import re
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

from proxy import Proxy


class Browser:
    """复用持久浏览器档案，并强制所有页面经过已验证代理桥。"""

    def __init__(
        self,
        baseDir,
        proxy,
        config,
        checkpoint,
        log=None,
        human=None,
    ):
        """初始化页面规则、联系方式规则和运行回调。"""
        self.baseDir = Path(baseDir)
        self.proxy = proxy
        self.config = config
        self.checkpoint = checkpoint
        self.log = log or (lambda message: None)
        self.human = human
        self.profileDir = self.baseDir / "browser_profiles"
        self.profileDir.mkdir(parents=True, exist_ok=True)
        self.pages = {}
        self.verifiedProfiles = set()
        self.pageWait = max(1, int(config.get("pageWait", 4)))
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
        self.blockedExtensions = {
            ".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".zip", ".mp4", ".mp3",
        }
        self.challengeMarkers = (
            "verify you are human", "checking your browser", "security check", "unusual traffic",
            "complete the security check", "captcha", "cloudflare ray id", "human verification",
        )

    def dependencyStatus(self):
        """检查当前 Python 环境是否可以启动 DrissionPage。"""
        try:
            from DrissionPage import ChromiumOptions, ChromiumPage

            del ChromiumOptions, ChromiumPage
            return True, "DrissionPage 可用"
        except Exception as error:
            return False, str(error)

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
            if len(set(digits)) == 1 or digits[3:6] == "555":
                continue
            if digits in {"1234567890", "0123456789", "5555555555"}:
                continue
            formatted = f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
            if formatted not in output:
                output.append(formatted)
        return output

    def isFacebook(self, url):
        """判断链接是否属于 Facebook。"""
        host = (urlparse(str(url or "")).hostname or "").lower()
        return host == "facebook.com" or host.endswith(".facebook.com")

    def usableFacebook(self, url):
        """排除登录、搜索、帖子和媒体等非主页链接。"""
        if not self.isFacebook(url):
            return False
        path = urlparse(url).path.lower()
        blocked = (
            "/login", "/search", "/share", "/groups/", "/events/", "/posts/",
            "/photos/", "/videos/",
        )
        return bool(path and path != "/" and not any(item in path for item in blocked))

    def isOrdinary(self, url):
        """判断链接是否适合进入普通二级页面提取。"""
        parsed = urlparse(str(url or ""))
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            return False
        host = parsed.hostname.lower()
        if host == "google.com" or host.endswith(".google.com"):
            return False
        if any(host == item or host.endswith("." + item) for item in self.socialHosts):
            return False
        return not any(parsed.path.lower().endswith(item) for item in self.blockedExtensions)

    def requestHuman(self, reason, url):
        """暂停自动流程，等待用户在可见浏览器中完成人机验证。"""
        if not self.human:
            raise RuntimeError(f"需要人工处理：{reason}，当前地址：{url}")
        self.human(reason, url)
        self.checkpoint()

    def ensurePage(self, profile):
        """按当前网络模式创建普通网页的可见浏览器。"""
        if profile in self.pages:
            return self.pages[profile]
        localProxy = self.proxy.start() if self.proxy else ""
        if self.proxy and self.proxy.required and not localProxy:
            raise RuntimeError("浏览器代理桥未启动，已禁止直连")
        try:
            from DrissionPage import ChromiumOptions, ChromiumPage
        except Exception as error:
            raise RuntimeError(f"DrissionPage 启动失败：{error}") from error
        options = ChromiumOptions()
        options.set_user_data_path(str(self.profileDir / profile))
        options.set_load_mode("eager")
        options.headless(False)
        if localProxy:
            options.set_proxy(localProxy)
            self.log(f"启动可见浏览器：{profile}，代理出口已锁定。")
        else:
            self.log(f"启动可见浏览器：{profile}，当前使用直接访问。")
        page = ChromiumPage(options)
        self.pages[profile] = page
        if localProxy:
            try:
                self.verifyExit(page, profile)
            except Exception:
                try:
                    page.quit()
                finally:
                    self.pages.pop(profile, None)
                raise
        return page

    def verifyExit(self, page, profile):
        """在 DrissionPage 内再次确认出口 IP 与代理桥一致。"""
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
        expected = str(self.proxy.exitIp if self.proxy else "")
        if not exitIp or not expected or exitIp != expected:
            raise RuntimeError("DrissionPage 出口 IP 与代理桥不一致，已禁止继续访问目标页面")
        self.verifiedProfiles.add(profile)

    def pageText(self, page):
        """读取当前页面可见正文。"""
        body = page.ele("tag:body", timeout=5)
        return str(body.text if body else getattr(page, "html", "") or "")

    def pageLinks(self, page, currentUrl):
        """收集当前页面全部可解析超链接。"""
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
        """打开普通网页并读取可见正文和链接。"""
        self.checkpoint()
        page = self.ensurePage("web")
        page.get(url)
        time.sleep(self.pageWait)
        self.checkpoint()
        text = self.pageText(page)
        currentUrl = str(getattr(page, "url", "") or url)
        challenged = any(marker in text.lower() for marker in self.challengeMarkers)
        if challenged:
            self.requestHuman("检测到人机验证，请在浏览器完成后点击继续", currentUrl)
            text = self.pageText(page)
            challenged = any(marker in text.lower() for marker in self.challengeMarkers)
        return {
            "url": currentUrl,
            "text": text,
            "links": self.pageLinks(page, currentUrl),
            "challenged": challenged,
        }

    def collect(
        self,
        payload,
        mode,
        candidate,
    ):
        """从搜索摘要和普通页面汇总联系方式，同时记录 Facebook 链接。"""
        organic = payload.get("organic_results") or []
        searchText = []
        sourceUrls = []
        facebookUrls = []
        ordinaryUrls = []
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
            elif self.isOrdinary(link):
                ordinaryUrls.append(link)

        emails = self.cleanEmails(" ".join(searchText))
        phones = self.cleanPhones(" ".join(searchText))
        detailUrls = []
        pageErrors = []
        maximumPages = max(0, int(self.config.get("maxDetailPages", 6)))
        browserOk, browserReason = self.dependencyStatus()
        if browserOk:
            for link in self.unique(ordinaryUrls)[:maximumPages]:
                self.checkpoint()
                try:
                    page = self.openPage(link)
                    if page["text"]:
                        detailUrls.append(page["url"])
                        emails.extend(self.cleanEmails(page["text"]))
                        phones.extend(self.cleanPhones(page["text"]))
                        facebookUrls.extend(
                            item for item in page["links"] if self.usableFacebook(item)
                        )
                except Exception as error:
                    pageErrors.append(f"{link}: {str(error)[:120]}")
                    self.log(f"二级页面抓取失败：{link}，{str(error)[:120]}")
        elif ordinaryUrls:
            pageErrors.append("浏览器不可用：" + browserReason[:160])

        emails = self.unique(emails)
        phones = self.unique(phones)
        return {
            "emails": emails,
            "phones": phones,
            "facebookUrls": self.unique(facebookUrls),
            "sourceUrls": self.unique(sourceUrls),
            "detailUrls": self.unique(detailUrls),
            "resultCount": len(organic),
            "contactStatus": "已找到联系方式" if emails or phones else "未找到联系方式",
            "pageErrors": pageErrors,
            "searchId": str((payload.get("search_metadata") or {}).get("id") or ""),
        }

    def close(self):
        """关闭全部浏览器页面，代理桥由 Main 统一停止。"""
        for page in list(self.pages.values()):
            try:
                page.quit()
            except Exception:
                pass
        self.pages.clear()
        self.verifiedProfiles.clear()
