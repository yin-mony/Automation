"""易得客客户端登录与 Amazon Seller Central 登录接管。"""

import ctypes
import os
import socket
import subprocess
import time
from pathlib import Path

import psutil
from DrissionPage import ChromiumPage
from pywinauto import Desktop


class YidekeLogin:
    """易得客客户端启动、登录与 Amazon Seller Central 登录接管。"""

    def __init__(self, username, password):
        """启动易得客客户端并准备登录账号。"""
        self.username = username
        self.password = password
        exePath = self.resolveEdecker()

        # 启动前关闭旧易得客进程，避免接管历史窗口
        for proc in psutil.process_iter(["name"]):
            if proc.info["name"] and proc.info["name"].lower() == "edecker.exe":
                proc.kill()

        # 清理固定调试端口，保证新客户端能正常开启远程调试
        self.killDebugPort(9222)
        time.sleep(1)
        userDataPath = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "User Data"
        self.edeckerProcess = subprocess.Popen([
            exePath,
            "--remote-debugging-port=9222",
            f"--user-data-dir={userDataPath}",
        ])

    def killDebugPort(self, port):
        """结束占用指定调试端口的浏览器进程，避免连到非易得客页面。"""
        flag = f"--remote-debugging-port={port}"
        for proc in psutil.process_iter(["pid", "cmdline"]):
            try:
                cmdline = proc.info.get("cmdline") or []
                if any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                pass

    def resolveEdecker(self):
        """按易得客默认安装目录查找 edecker.exe。"""
        localPath = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "Application"
        directPath = localPath / "edecker.exe"
        if directPath.exists():
            return str(directPath)
        if localPath.exists():
            versions = sorted(path for path in localPath.iterdir() if path.is_dir())
            if versions:
                exePath = versions[-1] / "edecker.exe"
                if exePath.exists():
                    return str(exePath)
        raise FileNotFoundError(f"edecker.exe not found under {localPath}")

    def login(self, maxRetries=3, retryInterval=2):
        """按顺序登录易得客客户端。"""
        lastError = None
        for attempt in range(1, maxRetries + 1):
            try:
                if attempt > 1:
                    try:
                        ChromiumPage("127.0.0.1:9222").refresh()
                    except Exception as refreshError:
                        print(f"重试前刷新浏览器失败，继续尝试登录: {refreshError}")

                # 第一步：等待易得客调试端口启动完成
                deadline = time.time() + 60
                while time.time() < deadline:
                    try:
                        with socket.create_connection(("127.0.0.1", 9222), timeout=2):
                            break
                    except OSError:
                        time.sleep(1)
                else:
                    raise RuntimeError("等待易得客调试端口 127.0.0.1:9222 超时")

                page = ChromiumPage("127.0.0.1:9222")

                # 第二步：如果直接进入网页版登录页，就在当前页输入账号密码
                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if "登录你的易得客" in pageText or "登录易得客" in pageText:
                    phoneInput = page.ele('x://input[@placeholder="请输入手机号"] | //input[@type="text"]', timeout=15)
                    passwordInput = page.ele('x://input[@placeholder="请输入登录密码"] | //input[@type="password"]', timeout=15)
                    if not phoneInput or not passwordInput:
                        raise RuntimeError("易得客网页登录页未找到手机号或密码输入框")
                    phoneInput.input(self.username, clear=True)
                    passwordInput.input(self.password, clear=True)
                    print("易得客网页登录账号密码输入成功", flush=True)
                    loginButton = page.ele(
                        'x://button[contains(normalize-space(),"登录易得客")]'
                        ' | //span[normalize-space()="登录易得客"]/ancestor::button[1]',
                        timeout=15,
                    )
                    if not loginButton:
                        raise RuntimeError("易得客网页登录页未找到登录按钮")
                    loginButton.click()
                    for waitCount in range(60):
                        time.sleep(1)
                        try:
                            currentText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                        except Exception:
                            currentText = ""
                        if "登录你的易得客" not in currentText and "登录易得客" not in currentText:
                            print("易得客网页登录完成", flush=True)
                            return
                    raise RuntimeError("易得客网页登录后未离开登录页")

                # 第三步：主页存在“登录”入口时先点击，点击后可能仍是网页版登录页
                loginElement = page.ele('x://span[text()="登录"]', timeout=60)
                if not loginElement:
                    raise RuntimeError(f'在易得客浏览器当前页面未找到 XPath: x://span[text()="登录"]，当前URL: {page.url}')
                loginElement.click()
                time.sleep(5)
                try:
                    page = page.browser.latest_tab
                except Exception:
                    pass

                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if "登录你的易得客" in pageText or "登录易得客" in pageText:
                    phoneInput = page.ele('x://input[@placeholder="请输入手机号"] | //input[@type="text"]', timeout=15)
                    passwordInput = page.ele('x://input[@placeholder="请输入登录密码"] | //input[@type="password"]', timeout=15)
                    if not phoneInput or not passwordInput:
                        raise RuntimeError("易得客网页登录页未找到手机号或密码输入框")
                    phoneInput.input(self.username, clear=True)
                    passwordInput.input(self.password, clear=True)
                    print("易得客网页登录账号密码输入成功", flush=True)
                    loginButton = page.ele(
                        'x://button[contains(normalize-space(),"登录易得客")]'
                        ' | //span[normalize-space()="登录易得客"]/ancestor::button[1]',
                        timeout=15,
                    )
                    if not loginButton:
                        raise RuntimeError("易得客网页登录页未找到登录按钮")
                    loginButton.click()
                    for waitCount in range(60):
                        time.sleep(1)
                        try:
                            currentText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                        except Exception:
                            currentText = ""
                        if "登录你的易得客" not in currentText and "登录易得客" not in currentText:
                            print("易得客网页登录完成", flush=True)
                            return
                    raise RuntimeError("易得客网页登录后未离开登录页")

                # 第四步：不是网页版登录页时，按 FBA 子项目方式处理 Windows 原生登录弹窗
                deadline = time.time() + 30
                dialog = None
                lastWindowError = None
                while time.time() < deadline:
                    try:
                        candidates = []
                        desktop = Desktop(backend="uia")
                        for win in desktop.windows(visible_only=False):
                            try:
                                title = win.window_text() or ""
                                if "易得客浏览器" not in title:
                                    continue
                                score = 0
                                if win.is_visible():
                                    score += 3
                                if win.child_window(title="登录易得客", control_type="Button").exists(timeout=0.2):
                                    score += 5
                                if win.child_window(title="手机号", control_type="Text").exists(timeout=0.2):
                                    score += 2
                                if win.child_window(title="密码", control_type="Text").exists(timeout=0.2):
                                    score += 2
                                if score >= 5:
                                    candidates.append((score, win))
                            except Exception:
                                pass
                        if not candidates:
                            raise RuntimeError("未找到包含登录控件的易得客窗口")
                        candidates.sort(key=lambda item: item[0], reverse=True)
                        dialog = candidates[0][1]
                        dialog.set_focus()
                        dialog.wait("visible ready", timeout=15)
                        break
                    except Exception as exc:
                        lastWindowError = exc
                        time.sleep(1)
                if not dialog:
                    raise RuntimeError("未找到包含登录控件的易得客登录窗口") from lastWindowError

                for labelTitle, text in (("手机号", self.username), ("密码", self.password)):
                    label = dialog.child_window(title=labelTitle, control_type="Text")
                    if not label.exists():
                        continue
                    labelRect = label.rectangle()
                    allEdits = dialog.descendants(control_type="Edit")
                    for edit in allEdits:
                        editRect = edit.rectangle()
                        if editRect.top > labelRect.bottom and abs(editRect.left - labelRect.left) < 50:
                            edit.set_text(text)
                            print(f"{labelTitle}输入成功")
                            break
                dialog.child_window(title="登录易得客", control_type="Button").click()
                return
            except Exception as exc:
                lastError = exc
                if attempt >= maxRetries:
                    break
                print(f"YidekeLogin 第{attempt}次尝试失败，{retryInterval}秒后重试: {exc}")
                time.sleep(retryInterval)
        raise RuntimeError(f"YidekeLogin 重试{maxRetries}次后仍失败") from lastError

    class AmazonSeller:
        """Amazon 卖家中心登录，接管店铺 profile 浏览器内的登录流程。"""

        def __init__(self, page, storePassword=None, siteName="United States"):
            """准备 Amazon 登录页、账号选择页和站点名称。"""
            self.page = page
            self.storePassword = storePassword or ""
            self.siteName = siteName or "United States"

        def fillCode(self):
            """通过易得客验证码插件窗口填入 Amazon 二步验证码。"""
            time.sleep(1)
            desktop = Desktop(backend="uia")
            success = False
            deadline = time.time() + 60
            for win in desktop.windows():
                try:
                    # 先找到验证码服务按钮，再轮询填入验证码或获取最新验证码按钮
                    for button in win.descendants(control_type="Button"):
                        if "二步验证码服务" in button.window_text():
                            button.click_input()
                            time.sleep(1.5)
                            while time.time() < deadline:
                                found = False
                                for codeButton in win.descendants(control_type="Button"):
                                    name = codeButton.window_text()
                                    if name == "填入验证码":
                                        codeButton.click_input()
                                        time.sleep(1.5)
                                        codeButton.click_input()
                                        found = True
                                        success = True
                                        break
                                    if name == "获取最新验证码":
                                        codeButton.click_input()
                                        time.sleep(1)
                                        codeButton.click_input()
                                        found = True
                                        break
                                if success:
                                    break
                                if not found:
                                    time.sleep(0.2)
                            break
                    if success:
                        break
                except Exception as exc:
                    print(exc)
            return success

        def waitPassword(self, passwordInput):
            """等待 Amazon 密码框自动填充，必要时提示用户手动确认。"""
            if not passwordInput:
                raise RuntimeError("未找到 Amazon 密码输入框")

            if self.storePassword:
                passwordInput.input(self.storePassword, clear=True)
                return True

            # 未配置 Amazon 密码时，等待浏览器保存的密码自动填充
            try:
                passwordInput.click()
            except Exception:
                pass
            for waitCount in range(20):
                value = passwordInput.attr("value") or ""
                if value:
                    return True
                time.sleep(0.5)

            # 自动填充未完成时，提示用户手动选择浏览器保存的密码
            ctypes.windll.user32.MessageBoxW(
                0,
                "Amazon 密码未自动填入。\n请在浏览器密码框中手动选择已保存密码或输入密码，完成后点击“确定”继续。",
                "Amazon 密码确认",
                0x40 | 0x40000,
            )

            # 用户确认后再次等待密码框有值，避免空密码直接提交
            for waitCount in range(240):
                value = passwordInput.attr("value") or ""
                if value:
                    return True
                time.sleep(0.5)
            return False

        def login(self, email=None, password=None, timeout=120):
            """按顺序处理 Amazon 登录页、账号选择页与二步验证码。"""
            if password:
                self.storePassword = password

            # 第一步：在已打开标签中挑选 Seller Central 或 Amazon 登录页
            page = self.page
            backendTab = None
            loginTab = None
            try:
                browser = page.browser
                for tabId in browser.tab_ids:
                    tab = browser.get_tab(tabId)
                    url = (tab.url or "").lower()
                    if url.startswith("chrome-extension://"):
                        continue
                    if "sellercentral.amazon" in url and "/ap/signin" not in url:
                        backendTab = tab
                        break
                    if "/ap/signin" in url or "sellercentral.amazon" in url:
                        loginTab = tab
                if backendTab or loginTab:
                    page = backendTab or loginTab
                    self.page = page
            except Exception:
                pass

            # 第二步：如果停留在邮箱页，先点击输入框，优先触发浏览器保存账号记录
            emailInput = page.ele('x://input[@id="ap_email"] | //input[@name="email"] | //input[@type="email"]', timeout=5)
            continueButton = page.ele('x://input[@id="continue"] | //button[@id="continue"]', timeout=5)
            if emailInput and continueButton:
                try:
                    emailInput.click()
                except Exception:
                    pass
                if email:
                    emailInput.input(email, clear=True)
                else:
                    emailValue = emailInput.attr("value") or ""
                    for waitCount in range(20):
                        if emailValue:
                            break
                        time.sleep(0.5)
                        emailValue = emailInput.attr("value") or ""
                    if not emailValue:
                        ctypes.windll.user32.MessageBoxW(
                            0,
                            "Amazon 账号未自动填入。\n请在当前 Amazon 邮箱/手机号输入框中选择浏览器保存记录或手动输入，完成后点击“确定”继续。",
                            "Amazon 账号确认",
                            0x40 | 0x40000,
                        )
                        for waitCount in range(240):
                            emailValue = emailInput.attr("value") or ""
                            if emailValue:
                                break
                            time.sleep(0.5)
                    if not emailValue:
                        raise RuntimeError("Amazon 账号未填入，已停止登录")
                continueButton.click()
                time.sleep(5)

            # 第三步：如果停留在密码页，补齐密码并提交登录
            passwordInput = page.ele('x://input[@type="password"]', timeout=10)
            if passwordInput:
                if not self.waitPassword(passwordInput):
                    raise RuntimeError("Amazon 密码未填入，已停止提交登录")
                time.sleep(0.78)
                signInButton = page.ele(
                    'x://input[@id="signInSubmit"] | //button[@id="signInSubmit"] | //input[@type="submit"]',
                    timeout=10,
                )
                if not signInButton:
                    raise RuntimeError("未找到 Amazon 登录提交按钮")
                signInButton.click()
                time.sleep(5)

            # 第四步：如果出现二步验证码或安全验证页，优先使用易得客插件，失败时交给人工
            try:
                pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
            except Exception:
                pageText = ""
            codeInput = page.ele(
                'x://input[contains(@id,"otp") or contains(@name,"otp") or contains(@id,"auth-mfa") or contains(@name,"code")]',
                timeout=5,
            )
            if codeInput or "Two-Step Verification" in pageText or "二步验证" in pageText or "验证码" in pageText or "安全验证" in pageText:
                if not self.fillCode():
                    ctypes.windll.user32.MessageBoxW(
                        0,
                        "易得客验证码插件未自动填入 Amazon 二步验证码。\n请在当前 Amazon 页面手动完成验证码或安全验证，完成后点击“确定”继续。",
                        "Amazon 二步验证确认",
                        0x40 | 0x40000,
                    )
                time.sleep(0.78)
                codeSubmit = page.ele('x://input[@type="submit"] | //button[@type="submit"]', timeout=5)
                if codeSubmit:
                    codeSubmit.click()
                    time.sleep(5)

            # 第五步：账号选择页按当前目标站点选择账号
            accountSearch = page.ele('x://kat-input[@placeholder="Search for an account"]', timeout=10)
            if not accountSearch:
                accountSearch = page.ele('x://*[@placeholder="Search for an account"]', timeout=5)
            if accountSearch:
                time.sleep(4)
                accountSearch.input(self.siteName, by_js=True)
                time.sleep(0.78)
                page.ele(f'x://span[text()="{self.siteName}"]').click()
                time.sleep(0.78)
                page.ele('x://kat-button[@label="Select account"]').click()
                time.sleep(5)

                # 选择账号后可能再次要求密码或二步验证，继续顺序处理
                passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                if passwordInput:
                    if not self.waitPassword(passwordInput):
                        raise RuntimeError("Amazon 密码未填入，已停止二次验证提交")
                    time.sleep(0.78)
                    confirmButton = page.ele('x://input[@type="submit"] | //button[@type="submit"]', timeout=5)
                    if confirmButton:
                        confirmButton.click()
                        time.sleep(5)

                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                codeInput = page.ele(
                    'x://input[contains(@id,"otp") or contains(@name,"otp") or contains(@id,"auth-mfa") or contains(@name,"code")]',
                    timeout=5,
                )
                if codeInput or "Two-Step Verification" in pageText or "二步验证" in pageText or "验证码" in pageText or "安全验证" in pageText:
                    if not self.fillCode():
                        ctypes.windll.user32.MessageBoxW(
                            0,
                            "易得客验证码插件未自动填入 Amazon 二步验证码。\n请在当前 Amazon 页面手动完成验证码或安全验证，完成后点击“确定”继续。",
                            "Amazon 二步验证确认",
                            0x40 | 0x40000,
                        )
                    time.sleep(0.78)
                    codeSubmit = page.ele('x://input[@type="submit"] | //button[@type="submit"]', timeout=5)
                    if codeSubmit:
                        codeSubmit.click()
                        time.sleep(5)

            # 第六步：确认已经进入 Seller Central；仍卡在登录页时给人工兜底
            for waitCount in range(timeout):
                url = (page.url or "").lower()
                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if (
                    "This site can" in pageText
                    or "ERR_TIMED_OUT" in pageText
                    or "ERR_PROXY" in pageText
                    or "无法访问此网站" in pageText
                    or "响应时间过长" in pageText
                ):
                    raise RuntimeError("当前店铺 profile 无法访问 Amazon Seller Central，请先检查店铺网络、代理或防火墙")
                if "sellercentral.amazon" in url and "/ap/signin" not in url:
                    self.page = page
                    return page
                try:
                    menuHost = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=1)
                    if menuHost:
                        self.page = page
                        return page
                except Exception:
                    pass
                time.sleep(1)

            ctypes.windll.user32.MessageBoxW(
                0,
                "Amazon 登录尚未确认完成。\n请在当前浏览器页面手动完成账号、密码、二步验证码或账号选择，完成并进入 Seller Central 后点击“确定”继续。",
                "Amazon 登录人工确认",
                0x40 | 0x40000,
            )
            for waitCount in range(timeout):
                url = (page.url or "").lower()
                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if (
                    "This site can" in pageText
                    or "ERR_TIMED_OUT" in pageText
                    or "ERR_PROXY" in pageText
                    or "无法访问此网站" in pageText
                    or "响应时间过长" in pageText
                ):
                    raise RuntimeError("当前店铺 profile 无法访问 Amazon Seller Central，请先检查店铺网络、代理或防火墙")
                if "sellercentral.amazon" in url and "/ap/signin" not in url:
                    self.page = page
                    return page
                try:
                    menuHost = page.ele('x://*[@data-test-tag="hamburger-menu"]', timeout=1)
                    if menuHost:
                        self.page = page
                        return page
                except Exception:
                    pass
                time.sleep(1)
            raise RuntimeError("Amazon 登录未完成，已停止后续报告请求流程")

    def amazonSellerLogin(self, page, email=None, password=None, timeout=120, siteEnglishName="United States"):
        """Amazon 卖家中心登录，委托内置 AmazonSeller。"""
        return self.AmazonSeller(page, storePassword=password, siteName=siteEnglishName).login(
            email,
            password,
            timeout=timeout,
        )


if __name__ == "__main__":
    config = {
        "username": "",
        "password": "",
        "shopPort": 8888,
        "amazonEmail": "",
        "amazonPassword": "",
        "testMode": "yideke",
    }

    testMode = config.get("testMode", "yideke")
    amazonEmail = config.get("amazonEmail") or None
    amazonPassword = config.get("amazonPassword") or None
    if not amazonEmail:
        amazonEmail = None
    if not amazonPassword:
        amazonPassword = None

    if testMode in ("yideke", "both"):
        if not config.get("username") or not config.get("password"):
            raise ValueError("请先在 __main__ config 中填写易得客测试账号和密码")
        yideke = YidekeLogin(config["username"], config["password"])
        time.sleep(2)
        yideke.login()
        print("易得客登录完成", flush=True)

    if testMode in ("amazon", "both"):
        page = ChromiumPage(f"127.0.0.1:{config['shopPort']}")
        if testMode == "both":
            yideke.amazonSellerLogin(page, amazonEmail, amazonPassword)
        else:
            YidekeLogin.AmazonSeller(page, storePassword=amazonPassword).login(
                amazonEmail,
                amazonPassword,
            )
        print("Amazon 登录流程结束", flush=True)
