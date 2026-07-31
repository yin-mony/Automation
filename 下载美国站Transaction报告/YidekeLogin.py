import time
from pathlib import Path
import psutil
import os
import subprocess
import socket
import ctypes
from pywinauto.application import Application
from pywinauto import Desktop
from DrissionPage import ChromiumPage,Chromium
class Specification:
    def __init__(self,username,password):

        exe_path = self.resolve_edecker()
        self.username = username
        self.password = password

        for proc in psutil.process_iter(["name"]):
            if proc.info["name"] and proc.info["name"].lower() == "edecker.exe":
                proc.kill()

        time.sleep(1)
        self.edecker_process = subprocess.Popen([
            exe_path,
            "--remote-debugging-port=9222",
            f"--user-data-dir={Path(os.environ['LOCALAPPDATA']) / 'eDecker6' / 'User Data'}"
        ])


    def resolve_edecker(self):
        local = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "Application"
        if local.exists():
            versions = sorted(p for p in local.iterdir() if p.is_dir())
            if versions:
                exe = versions[-1] / "edecker.exe"
                if exe.exists():
                    return str(exe)
        # fallback: search common dirs
        for root in [Path(r"C:\Program Files"), Path(r"C:\Program Files (x86)"), Path(os.environ["LOCALAPPDATA"])]:
            for path in root.rglob("edecker.exe"):
                return str(path)
        raise FileNotFoundError("edecker.exe not found")

    def YidekeLogin(self, max_retries=3, retry_interval=2):
        last_error = None
        for attempt in range(1, max_retries + 1):
            try:
                if attempt > 1:
                    try:
                        ChromiumPage("127.0.0.1:9222").refresh()
                    except Exception as e_refresh:
                        print(f"重试前刷新浏览器失败，继续尝试登录: {e_refresh}")

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
                login_ele = page.ele('x://span[text()="登录"]', timeout=60)
                if not login_ele:
                    raise RuntimeError(f'在易得客浏览器当前页面未找到 XPath: x://span[text()="登录"]，当前URL: {page.url}')
                login_ele.click()
                time.sleep(5)

                deadline = time.time() + 30
                last_window_error = None
                while time.time() < deadline:
                    try:
                        win_app = Application(backend='win32').connect(title_re="易得客浏览器", visible_only=False)
                        hwnd = win_app.window(title_re="易得客浏览器").handle
                        app = Application(backend='uia').connect(handle=hwnd)
                        dlg = app.window(handle=hwnd)
                        dlg.wait("visible ready", timeout=15)
                        break
                    except Exception as e:
                        last_window_error = e
                        time.sleep(1)
                else:
                    raise RuntimeError("未找到标题精确匹配“易得客浏览器”的登录窗口") from last_window_error
                # dlg.print_control_identifiers()

                phone_label = dlg.child_window(title="手机号", control_type="Text")
                if phone_label.exists():
                    # 获取"手机号"标签的位置
                    label_rect = phone_label.rectangle()

                    # 在标签下方查找Edit控件
                    all_edits = dlg.descendants(control_type="Edit")
                    for edit in all_edits:
                        edit_rect = edit.rectangle()
                        # 如果Edit控件在标签下方且x坐标相近
                        if (edit_rect.top > label_rect.bottom and
                                abs(edit_rect.left - label_rect.left) < 50):
                            edit.set_text(self.username)
                            print("手机号输入成功")
                            break

                # 定位密码输入框
                password_label = dlg.child_window(title="密码", control_type="Text")
                if password_label.exists():
                    # 获取"密码"标签的位置
                    label_rect = password_label.rectangle()

                    # 在标签下方查找Edit控件
                    all_edits = dlg.descendants(control_type="Edit")
                    for edit in all_edits:
                        edit_rect = edit.rectangle()
                        # 如果Edit控件在标签下方且x坐标相近
                        if (edit_rect.top > label_rect.bottom and
                                abs(edit_rect.left - label_rect.left) < 50):
                            edit.set_text(self.password)  # 替换为实际密码
                            print("密码输入成功")
                            break
                dlg.child_window(title="登录易得客", control_type="Button").click()
                return
            except Exception as e:
                last_error = e
                if attempt >= max_retries:
                    break
                print(f"YidekeLogin 第{attempt}次尝试失败，{retry_interval}秒后重试: {e}")
                time.sleep(retry_interval)
        raise RuntimeError(f"YidekeLogin 重试{max_retries}次后仍失败") from last_error

    class AmazonSeller:
        """Amazon 卖家中心登录（店铺 profile 浏览器内）"""

        def __init__(self, page, store_password=None, siteName="United States", siteChineseName="美国"):
            # 店铺 profile 页面由易得客打开，Amazon 登录逻辑只接管该页面
            self.page = page
            self.StorePassword = store_password or ""
            self.siteName = siteName or "United States"
            self.siteChineseName = siteChineseName or "美国"

        def pickSellerTab(self, page):
            """在多标签中选取 Amazon 卖家相关页面"""
            backend_tab = None
            login_tab = None
            try:
                browser = page.browser
                for tab_id in browser.tab_ids:
                    tab = browser.get_tab(tab_id)
                    url = (tab.url or "").lower()
                    if url.startswith("chrome-extension://"):
                        continue
                    if "sellercentral.amazon" in url and "/ap/signin" not in url:
                        backend_tab = tab
                        break
                    if "/ap/signin" in url or "sellercentral.amazon" in url:
                        login_tab = tab
                picked = backend_tab or login_tab
                if picked:
                    self.page = picked
                    return picked
            except Exception:
                pass
            return page

        def Code(self):
            """通过易得客验证码插件窗口填入 Amazon 二步验证码"""
            time.sleep(1)
            desktop = Desktop(backend="uia")
            success = False
            for win in desktop.windows():
                try:
                    # 先找到验证码服务按钮，再轮询填入验证码或获取最新验证码按钮
                    for btn in win.descendants(control_type="Button"):
                        if "二步验证码服务" in btn.window_text():
                            btn.click_input()
                            time.sleep(1.5)
                            while True:
                                found = False
                                for button in win.descendants(control_type="Button"):
                                    name = button.window_text()
                                    if name == "填入验证码":
                                        button.click_input()
                                        time.sleep(1.5)
                                        button.click_input()
                                        found = True
                                        success = True
                                        break
                                    if name == "获取最新验证码":
                                        button.click_input()
                                        time.sleep(1)
                                        button.click_input()
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
            """等待 Amazon 密码框自动填充或提示用户手动填入"""
            if not passwordInput:
                raise RuntimeError("未找到 Amazon 密码输入框")

            if self.StorePassword:
                passwordInput.input(self.StorePassword, clear=True)
                return True

            # 未配置 Amazon 密码时，等待浏览器保存的密码自动填充
            for _ in range(20):
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
            for _ in range(240):
                value = passwordInput.attr("value") or ""
                if value:
                    return True
                time.sleep(0.5)
            return False

        def submitPassword(self, page, passwordInput):
            """确认 Amazon 密码已填入后提交登录"""
            if not self.waitPassword(passwordInput):
                raise RuntimeError("Amazon 密码未填入，已停止提交登录")
            time.sleep(0.78)
            page.ele('x://input[@id="signInSubmit"]').click()
            time.sleep(5)

        def Login(self, page):
            """处理 Amazon 登录页、店铺账户选择页与二步验证码"""
            # 登录按钮存在时说明当前停留在 Amazon 密码确认页
            login = page.ele('x://input[@id="continue"]')
            # 账户搜索框存在时说明当前停留在 Seller Central 账户选择页
            SFA = page.ele(
                'x://kat-input[@placeholder="搜索账户" '
                'or @placeholder="搜索账号" '
                'or @placeholder="Search for an account"]',
                timeout=5,
            )
            # 密码框存在时说明当前需要补充 Amazon 密码
            passwordInput = page.ele('x://input[@type="password"]', timeout=5)
            if login:
                login.click()
                time.sleep(5)
                passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                self.submitPassword(page, passwordInput)
                self.Code()
                time.sleep(0.78)
                page.ele('x://input[@type="submit"]').click()
                SFA = page.ele(
                    'x://*[@placeholder="搜索账户" '
                    'or @placeholder="搜索账号" '
                    'or @placeholder="Search for an account"]'
                )
            elif passwordInput:
                self.submitPassword(page, passwordInput)
                self.Code()
                time.sleep(0.78)
                codeSubmit = page.ele('x://input[@type="submit"]', timeout=5)
                if codeSubmit:
                    codeSubmit.click()
                SFA = page.ele(
                    'x://*[@placeholder="搜索账户" '
                    'or @placeholder="搜索账号" '
                    'or @placeholder="Search for an account"]',
                    timeout=10,
                )
            if SFA:
                time.sleep(4)
                searchPlaceholder = SFA.attr("placeholder") or ""
                searchSiteName = self.siteChineseName if "搜索" in searchPlaceholder else self.siteName
                SFA.input(searchSiteName, by_js=True)
                time.sleep(0.78)
                page.ele(
                    f'x://span[normalize-space()="{self.siteName}" '
                    f'or normalize-space()="{self.siteChineseName}"]'
                ).click()
                time.sleep(0.78)
                page.ele(
                    'x://kat-button[@label="选择账户" or @label="Select account"]'
                    ' | //button[normalize-space()="选择账户" or normalize-space()="Select account"]'
                ).click()
                time.sleep(5)
                try:
                    sub = page.ele('x://input[@type="submit"]').click()
                    if sub:
                        passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                        if not self.waitPassword(passwordInput):
                            raise RuntimeError("Amazon 密码未填入，已停止二次验证提交")
                        time.sleep(0.78)
                        sub.click()
                        self.Code()
                except Exception:
                    pass

        def login(self, email=None, password=None, timeout=120):
            """切换 Amazon 标签后委托 Login() 完成初次登录"""
            if password:
                self.StorePassword = password
            page = self.pickSellerTab(self.page)
            self.Login(page)
            self.page = page
            return page

    def amazonSellerLogin(
        self,
        page,
        email=None,
        password=None,
        timeout=120,
        siteEnglishName="United States",
        siteChineseName="美国",
    ):
        """Amazon 卖家中心登录，委托内置 AmazonSeller"""
        return self.AmazonSeller(
            page,
            store_password=password,
            siteName=siteEnglishName,
            siteChineseName=siteChineseName,
        ).login(
            email, password, timeout=timeout,
        )
