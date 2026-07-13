import time
from pathlib import Path
import psutil
import os
import subprocess
import socket
from pywinauto.application import Application
from DrissionPage import ChromiumPage,Chromium
class Specification:
    """易得客客户端启动、登录与 Amazon Seller Central 登录接管"""

    def __init__(self,username,password):

        # 定位本机易得客可执行文件，后续通过调试端口接管浏览器
        exe_path = self.resolve_edecker()
        self.username = username
        self.password = password

        # 启动前先关闭旧易得客进程，避免接管到历史窗口
        for proc in psutil.process_iter(["name"]):
            if proc.info["name"] and proc.info["name"].lower() == "edecker.exe":
                proc.kill()

        # 清理固定调试端口，保证新客户端能正常开启远程调试
        self.kill_debug_port(9222)
        time.sleep(1)
        # 使用易得客真实用户数据目录启动，保留店铺与浏览器环境
        self.edecker_process = subprocess.Popen([
            exe_path,
            "--remote-debugging-port=9222",
            f"--user-data-dir={Path(os.environ['LOCALAPPDATA']) / 'eDecker6' / 'User Data'}"
        ])


    def kill_debug_port(self, port):
        """结束占用指定调试端口的浏览器进程，避免连到非易得客页面"""
        flag = f"--remote-debugging-port={port}"
        for proc in psutil.process_iter(["pid", "cmdline"]):
            try:
                cmdline = proc.info.get("cmdline") or []
                if any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                pass

    def resolve_edecker(self):
        """按易得客默认安装目录查找 edecker.exe"""
        # 先查找固定入口，再按版本目录取最新版本
        local = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "Application"
        direct = local / "edecker.exe"
        if direct.exists():
            return str(direct)
        if local.exists():
            versions = sorted(p for p in local.iterdir() if p.is_dir())
            if versions:
                exe = versions[-1] / "edecker.exe"
                if exe.exists():
                    return str(exe)
        raise FileNotFoundError(f"edecker.exe not found under {local}")



    def YidekeLogin(self, max_retries=3, retry_interval=2):
        """登录易得客客户端，并返回可接管的浏览器页面"""
        last_error = None
        for attempt in range(1, max_retries + 1):
            try:
                # 重试时刷新浏览器，避免停留在上一次失败的登录状态
                if attempt > 1:
                    try:
                        ChromiumPage("127.0.0.1:9222").refresh()
                    except Exception as e_refresh:
                        print(f"重试前刷新浏览器失败，继续尝试登录: {e_refresh}")

                # 等待远程调试端口可用，再创建 DrissionPage 连接
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

                # 易得客登录弹窗是 Windows 原生窗口，使用 pywinauto 填入账号密码
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

        def __init__(self, page, store_password=None, siteName="United States"):
            # 店铺 profile 页面由易得客打开，Amazon 登录逻辑只接管该页面
            self.page = page
            self.StorePassword = store_password or ""
            self.siteName = siteName or "United States"

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
            import time
            from pywinauto import Desktop
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
                                for b in win.descendants(control_type="Button"):
                                    name = b.window_text()
                                    if name == "填入验证码":
                                        b.click_input()
                                        time.sleep(1.5)
                                        b.click_input()
                                        found = True
                                        success = True
                                        break
                                    elif name == "获取最新验证码":
                                        b.click_input()
                                        time.sleep(1)
                                        b.click_input()
                                        found = True
                                        break
                                if success:
                                    break
                                if not found:
                                    time.sleep(0.2)
                            break
                    if success:
                        break
                except Exception as e:
                    print(e)

            return success

        def Login(self, page):
            """处理 Amazon 登录页、店铺账户选择页与二步验证码"""
            # 登录按钮存在时说明当前停留在 Amazon 密码确认页
            login = page.ele('x://input[@id="continue"]')
            # 账户搜索框存在时说明当前停留在 Seller Central 账户选择页
            SFA = page.ele('x://kat-input[@placeholder="Search for an account"]', timeout=5)
            # 密码框存在时说明当前需要补充 Amazon 密码
            passwordInput = page.ele('x://input[@type="password"]', timeout=5)
            if login:
                login.click()
                time.sleep(5)
                page.ele('x://input[@type="password"]').input(self.StorePassword, clear=True)
                time.sleep(0.78)
                page.ele('x://input[@id="signInSubmit"]').click()
                time.sleep(5)
                self.Code()  # 点击验证码插件
                time.sleep(0.78)
                page.ele('x://input[@type="submit"]').click()  # 填入验证码
                SFA = page.ele('x://*[@placeholder="Search for an account"]')
            elif passwordInput:
                passwordInput.input(self.StorePassword, clear=True)
                time.sleep(0.78)
                page.ele('x://input[@id="signInSubmit"]').click()
                time.sleep(5)
                self.Code()  # 点击验证码插件
                time.sleep(0.78)
                codeSubmit = page.ele('x://input[@type="submit"]', timeout=5)
                if codeSubmit:
                    codeSubmit.click()  # 填入验证码
                SFA = page.ele('x://*[@placeholder="Search for an account"]', timeout=10)
            if SFA:
                time.sleep(4)
                SFA.input(self.siteName, by_js=True)
                time.sleep(0.78)
                page.ele(f'x://span[text()="{self.siteName}"]').click()  # 选择目标站点
                time.sleep(0.78)
                page.ele('x://kat-button[@label="Select account"]').click()
                time.sleep(5)
                try:
                    sub = page.ele('x://input[@type="submit"]').click()
                    if sub:
                        page.ele('x://input[@type="password"]').input(self.StorePassword, clear=True)
                        time.sleep(0.78)
                        sub.click()
                        self.Code()  # 点击验证码插件
                except:
                    pass

        def login(self, email=None, password=None, timeout=120):
            """切换 Amazon 标签后委托 Login() 完成初次登录"""
            if password:
                self.StorePassword = password
            page = self.pickSellerTab(self.page)
            self.Login(page)
            self.page = page
            return page

    def amazonSellerLogin(self, page, email=None, password=None, timeout=120, siteEnglishName="United States"):
        """Amazon 卖家中心登录，委托内置 AmazonSeller"""
        return self.AmazonSeller(page, store_password=password, siteName=siteEnglishName).login(
            email, password, timeout=timeout,
        )
    
    


if __name__ == "__main__":
    config = {
        "username": "",
        "password": "",
        "shopPort": 9999,
        "amazonEmail": "",
        "amazonPassword": "",
        "testMode": "yideke",
    }

    test_mode = config.get("testMode", "yideke")
    amazon_email = config.get("amazonEmail") or config.get("amazon_email") or None
    amazon_password = config.get("amazonPassword") or config.get("amazon_password") or None
    if not amazon_email:
        amazon_email = None
    if not amazon_password:
        amazon_password = None

    if test_mode in ("yideke", "both"):
        sp = Specification(config["username"], config["password"])
        time.sleep(2)
        sp.YidekeLogin()
        print("易得客登录完成", flush=True)

    if test_mode in ("amazon", "both"):
        page = ChromiumPage(f"127.0.0.1:{config['shopPort']}")
        if test_mode == "both":
            sp.amazonSellerLogin(page, amazon_email, amazon_password)
        else:
            Specification.AmazonSeller(page, store_password=amazon_password).login(
                amazon_email, amazon_password,
            )
        print("Amazon 登录流程结束", flush=True)
