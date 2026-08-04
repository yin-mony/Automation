import time
from pathlib import Path
import psutil
import os
import subprocess
import socket
import ctypes
from pywinauto import Desktop
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
                # 当前如果已停在易得客网页登录页，直接通过页面 DOM 填入账号密码
                pageText = ""
                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if "登录你的易得客" in pageText or "登录易得客" in pageText:
                    phoneInput = page.ele('x://input[@placeholder="请输入手机号码"] | //input[@type="text"]', timeout=15)
                    passwordInput = page.ele('x://input[@placeholder="请输入登录密码"] | //input[@type="password"]', timeout=15)
                    if not phoneInput or not passwordInput:
                        raise RuntimeError("易得客网页登录页未找到手机号或密码输入框")
                    phoneInput.input(self.username, clear=True)
                    passwordInput.input(self.password, clear=True)
                    print("易得客网页登录账号密码输入成功", flush=True)
                    loginBtn = page.ele(
                        'x://button[contains(normalize-space(),"登录易得客")]'
                        ' | //span[normalize-space()="登录易得客"]/ancestor::button[1]',
                        timeout=15,
                    )
                    if not loginBtn:
                        raise RuntimeError("易得客网页登录页未找到登录按钮")
                    loginBtn.click()
                    for waitIndex in range(60):
                        time.sleep(1)
                        try:
                            currentText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                        except Exception:
                            currentText = ""
                        if "登录你的易得客" not in currentText and "登录易得客" not in currentText:
                            print("易得客网页登录完成", flush=True)
                            return
                    raise RuntimeError("易得客网页登录后未离开登录页")

                login_ele = page.ele('x://span[text()="登录"]', timeout=60)
                if not login_ele:
                    raise RuntimeError(f'在易得客浏览器当前页面未找到 XPath: x://span[text()="登录"]，当前URL: {page.url}')
                login_ele.click()
                time.sleep(5)
                # 主界面点击登录后会打开新的登录标签页，切到最新标签页再处理网页登录
                try:
                    page = page.browser.latest_tab
                except Exception:
                    pass
                try:
                    pageText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                except Exception:
                    pageText = ""
                if "登录你的易得客" in pageText or "登录易得客" in pageText:
                    phoneInput = page.ele('x://input[@placeholder="请输入手机号码"] | //input[@type="text"]', timeout=15)
                    passwordInput = page.ele('x://input[@placeholder="请输入登录密码"] | //input[@type="password"]', timeout=15)
                    if not phoneInput or not passwordInput:
                        raise RuntimeError("易得客网页登录页未找到手机号或密码输入框")
                    phoneInput.input(self.username, clear=True)
                    passwordInput.input(self.password, clear=True)
                    print("易得客网页登录账号密码输入成功", flush=True)
                    loginBtn = page.ele(
                        'x://button[contains(normalize-space(),"登录易得客")]'
                        ' | //span[normalize-space()="登录易得客"]/ancestor::button[1]',
                        timeout=15,
                    )
                    if not loginBtn:
                        raise RuntimeError("易得客网页登录页未找到登录按钮")
                    loginBtn.click()
                    for waitIndex in range(60):
                        time.sleep(1)
                        try:
                            currentText = page.run_js('return document.body ? document.body.innerText : ""') or ""
                        except Exception:
                            currentText = ""
                        if "登录你的易得客" not in currentText and "登录易得客" not in currentText:
                            print("易得客网页登录完成", flush=True)
                            return
                    raise RuntimeError("易得客网页登录后未离开登录页")

                # 易得客登录弹窗是 Windows 原生窗口，按控件特征选择真实登录框
                deadline = time.time() + 30
                last_window_error = None
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
                        dlg = candidates[0][1]
                        dlg.set_focus()
                        dlg.wait("visible ready", timeout=15)
                        break
                    except Exception as e:
                        last_window_error = e
                        time.sleep(1)
                else:
                    raise RuntimeError("未找到包含登录控件的易得客登录窗口") from last_window_error
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
            SFA = page.ele('x://kat-input[@placeholder="Search for an account"]', timeout=5)
            # 密码框存在时说明当前需要补充 Amazon 密码
            passwordInput = page.ele('x://input[@type="password"]', timeout=5)
            if login:
                login.click()
                time.sleep(5)
                passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                self.submitPassword(page, passwordInput)
                self.Code()  # 点击验证码插件
                time.sleep(0.78)
                page.ele('x://input[@type="submit"]').click()  # 填入验证码
                SFA = page.ele('x://*[@placeholder="Search for an account"]')
            elif passwordInput:
                self.submitPassword(page, passwordInput)
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
                        passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                        if not self.waitPassword(passwordInput):
                            raise RuntimeError("Amazon 密码未填入，已停止二次验证提交")
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
        "username": "19167561839",
        "password": "yxh643208yang",
        "shopPort": 8888,
        "amazonEmail": "happymike9@outlook.com",
        "amazonPassword": "Happylife989.",
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
