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
    """易得客客户端启动、登录与调试端口管理"""

    class AmazonSeller:
        """Amazon 卖家中心登录（店铺 profile 浏览器内）"""

        def __init__(self, page, store_password=None, siteName="United States"):
            # 店铺 profile 页面由易得客打开，Amazon 登录逻辑只接管该页面
            self.page = page
            self.StorePassword = store_password or ""
            self.siteName = siteName or "United States"
            self.codeExtensionId = "ahglfiiniifnanngnmafhkmnpafpcjhn"

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

        def clickControl(self, control):
            """优先用 FBA 的 click_input 点击控件，失败时尝试 UIA click"""
            try:
                control.click_input()
                return True
            except Exception:
                pass

            try:
                control.click()
                return True
            except Exception:
                pass
            return False

        def clickCodeByWindow(self, timeout=12):
            """按 FBA 方式从易得客窗口点击二步验证码服务和填入验证码"""
            time.sleep(1)
            desktop = Desktop(backend="uia")
            deadline = time.time() + timeout
            serviceClicked = False

            while time.time() < deadline:
                try:
                    for win in desktop.windows():
                        for btn in win.descendants(control_type="Button"):
                            name = (btn.window_text() or "").strip()
                            if not serviceClicked and "二步验证码服务" in name:
                                serviceClicked = self.clickControl(btn)
                                time.sleep(1.5)
                                continue

                            if name == "填入验证码" or "填入验证码" in name:
                                if self.clickControl(btn):
                                    time.sleep(1.5)
                                    self.clickControl(btn)
                                    return True

                            if name == "获取最新验证码" or "获取最新验证码" in name:
                                if self.clickControl(btn):
                                    time.sleep(1)
                                    self.clickControl(btn)
                except Exception as e:
                    print(e)

                time.sleep(0.2)

            return False

        def callDevtools(self, websocketUrl, method, params=None, counter=None, connection=None):
            """调用 Chrome DevTools Protocol 并返回本次响应"""
            import json

            messageId = next(counter)
            connection.send(json.dumps({
                "id": messageId,
                "method": method,
                "params": params or {},
            }))
            while True:
                data = json.loads(connection.recv())
                if data.get("id") == messageId:
                    return data

        def inputCodeByExtension(self, page):
            """直接触发易得客二步验证码扩展的填码逻辑"""
            import itertools
            import json

            import requests
            import websocket

            try:
                browser = page.browser
                address = browser.address
                targets = requests.get(f"http://{address}/json/list", timeout=5).json()
                amazonTarget = None
                workerTarget = None

                for target in targets:
                    url = target.get("url", "")
                    if target.get("type") == "page" and ("/ap/mfa" in url or "/a/settings/approval" in url):
                        amazonTarget = target
                    if target.get("type") == "service_worker" and self.codeExtensionId in url:
                        workerTarget = target

                if not amazonTarget or not workerTarget:
                    return False

                browser.activate_tab(amazonTarget["id"])
                time.sleep(0.5)

                connection = websocket.create_connection(
                    workerTarget["webSocketDebuggerUrl"],
                    timeout=15,
                    suppress_origin=True,
                )
                counter = itertools.count(1)
                self.callDevtools(
                    workerTarget["webSocketDebuggerUrl"],
                    "Runtime.enable",
                    counter=counter,
                    connection=connection,
                )

                script = """
new Promise(async resolve => {
  try {
    const getValue = (path, location='global') => pingpong.getValueForPath({location, path});
    const token = await getValue('token');
    const shop = await getValue('shopInfo', 'local');
    if (!shop || !shop.shopId) return resolve(JSON.stringify({status:500, message:'no shopInfo'}));
    const resp = await fetch(`${pingpong.metadata.settings.EDECKER_WORKSTATION}/api/secondaryCheck/cloudCode/${shop.shopId}`, {headers:{Authorization:token}});
    const body = await resp.json();
    const code = body && body.data && body.data.otpCode;
    if (!code) return resolve(JSON.stringify({status:500, message:'no otp'}));
    chrome.tabs.query({active:true, highlighted:true}, ([tab]) => {
      if (!tab || !tab.url || (!tab.url.includes('/ap/mfa') && !tab.url.includes('/a/settings/approval'))) {
        return resolve(JSON.stringify({status:500, message:'not mfa'}));
      }
      chrome.scripting.executeScript({
        target:{tabId:tab.id},
        func:(otp) => {
          let filled = false;
          const one = document.querySelector('#auth-mfa-otpcode');
          if (one) {
            one.value = `${otp || ''}`;
            one.dispatchEvent(new Event('input', {bubbles:true}));
            one.dispatchEvent(new Event('change', {bubbles:true}));
            filled = true;
          }
          const two = document.querySelector('#ch-auth-app-code-input');
          if (two) {
            two.value = `${otp || ''}`;
            two.dispatchEvent(new Event('input', {bubbles:true}));
            two.dispatchEvent(new Event('change', {bubbles:true}));
            filled = true;
          }
          return filled;
        },
        args:[code]
      }, async result => {
        try {
          await fetch(`${pingpong.metadata.settings.EDECKER_WORKSTATION}/api/secondaryCheck/cloudCodeLog/${shop.shopId}`, {
            method:'POST',
            headers:{Authorization:token}
          });
        } catch (e) {}
        const err = chrome.runtime.lastError && chrome.runtime.lastError.message;
        resolve(JSON.stringify({status: err ? 500 : 200, message: err || 'ok', filled: result && result[0] && result[0].result === true}));
      });
    });
  } catch (e) {
    resolve(JSON.stringify({status:500, message:e.message}));
  }
})
"""
                result = self.callDevtools(
                    workerTarget["webSocketDebuggerUrl"],
                    "Runtime.evaluate",
                    {
                        "expression": script,
                        "awaitPromise": True,
                        "returnByValue": True,
                        "timeout": 120000,
                    },
                    counter=counter,
                    connection=connection,
                )
                connection.close()
                value = result.get("result", {}).get("result", {}).get("value") or "{}"
                payload = json.loads(value)
                if payload.get("status") == 200 and payload.get("filled"):
                    return True
                print(f"易得客验证码扩展填入失败: {payload.get('message')}")
            except Exception as exc:
                print(f"易得客验证码扩展填入异常: {exc}")
            return False

        def Code(self, page=None):
            """通过易得客验证码插件填入 Amazon 二步验证码"""
            page = page or self.page
            if self.clickCodeByWindow():
                return True
            print("易得客验证码窗口点击未完成，改用扩展接口填入验证码。")
            return self.inputCodeByExtension(page)

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
            # 验证码框存在时说明当前已经停留在 Amazon MFA 页面
            codeInput = page.ele('x://input[@id="auth-mfa-otpcode"] | //input[@id="ch-auth-app-code-input"]', timeout=2)
            if login:
                login.click()
                time.sleep(5)
                passwordInput = page.ele('x://input[@type="password"]', timeout=10)
                self.submitPassword(page, passwordInput)
                if not self.Code(page):  # 点击验证码插件
                    raise RuntimeError("Amazon 二步验证码未填入")
                time.sleep(0.78)
                page.ele('x://input[@type="submit"]').click()  # 填入验证码
                SFA = page.ele('x://*[@placeholder="Search for an account"]')
            elif passwordInput:
                self.submitPassword(page, passwordInput)
                if not self.Code(page):  # 点击验证码插件
                    raise RuntimeError("Amazon 二步验证码未填入")
                time.sleep(0.78)
                codeSubmit = page.ele('x://input[@type="submit"]', timeout=5)
                if codeSubmit:
                    codeSubmit.click()  # 填入验证码
                SFA = page.ele('x://*[@placeholder="Search for an account"]', timeout=10)
            elif codeInput:
                if not self.Code(page):  # 当前已在 MFA 页时补填验证码
                    raise RuntimeError("Amazon 二步验证码未填入")
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
                        if not self.Code(page):  # 点击验证码插件
                            raise RuntimeError("Amazon 二步验证码未填入")
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

    def __init__(self, username="", password=""):
        # 易得客账号
        self.username = username
        # 易得客密码
        self.password = password
        # 易得客管理端口
        self.debugPort = 9222
        # 易得客进程对象
        self.edeckerProcess = None

    def killDebugPort(self, port):
        """结束占用指定调试端口的进程"""
        # 拼出调试端口参数
        flag = f"--remote-debugging-port={port}"
        # 遍历系统进程，查找占用该调试端口的命令行
        for proc in psutil.process_iter(["pid", "cmdline"]):
            try:
                cmdline = proc.info.get("cmdline") or []
                if any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                # 进程无权限或已退出时跳过
                pass

    def killEdecker(self):
        """启动管理窗口前关闭历史易得客进程"""
        # 遍历并关闭旧 edecker 进程，避免接管到历史窗口
        for proc in psutil.process_iter(["name"]):
            try:
                if proc.info["name"] and proc.info["name"].lower() == "edecker.exe":
                    proc.kill()
            except Exception:
                # 进程状态变化时跳过
                pass

    def resolveEdecker(self):
        """按易得客默认安装目录查找 edecker.exe"""
        # 先检查默认 Application 目录
        localDir = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "Application"
        # 新版可能直接放在 Application 根目录
        directExe = localDir / "edecker.exe"
        if directExe.exists():
            return str(directExe)

        # 旧版可能按版本号建子目录
        if localDir.exists():
            versions = sorted(path for path in localDir.iterdir() if path.is_dir())
            if versions:
                exePath = versions[-1] / "edecker.exe"
                if exePath.exists():
                    return str(exePath)

        raise FileNotFoundError(f"edecker.exe not found under {localDir}")

    def startClient(self):
        """启动易得客管理浏览器并开启调试端口"""
        # 查找易得客可执行文件
        exePath = self.resolveEdecker()
        # 启动前清理旧进程和固定调试端口
        self.killEdecker()
        self.killDebugPort(self.debugPort)
        time.sleep(1)

        # 使用真实用户数据目录启动，保留店铺环境
        userDataDir = Path(os.environ["LOCALAPPDATA"]) / "eDecker6" / "User Data"
        self.edeckerProcess = subprocess.Popen([
            exePath,
            f"--remote-debugging-port={self.debugPort}",
            f"--user-data-dir={userDataDir}",
        ])

    def waitDebugPort(self, port, timeout=60):
        """等待调试端口可连接"""
        # 按超时时间持续探测本地端口
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=2):
                    return
            except OSError:
                time.sleep(1)

        raise RuntimeError(f"等待易得客调试端口 127.0.0.1:{port} 超时")

    def getPageText(self, page):
        """读取页面正文文本，失败时返回空字符串"""
        try:
            return page.run_js('return document.body ? document.body.innerText : ""') or ""
        except Exception:
            # 页面脚本执行失败时按空页面处理
            return ""

    def loginWebPage(self, page):
        """在易得客网页登录页输入账号密码并提交"""
        # 读取页面文本判断是否为网页登录页
        pageText = self.getPageText(page)
        if "登录你的易得客" not in pageText and "登录易得客" not in pageText:
            return False

        # 定位手机号和密码输入框
        phoneInput = page.ele('x://input[@placeholder="请输入手机号码"] | //input[@type="text"]', timeout=15)
        passwordInput = page.ele('x://input[@placeholder="请输入登录密码"] | //input[@type="password"]', timeout=15)
        if not phoneInput or not passwordInput:
            raise RuntimeError("易得客网页登录页未找到手机号或密码输入框")

        # 填入易得客账号密码
        phoneInput.input(self.username, clear=True)
        passwordInput.input(self.password, clear=True)
        print("易得客网页登录账号密码输入成功", flush=True)

        # 定位并点击登录按钮
        loginBtn = page.ele(
            'x://button[contains(normalize-space(),"登录易得客")]'
            ' | //span[normalize-space()="登录易得客"]/ancestor::button[1]',
            timeout=15,
        )
        if not loginBtn:
            raise RuntimeError("易得客网页登录页未找到登录按钮")
        loginBtn.click()

        # 等待真正进入工作台，不能只以离开登录页作为成功条件
        for _ in range(60):
            time.sleep(1)
            page = self.pickWorkPage(page)
            if self.isLoggedIn(page):
                print("易得客网页登录完成", flush=True)
                return True

        return False

    def isLoggedIn(self, page):
        """判断易得客是否已经处于登录后的工作台页面"""
        # 检查当前标签与浏览器中其他标签
        pages = [page]
        try:
            browser = page.browser
            for tabId in browser.tab_ids:
                pages.append(browser.get_tab(tabId))
        except Exception:
            pass

        for item in pages:
            try:
                url = (item.url or "").lower()
                text = self.getPageText(item)
                if "new-page-guest" in url or ("注册" in text and "登录" in text):
                    continue
                if "登录易得客" in text or "登录你的易得客" in text:
                    continue
                if (
                    "selleros.cn" in url
                    or "work-station" in url
                    or "shops.edecker.cn" in url
                    or "workbench" in url
                ) and ("店铺" in text or "访问" in text):
                    return True
                if "店铺" in text and "访问" in text:
                    return True
            except Exception:
                # 单个标签读取失败时继续检查其他标签
                pass
        return False

    def pickWorkPage(self, page):
        """从浏览器标签中选取易得客工作台页面"""
        # 当前页是侧栏扩展时，优先切换到 selleros 工作台标签
        try:
            browser = page.browser
            for tabId in browser.tab_ids:
                tab = browser.get_tab(tabId)
                url = (tab.url or "").lower()
                if (
                    "selleros.cn" in url
                    or "work-station" in url
                    or "shops.edecker.cn" in url
                    or "workbench" in url
                ):
                    return tab
        except Exception:
            pass
        return page

    def loginNativeWindow(self):
        """在易得客原生登录窗口输入账号密码并提交"""
        # 轮询查找包含登录控件的易得客窗口
        deadline = time.time() + 30
        lastError = None
        while time.time() < deadline:
            try:
                candidates = []
                desktop = Desktop(backend="uia")
                for win in desktop.windows(visible_only=False):
                    try:
                        title = win.window_text() or ""
                        if "易得客浏览器" not in title:
                            continue

                        # 按关键控件给窗口评分，选出真正登录窗
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
                        # 单个窗口读取失败时跳过
                        pass

                if not candidates:
                    raise RuntimeError("未找到包含登录控件的易得客窗口")

                # 选择评分最高的窗口作为登录窗口
                candidates.sort(key=lambda item: item[0], reverse=True)
                dialog = candidates[0][1]
                dialog.set_focus()
                dialog.wait("visible ready", timeout=15)
                break
            except Exception as exc:
                lastError = exc
                time.sleep(1)
        else:
            raise RuntimeError("未找到包含登录控件的易得客登录窗口") from lastError

        # 按“手机号”标签位置找到对应输入框
        phoneLabel = dialog.child_window(title="手机号", control_type="Text")
        if phoneLabel.exists():
            labelRect = phoneLabel.rectangle()
            for edit in dialog.descendants(control_type="Edit"):
                editRect = edit.rectangle()
                if editRect.top > labelRect.bottom and abs(editRect.left - labelRect.left) < 50:
                    edit.set_text(self.username)
                    print("手机号输入成功")
                    break

        # 按“密码”标签位置找到对应输入框
        passwordLabel = dialog.child_window(title="密码", control_type="Text")
        if passwordLabel.exists():
            labelRect = passwordLabel.rectangle()
            for edit in dialog.descendants(control_type="Edit"):
                editRect = edit.rectangle()
                if editRect.top > labelRect.bottom and abs(editRect.left - labelRect.left) < 50:
                    edit.set_text(self.password)
                    print("密码输入成功")
                    break

        # 点击原生窗口登录按钮
        dialog.child_window(title="登录易得客", control_type="Button").click()

    def run(self, maxRetry=3, retryInterval=2):
        """启动并登录易得客客户端"""
        # 启动易得客管理浏览器
        self.startClient()
        # 记录最后一次异常，便于重试失败后抛出
        lastError = None

        for attempt in range(1, maxRetry + 1):
            try:
                # 重试前刷新页面，避免停留在失败状态
                if attempt > 1:
                    try:
                        ChromiumPage(f"127.0.0.1:{self.debugPort}").refresh()
                    except Exception as refreshError:
                        print(f"重试前刷新浏览器失败，继续尝试登录: {refreshError}")

                # 等待管理端口可用后接管页面
                self.waitDebugPort(self.debugPort, timeout=60)
                page = ChromiumPage(f"127.0.0.1:{self.debugPort}")
                page = self.pickWorkPage(page)

                # 优先处理网页登录页
                if self.loginWebPage(page):
                    return

                # 已经登录时直接返回，避免反复等待“登录”按钮
                if self.isLoggedIn(page):
                    print("易得客已处于登录状态", flush=True)
                    return

                # 点击管理页登录按钮
                loginEle = page.ele('x://span[text()="登录"]', timeout=60)
                if not loginEle:
                    raise RuntimeError(f'在易得客浏览器当前页面未找到 XPath: x://span[text()="登录"]，当前URL: {page.url}')
                loginEle.click()
                time.sleep(5)

                # 登录页可能打开在最新标签页
                try:
                    page = page.browser.latest_tab
                except Exception:
                    pass

                # 再次尝试网页登录
                if self.loginWebPage(page):
                    return

                # 最后回退到原生窗口登录
                self.loginNativeWindow()
                return
            except Exception as exc:
                lastError = exc
                if attempt >= maxRetry:
                    break
                print(f"YidekeLogin 第{attempt}次尝试失败，{retryInterval}秒后重试: {exc}")
                time.sleep(retryInterval)

        raise RuntimeError(f"YidekeLogin 重试{maxRetry}次后仍失败") from lastError

    def loginSeller(self, page, email="", password="", timeout=120, siteName="United States"):
        """委托 AmazonSeller 完成 Amazon 卖家中心登录"""
        return self.AmazonSeller(page, store_password=password, siteName=siteName).login(
            email, password, timeout=timeout,
        )


if __name__ == "__main__":
    # 本文件独立调试配置

    config = {
        "username": os.getenv("YIDEKE_USERNAME", ""),
        "password": os.getenv("YIDEKE_PASSWORD", ""),
    }

    # 未配置账号密码时只提示调试方式，避免误启动真实登录
    if not config["username"] or not config["password"]:
        print("请在 YidekeLogin.py 的 main 配置中填写易得客账号密码后再调试。")
    else:
        login = YidekeLogin(config["username"], config["password"])
        login.run()
