import time
from pathlib import Path
import psutil
import os
import subprocess
import socket
from pywinauto.application import Application
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
