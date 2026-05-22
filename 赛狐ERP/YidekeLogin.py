import time
from pathlib import Path
import psutil
import os
import subprocess
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
        subprocess.Popen([
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
    def YidekeLogin(self):

        page = ChromiumPage("127.0.0.1:9222")
        page.ele('x://span[text()="登录"]').click()
        time.sleep(5)


        win_app = Application(backend='win32').connect(title_re="易得客浏览器")
        hwnd = win_app.window(title_re="易得客浏览器").handle
        app = Application(backend='uia').connect(handle=hwnd)
        dlg = app.window(handle=hwnd)
        dlg.wait("visible ready", timeout=15)
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


