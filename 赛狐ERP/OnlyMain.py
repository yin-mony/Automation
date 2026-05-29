# 唯一的整体调用入口
from pathlib import Path
import os

from DrissionPage import ChromiumOptions
from DrissionPage import ChromiumPage
from SaihuERPLogin import SaihuERPLogin


class SaiHuMain:
    DEFAULT_CHROME_DEBUG_ADDR = "127.0.0.1:9222"
    _shared_page = None

    def __init__(self, username=None, password=None):
        self.username = username
        self.password = password
        self.page = None
        self.reused_session = False

    def _is_page_alive(self, page):
        if page is None:
            return False
        try:
            _ = page.url
            return True
        except Exception:
            return False

    def _connect_chrome_with_ap(self):
        # 同一进程内优先复用已连接的 Chrome 页面对象。
        if self._is_page_alive(self.__class__._shared_page):
            self.page = self.__class__._shared_page
            return self.page, "shared-session"

        # 与“审核请款单/test.py”保持一致：优先读取 EDGE_DEBUG_ADDR，并通过调试端口连接。
        debug_addr = (
            os.getenv("EDGE_DEBUG_ADDR", "").strip()
            or os.getenv("CHROME_DEBUG_ADDR", "").strip()
            or self.DEFAULT_CHROME_DEBUG_ADDR
        )
        options = ChromiumOptions().set_address(debug_addr)
        try:
            self.page = ChromiumPage(options)
        except Exception as exc:
            raise RuntimeError(
                f"无法连接 Chrome 调试实例 {debug_addr}。"
                "请先启动 Chrome 并开启远程调试端口，例如："
                "chrome.exe --remote-debugging-port=9222"
            ) from exc
        self.page.set.window.max()
        self.__class__._shared_page = self.page
        print(f"已连接 Chrome 调试实例: {debug_addr}", flush=True)
        return self.page, debug_addr

    def login(self):
        page, debug_addr = self._connect_chrome_with_ap()
        if debug_addr != "shared-session":
            print(f"已连接 Chrome 实例: {debug_addr}", flush=True)

        login_client = SaihuERPLogin(
            page=page,
            username=self.username,
            password=self.password,
            img_dir=str(Path(__file__).resolve().parent),
        )
        # 按参考目录 test.py 保持一致：固定使用 prefer_entry_check=True。
        login_client.login(prefer_entry_check=True)
        print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)
        self.reused_session = False
        return page

    def run(self, mode, excel_file_path=None, sheet_name=None):
        page = self.login()
        normalized_mode = str(mode or "").strip().lower()

        if normalized_mode in ("dew", "pure", "online"):
            from DewMain import DewMainWorkflow

            workflow_kwargs = {
                "excel_file_path": excel_file_path,
                "username": self.username,
                "password": self.password,
            }
            if sheet_name:
                workflow_kwargs["sheet_name"] = sheet_name

            workflow = DewMainWorkflow(**workflow_kwargs)
            return workflow.run(page=page)

        if normalized_mode in ("low", "cheap"):
            from LowMain import LowMainWorkflow

            workflow = LowMainWorkflow(
                excel_file_path=excel_file_path,
                username=self.username,
                password=self.password,
            )
            return workflow.run(page=page, skip_close_popups=self.reused_session)

        raise ValueError(f"不支持的模式: {mode}，可选: low / dew")









