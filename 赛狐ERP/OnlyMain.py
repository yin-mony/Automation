# 唯一的整体调用入口
from pathlib import Path
import os

from DrissionPage import ChromiumOptions as EdgeOptions
from DrissionPage import ChromiumPage as EdgePage


class SaiHuMain:
    INITIAL_URL = "https://www.sellfox.com/amzup-web-main/web/dashboard.html"
    LOGIN_URL = "https://www.sellfox.com/amzup-web-main/login.html"
    PROFILE_DIR_NAME = "EdgeDebugProfile_SaiHuMain"
    _shared_page = None

    def __init__(self, username=None, password=None):
        self.initial_url = self.INITIAL_URL
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

    def _start_edge(self):
        # 同一进程内优先复用已打开的 Edge 实例，避免每次执行都新开窗口。
        if self._is_page_alive(self.__class__._shared_page):
            self.page = self.__class__._shared_page
            return self.page

        edge_candidates = [
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        ]
        edge_path = next((p for p in edge_candidates if p.exists()), None)
        if not edge_path:
            raise FileNotFoundError("未找到本机 Edge 浏览器 msedge.exe")

        user_data_dir = Path(os.environ.get("LOCALAPPDATA", "")) / self.PROFILE_DIR_NAME
        user_data_dir.mkdir(parents=True, exist_ok=True)

        options = EdgeOptions()
        options.set_browser_path(str(edge_path))
        options.auto_port()
        options.set_argument(f"--user-data-dir={user_data_dir}")
        options.set_argument("--new-window")
        options.set_argument("--no-first-run")
        options.set_argument("--no-default-browser-check")
        options.set_argument("--disable-sync")

        self.page = EdgePage(options)
        self.__class__._shared_page = self.page
        return self.page

    def _can_enter_initial(self):
        self.page.get(self.initial_url)
        return bool(self.page.ele('x://span[text()="商品"]', timeout=1.2))

    def login(self):
        from SaihuERPLogin import SaihuERPLogin

        page = self._start_edge()
        if self._can_enter_initial():
            self.reused_session = True
            print("检测到当前会话可直接进入 INITIAL_URL。", flush=True)
            return page

        self.reused_session = False
        print("INITIAL_URL 无法直接进入，调用 SaihuERPLogin 执行稳定登录。", flush=True)
        login_client = SaihuERPLogin(
            page=page,
            username=self.username,
            password=self.password,
            img_dir=str(Path(__file__).resolve().parent),
        )
        login_client.login()
        page.get(self.initial_url)
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









