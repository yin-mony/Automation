import os
import socket
import subprocess
import time
import urllib.error
import urllib.request
from pathlib import Path

import psutil
from DrissionPage import ChromiumPage


class EdgeBrowserRunner:
    """Edge 浏览器拉起与调试连接公共类。"""

    DEFAULT_DEBUG_PORT = 9333

    @staticmethod
    def resolve_edge_path():
        """定位本机 Edge 可执行文件路径。"""
        edge_candidates = [
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        ]
        edge_path = next((p for p in edge_candidates if p.exists()), None)
        if not edge_path:
            raise FileNotFoundError("未找到本机 Edge 浏览器 msedge.exe")
        return edge_path

    @staticmethod
    def is_edge_running():
        """判断本机是否已有 Edge 进程。"""
        for proc in psutil.process_iter(["name"]):
            name = (proc.info.get("name") or "").lower()
            if name == "msedge.exe":
                return True
        return False

    @staticmethod
    def connect_existing_debug(debug_port):
        """尝试连接已存在的 Edge 调试端口。"""
        try:
            page = ChromiumPage(f"127.0.0.1:{debug_port}")
            _ = page.url
            return page
        except Exception:
            return None

    @staticmethod
    def is_port_open(port):
        """检查本机端口是否被占用。"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            return sock.connect_ex(("127.0.0.1", int(port))) == 0

    @staticmethod
    def is_edge_debug_port(debug_port):
        """判断调试端口是否由 Edge 提供。"""
        version_url = f"http://127.0.0.1:{debug_port}/json/version"
        try:
            with urllib.request.urlopen(version_url, timeout=1.2) as resp:
                body = resp.read().decode("utf-8", errors="ignore")
            return "Edg/" in body or "Microsoft Edge" in body
        except (urllib.error.URLError, TimeoutError, OSError):
            return False

    @classmethod
    def find_existing_edge_debug_port(cls, preferred_port):
        """扫描常见端口，寻找已存在的 Edge 调试端口。"""
        candidates = [preferred_port, 9222, 9333]
        candidates.extend(range(preferred_port - 10, preferred_port + 11))
        checked = set()
        for port in candidates:
            if port <= 0 or port in checked:
                continue
            checked.add(port)
            if cls.is_port_open(port) and cls.is_edge_debug_port(port):
                return port
        return None

    @classmethod
    def choose_debug_port(cls, preferred_port):
        """优先使用给定端口，如冲突则自动选择可用端口。"""
        if not cls.is_port_open(preferred_port):
            return preferred_port
        if cls.is_edge_debug_port(preferred_port):
            return preferred_port

        for port in range(preferred_port + 1, preferred_port + 31):
            if not cls.is_port_open(port):
                print(f"调试端口 {preferred_port} 已被占用，自动切换到 {port}。", flush=True)
                return port
        raise RuntimeError(f"未找到可用调试端口，请释放 {preferred_port} 附近端口后重试。")

    @staticmethod
    def open_url_in_new_tab(page, url):
        """在当前浏览器尽量新建标签页打开 URL。"""
        try:
            tab = page.new_tab(url)
            if tab:
                return tab
        except Exception:
            pass

        try:
            page.run_js(f'window.open("{url}", "_blank");')
            time.sleep(0.5)
        except Exception:
            pass

        page.get(url)
        return page

    @classmethod
    def enable_debug_on_running_edge(cls, debug_port, start_url, wait_seconds):
        """尝试在当前用户态 Edge 上启用调试端口并接管。"""
        edge_path = cls.resolve_edge_path()
        subprocess.Popen(
            [
                str(edge_path),
                f"--remote-debugging-port={debug_port}",
                "--new-window",
                start_url,
            ]
        )
        time.sleep(wait_seconds)
        return cls.is_edge_debug_port(debug_port)

    @classmethod
    def start_edge_and_connect(
        cls,
        debug_port=None,
        start_url=None,
        fresh_profile=False,
        wait_seconds=3,
    ):
        """启动本机 Edge 浏览器并连接调试端口，返回 ChromiumPage。"""
        debug_port = cls.choose_debug_port(debug_port or cls.DEFAULT_DEBUG_PORT)
        if not start_url:
            raise ValueError("start_url 不能为空")
        print("已配置为直接新开 Edge 实例（不复用、不接管）。", flush=True)

        edge_path = cls.resolve_edge_path()

        profile_name = "EdgeDebugProfile"
        if fresh_profile:
            profile_name = f"EdgeDebugProfile_{int(time.time())}"
        user_data_dir = Path(os.environ.get("LOCALAPPDATA", "")) / profile_name
        user_data_dir.mkdir(parents=True, exist_ok=True)

        subprocess.Popen(
            [
                str(edge_path),
                f"--remote-debugging-port={debug_port}",
                f"--user-data-dir={user_data_dir}",
                "--new-window",
                start_url,
            ]
        )
        time.sleep(wait_seconds)

        if not cls.is_edge_debug_port(debug_port):
            raise RuntimeError(
                f"Edge 调试端口 {debug_port} 未就绪，请确认浏览器启动参数是否生效。"
            )

        page = ChromiumPage(f"127.0.0.1:{debug_port}")
        page.get(start_url)
        return page
