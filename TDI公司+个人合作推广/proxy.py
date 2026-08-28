"""DrissionPage 浏览器专用的认证 SOCKS5 本地代理桥。"""

import ast
import asyncio
import importlib.util
import json
import socket
import sys
import threading
import time
from pathlib import Path
from urllib.parse import unquote, urlparse
from urllib.request import ProxyHandler, Request, build_opener


class Proxy:
    """把 ipfiy 的认证 SOCKS5 转为仅监听本机的 HTTP 代理。"""

    def __init__(
        self,
        baseDir,
        outputDir,
        config,
        log=None,
    ):
        """初始化代理来源、桥接进程和公开状态。"""
        self.baseDir = Path(baseDir)
        self.outputDir = Path(outputDir)
        self.outputDir.mkdir(parents=True, exist_ok=True)
        self.config = config
        self.log = log or (lambda message: None)
        self.required = bool(config.get("proxyRequired", True))
        self.thread = None
        self.loop = None
        self.server = None
        self.startedEvent = threading.Event()
        self.startError = ""
        self.localUrl = ""
        self.exitIp = ""
        self.upstreamData = None

    def readIpfiy(self, path):
        """用 AST 读取 set_proxy 参数，不执行 ipfiy.py 中的测试代码。"""
        sourcePath = Path(path)
        if not sourcePath.exists():
            return {}
        tree = ast.parse(sourcePath.read_text(encoding="utf-8"), filename=str(sourcePath))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if not isinstance(node.func, ast.Attribute) or node.func.attr != "set_proxy":
                continue
            if len(node.args) < 3:
                continue
            try:
                host = str(ast.literal_eval(node.args[1]))
                port = int(ast.literal_eval(node.args[2]))
            except (ValueError, TypeError):
                continue
            values = {item.arg: ast.literal_eval(item.value) for item in node.keywords if item.arg}
            return {
                "scheme": "socks5",
                "host": host,
                "port": port,
                "username": str(values.get("username") or ""),
                "password": str(values.get("password") or ""),
                "source": "ipfiy.py",
            }
        return {}

    def ipfiyPaths(self):
        """返回外置与 PyInstaller 内置的 ipfiy.py 候选路径。"""
        candidates = [self.baseDir / "ipfiy.py"]
        bundleValue = str(getattr(sys, "_MEIPASS", "") or "").strip()
        if bundleValue:
            bundleRoot = Path(bundleValue)
            candidates.extend([bundleRoot / "ipfiy.py", bundleRoot / "file" / "ipfiy.py"])
        candidates.append(Path(__file__).resolve().parent / "ipfiy.py")

        output = []
        seen = set()
        for path in candidates:
            value = str(path)
            if value not in seen:
                seen.add(value)
                output.append(path)
        return output

    def readConfig(self):
        """读取本地配置中的 SOCKS5 地址，缺失时兼容 ipfiy.py。"""
        proxyUrl = str(self.config.get("proxyUrl") or "").strip()
        if proxyUrl:
            if "://" not in proxyUrl:
                proxyUrl = "socks5://" + proxyUrl
            parsed = urlparse(proxyUrl)
            return {
                "scheme": parsed.scheme.lower(),
                "host": str(parsed.hostname or ""),
                "port": int(parsed.port or 0),
                "username": str(
                    self.config.get("proxyUsername") or unquote(parsed.username or "")
                ),
                "password": str(
                    self.config.get("proxyPassword") or unquote(parsed.password or "")
                ),
                "source": "config.local.json",
            }
        for path in self.ipfiyPaths():
            value = self.readIpfiy(path)
            if value:
                return value
        return {}

    def upstream(self):
        """返回缓存后的上游代理配置。"""
        if self.upstreamData is None:
            self.upstreamData = self.readConfig()
        return dict(self.upstreamData)

    def validationError(self):
        """验证上游协议、主机、端口和认证字段。"""
        upstream = self.upstream()
        if not upstream:
            return "未找到 ipfiy.py 或 config.local.json 代理配置" if self.required else ""
        if upstream.get("scheme") not in {"socks5", "socks5h", "socks"}:
            return "浏览器上游代理必须是 SOCKS5"
        if not upstream.get("host") or not int(upstream.get("port") or 0):
            return "代理主机或端口为空"
        if bool(upstream.get("username")) != bool(upstream.get("password")):
            return "代理账号和密码必须同时配置"
        return ""

    def configured(self):
        """判断代理是否具备可启动条件。"""
        return not self.validationError()

    def proxyHealthCheckTargets(self):
        """返回 SOCKS5 上游应能访问的 HTTPS 健康检查目标。"""
        rawTargets = self.config.get("proxyHealthCheckTargets") or [
            "api.ipify.org:443",
            "www.google.com:443",
        ]
        targets = []
        for item in rawTargets:
            text = str(item or "").strip()
            if not text:
                continue
            host, _, port = text.rpartition(":")
            if not host:
                host, port = text, "443"
            try:
                targets.append((host, int(port or 443)))
            except ValueError:
                continue
        return targets or [("api.ipify.org", 443)]

    def testUpstream(self, timeout = 15):
        """先通过 SOCKS5 建立 HTTPS TCP 连接，失败时禁止继续运行。"""
        upstream = self.upstream()
        try:
            import socks
        except Exception as error:
            raise RuntimeError("缺少 PySocks，无法测试 SOCKS5 代理") from error
        errors = []
        for host, port in self.proxyHealthCheckTargets():
            client = socks.socksocket()
            client.settimeout(timeout)
            try:
                client.set_proxy(
                    socks.SOCKS5,
                    str(upstream["host"]),
                    int(upstream["port"]),
                    rdns=True,
                    username=str(upstream.get("username") or "") or None,
                    password=str(upstream.get("password") or "") or None,
                )
                client.connect((host, port))
                return
            except Exception as error:
                errors.append(f"{host}:{port} -> {str(error)[:120]}")
            finally:
                client.close()
        raise RuntimeError(f"上游 SOCKS5 代理测试失败：{'；'.join(errors)[:260]}")

    def findPort(self):
        """优先使用固定本地端口，占用时改用系统分配端口。"""
        preferred = max(0, int(self.config.get("proxyBridgePort", 8899)))
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            try:
                client.bind(("127.0.0.1", preferred))
            except OSError:
                client.bind(("127.0.0.1", 0))
            return int(client.getsockname()[1])
        finally:
            client.close()

    def waitPort(self, port, timeout = 15):
        """等待本地 HTTP 代理开始监听。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.startError:
                raise RuntimeError(self.startError)
            if self.thread and not self.thread.is_alive():
                raise RuntimeError("本地代理桥启动失败")
            try:
                connection = socket.create_connection(("127.0.0.1", port), timeout=0.5)
                connection.close()
                return
            except OSError:
                time.sleep(0.2)
        raise RuntimeError("本地代理桥启动超时")

    def runBridge(self, localUri, remoteUri):
        """在独立事件循环中运行 pproxy，不把认证信息写入命令行。"""
        loop = asyncio.new_event_loop()
        self.loop = loop
        asyncio.set_event_loop(loop)
        try:
            from pproxy.server import proxies_by_uri

            localOption = proxies_by_uri(localUri)
            remoteOption = proxies_by_uri(remoteUri)
            arguments = {
                "rserver": [remoteOption],
                "debug": 0,
                "authtime": 86400 * 30,
                "block": None,
                "salgorithm": "fa",
                "ruport": False,
            }
            self.server = loop.run_until_complete(localOption.start_server(arguments))
            self.startedEvent.set()
            loop.run_forever()
        except Exception as error:
            self.startError = f"本地代理桥启动失败：{str(error)[:160]}"
            self.startedEvent.set()
        finally:
            if self.server:
                self.server.close()
                if hasattr(self.server, "wait_closed"):
                    loop.run_until_complete(self.server.wait_closed())
            tasks = list(asyncio.all_tasks(loop))
            for task in tasks:
                task.cancel()
            if tasks:
                loop.run_until_complete(asyncio.gather(*tasks, return_exceptions=True))
            loop.run_until_complete(loop.shutdown_asyncgens())
            loop.close()
            self.server = None
            self.loop = None

    def testBridge(self, timeout = 30):
        """通过本地 HTTP 代理读取出口 IP，确认没有直连回退。"""
        if not self.localUrl:
            raise RuntimeError("本地代理桥尚未启动")
        opener = build_opener(ProxyHandler({"http": self.localUrl, "https": self.localUrl}))
        request = Request(
            "https://api.ipify.org?format=json",
            headers={"Accept": "application/json", "User-Agent": "TDI-Automation/1.0"},
        )
        try:
            response = opener.open(request, timeout=timeout)
            payload = json.loads(response.read().decode("utf-8"))
            exitIp = str(payload.get("ip") or "").strip()
        except Exception as error:
            raise RuntimeError(f"本地代理桥出口验证失败：{str(error)[:160]}") from error
        if not exitIp:
            raise RuntimeError("本地代理桥未返回出口 IP")
        self.exitIp = exitIp
        return exitIp

    def start(self):
        """启动仅监听 127.0.0.1 的 pproxy 桥并验证出口。"""
        if self.thread and self.thread.is_alive() and self.localUrl:
            return self.localUrl
        error = self.validationError()
        if error:
            if self.required:
                raise RuntimeError(error)
            return ""
        if importlib.util.find_spec("pproxy") is None:
            raise RuntimeError("缺少 pproxy，无法建立浏览器代理桥")

        # 先验证上游，再启动本地桥；任一步失败都不会让浏览器直连。
        self.testUpstream()
        upstream = self.upstream()
        port = self.findPort()
        self.localUrl = f"http://127.0.0.1:{port}"
        remoteUrl = f"socks5://{upstream['host']}:{upstream['port']}/"
        username = str(upstream.get("username") or "")
        password = str(upstream.get("password") or "")
        if username:
            remoteUrl += f"#{username}:{password}"
        self.startedEvent.clear()
        self.startError = ""
        try:
            self.thread = threading.Thread(
                target=self.runBridge,
                args=(self.localUrl + "/", remoteUrl),
                name="TdiProxyBridge",
                daemon=True,
            )
            self.thread.start()
            self.startedEvent.wait(timeout=10)
            if self.startError:
                raise RuntimeError(self.startError)
            self.waitPort(port)
            self.testBridge()
        except Exception:
            self.stop()
            raise
        self.log(f"浏览器代理桥已就绪：127.0.0.1:{port}，出口验证通过。")
        return self.localUrl

    def browserUrl(self):
        """返回 DrissionPage 应使用的本地 HTTP 代理地址。"""
        if self.required and not self.localUrl:
            raise RuntimeError("浏览器代理桥尚未通过验证")
        return self.localUrl

    def publicLabel(self):
        """返回不含账号密码的代理状态文本。"""
        upstream = self.upstream()
        if not upstream:
            return "未配置"
        source = str(upstream.get("source") or "本地配置")
        if self.localUrl and self.exitIp:
            return f"{source} · 已验证出口 {self.exitIp}"
        return f"{source} · 待验证"

    def status(self):
        """返回可安全展示的代理状态。"""
        return {
            "configured": self.configured(),
            "required": self.required,
            "running": bool(self.thread and self.thread.is_alive()),
            "localUrl": self.localUrl,
            "exitIp": self.exitIp,
            "label": self.publicLabel(),
            "error": self.validationError(),
        }

    def stop(self):
        """停止本地代理桥事件循环并等待线程退出。"""
        loop = self.loop
        thread = self.thread
        self.thread = None
        if loop and loop.is_running():
            loop.call_soon_threadsafe(loop.stop)
        if thread and thread.is_alive() and thread is not threading.current_thread():
            thread.join(timeout=5)
        self.localUrl = ""
