"""
代理 IP 池模块 — 给 DrissionPage/Chrome 提供轮转代理。

用法:
    pool = ProxyPool()
    pool.load_from_file("proxies.txt")
    proxy = pool.get_proxy()       # {"http": "http://user:pass@1.2.3.4:8080", "https": ...}
    pool.release_proxy(proxy)      # 用完归还（空操作）
    pool.mark_bad(proxy)           # 标记不可用，自动换下一个
"""

from __future__ import annotations

import random
import re
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

# 代理格式:  protocol://[user:pass@]host[:port]
_PROXY_RE = re.compile(
    r"^(?P<proto>https?|socks5|socks5h)://"
    r"(?:(?P<user>[^:]+):(?P<pass>[^@]+)@)?"
    r"(?P<host>[^:]+)"
    r"(?::(?P<port>\d+))?$"
)


@dataclass
class ProxyEntry:
    raw: str = ""
    proto: str = "http"
    host: str = ""
    port: int = 0
    user: Optional[str] = None
    password: Optional[str] = None
    failures: int = 0
    last_used: float = 0.0
    disabled: bool = False


class ProxyPool:
    """线程安全的代理轮转池。"""

    def __init__(self, max_failures: int = 3, cooldown_seconds: float = 2.0) -> None:
        self._entries: List[ProxyEntry] = []
        self._lock = threading.Lock()
        self._max_failures = max_failures
        self._cooldown = cooldown_seconds

    # ── 公开 API ──────────────────────────────────────────────

    def load_from_file(self, path: str | Path, encoding: str = "utf-8") -> int:
        """从文件加载代理，每行一条，跳过空行/注释。"""
        p = Path(path)
        if not p.exists():
            return 0
        lines = p.read_text(encoding=encoding).splitlines()
        count = 0
        with self._lock:
            for line in lines:
                stripped = line.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                entry = self._parse_line(stripped)
                if entry is not None:
                    self._entries.append(entry)
                    count += 1
        return count

    def add_proxy(self, raw: str) -> bool:
        """手动添加一条代理。"""
        entry = self._parse_line(raw)
        if entry is None:
            return False
        with self._lock:
            self._entries.append(entry)
        return True

    def get_proxy(self) -> Optional[Dict[str, str]]:
        """取一个可用代理，返回 DrissionPage proxy 字典，无可用代理时返回 None（直连）。"""
        entries = self._alive_entries()
        if not entries:
            return None

        weights = [1.0 / (1.0 + e.failures) for e in entries]
        with self._lock:
            chosen: ProxyEntry = random.choices(entries, weights=weights, k=1)[0]
            chosen.last_used = time.time()

        return self._to_dict(chosen)

    def release_proxy(self, proxy: Optional[Dict[str, str]]) -> None:
        """归还代理（占位）。"""
        pass

    def mark_bad(self, proxy: Optional[Dict[str, str]]) -> None:
        """标记代理不可用，失败计数 +1，超阈值则禁用。"""
        if proxy is None:
            return
        raw = (proxy.get("http") or proxy.get("https") or "").strip()
        with self._lock:
            for entry in self._entries:
                if entry.raw == raw:
                    entry.failures += 1
                    if entry.failures >= self._max_failures:
                        entry.disabled = True
                    break

    def mark_ok(self, proxy: Optional[Dict[str, str]]) -> None:
        """标记代理可用，重置失败计数。"""
        if proxy is None:
            return
        raw = (proxy.get("http") or proxy.get("https") or "").strip()
        with self._lock:
            for entry in self._entries:
                if entry.raw == raw:
                    entry.failures = 0
                    entry.disabled = False
                    break

    def stats(self) -> Dict[str, object]:
        """池状态统计。"""
        with self._lock:
            total = len(self._entries)
            alive = sum(1 for e in self._entries if not e.disabled)
        return {"total": total, "alive": alive, "dead": total - alive}

    def clear(self) -> None:
        """清空池。"""
        with self._lock:
            self._entries.clear()

    # ── 内部 ──────────────────────────────────────────────────

    def _parse_line(self, raw: str) -> Optional[ProxyEntry]:
        m = _PROXY_RE.match(raw)
        if not m:
            return None
        proto = m.group("proto")
        host = m.group("host")
        port_str = m.group("port")
        port = int(port_str) if port_str else (443 if proto == "https" else 80)
        user = m.group("user")
        password = m.group("pass")
        return ProxyEntry(
            raw=raw.strip(),
            proto=proto,
            host=host,
            port=port,
            user=user,
            password=password,
        )

    def _alive_entries(self) -> List[ProxyEntry]:
        now = time.time()
        with self._lock:
            alive = []
            for e in self._entries:
                if e.disabled:
                    continue
                if e.last_used and (now - e.last_used) < self._cooldown:
                    continue
                alive.append(e)
        return alive

    def _to_dict(self, entry: ProxyEntry) -> Dict[str, str]:
        auth = f"{entry.user}:{entry.password}@" if entry.user and entry.password else ""
        url = f"{entry.proto}://{auth}{entry.host}:{entry.port}"
        return {"http": url, "https": url}


def quick_check(proxy_file: str | Path, sample_size: int = 5,
                timeout: int = 10) -> Dict[str, object]:
    """快速检查代理文件前 sample_size 条是否可连通。"""
    pool = ProxyPool()
    count = pool.load_from_file(proxy_file)
    results: List[Dict[str, object]] = []
    alive_count = 0
    for _ in range(min(sample_size, count)):
        proxy = pool.get_proxy()
        if proxy is None:
            break
        # 简单 TCP/SSL 连通性检查
        import socket
        import ssl
        from urllib.parse import urlparse
        parsed = urlparse(proxy.get("http", ""))
        host = parsed.hostname or ""
        port = parsed.port or 443
        ok = False
        try:
            ctx = ssl.create_default_context()
            with socket.create_connection((host, port), timeout=timeout) as sock:
                with ctx.wrap_socket(sock, server_hostname=host):
                    ok = True
        except Exception:
            ok = False
        results.append({"proxy": proxy["http"], "alive": ok})
        if ok:
            alive_count += 1
    return {
        "loaded": count,
        "checked": len(results),
        "alive": alive_count,
        "results": results,
    }
