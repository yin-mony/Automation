"""
1688 开放平台 API 调用示例（独立模块，暂不关联浏览器抓包逻辑）

使用前请在开放平台完成：
1. 应用 → API 集成 → 申请「商品详情」相关接口权限（如 alibaba.product.get）
2. 配置「授权回调地址」为 https://127.0.0.1:9527/callback（与 REDIRECT_URI 一致）
3. 完成 OAuth 授权；token 自动保存到同目录 `.1688_token.json`（已加入 .gitignore）
4. 申请「根据商品ID列表获取商品(卖家)」接口权限：alibaba.product.getByIdList
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import ssl
import subprocess
import threading
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import requests

DEFAULT_CALLBACK_HOST = "127.0.0.1"
DEFAULT_CALLBACK_PORT = 9527
DEFAULT_CALLBACK_PATH = "/callback"
DEFAULT_REDIRECT_URI = (
    f"https://{DEFAULT_CALLBACK_HOST}:{DEFAULT_CALLBACK_PORT}{DEFAULT_CALLBACK_PATH}"
)
SCRIPT_DIR = Path(__file__).resolve().parent
CERT_DIR = SCRIPT_DIR / ".oauth_certs"
TOKEN_FILE = SCRIPT_DIR / ".1688_token.json"
OPENSSL_CANDIDATES = (
    "openssl",
    r"C:\Program Files\Git\usr\bin\openssl.exe",
    r"C:\Program Files (x86)\Git\usr\bin\openssl.exe",
)


def _find_openssl() -> str:
    for candidate in OPENSSL_CANDIDATES:
        if candidate == "openssl":
            continue
        if Path(candidate).exists():
            return candidate
    return "openssl"


class Ali1688Client:
    """1688 开放平台 param2 接口客户端"""

    GATEWAY = "https://gw.open.1688.com"
    AUTH_URL = "https://auth.1688.com/oauth/authorize"

    def __init__(self, app_key: str, app_secret: str, redirect_uri: str):
        self.app_key = app_key
        self.app_secret = app_secret
        self.redirect_uri = redirect_uri

    @staticmethod
    def extract_product_id(url_or_id: str) -> str:
        """从 1688 链接或纯数字中提取商品 ID（offerId / productId）"""
        text = str(url_or_id).strip()
        match = re.search(r"offer/(\d+)", text)
        if match:
            return match.group(1)
        if text.isdigit():
            return text
        raise ValueError(f"无法解析商品 ID: {url_or_id}")

    def _aop_signature(self, url_path: str, params: dict[str, Any]) -> str:
        """
        1688 param2 签名：HMAC-SHA1
        url_path 示例: param2/1/com.alibaba.product/alibaba.product.get/2414217
        """
        sign_params = {
            k: v for k, v in params.items()
            if k != "_aop_signature" and v is not None and str(v) != ""
        }
        pieces = sorted(f"{k}{v}" for k, v in sign_params.items())
        sign_content = url_path + "".join(pieces)
        digest = hmac.new(
            self.app_secret.encode("utf-8"),
            sign_content.encode("utf-8"),
            hashlib.sha1,
        ).digest()
        return digest.hex().upper()

    def get_authorize_url(self, state: str = "1688_demo") -> str:
        """生成 OAuth 授权页 URL，浏览器打开后登录并同意授权"""
        params = {
            "client_id": self.app_key,
            "site": "1688",
            "redirect_uri": self.redirect_uri,
            "response_type": "code",
            "state": state,
        }
        return f"{self.AUTH_URL}?{urllib.parse.urlencode(params)}"

    def get_access_token(self, code: str) -> dict:
        """用授权码换取 access_token（code 有效期约 2 分钟，一次性使用）"""
        url = f"{self.GATEWAY}/openapi/http/1/system.oauth2/getToken/{self.app_key}"
        data = {
            "grant_type": "authorization_code",
            "need_refresh_token": "true",
            "client_id": self.app_key,
            "client_secret": self.app_secret,
            "redirect_uri": self.redirect_uri,
            "code": code,
        }
        resp = requests.post(url, data=data, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def refresh_access_token(self, refresh_token: str) -> dict:
        """刷新 access_token"""
        url = f"{self.GATEWAY}/openapi/http/1/system.oauth2/getToken/{self.app_key}"
        data = {
            "grant_type": "refresh_token",
            "client_id": self.app_key,
            "client_secret": self.app_secret,
            "refresh_token": refresh_token,
        }
        resp = requests.post(url, data=data, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def call_api(
        self,
        namespace: str,
        api_name: str,
        access_token: str,
        version: str = "1",
        **biz_params: Any,
    ) -> dict:
        """通用 param2 接口调用"""
        url_path = f"param2/1/{namespace}/{api_name}/{self.app_key}"
        api_url = f"{self.GATEWAY}/openapi/{url_path}"

        params: dict[str, Any] = {
            "access_token": access_token,
            **biz_params,
        }
        params["_aop_timestamp"] = str(int(time.time() * 1000))
        params["_aop_signature"] = self._aop_signature(url_path, params)

        resp = requests.post(api_url, data=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    @staticmethod
    def parse_product_id_list(raw: str | list[str]) -> list[str]:
        """解析商品 ID 列表，支持逗号/换行/空格分隔，或 JSON 数组字符串"""
        if isinstance(raw, list):
            items = raw
        else:
            text = str(raw).strip()
            if not text:
                return []
            if text.startswith("["):
                parsed = json.loads(text)
                if not isinstance(parsed, list):
                    raise ValueError("productIdList JSON 必须是数组")
                items = [str(x).strip() for x in parsed]
            else:
                items = re.split(r"[\s,;]+", text)

        product_ids: list[str] = []
        for item in items:
            value = str(item).strip()
            if not value:
                continue
            if "offer/" in value or "http" in value:
                value = Ali1688Client.extract_product_id(value)
            if not value.isdigit():
                raise ValueError(f"无效商品 ID: {item}")
            product_ids.append(value)
        return product_ids

    def get_product(self, product_id: str, access_token: str) -> dict:
        """
        获取商品详情（含标题、价格、SKU 规格等）
        接口: com.alibaba.product:alibaba.product.get
        """
        return self.call_api(
            namespace="com.alibaba.product",
            api_name="alibaba.product.get",
            access_token=access_token,
            productID=str(product_id),
            webSite="1688",
        )

    def get_products_by_id_list(
        self,
        product_ids: list[str],
        access_token: str,
        website: str = "1688",
    ) -> dict:
        """
        根据商品 ID 列表获取商品（卖家接口）
        接口: com.alibaba.product:alibaba.product.getByIdList
        注意：只能查询当前授权卖家账号下的商品
        """
        if not product_ids:
            raise ValueError("product_ids 不能为空")
        if len(product_ids) > 20:
            raise ValueError("单次最多查询 20 个商品 ID")

        return self.call_api(
            namespace="com.alibaba.product",
            api_name="alibaba.product.getByIdList",
            access_token=access_token,
            productIdList=json.dumps([int(pid) for pid in product_ids]),
            webSite=website,
        )


def load_token_file() -> dict[str, Any] | None:
    """从本地 `.1688_token.json` 读取 token"""
    if not TOKEN_FILE.exists():
        return None
    with TOKEN_FILE.open(encoding="utf-8") as f:
        return json.load(f)


def save_token_file(token_info: dict[str, Any]) -> None:
    """保存 token 到本地文件，供后续运行和自动刷新使用"""
    data = {
        **token_info,
        "saved_at": int(time.time()),
    }
    with TOKEN_FILE.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"token 已保存到: {TOKEN_FILE}")


def is_access_token_expired(token_info: dict[str, Any], buffer_seconds: int = 300) -> bool:
    """判断 access_token 是否即将或已经过期（默认提前 5 分钟刷新）"""
    saved_at = token_info.get("saved_at")
    expires_in = token_info.get("expires_in")
    if saved_at is None or expires_in is None:
        return False
    try:
        return time.time() >= int(saved_at) + int(expires_in) - buffer_seconds
    except (TypeError, ValueError):
        return False


def resolve_access_token(client: Ali1688Client) -> str:
    """
    按优先级获取可用 access_token：
    1. 环境变量 ALI1688_ACCESS_TOKEN
    2. 本地 `.1688_token.json`（过期则自动 refresh）
    """
    env_token = os.getenv("ALI1688_ACCESS_TOKEN", "").strip()
    if env_token:
        return env_token

    token_info = load_token_file()
    if not token_info:
        return ""

    access_token = str(token_info.get("access_token", "")).strip()
    if not access_token:
        return ""

    if not is_access_token_expired(token_info):
        return access_token

    refresh_token = (
        os.getenv("ALI1688_REFRESH_TOKEN", "").strip()
        or str(token_info.get("refresh_token", "")).strip()
    )
    if not refresh_token:
        print("access_token 已过期，且未找到 refresh_token，请重新授权")
        return ""

    print("access_token 即将或已经过期，正在用 refresh_token 刷新…")
    try:
        new_info = client.refresh_access_token(refresh_token)
    except requests.HTTPError as exc:
        print(f"刷新 token 失败: {exc}")
        if exc.response is not None:
            print(exc.response.text)
        return ""

    save_token_file(new_info)
    return str(new_info.get("access_token", "")).strip()


def ensure_local_ssl_cert(host: str = DEFAULT_CALLBACK_HOST) -> tuple[Path, Path]:
    """生成本地 OAuth 回调用的自签名 HTTPS 证书（首次自动创建）"""
    CERT_DIR.mkdir(exist_ok=True)
    cert_file = CERT_DIR / "localhost.pem"
    key_file = CERT_DIR / "localhost.key"
    if cert_file.exists() and key_file.exists():
        return cert_file, key_file

    cmd = [
        _find_openssl(),
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-keyout",
        str(key_file),
        "-out",
        str(cert_file),
        "-days",
        "3650",
        "-nodes",
        "-subj",
        f"/CN={host}",
        "-addext",
        f"subjectAltName=DNS:localhost,IP:{host}",
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
    except FileNotFoundError as exc:
        raise RuntimeError(
            "本地 HTTPS 回调需要 OpenSSL 生成证书。"
            "可安装 Git for Windows（自带 openssl），或手动执行：\n"
            f"  openssl req -x509 -newkey rsa:2048 -keyout {key_file} "
            f"-out {cert_file} -days 3650 -nodes -subj /CN={host} "
            f'-addext "subjectAltName=DNS:localhost,IP:{host}"'
        ) from exc
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or "").strip()
        raise RuntimeError(f"生成自签名证书失败: {stderr or exc}") from exc

    return cert_file, key_file


def redirect_uri_uses_https(redirect_uri: str) -> bool:
    return redirect_uri.lower().startswith("https://")


def wait_for_oauth_callback(
    host: str = DEFAULT_CALLBACK_HOST,
    port: int = DEFAULT_CALLBACK_PORT,
    path: str = DEFAULT_CALLBACK_PATH,
    timeout: int = 300,
    use_https: bool | None = None,
    redirect_uri: str | None = None,
) -> str:
    """启动本地回调服务，等待 1688 OAuth 回调并返回授权 code"""
    if use_https is None:
        use_https = redirect_uri_uses_https(redirect_uri or DEFAULT_REDIRECT_URI)

    done = threading.Event()
    result: dict[str, str | None] = {"code": None, "error": None}

    class OAuthCallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path != path:
                self.send_response(404)
                self.end_headers()
                return

            params = parse_qs(parsed.query)
            if "error" in params:
                result["error"] = params["error"][0]
            else:
                result["code"] = params.get("code", [None])[0]

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            if result["code"]:
                body = (
                    "<html><body><h2>授权成功</h2>"
                    "<p>可关闭此页面，返回终端查看 token 信息。</p></body></html>"
                )
            else:
                body = (
                    f"<html><body><h2>授权失败</h2>"
                    f"<p>{result.get('error') or '未收到 code'}</p></body></html>"
                )
            self.wfile.write(body.encode("utf-8"))
            done.set()
            threading.Thread(target=self.server.shutdown, daemon=True).start()

        def log_message(self, format, *args):
            return

    server = HTTPServer((host, port), OAuthCallbackHandler)
    if use_https:
        cert_file, key_file = ensure_local_ssl_cert(host)
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(certfile=str(cert_file), keyfile=str(key_file))
        server.socket = context.wrap_socket(server.socket, server_side=True)

    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    if not done.wait(timeout):
        server.shutdown()
        raise TimeoutError(f"等待 OAuth 回调超时（{timeout}s），请重试")

    if result["error"]:
        raise RuntimeError(f"OAuth 授权失败: {result['error']}")
    if not result["code"]:
        raise RuntimeError("OAuth 回调未包含 code")
    return result["code"]


def print_products_summary(result: dict) -> None:
    """打印 getByIdList 返回的商品摘要"""
    if not result.get("success", True) and result.get("message"):
        print(f"接口提示: {result['message']}")

    products = (
        result.get("productInfos")
        or result.get("result", {}).get("productInfos")
        or result.get("productInfo")
        or []
    )
    if isinstance(products, dict):
        products = [products]
    if not products:
        return

    print("\n" + "=" * 60)
    print(f"共返回 {len(products)} 个商品：")
    for i, product in enumerate(products, 1):
        subject = product.get("subject") or product.get("title") or "—"
        product_id = product.get("productID") or product.get("productId") or "—"
        status = product.get("status") or "—"
        print(f"  [{i}] ID={product_id}  状态={status}  标题={subject}")


def main():
    """1688 官方 API 调用示例入口"""
    app_key = os.getenv("ALI1688_APP_KEY", "").strip()
    app_secret = os.getenv("ALI1688_APP_SECRET", "").strip()
    if not app_key or not app_secret:
        raise ValueError(
            "请设置环境变量 ALI1688_APP_KEY 与 ALI1688_APP_SECRET（勿将密钥写入代码或提交 Git）"
        )
    redirect_uri = os.getenv("ALI1688_REDIRECT_URI", DEFAULT_REDIRECT_URI)

    # 示例商品 ID 列表（逗号分隔，仅支持当前授权卖家自己的商品）
    demo_products = os.getenv(
        "ALI1688_PRODUCT_IDS",
        "41618176125",
    )

    client = Ali1688Client(app_key, app_secret, redirect_uri)

    # ---------- 第一步：OAuth 授权（首次使用时执行） ----------
    access_token = resolve_access_token(client)
    if not access_token:
        auth_code = os.getenv("ALI1688_AUTH_CODE", "")
        print("=" * 60)
        print("尚未配置 access_token，请先完成 OAuth 授权：")
        print("1. 在开放平台应用设置中配置回调地址，需与 REDIRECT_URI 完全一致")
        print(f"   当前 REDIRECT_URI: {redirect_uri}")
        print("2. 先运行本脚本并保持终端窗口不关闭，再打开开放平台测试链接完成授权")
        print("   官方测试页会回调 https://127.0.0.1:9527/callback，需本地 HTTPS 服务已启动")
        print("=" * 60)

        if not auth_code:
            auth_url = client.get_authorize_url()
            callback_host = os.getenv("ALI1688_CALLBACK_HOST", DEFAULT_CALLBACK_HOST)
            callback_port = int(os.getenv("ALI1688_CALLBACK_PORT", str(DEFAULT_CALLBACK_PORT)))
            callback_path = os.getenv("ALI1688_CALLBACK_PATH", DEFAULT_CALLBACK_PATH)
            use_https = redirect_uri_uses_https(redirect_uri)
            scheme = "https" if use_https else "http"
            skip_open_browser = os.getenv("ALI1688_SKIP_OPEN_BROWSER", "").lower() in {
                "1", "true", "yes",
            }

            print(f"本地回调服务: {scheme}://{callback_host}:{callback_port}{callback_path}")
            if skip_open_browser:
                print("已跳过自动打开浏览器，请手动打开开放平台测试链接或授权 URL")
            else:
                print("正在启动回调服务并打开授权页…")
            print(auth_url)
            print()

            auth_result: dict[str, Any] = {}

            def _capture_callback() -> None:
                try:
                    auth_result["code"] = wait_for_oauth_callback(
                        host=callback_host,
                        port=callback_port,
                        path=callback_path,
                        use_https=use_https,
                        redirect_uri=redirect_uri,
                    )
                except Exception as exc:
                    auth_result["error"] = exc

            callback_thread = threading.Thread(target=_capture_callback, daemon=True)
            callback_thread.start()
            time.sleep(0.5)
            if not skip_open_browser:
                webbrowser.open(auth_url)
            callback_thread.join()

            if auth_result.get("error"):
                print(f"OAuth 回调失败: {auth_result['error']}")
                return

            auth_code = auth_result.get("code", "")
            if not auth_code:
                print("未收到授权 code，请重试")
                return
            print(f"已收到授权 code: {auth_code[:8]}…")

        token_info = client.get_access_token(auth_code)
        print("授权成功，token 信息：")
        print(json.dumps(token_info, ensure_ascii=False, indent=2))
        access_token = token_info.get("access_token", "")
        if not access_token:
            print("未获取到 access_token，请检查 code 是否有效或已过期")
            return
        save_token_file(token_info)
        print("后续可直接运行本脚本，无需再次授权（过期会自动刷新）\n")

    # ---------- 第二步：调用 getByIdList 批量获取卖家商品 ----------
    try:
        product_ids = client.parse_product_id_list(demo_products)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"商品 ID 解析失败: {exc}")
        return

    print(f"正在查询商品 ID 列表 ({len(product_ids)} 个): {', '.join(product_ids)}")

    try:
        result = client.get_products_by_id_list(product_ids, access_token)
    except requests.HTTPError as exc:
        print(f"HTTP 请求失败: {exc}")
        if exc.response is not None:
            print(exc.response.text)
        return

    print("=" * 60)
    print("API 原始响应：")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print_products_summary(result)


if __name__ == "__main__":
    main()
