import base64
import hashlib
import json
import logging
import os
import struct
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from logging.handlers import RotatingFileHandler
from urllib.parse import parse_qs, urlparse

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
LOG_DIR = os.path.join(BASE_DIR, "logs")


def load_env(path):
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("qyvx_callback")
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "callback.log"),
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger


load_env(ENV_PATH)
LOGGER = setup_logging()


def get_config():
    return {
        "token": os.getenv("QYWX_CALLBACK_TOKEN", "").strip(),
        "aes_key": os.getenv("QYWX_ENCODING_AES_KEY", "").strip(),
        "corp_id": os.getenv("QYWX_CORP_ID", "").strip(),
        "host": os.getenv("QYWX_BIND_HOST", "127.0.0.1").strip(),
        "port": int(os.getenv("QYWX_PORT", "8600")),
    }


def first_value(query, name):
    values = query.get(name, [""])
    return values[0] if values else ""


def signature(token, timestamp, nonce, encrypted):
    payload = "".join(sorted([token, timestamp, nonce, encrypted]))
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def decrypt_wecom_message(encrypted):
    config = get_config()
    if not config["token"] or not config["aes_key"]:
        raise ValueError("missing QYWX_CALLBACK_TOKEN or QYWX_ENCODING_AES_KEY")
    if len(config["aes_key"]) != 43:
        raise ValueError("QYWX_ENCODING_AES_KEY must be 43 characters")

    aes_key = base64.b64decode(config["aes_key"] + "=")
    cipher = Cipher(algorithms.AES(aes_key), modes.CBC(aes_key[:16]))
    decryptor = cipher.decryptor()
    plain = decryptor.update(base64.b64decode(encrypted)) + decryptor.finalize()

    padding = plain[-1]
    if padding < 1 or padding > 32:
        raise ValueError("invalid PKCS7 padding")
    plain = plain[:-padding]

    if len(plain) < 20:
        raise ValueError("decrypted payload is too short")
    message_len = struct.unpack("!I", plain[16:20])[0]
    message = plain[20:20 + message_len]
    receive_id = plain[20 + message_len:].decode("utf-8", errors="ignore")

    expected_corp_id = config["corp_id"]
    if expected_corp_id and receive_id and receive_id != expected_corp_id:
        raise ValueError("receive id does not match QYWX_CORP_ID")
    return message.decode("utf-8", errors="replace"), receive_id


def verify_and_decrypt(query, encrypted):
    config = get_config()
    timestamp = first_value(query, "timestamp")
    nonce = first_value(query, "nonce")
    msg_signature = first_value(query, "msg_signature")

    if not timestamp or not nonce or not msg_signature:
        raise ValueError("missing msg_signature, timestamp, or nonce")
    expected = signature(config["token"], timestamp, nonce, encrypted)
    if expected != msg_signature:
        raise ValueError("invalid msg_signature")
    return decrypt_wecom_message(encrypted)


def xml_encrypt_value(body):
    root = ET.fromstring(body)
    encrypt = root.findtext("Encrypt")
    if not encrypt:
        raise ValueError("missing Encrypt node")
    return encrypt


class CallbackHandler(BaseHTTPRequestHandler):
    server_version = "QyvxCallback/1.0"

    def log_message(self, fmt, *args):
        LOGGER.info("%s - %s", self.client_address[0], fmt % args)

    def send_text(self, status, body, content_type="text/plain; charset=utf-8"):
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, status, payload):
        self.send_text(
            status,
            json.dumps(payload, ensure_ascii=False),
            "application/json; charset=utf-8",
        )

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/healthz":
            config = get_config()
            self.send_json(200, {
                "ok": True,
                "config_ready": bool(config["token"] and config["aes_key"]),
                "service": "qyvx_callback",
            })
            return

        if parsed.path.rstrip("/") != "/qyvx/callback":
            self.send_text(404, "not found")
            return

        query = parse_qs(parsed.query)
        echostr = first_value(query, "echostr")
        if not echostr:
            self.send_text(200, "ok")
            return

        try:
            decrypted, receive_id = verify_and_decrypt(query, echostr)
            LOGGER.info("verified callback url receive_id=%s", receive_id)
            self.send_text(200, decrypted)
        except Exception as exc:
            LOGGER.exception("verify failed: %s", exc)
            self.send_text(400, "verify failed")

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path.rstrip("/") != "/qyvx/callback":
            self.send_text(404, "not found")
            return

        body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
        body_text = body.decode("utf-8", errors="replace")
        query = parse_qs(parsed.query)

        try:
            if "<Encrypt>" in body_text:
                encrypted = xml_encrypt_value(body_text)
                decrypted, receive_id = verify_and_decrypt(query, encrypted)
                LOGGER.info("received encrypted callback receive_id=%s body=%s", receive_id, decrypted)
            else:
                LOGGER.info("received plain callback body=%s", body_text)
            self.send_text(200, "success")
        except Exception as exc:
            LOGGER.exception("callback failed: %s", exc)
            self.send_text(400, "callback failed")


def main():
    config = get_config()
    server = ThreadingHTTPServer((config["host"], config["port"]), CallbackHandler)
    LOGGER.info("starting qyvx callback service on %s:%s", config["host"], config["port"])
    server.serve_forever()


if __name__ == "__main__":
    main()
