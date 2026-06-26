"""
亚马逊评论下载 - GUI 入口
调用 main.Comment(config).run()，不修改 main.py 中的 Comment 类。
"""

import logging
import queue
import re
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from main import Comment

# 与 main.py 中 if __name__ == "__main__" 内 config 默认值保持一致（未改动 main.py，此处手工同步）
DEFAULT_CONFIG = {
    "username": "19944318805",
    "password": "DY0924DY0924",
    "ip": ["35.82.248.104"],
    "port": [8945],
    "experts": ["B0963P4V3B", "B09YVFYTGX"],
    "file_path": r"C:\RPA流程\亚马逊评论分析\flie",
}


def _default_shop_ip_text():
    """将默认 config 中的 IP 列表格式化为界面初始文本。"""
    return ", ".join(DEFAULT_CONFIG["ip"])


def _default_shop_port_text():
    """将默认 config 中的端口列表格式化为界面初始文本。"""
    return ", ".join(str(p) for p in DEFAULT_CONFIG["port"])


def _default_experts_text():
    """将默认 ASIN 列表格式化为多行文本框初始内容。"""
    return ",\n".join(DEFAULT_CONFIG["experts"])


def get_app_base_dir():
    """打包为 exe 时取 exe 所在目录，否则取本脚本目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_log_file_path():
    """日志文件路径：{应用目录}/logs/亚马逊评论下载.log。"""
    log_dir = get_app_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "亚马逊评论下载.log"


class QueueHandler(logging.Handler):
    """将 logging 记录写入队列，供界面定时刷新显示。"""

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        """将单条日志记录放入队列。"""
        self.log_queue.put(self.format(record))


def _parse_list_csv(text):
    """逗号/换行分隔文本拆成列表。"""
    return [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]


def _parse_experts_block(text):
    """解析 ASIN 多行/逗号混排输入。"""
    out = []
    for line in text.splitlines():
        for part in line.split(","):
            s = part.strip()
            if s:
                out.append(s)
    return out


_IPV4_OCTET = r"(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)"
IPV4_PATTERN = re.compile(rf"^{_IPV4_OCTET}(\.{_IPV4_OCTET}){{3}}$")  # 单段 IPv4 校验


def is_valid_ipv4(token: str) -> bool:
    """校验是否为合法 IPv4 字符串。"""
    return bool(IPV4_PATTERN.fullmatch((token or "").strip()))


def parse_shop_ipv4_list(text: str):
    """从店铺 IP 输入框解析 IPv4 列表。"""
    return [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]

_IP_FIELD_CHARS = re.compile(r"^[\d., ]*$")  # IP 输入框允许的字符集


def validate_shop_ip_field_typing(proposed: str) -> bool:
    """IP 输入框按键校验：仅允许数字、点、逗号与空格。"""
    if proposed == "":
        return True
    if any(ord(c) > 127 for c in proposed):
        return False
    return _IP_FIELD_CHARS.fullmatch(proposed) is not None


def normalize_shop_ip_punctuation(s: str) -> str:
    """将全角句号等替换为英文点，便于解析 IP。"""
    for full in ("\uff0e", "\u3002", "\uff61"):
        s = s.replace(full, ".")
    return s


class CommentDownloadApp:
    """评论下载 GUI：配置表单、后台线程、日志与强制停止。"""

    def __init__(self, root):
        """初始化窗口、配置变量、日志与界面。"""
        self.root = root
        self.root.title("亚马逊评论下载工具")
        self.root.geometry("900x680")
        self.root.minsize(720, 560)

        self.is_running = False
        self.current_thread = None
        self.comment_instance = None
        self.log_queue = queue.Queue()

        self.username = tk.StringVar(value=DEFAULT_CONFIG["username"])
        self.password = tk.StringVar(value=DEFAULT_CONFIG["password"])
        self.shop_ip = tk.StringVar(value=_default_shop_ip_text())
        self.shop_port = tk.StringVar(value=_default_shop_port_text())
        self.file_path = tk.StringVar(value=DEFAULT_CONFIG["file_path"])

        style = ttk.Style()
        style.theme_use("clam")

        self.setup_logging()
        self._build_ui()

        self.load_config()
        self._shop_ip_focus_out()
        self.process_log_queue()
        logging.info("界面已就绪，填写配置后点击「开始下载评论」。")

    def _build_ui(self):
        """构建配置表单、操作按钮与日志区。"""
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        form = ttk.Frame(outer, padding=8)
        form.pack(fill="both", expand=True)

        row = 0
        ttk.Label(form, text="易得客账号:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.username, width=50).grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
        )
        row += 1

        ttk.Label(form, text="易得客密码:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.password, width=50, show="*").grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
        )
        row += 1

        ttk.Label(form, text="店铺 IP（英文逗号分隔）:").grid(
            row=row, column=0, sticky="w", pady=4
        )
        self.shop_ip_entry = ttk.Entry(form, textvariable=self.shop_ip, width=72)
        self.shop_ip_entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)
        vcmd = (self.root.register(validate_shop_ip_field_typing), "%P")
        self.shop_ip_entry.configure(validate="key", validatecommand=vcmd)
        self.shop_ip_entry.bind("<FocusOut>", self._shop_ip_focus_out)
        self.shop_ip_entry.bind("<<Paste>>", self._shop_ip_paste)
        row += 1

        ttk.Label(form, text="调试端口（与 IP 个数一致）:").grid(
            row=row, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self.shop_port, width=72).grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
        )
        row += 1

        ttk.Label(form, text="评论保存目录:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.file_path, width=62).grid(
            row=row, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Button(form, text="浏览", width=10, command=self.select_file_path).grid(
            row=row, column=2, sticky="e", pady=4
        )
        row += 1

        ttk.Label(form, text="商品 ASIN（逗号或换行）:").grid(
            row=row, column=0, sticky="nw", pady=4
        )
        asin_frame = ttk.Frame(form)
        asin_frame.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=8, pady=4)
        self.asin_text = tk.Text(
            asin_frame, height=8, width=70, font=("Consolas", 9), wrap=tk.WORD
        )
        asin_scroll = ttk.Scrollbar(
            asin_frame, orient=tk.VERTICAL, command=self.asin_text.yview
        )
        self.asin_text.configure(yscrollcommand=asin_scroll.set)
        self.asin_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        asin_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.asin_text.insert("1.0", _default_experts_text())
        row += 1

        btn_row = ttk.Frame(form)
        btn_row.grid(row=row, column=0, columnspan=3, pady=12)
        self.run_btn = ttk.Button(
            btn_row, text="开始下载评论", command=self.run_download, width=18
        )
        self.run_btn.pack(side=tk.LEFT, padx=6)
        self.stop_btn = ttk.Button(
            btn_row, text="强制停止", command=self.force_stop, width=12, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=6)
        row += 1

        ttk.Label(
            form,
            text="说明：将依次登录易得客、启动店铺浏览器并抓取各星级评论，导出为「亚马逊评论.xlsx」。",
            foreground="gray",
            wraplength=780,
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=6)

        form.columnconfigure(1, weight=1)
        form.rowconfigure(6, weight=1)

        ttk.Separator(outer).pack(fill="x", pady=8)

        ctrl = ttk.Frame(outer)
        ctrl.pack(fill="x")
        self.status_label = ttk.Label(ctrl, text="就绪", foreground="green")
        self.status_label.pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(8, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=14, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _shop_ip_focus_out(self, _evt=None):
        """失焦时规范化店铺 IP 标点。"""
        self.shop_ip.set(normalize_shop_ip_punctuation(self.shop_ip.get()))

    def _shop_ip_paste(self, event):
        """粘贴时过滤非法字符并规范化 IP 文本。"""
        try:
            clip = self.root.clipboard_get()
        except tk.TclError:
            return "break"
        norm = normalize_shop_ip_punctuation(clip)
        norm = "".join(c for c in norm if ord(c) < 128 and c in "0123456789., ")
        norm = norm.replace("\n", ",").replace("\r", "").strip()
        if not norm:
            return "break"
        widget = event.widget
        try:
            widget.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        widget.insert("insert", norm)
        return "break"

    def select_file_path(self):
        """选择评论 Excel 保存目录。"""
        path = filedialog.askdirectory(title="选择评论 Excel 保存目录")
        if path:
            self.file_path.set(path)

    def setup_logging(self):
        """配置队列日志、滚动文件日志，并将 stdout 重定向到 logging。"""
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setLevel(logging.INFO)
        queue_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt="%H:%M:%S")
        )
        file_handler = RotatingFileHandler(
            get_log_file_path(),
            maxBytes=2 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        )

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        root_logger.addHandler(queue_handler)
        root_logger.addHandler(file_handler)
        logging.info("日志文件: %s", get_log_file_path())

        self._stdout = sys.stdout
        sys.stdout = self

    def write(self, text):
        """兼容 sys.stdout 重定向：print 内容写入 logging。"""
        if text and text.strip():
            logging.info(text.rstrip())

    def flush(self):
        """stdout 重定向接口占位，无缓冲需刷新。"""
        pass

    def process_log_queue(self):
        """定时从队列取日志写入文本框。"""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._insert_log(msg + "\n")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_log_queue)

    def _insert_log(self, message):
        """追加日志并滚动到底部。"""
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def _set_running(self, running):
        """切换运行中 UI 状态（按钮、状态栏）。"""
        self.is_running = running
        if running:
            self.run_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.status_label.config(text="下载运行中...", foreground="orange")
        else:
            self.run_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.status_label.config(text="就绪", foreground="green")
            self.comment_instance = None

    def _build_config(self):
        """校验界面输入，组装 main.Comment 所需的 config 字典。"""
        self._shop_ip_focus_out()
        ips = parse_shop_ipv4_list(self.shop_ip.get())
        ports_raw = _parse_list_csv(self.shop_port.get())
        file_path = self.file_path.get().strip()
        experts = _parse_experts_block(self.asin_text.get("1.0", tk.END))
        username = self.username.get().strip()
        password = self.password.get()

        if not username or not password:
            raise ValueError("请填写易得客账号与密码。")
        if not ips:
            raise ValueError("请填写至少一个店铺 IP。")
        bad_ips = [ip for ip in ips if not is_valid_ipv4(ip)]
        if bad_ips:
            raise ValueError(
                "以下为无效 IPv4：\n" + "\n".join(bad_ips[:8]) + ("\n…" if len(bad_ips) > 8 else "")
            )
        try:
            ports = [int(x) for x in ports_raw]
        except ValueError as exc:
            raise ValueError("端口必须为整数，多个时用英文逗号分隔。") from exc
        if len(ports) != len(ips):
            raise ValueError(f"IP 数量 ({len(ips)}) 与端口数量 ({len(ports)}) 不一致。")
        if not file_path:
            raise ValueError("请选择评论保存目录。")
        Path(file_path).mkdir(parents=True, exist_ok=True)
        if not experts:
            raise ValueError("请填写至少一个商品 ASIN。")

        return {
            "username": username,
            "password": password,
            "ip": ips,
            "port": ports,
            "experts": experts,
            "file_path": file_path,
        }

    def save_config(self):
        """将当前界面配置写入 comment_download_gui_config.txt。"""
        cfg = get_app_base_dir() / "comment_download_gui_config.txt"
        try:
            experts = self.asin_text.get("1.0", tk.END).strip()
            with open(cfg, "w", encoding="utf-8") as f:
                f.write(f"username={self.username.get()}\n")
                f.write(f"password={self.password.get()}\n")
                f.write(f"shop_ip={self.shop_ip.get()}\n")
                f.write(f"shop_port={self.shop_port.get()}\n")
                f.write(f"file_path={self.file_path.get()}\n")
                f.write(f"experts={experts.replace(chr(10), '\\n')}\n")
            logging.info("已保存界面配置")
        except OSError as e:
            logging.warning("保存配置失败: %s", e)

    def load_config(self):
        """启动时从 comment_download_gui_config.txt 恢复上次配置。"""
        cfg = get_app_base_dir() / "comment_download_gui_config.txt"
        if not cfg.exists():
            logging.info("未找到本地配置，使用 main.py 默认 config 值。")
            return
        try:
            experts_text = None
            with open(cfg, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("username="):
                        self.username.set(line.split("=", 1)[1].strip())
                    elif line.startswith("password="):
                        self.password.set(line.split("=", 1)[1].strip())
                    elif line.startswith("shop_ip="):
                        self.shop_ip.set(line.split("=", 1)[1].strip())
                    elif line.startswith("shop_port="):
                        self.shop_port.set(line.split("=", 1)[1].strip())
                    elif line.startswith("file_path="):
                        self.file_path.set(line.split("=", 1)[1].strip())
                    elif line.startswith("experts="):
                        experts_text = line.split("=", 1)[1].strip().replace("\\n", "\n")
            if experts_text is not None:
                self.asin_text.delete("1.0", tk.END)
                self.asin_text.insert("1.0", experts_text)
            logging.info("已加载本地配置（覆盖默认值）")
        except OSError as e:
            logging.warning("加载配置失败，保留 main.py 默认值: %s", e)

    def run_download(self):
        """校验配置后在后台线程执行 Comment(config).run()。"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在运行中，请稍候。")
            return

        try:
            config = self._build_config()
        except ValueError as e:
            messagebox.showwarning("参数错误", str(e))
            return

        self.save_config()
        self.log_text.delete("1.0", tk.END)
        self._set_running(True)

        def target():
            """后台线程：实例化 Comment 并执行 run()。"""
            try:
                logging.info("=" * 50)
                logging.info("开始下载：Comment(config).run()")
                logging.info("店铺 IP: %s | 端口: %s", config["ip"], config["port"])
                logging.info("保存目录: %s", config["file_path"])
                logging.info("ASIN 数量: %d", len(config["experts"]))
                logging.info("=" * 50)

                comment = Comment(config)
                self.comment_instance = comment
                comment.run()

                if self.is_running:
                    logging.info("评论下载流程已结束。")
                    self.root.after(0, lambda: self.on_finish(True, "评论下载完成，已导出 Excel。"))
                else:
                    self.root.after(0, lambda: self.on_finish(False, "任务已中止"))
            except Exception as e:
                logging.exception("下载任务出错: %s", e)
                err_msg = str(e)
                self.root.after(0, lambda msg=err_msg: self.on_finish(False, msg))

        self.current_thread = threading.Thread(target=target, daemon=True)
        self.current_thread.start()

    def force_stop(self):
        """调用 Comment.stop_program() 终止 Chrome 并强制退出。"""
        if not self.is_running:
            return
        if not messagebox.askyesno(
            "确认停止",
            "将终止 Chrome 进程并强制退出程序，确定继续？",
        ):
            return
        self.is_running = False
        logging.warning("用户请求强制停止...")
        if self.comment_instance is not None:
            try:
                self.comment_instance.stop_program()
            except Exception as e:
                logging.exception("强制停止失败: %s", e)
        else:
            logging.warning("任务实例尚未创建，仅标记停止。")

    def on_finish(self, success, message):
        """任务结束：恢复 UI 并弹窗提示。"""
        self._set_running(False)
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("错误", message)


def main():
    """启动 Tk 应用。"""
    root = tk.Tk()
    CommentDownloadApp(root)
    root.mainloop()


# GUI 入口
if __name__ == "__main__":
    main()
