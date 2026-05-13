"""
月度销售额统计 - 项目主入口（GUI）
单页整合：合并汇总表路径 + 易得客下载配置；导出目录与合并读取目录为同一路径。
MergeData 参数顺序与 analysis 一致：(导出文件夹, 汇总表 xlsx)。

易得客下载：main.Automation(config).Start()（不修改 main.py）。
默认值常量 MAIN_DEFAULT_* 与 main.py 中 __main__ 内 config 保持一致。

打包为 exe：在项目目录执行 build_exe.bat（需已安装 requirements.txt 依赖）。
"""

import logging
import queue
import re
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from analysis import ExcelUtil

# 与 main.py 中 if __name__ == "__main__" 内 config 默认值保持一致（未改动 main.py，此处手工同步）
# 请勿在仓库中提交真实账号；本地可在 GUI 中填写或写入 monthly_sales_gui_config.txt（该文件已 .gitignore）
MAIN_DEFAULT_USERNAME = ""
MAIN_DEFAULT_PASSWORD = ""
MAIN_DEFAULT_FILE_PATH = ""


def get_app_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_log_file_path():
    log_dir = get_app_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "月度销售额统计.log"


class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


def _parse_list_csv(text):
    return [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]


def _parse_experts_block(text):
    out = []
    for line in text.splitlines():
        for p in line.split(","):
            s = p.strip()
            if s:
                out.append(s)
    return out


# IPv4：四段十进制，段内仅允许英文句点 U+002E 作为分隔（不允许全角句号等）
_IPV4_OCTET = r"(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)"
IPV4_PATTERN = re.compile(rf"^{_IPV4_OCTET}(\.{_IPV4_OCTET}){{3}}$")


def is_valid_ipv4(token: str) -> bool:
    s = (token or "").strip()
    return bool(IPV4_PATTERN.fullmatch(s))


def parse_shop_ipv4_list(text: str):
    """英文逗号分隔多个 IPv4；单项须为合法 IPv4。"""
    parts = [x.strip() for x in text.replace("\n", ",").split(",") if x.strip()]
    return parts


_IP_FIELD_CHARS = re.compile(r"^[\d., ]*$")


def validate_shop_ip_field_typing(proposed: str) -> bool:
    """仅允许数字、英文逗号、英文句号、空格（禁止全角标点与非 ASCII）。"""
    if proposed == "":
        return True
    if any(ord(c) > 127 for c in proposed):
        return False
    return _IP_FIELD_CHARS.fullmatch(proposed) is not None


def normalize_shop_ip_punctuation(s: str) -> str:
    """将常见全角句点类字符规范为英文句点，便于用户误输入后修正。"""
    for full in ("\uff0e", "\u3002", "\uff61"):
        s = s.replace(full, ".")
    return s


class MonthlySalesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("月度销售额统计工具")
        self.root.geometry("920x680")
        self.root.minsize(720, 560)

        self.is_running = False
        self.current_thread = None
        self.log_queue = queue.Queue()

        self.summary_path = tk.StringVar(value="")
        # 易得客下载目录与「合并」读取的数据文件夹为同一路径（默认与 main.config["file_path"] 一致）
        self.export_dir = tk.StringVar(value=MAIN_DEFAULT_FILE_PATH)

        # main.Automation 所需配置（与 main.py 中 config 键一致）
        self.auto_username = tk.StringVar(value=MAIN_DEFAULT_USERNAME)
        self.auto_password = tk.StringVar(value=MAIN_DEFAULT_PASSWORD)
        self.auto_ip = tk.StringVar(value="")
        self.auto_port = tk.StringVar(value="")
        default_experts = ""

        style = ttk.Style()
        style.theme_use("clam")

        self.setup_logging()

        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        tab = ttk.Frame(outer, padding=8)
        tab.pack(fill="both", expand=True)

        row = 0
        ttk.Label(tab, text="汇总表 Excel:").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(tab, textvariable=self.summary_path, width=72).grid(
            row=row, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(tab, text="浏览", width=10, command=self.select_summary).grid(
            row=row, column=2, sticky="e", pady=6
        )
        row += 1

        ttk.Label(tab, text="下载保存目录:").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(tab, textvariable=self.export_dir, width=62).grid(
            row=row, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(tab, text="浏览", width=10, command=self.select_export_dir).grid(
            row=row, column=2, sticky="e", pady=6
        )
        row += 1

        ttk.Label(tab, text="数据文件夹（只读）:").grid(row=row, column=0, sticky="w", pady=6)
        mirror = ttk.Label(tab, textvariable=self.export_dir, anchor="w", relief="groove")
        mirror.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        ttk.Label(tab, text="与下载目录一致", foreground="gray").grid(
            row=row, column=2, sticky="e", pady=6
        )
        row += 1

        ttk.Separator(tab).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1

        ttk.Label(tab, text="易得客账号:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(tab, textvariable=self.auto_username, width=50).grid(
            row=row, column=1, columnspan=2, sticky="w", padx=8, pady=4
        )
        row += 1
        ttk.Label(tab, text="易得客密码:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(tab, textvariable=self.auto_password, width=50, show="*").grid(
            row=row, column=1, columnspan=2, sticky="w", padx=8, pady=4
        )
        row += 1
        ttk.Label(tab, text="店铺 IP（英文逗号分隔多个；段内仅英文句号 .）:").grid(
            row=row, column=0, sticky="w", pady=4
        )
        self.auto_ip_entry = ttk.Entry(tab, textvariable=self.auto_ip, width=72)
        self.auto_ip_entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)
        _vcmd = (self.root.register(validate_shop_ip_field_typing), "%P")
        self.auto_ip_entry.configure(validate="key", validatecommand=_vcmd)
        self.auto_ip_entry.bind("<FocusOut>", self._shop_ip_focus_out)
        self.auto_ip_entry.bind("<<Paste>>", self._shop_ip_paste)
        row += 1
        ttk.Label(tab, text="调试端口（与 IP 个数一致）:").grid(row=row, column=0, sticky="nw", pady=4)
        ttk.Entry(tab, textvariable=self.auto_port, width=72).grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
        )
        row += 1
        ttk.Label(tab, text="达人列表（逗号或换行）:").grid(row=row, column=0, sticky="nw", pady=4)
        experts_frame = ttk.Frame(tab)
        experts_frame.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=8, pady=4)
        self.auto_experts_text = tk.Text(
            experts_frame, height=8, width=70, font=("Consolas", 9), wrap=tk.WORD
        )
        exp_scroll = ttk.Scrollbar(
            experts_frame, orient=tk.VERTICAL, command=self.auto_experts_text.yview
        )
        self.auto_experts_text.configure(yscrollcommand=exp_scroll.set)
        self.auto_experts_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        exp_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.auto_experts_text.insert("1.0", default_experts.replace(",", ",\n"))
        row += 1

        btn_row = ttk.Frame(tab)
        btn_row.grid(row=row, column=0, columnspan=3, pady=12)
        self.run_btn = ttk.Button(
            btn_row, text="开始合并数据", command=self.run_merge, width=18
        )
        self.run_btn.pack(side=tk.LEFT, padx=6)
        self.download_btn = ttk.Button(
            btn_row, text="一键下载 (Automation.Start)", command=self.run_download, width=26
        )
        self.download_btn.pack(side=tk.LEFT, padx=6)
        row += 1

        ttk.Label(
            tab,
            text="说明：下载目录与合并用的「数据文件夹」为同一路径；修改下载目录后只读行同步显示。"
            "两任务请勿并行。易得客默认账号密码与下载目录与 main.py 中 config 一致。",
            foreground="gray",
            wraplength=820,
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=6)

        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(8, weight=1)

        # ---------- 底部：状态 + 停止 + 日志 ----------
        ttk.Separator(outer).pack(fill="x", pady=8)

        ctrl = ttk.Frame(outer)
        ctrl.pack(fill="x")
        self.stop_btn = ttk.Button(
            ctrl, text="停止（仅标记）", command=self.stop_task, width=14, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 12))
        self.status_label = ttk.Label(ctrl, text="就绪", foreground="green")
        self.status_label.pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(8, 0))

        self.log_text = tk.Text(log_frame, height=12, wrap=tk.WORD, font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(
            log_frame, orient=tk.VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.load_config()
        self._shop_ip_focus_out()
        self.process_log_queue()
        logging.info("界面已就绪：单页内可合并汇总或易得客下载；导出目录已统一。")

    def _shop_ip_focus_out(self, _evt=None):
        self.auto_ip.set(normalize_shop_ip_punctuation(self.auto_ip.get()))

    def _shop_ip_paste(self, event):
        """粘贴时先将全角句点等转为英文句点，并过滤非法字符。"""
        try:
            clip = self.root.clipboard_get()
        except tk.TclError:
            return "break"
        norm = normalize_shop_ip_punctuation(clip)
        norm = "".join(c for c in norm if ord(c) < 128 and c in "0123456789., ")
        norm = norm.replace("\n", ",").replace("\r", "").strip()
        if not norm:
            return "break"
        w = event.widget
        try:
            w.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        w.insert("insert", norm)
        return "break"

    def select_summary(self):
        path = filedialog.askopenfilename(
            title="选择汇总表",
            filetypes=[("Excel", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.summary_path.set(path)

    def select_export_dir(self):
        path = filedialog.askdirectory(title="选择下载/导出数据目录（合并与此目录同源）")
        if path:
            self.export_dir.set(path)

    def setup_logging(self):
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
        for h in root_logger.handlers[:]:
            root_logger.removeHandler(h)
        root_logger.addHandler(queue_handler)
        root_logger.addHandler(file_handler)
        logging.info("日志文件: %s", get_log_file_path())

        self._stdout = sys.stdout
        sys.stdout = self

    def write(self, text):
        if text and text.strip():
            logging.info(text.rstrip())

    def flush(self):
        pass

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._insert_log(msg + "\n")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_log_queue)

    def _insert_log(self, message):
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def _set_task_running(self, running, label_text, busy_color):
        self.is_running = running
        if running:
            self.run_btn.config(state=tk.DISABLED)
            self.download_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.status_label.config(text=label_text, foreground=busy_color)
        else:
            self.run_btn.config(state=tk.NORMAL)
            self.download_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.status_label.config(text="就绪", foreground="green")

    def save_config(self):
        cfg = get_app_base_dir() / "monthly_sales_gui_config.txt"
        try:
            with open(cfg, "w", encoding="utf-8") as f:
                f.write(f"summary_path={self.summary_path.get()}\n")
                f.write(f"export_dir={self.export_dir.get()}\n")
                f.write(f"auto_username={self.auto_username.get()}\n")
                f.write(f"auto_password={self.auto_password.get()}\n")
                f.write(f"auto_ip={self.auto_ip.get()}\n")
                f.write(f"auto_port={self.auto_port.get()}\n")
                experts = self.auto_experts_text.get("1.0", tk.END).strip()
                f.write(f"auto_experts={experts.replace(chr(10), '\\n')}\n")
            logging.info("已保存界面配置")
        except OSError as e:
            logging.warning("保存配置失败: %s", e)

    def load_config(self):
        cfg = get_app_base_dir() / "monthly_sales_gui_config.txt"
        default_summary = ""
        try:
            if cfg.exists():
                with open(cfg, "r", encoding="utf-8") as f:
                    content = f.read()
                self.summary_path.set(default_summary)
                self.export_dir.set(MAIN_DEFAULT_FILE_PATH)
                experts_text = None
                ed_export = None
                ed_auto = None
                ed_data = None
                for part in content.splitlines():
                    if part.startswith("summary_path="):
                        self.summary_path.set(part.split("=", 1)[1])
                    elif part.startswith("export_dir="):
                        ed_export = part.split("=", 1)[1]
                    elif part.startswith("auto_file_path="):
                        ed_auto = part.split("=", 1)[1]
                    elif part.startswith("data_path="):
                        ed_data = part.split("=", 1)[1]
                    elif part.startswith("auto_username="):
                        self.auto_username.set(part.split("=", 1)[1])
                    elif part.startswith("auto_password="):
                        self.auto_password.set(part.split("=", 1)[1])
                    elif part.startswith("auto_ip="):
                        self.auto_ip.set(part.split("=", 1)[1])
                    elif part.startswith("auto_port="):
                        self.auto_port.set(part.split("=", 1)[1])
                    elif part.startswith("auto_experts="):
                        raw = part.split("=", 1)[1]
                        experts_text = raw.replace("\\n", "\n")
                merged_export = ed_export if ed_export is not None else (
                    ed_auto if ed_auto is not None else ed_data
                )
                if merged_export is not None:
                    self.export_dir.set(merged_export)
                if experts_text is not None:
                    self.auto_experts_text.delete("1.0", tk.END)
                    self.auto_experts_text.insert("1.0", experts_text)
                logging.info("已加载本地配置")
            else:
                self.summary_path.set(default_summary)
                self.export_dir.set(MAIN_DEFAULT_FILE_PATH)
        except OSError as e:
            logging.warning("加载配置失败: %s", e)
            self.summary_path.set(default_summary)
            self.export_dir.set(MAIN_DEFAULT_FILE_PATH)

    def stop_task(self):
        self.is_running = False
        logging.info("已请求停止（下载/合并任务无法强制中断 main.Automation，仅作状态标记）。")
        self.status_label.config(text="已请求停止", foreground="orange")

    def run_merge(self):
        if self.is_running:
            messagebox.showwarning("提示", "任务正在运行中，请稍候。")
            return

        summary = self.summary_path.get().strip()
        folder = self.export_dir.get().strip()

        if not summary:
            messagebox.showwarning("提示", "请选择汇总表 Excel 文件。")
            return
        if not summary.lower().endswith((".xlsx", ".xls")):
            messagebox.showwarning("提示", "汇总表请选择 .xlsx 或 .xls 文件。")
            return
        if not Path(summary).is_file():
            messagebox.showwarning("提示", f"汇总表不存在：\n{summary}")
            return
        if not folder:
            messagebox.showwarning("提示", "请选择导出数据所在文件夹。")
            return
        if not Path(folder).is_dir():
            messagebox.showwarning("提示", f"数据文件夹不存在：\n{folder}")
            return

        self.save_config()
        self.log_text.delete("1.0", tk.END)
        self._set_task_running(True, "合并中...", "orange")

        def target():
            try:
                logging.info("=" * 50)
                logging.info("开始合并：文件夹 → 汇总表")
                logging.info("数据文件夹: %s", folder)
                logging.info("汇总表: %s", summary)
                logging.info("=" * 50)

                excel = ExcelUtil()
                excel.MergeData(folder, summary)

                if self.is_running:
                    logging.info("合并流程执行完毕。")
                    self.root.after(0, lambda: self.on_finish(True, "数据处理完成"))
                else:
                    self.root.after(0, lambda: self.on_finish(False, "任务已中止标记"))
            except Exception as e:
                logging.exception("合并出错: %s", e)
                self.root.after(0, lambda: self.on_finish(False, str(e)))

        self.current_thread = threading.Thread(target=target, daemon=True)
        self.current_thread.start()

    def run_download(self):
        if self.is_running:
            messagebox.showwarning("提示", "任务正在运行中，请稍候。")
            return

        username = self.auto_username.get().strip()
        password = self.auto_password.get()
        self._shop_ip_focus_out()
        ips = parse_shop_ipv4_list(self.auto_ip.get())
        ports_raw = _parse_list_csv(self.auto_port.get())
        file_path = self.export_dir.get().strip()
        experts = _parse_experts_block(self.auto_experts_text.get("1.0", tk.END))

        if not username or not password:
            messagebox.showwarning("提示", "请填写易得客账号与密码。")
            return
        if not ips:
            messagebox.showwarning("提示", "请填写至少一个店铺 IP。")
            return
        bad = [p for p in ips if not is_valid_ipv4(p)]
        if bad:
            messagebox.showwarning(
                "提示",
                "以下为无效 IPv4，请使用四段十进制数字且段之间为英文句点（半角 .）：\n"
                + "\n".join(bad[:8])
                + ("\n…" if len(bad) > 8 else ""),
            )
            return
        try:
            ports = [int(x) for x in ports_raw]
        except ValueError:
            messagebox.showwarning("提示", "端口必须为整数，多个时用英文逗号分隔。")
            return
        if len(ports) != len(ips):
            messagebox.showwarning(
                "提示", f"IP 数量 ({len(ips)}) 与端口数量 ({len(ports)}) 不一致。"
            )
            return
        if not file_path:
            messagebox.showwarning("提示", "请填写下载保存目录。")
            return
        try:
            Path(file_path).mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showwarning("提示", f"无法创建或使用下载目录：{e}")
            return
        if not experts:
            messagebox.showwarning("提示", "请填写至少一位达人账号。")
            return

        self.save_config()
        self.log_text.delete("1.0", tk.END)
        self._set_task_running(True, "易得客下载运行中...", "orange")

        config = {
            "username": username,
            "password": password,
            "ip": ips,
            "port": ports,
            "experts": experts,
            "file_path": file_path,
        }

        def target():
            try:
                logging.info("=" * 50)
                logging.info("一键下载：main.Automation(config).Start()")
                logging.info("IP: %s | 端口: %s", ips, ports)
                logging.info("下载目录: %s", file_path)
                logging.info("达人数: %d", len(experts))
                logging.info("=" * 50)

                logging.info("正在导入 main.Automation ...")
                from main import Automation

                automation = Automation(config)
                automation.Start()

                if self.is_running:
                    logging.info("Automation.Start() 已返回。")
                    self.root.after(0, lambda: self.on_finish(True, "易得客下载流程已结束"))
                else:
                    self.root.after(0, lambda: self.on_finish(False, "任务已中止标记"))
            except Exception as e:
                logging.exception("下载任务出错: %s", e)
                self.root.after(0, lambda: self.on_finish(False, str(e)))

        self.current_thread = threading.Thread(target=target, daemon=True)
        self.current_thread.start()

    def on_finish(self, success, message):
        self._set_task_running(False, "", "")
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("错误", message)


def main():
    root = tk.Tk()
    MonthlySalesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
