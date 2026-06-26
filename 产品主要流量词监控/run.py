"""
产品主要流量词监控 — Tkinter GUI 入口。

从界面收集 config，后台线程调用 main.Comment.run()，
stdout/stderr 重定向到日志区实时显示。
"""

import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from main import Comment


# 界面启动时的默认填表值（实际运行以界面输入为准）
DEFAULT_CONFIG = {
    "username": "13778451825",
    "password": "wjh12345.",
    "number": "18280194086",
    "ip": "35.84.243.7",
    "port": "9121",
    "asin": "B0963P4V3B,B09YVFYTGX",
    "file_path": r"C:\Users\admin\Desktop\产品流量词监控",
    "wechat_webhook": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=c4e9f21f-c365-496f-8c05-c56dd358926e",
}


class QueueWriter:
    """将 print 输出写入队列，供主线程刷新到日志文本框。"""

    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, text):
        if text:
            self.log_queue.put(text)

    def flush(self):
        pass


class App(tk.Tk):
    """主窗口：配置表单、开始运行、日志展示。"""

    def __init__(self):
        super().__init__()
        self.title("产品主要流量词监控")
        self.geometry("820x640")
        self.minsize(760, 560)

        self.log_queue = queue.Queue()
        self.worker = None
        self.vars = {key: tk.StringVar(value=value) for key, value in DEFAULT_CONFIG.items()}

        self._build_ui()
        self.after(100, self._flush_logs)

    def _build_ui(self):
        """构建配置表单、操作按钮与日志区。"""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        form = ttk.Frame(self, padding=14)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        self._add_entry(form, "易得客账号", "username", 0, 0)
        self._add_entry(form, "易得客密码", "password", 0, 2, show="*")
        self._add_entry(form, "企业微信手机号", "number", 1, 0)
        self._add_entry(form, "店铺 IP", "ip", 1, 2)
        self._add_entry(form, "端口", "port", 2, 0)
        self._add_entry(form, "ASIN", "asin", 2, 2)

        ttk.Label(form, text="保存目录").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
        path_frame = ttk.Frame(form)
        path_frame.grid(row=3, column=1, columnspan=3, sticky="ew", pady=6)
        path_frame.columnconfigure(0, weight=1)
        ttk.Entry(path_frame, textvariable=self.vars["file_path"]).grid(row=0, column=0, sticky="ew")
        ttk.Button(path_frame, text="选择", command=self._choose_folder).grid(row=0, column=1, padx=(8, 0))

        self._add_entry(form, "Webhook", "wechat_webhook", 4, 0, columnspan=3)

        buttons = ttk.Frame(form)
        buttons.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        buttons.columnconfigure(0, weight=1)
        self.start_button = ttk.Button(buttons, text="开始运行", command=self._start)
        self.start_button.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(buttons, text="清空日志", command=self._clear_logs).grid(row=0, column=2)

        log_frame = ttk.Frame(self, padding=(14, 0, 14, 14))
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _add_entry(self, parent, label, key, row, column, show=None, columnspan=1):
        """在表单中增加标签 + 输入框一行。"""
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=6)
        entry = ttk.Entry(parent, textvariable=self.vars[key], show=show)
        entry.grid(row=row, column=column + 1, columnspan=columnspan, sticky="ew", pady=6)
        return entry

    def _choose_folder(self):
        """选择西柚导出文件保存目录。"""
        folder = filedialog.askdirectory(initialdir=self.vars["file_path"].get() or str(Path.home()))
        if folder:
            self.vars["file_path"].set(folder)

    def _clear_logs(self):
        """清空日志文本框。"""
        self.log_text.delete("1.0", tk.END)

    def _build_config(self):
        """校验界面输入，组装 main.Comment 所需的 config 字典。"""
        ips = self._split_values(self.vars["ip"].get())
        asins = self._split_values(self.vars["asin"].get())
        ports = []

        for value in self._split_values(self.vars["port"].get()):
            try:
                ports.append(int(value))
            except ValueError as exc:
                raise ValueError(f"端口必须是数字：{value}") from exc

        if not ips:
            raise ValueError("请填写店铺 IP")
        if not ports:
            raise ValueError("请填写端口")
        if len(ips) != len(ports):
            raise ValueError("店铺 IP 和端口数量必须一一对应")
        if not asins:
            raise ValueError("请填写 ASIN")

        config = {
            "username": self.vars["username"].get().strip(),
            "password": self.vars["password"].get(),
            "number": self.vars["number"].get().strip(),
            "ip": ips,
            "port": ports,
            "asin": asins,
            "file_path": self.vars["file_path"].get().strip(),
        }

        webhook = self.vars["wechat_webhook"].get().strip()
        if webhook:
            config["wechat_webhook"] = webhook

        return config

    @staticmethod
    def _split_values(text):
        """将逗号/换行分隔的文本拆成列表（支持中英文逗号）。"""
        return [item.strip() for item in text.replace("\n", ",").replace("，", ",").split(",") if item.strip()]

    def _start(self):
        """校验配置后在后台线程启动 Comment.run()。"""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", "任务正在运行中，请等待完成。")
            return

        try:
            config = self._build_config()
        except ValueError as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        Path(config["file_path"]).mkdir(parents=True, exist_ok=True)
        self.start_button.configure(state="disabled")
        self._log("开始运行自动化流程...\n")

        self.worker = threading.Thread(target=self._run_comment, args=(config,), daemon=True)
        self.worker.start()

    def _run_comment(self, config):
        """工作线程：重定向 stdout/stderr 并执行完整采集流程。"""
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.log_queue)
        sys.stdout = writer
        sys.stderr = writer

        try:
            Comment(config).run()
            self.log_queue.put("\n任务运行完成。\n")
        except Exception:
            self.log_queue.put("\n任务运行出错：\n")
            self.log_queue.put(traceback.format_exc())
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.log_queue.put(("__DONE__", None))  # 通知主线程恢复「开始运行」按钮

    def _flush_logs(self):
        """定时从队列取日志写入文本框；收到 __DONE__ 时恢复按钮。"""
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    self.start_button.configure(state="normal")
                    continue
                self._log(item)
        except queue.Empty:
            pass

        self.after(100, self._flush_logs)

    def _log(self, text):
        """追加一行日志并滚动到底部。"""
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)


if __name__ == "__main__":
    App().mainloop()
