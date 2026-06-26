"""
下载全站点汇总报告 — Tkinter GUI 入口。

从界面收集 config（含可选站点），后台线程调用 main.TestPage.run()，
stdout/stderr 重定向到日志区。
"""

# -*- coding: utf-8 -*-
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from main import TestPage


SITES = ("United States", "Canada", "Mexico", "Brazil")  # 可选站点列表


class QueueWriter:
    """将 print 输出写入队列，供主线程刷新到日志文本框。"""

    def __init__(self, log_queue):
        self.log_queue = log_queue

    def write(self, text):
        if text:
            self.log_queue.put(text)

    def flush(self):
        pass


class AutomationGui:
    """主窗口：易得客配置、站点多选、开始运行、日志展示。"""

    def __init__(self, root):
        self.root = root
        self.root.title("下载全站点汇总报告")
        self.root.geometry("760x620")
        self.root.minsize(680, 560)

        self.log_queue = queue.Queue()
        self.worker = None
        self.site_vars = {}

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        """构建配置表单（含站点勾选）、操作按钮与日志区。"""
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(container, text="运行配置", padding=12)
        form.pack(fill=tk.X)

        ttk.Label(form, text="账号").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.username_var = tk.StringVar(value="13281439638")
        ttk.Entry(form, textvariable=self.username_var).grid(row=0, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="密码").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.password_var = tk.StringVar(value="13281439638@MM")
        ttk.Entry(form, textvariable=self.password_var, show="*").grid(row=1, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="店铺 IP").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.ip_var = tk.StringVar(value="54.70.92.80")
        ttk.Entry(form, textvariable=self.ip_var).grid(row=2, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="端口").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.port_var = tk.StringVar(value="8888")
        ttk.Entry(form, textvariable=self.port_var).grid(row=3, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="站点").grid(row=4, column=0, sticky=tk.NW, padx=(0, 8), pady=6)
        site_frame = ttk.Frame(form)
        site_frame.grid(row=4, column=1, sticky=tk.W, pady=6)
        for index, site in enumerate(SITES):
            var = tk.BooleanVar(value=True)
            self.site_vars[site] = var
            ttk.Checkbutton(site_frame, text=site, variable=var).grid(
                row=index // 2,
                column=index % 2,
                sticky=tk.W,
                padx=(0, 24),
                pady=2,
            )

        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X, pady=(12, 8))

        self.start_button = ttk.Button(actions, text="开始运行", command=self.start)
        self.start_button.pack(side=tk.LEFT)

        ttk.Button(actions, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=(8, 0))

        self.status_var = tk.StringVar(value="待运行")
        ttk.Label(actions, textvariable=self.status_var).pack(side=tk.RIGHT)

        log_box = ttk.LabelFrame(container, text="运行日志", padding=8)
        log_box.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_box, height=18, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _parse_list(self, value):
        """将逗号/换行分隔的文本拆成列表。"""
        items = []
        for item in value.replace("\n", ",").split(","):
            item = item.strip()
            if item:
                items.append(item)
        return items

    def _build_config(self):
        """校验界面输入，组装 main.TestPage 所需的 config 字典。"""
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        ips = self._parse_list(self.ip_var.get())
        port_values = self._parse_list(self.port_var.get())
        sites = [site for site, var in self.site_vars.items() if var.get()]

        if not username:
            raise ValueError("账号不能为空")
        if not password:
            raise ValueError("密码不能为空")
        if not ips:
            raise ValueError("至少填写一个店铺 IP")
        if not port_values:
            raise ValueError("至少填写一个端口")
        if not sites:
            raise ValueError("至少选择一个站点")

        ports = [int(port) for port in port_values]
        # 只填一个端口时，复用到全部 IP
        if len(ports) == 1 and len(ips) > 1:
            ports = ports * len(ips)
        if len(ports) != len(ips):
            raise ValueError("端口数量需要和 IP 数量一致，或只填写一个端口")

        return {
            "username": username,
            "password": password,
            "ip": ips,
            "port": ports,
            "data": sites,
        }

    def start(self):
        """校验配置后在后台线程启动 TestPage.run()。"""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "任务正在运行中")
            return

        try:
            config = self._build_config()
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.start_button.config(state=tk.DISABLED)
        self.status_var.set("运行中")
        self._append_log("开始运行...\n")

        self.worker = threading.Thread(target=self._run_task, args=(config,), daemon=True)
        self.worker.start()

    def _run_task(self, config):
        """工作线程：重定向 stdout/stderr 并执行完整流程。"""
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        writer = QueueWriter(self.log_queue)
        sys.stdout = writer
        sys.stderr = writer

        try:
            TestPage(config).run()
            self.log_queue.put("\n任务完成。\n")
            self.root.after(0, lambda: self.status_var.set("已完成"))
        except Exception:
            self.log_queue.put("\n运行失败：\n")
            self.log_queue.put(traceback.format_exc())
            self.root.after(0, lambda: self.status_var.set("运行失败"))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))

    def _poll_log(self):
        """定时从队列取日志写入文本框。"""
        while True:
            try:
                text = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(text)
        self.root.after(100, self._poll_log)

    def _append_log(self, text):
        """追加日志并滚动到底部。"""
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)

    def clear_log(self):
        """清空日志文本框。"""
        self.log_text.delete("1.0", tk.END)


def main():
    """启动 Tk 应用。"""
    root = tk.Tk()
    AutomationGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
