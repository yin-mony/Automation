import sys
import traceback
import queue
import threading
import tkinter as tk
from pathlib import Path
from typing import Optional
from tkinter import ttk, filedialog, messagebox

from main import DEFAULT_CONFIG, run_mode


class _QueueWriter:
    def __init__(self, log_queue: queue.Queue):
        self._log_queue = log_queue

    def write(self, text: str) -> None:
        if text:
            self._log_queue.put(text)

    def flush(self) -> None:
        pass


class SaihuERPApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("赛狐 ERP 自动化")
        self.root.geometry("720x520")
        self.root.minsize(640, 480)

        self._log_queue: queue.Queue = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._running = False

        self.mode_var = tk.StringVar(value=DEFAULT_CONFIG["mode"])
        self.username_var = tk.StringVar(value=DEFAULT_CONFIG["username"])
        self.password_var = tk.StringVar(value=DEFAULT_CONFIG["password"])
        self.excel_path_var = tk.StringVar(value=DEFAULT_CONFIG["excel_path"])

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        mode_frame = ttk.LabelFrame(main, text="运行模式", padding=10)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Radiobutton(
            mode_frame,
            text="模式一：纯新品",
            variable=self.mode_var,
            value="mode1",
        ).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(
            mode_frame,
            text="模式二：横向变体",
            variable=self.mode_var,
            value="mode2",
        ).pack(side=tk.LEFT)

        form = ttk.LabelFrame(main, text="配置", padding=10)
        form.pack(fill=tk.X, pady=(0, 10))
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="账号").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.username_var).grid(
            row=0, column=1, sticky=tk.EW, padx=(8, 0), pady=4
        )

        ttk.Label(form, text="密码").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.password_var, show="*").grid(
            row=1, column=1, sticky=tk.EW, padx=(8, 0), pady=4
        )

        ttk.Label(form, text="Excel 路径").grid(row=2, column=0, sticky=tk.W, pady=4)
        path_row = ttk.Frame(form)
        path_row.grid(row=2, column=1, sticky=tk.EW, padx=(8, 0), pady=4)
        path_row.columnconfigure(0, weight=1)
        ttk.Entry(path_row, textvariable=self.excel_path_var).grid(
            row=0, column=0, sticky=tk.EW
        )
        ttk.Button(path_row, text="浏览", command=self._browse_excel).grid(
            row=0, column=1, padx=(8, 0)
        )

        btn_row = ttk.Frame(main)
        btn_row.pack(fill=tk.X, pady=(0, 10))
        self.start_btn = ttk.Button(btn_row, text="开始运行", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT)
        ttk.Button(btn_row, text="清空日志", command=self._clear_log).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        log_frame = ttk.LabelFrame(main, text="运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 10),
        )
        scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        scroll.grid(row=0, column=1, sticky=tk.NS)

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.excel_path_var.set(path)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _poll_log_queue(self) -> None:
        try:
            while True:
                msg = self._log_queue.get_nowait()
                if msg == "__DONE__":
                    self._on_worker_finished(success=True)
                    continue
                if msg == "__ERROR__":
                    self._on_worker_finished(success=False)
                    continue
                self._append_log(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _set_running(self, running: bool) -> None:
        self._running = running
        self.start_btn.configure(state=tk.DISABLED if running else tk.NORMAL)

    def _on_worker_finished(self, success: bool) -> None:
        self._set_running(False)
        if success:
            self._append_log("\n任务执行完成。\n")
            messagebox.showinfo("完成", "任务执行完成。")
        else:
            self._append_log("\n任务执行失败，请查看日志。\n")
            messagebox.showerror("错误", "任务执行失败，请查看日志。")

    def _validate(self) -> Optional[dict]:
        username = self.username_var.get().strip()
        password = self.password_var.get()
        excel_path = self.excel_path_var.get().strip()

        if not username:
            messagebox.showwarning("提示", "请填写赛狐账号。")
            return None
        if not password:
            messagebox.showwarning("提示", "请填写赛狐密码。")
            return None
        if not excel_path:
            messagebox.showwarning("提示", "请选择 Excel 文件路径。")
            return None
        if not Path(excel_path).exists():
            messagebox.showwarning("提示", f"Excel 文件不存在：\n{excel_path}")
            return None

        return {
            "mode": self.mode_var.get(),
            "username": username,
            "password": password,
            "excel_path": excel_path,
        }

    def _run_worker(self, config: dict) -> None:
        old_stdout = sys.stdout
        sys.stdout = _QueueWriter(self._log_queue)
        try:
            run_mode(config)
            self._log_queue.put("__DONE__")
        except Exception:
            self._log_queue.put(traceback.format_exc())
            self._log_queue.put("__ERROR__")
        finally:
            sys.stdout = old_stdout

    def _on_start(self) -> None:
        if self._running:
            return

        config = self._validate()
        if not config:
            return

        mode_name = "纯新品" if config["mode"] == "mode1" else "横向变体"
        self._append_log(
            f"\n{'=' * 50}\n"
            f"开始运行：{mode_name}\n"
            f"账号：{config['username']}\n"
            f"Excel：{config['excel_path']}\n"
            f"{'=' * 50}\n"
        )
        self._set_running(True)

        self._worker = threading.Thread(
            target=self._run_worker,
            args=(config,),
            daemon=True,
        )
        self._worker.start()


def main() -> None:
    root = tk.Tk()
    SaihuERPApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
