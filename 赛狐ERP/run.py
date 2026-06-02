import sys
import traceback
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from DrissionPage import ChromiumPage
from NewSet import NewSetPage
from LowPrice import LowPricePage
# import main as main_workflow

MODE_ONE = "mode_one"
MODE_TWO = "mode_two"
MODE_ONE_TITLE = "纯新品列表（Excel文件）创建商品并在线配对"
MODE_TWO_TITLE = "低价商城列表（Excel文件）直接在商品列表创建SKU创建商品并在线配对"


class _TkLogStream:
    def __init__(self, event_queue, prefix=""):
        self.event_queue = event_queue
        self.prefix = prefix
        self.buffer = ""

    def write(self, text):
        if not text:
            return
        self.buffer += str(text)
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            line = line.rstrip()
            if line:
                self.event_queue.put(("log", f"{self.prefix}{line}"))

    def flush(self):
        line = self.buffer.rstrip()
        if line:
            self.event_queue.put(("log", f"{self.prefix}{line}"))
        self.buffer = ""


class RunnerThread(threading.Thread):
    def __init__(self, mode, excel_path, username, password, event_queue):
        super().__init__(daemon=True)
        self.mode = mode
        self.excel_path = excel_path
        self.username = username
        self.password = password
        self.event_queue = event_queue

    def run(self):
        stdout_stream = _TkLogStream(self.event_queue)
        stderr_stream = _TkLogStream(self.event_queue, "[stderr] ")
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = stdout_stream
        sys.stderr = stderr_stream
        try:
            mode_text = MODE_ONE_TITLE if self.mode == MODE_ONE else MODE_TWO_TITLE
            self.event_queue.put(("log", f"[启动] {mode_text}"))
            self.event_queue.put(("log", f"[参数] 账号: {self.username}"))
            self.event_queue.put(("log", f"[参数] 文件: {self.excel_path}"))

            page = ChromiumPage()
            if self.mode == MODE_ONE:
                # main_workflow.new_set_pairing(username=self.username, password=self.password, path=self.excel_path)
                run = NewSetPage(
                    page=page,
                    username=self.username,
                    password=self.password,
                    excel_path=self.excel_path
                )
                run.main()
            else:
                # main_workflow.low_price_pairing(username=self.username, password=self.password, path=self.excel_path)
                run = LowPricePage(
                    page=page,
                    username=self.username,
                    password=self.password,
                    excel_path=self.excel_path
                )
                run.main()

            stdout_stream.flush()
            stderr_stream.flush()
            self.event_queue.put(("done", True, f"{mode_text} 运行完成"))
        except Exception as exc:
            traceback.print_exc()
            stdout_stream.flush()
            stderr_stream.flush()
            self.event_queue.put(("done", False, str(exc)))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class RunnerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("赛狐ERP 运行面板（run.py）")
        self.root.geometry("860x620")

        self.current_mode = MODE_ONE
        self.paths_by_mode = {MODE_ONE: "", MODE_TWO: ""}
        self.event_queue = queue.Queue()
        self.worker = None

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.mode_var = tk.StringVar(value=MODE_ONE)
        self.path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="运行状态：就绪")
        self.mode_label_var = tk.StringVar(value=f"当前模式：{MODE_ONE_TITLE}")

        self._build_ui()
        self._switch_mode(MODE_ONE)
        self._poll_events()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="赛狐账号：").pack(anchor="w")
        ttk.Entry(container, textvariable=self.username_var).pack(fill="x", pady=(0, 8))

        ttk.Label(container, text="赛狐密码：").pack(anchor="w")
        ttk.Entry(container, textvariable=self.password_var, show="*").pack(fill="x", pady=(0, 8))

        status_label = ttk.Label(container, textvariable=self.status_var, relief="solid", padding=6)
        status_label.pack(fill="x", pady=(0, 8))

        mode_row = ttk.Frame(container)
        mode_row.pack(fill="x", pady=(0, 6))
        ttk.Label(mode_row, text="模式选择：").pack(side="left")
        self.mode_one_btn = ttk.Radiobutton(
            mode_row,
            text="模式一",
            variable=self.mode_var,
            value=MODE_ONE,
            command=lambda: self._switch_mode(MODE_ONE),
        )
        self.mode_two_btn = ttk.Radiobutton(
            mode_row,
            text="模式二",
            variable=self.mode_var,
            value=MODE_TWO,
            command=lambda: self._switch_mode(MODE_TWO),
        )
        self.mode_one_btn.pack(side="left", padx=(8, 0))
        self.mode_two_btn.pack(side="left", padx=(8, 0))

        ttk.Label(container, textvariable=self.mode_label_var).pack(fill="x", pady=(0, 8))

        path_row = ttk.Frame(container)
        path_row.pack(fill="x", pady=(0, 8))
        ttk.Label(path_row, text="模式文件：").pack(side="left")
        self.path_entry = ttk.Entry(path_row, textvariable=self.path_var)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(8, 8))
        self.upload_btn = ttk.Button(path_row, text="上传文件", command=self._choose_file_for_current_mode)
        self.upload_btn.pack(side="left", padx=(0, 8))
        self.run_btn = ttk.Button(path_row, text="运行当前模式", command=self._run_mode)
        self.run_btn.pack(side="left")

        ttk.Label(container, text="运行日志：").pack(anchor="w")
        log_frame = ttk.Frame(container)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    def _switch_mode(self, mode):
        self.current_mode = mode
        if mode == MODE_ONE:
            self.mode_label_var.set(f"当前模式：{MODE_ONE_TITLE}")
            self.upload_btn.config(text="上传当前模式文件")
            self.run_btn.config(text="运行当前模式")
        else:
            self.mode_label_var.set(f"当前模式：{MODE_TWO_TITLE}")
            self.upload_btn.config(text="上传当前模式文件")
            self.run_btn.config(text="运行当前模式")
        self.path_var.set(self.paths_by_mode.get(mode, ""))

    def _choose_file_for_current_mode(self):
        title = f"选择文件：{MODE_ONE_TITLE}" if self.current_mode == MODE_ONE else f"选择文件：{MODE_TWO_TITLE}"
        file_path = filedialog.askopenfilename(
            title=title,
            initialdir=str(Path.cwd()),
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        self.paths_by_mode[self.current_mode] = file_path
        self.path_var.set(file_path)

    def _append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        text = message.strip()
        if text:
            self.status_var.set(f"运行状态：{text}")

    def _set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.upload_btn.config(state=state)
        self.run_btn.config(state=state)
        self.mode_one_btn.config(state=state)
        self.mode_two_btn.config(state=state)

    def _run_mode(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "已有任务在运行，请稍后再试。")
            return

        username = self.username_var.get().strip()
        password = self.password_var.get()
        excel_path = self.path_var.get().strip()

        if not username:
            messagebox.showwarning("参数错误", "请填写赛狐账号。")
            return
        if not password:
            messagebox.showwarning("参数错误", "请填写赛狐密码。")
            return
        if not excel_path:
            messagebox.showwarning("参数错误", "请先选择文件路径。")
            return
        if not Path(excel_path).exists():
            messagebox.showwarning("参数错误", f"文件不存在：{excel_path}")
            return

        self.paths_by_mode[self.current_mode] = excel_path
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._append_log("------------------------------------------------------------")
        self._append_log("开始执行任务...")
        self.status_var.set("运行状态：任务启动中...")
        self._set_controls_enabled(False)

        self.worker = RunnerThread(
            mode=self.current_mode,
            excel_path=excel_path,
            username=username,
            password=password,
            event_queue=self.event_queue,
        )
        self.worker.start()

    def _on_done(self, success, message):
        self._set_controls_enabled(True)
        self._append_log("------------------------------------------------------------")
        self._append_log(message)
        self.status_var.set(f"运行状态：{'执行完成' if success else '执行失败'}")
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showwarning("失败", message)

    def _poll_events(self):
        try:
            while True:
                event = self.event_queue.get_nowait()
                if not event:
                    continue
                if event[0] == "log":
                    self._append_log(event[1])
                elif event[0] == "done":
                    self._on_done(event[1], event[2])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def run(self):
        self.root.mainloop()


def main():
    app = RunnerApp()
    app.run()


if __name__ == "__main__":
    main()
