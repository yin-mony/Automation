import json
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from DrissionPage import ChromiumPage

from main import MODE_ONE, MODE_TWO, SaihuERP


MODE_ONE_TITLE = "纯新品列表（Excel文件）创建商品并在线配对"
MODE_TWO_TITLE = "低价商城列表（Excel文件）直接在商品列表创建SKU创建商品并在线配对"


# 赛狐 ERP Tkinter 统一入口
class SaihuERPRun:
    def __init__(self, config):
        self.config = config
        self.base_dir = Path(config.get("base_dir") or Path(__file__).resolve().parent)
        self.config_path = Path(config.get("config_path") or self.base_dir / "onlyrun_config.json")
        self.saved_config = self.load_config()

        self.root = tk.Tk()
        self.root.title("赛狐ERP 运行面板（run.py）")
        self.root.geometry("860x620")

        self.current_mode = MODE_ONE
        self.paths_by_mode = {
            MODE_ONE: str(self.saved_config.get("last_excel_dew") or ""),
            MODE_TWO: str(self.saved_config.get("last_excel_low") or ""),
        }
        self.event_queue = queue.Queue()
        self.worker_log_buffer = ""
        self.worker = None

        self.username_var = tk.StringVar(value=str(self.saved_config.get("last_username") or os.getenv("SAIHU_USERNAME", "")))
        self.password_var = tk.StringVar(value=str(self.saved_config.get("last_password") or os.getenv("SAIHU_PASSWORD", "")))
        self.mode_var = tk.StringVar(value=MODE_ONE)
        self.path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="运行状态：就绪")
        self.mode_label_var = tk.StringVar(value=f"当前模式：{MODE_ONE_TITLE}")

        self.build_ui()
        self.switch_mode(MODE_ONE)
        self.poll_events()

    def write(self, text):
        if not text:
            return
        self.worker_log_buffer += str(text)
        while "\n" in self.worker_log_buffer:
            line, self.worker_log_buffer = self.worker_log_buffer.split("\n", 1)
            line = line.rstrip()
            if line:
                self.event_queue.put(("log", line))

    def flush(self):
        line = self.worker_log_buffer.rstrip()
        if line:
            self.event_queue.put(("log", line))
        self.worker_log_buffer = ""

    def load_config(self):
        if not self.config_path.exists():
            return {}
        try:
            return json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"读取配置失败，将使用默认配置: {exc}")
            return {}

    def save_config(self):
        data = {
            "last_username": self.username_var.get().strip(),
            "last_password": self.password_var.get(),
            "last_excel_dew": self.paths_by_mode.get(MODE_ONE, ""),
            "last_excel_low": self.paths_by_mode.get(MODE_TWO, ""),
        }
        self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def build_ui(self):
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
            command=lambda: self.switch_mode(MODE_ONE),
        )
        self.mode_two_btn = ttk.Radiobutton(
            mode_row,
            text="模式二",
            variable=self.mode_var,
            value=MODE_TWO,
            command=lambda: self.switch_mode(MODE_TWO),
        )
        self.mode_one_btn.pack(side="left", padx=(8, 0))
        self.mode_two_btn.pack(side="left", padx=(8, 0))

        ttk.Label(container, textvariable=self.mode_label_var).pack(fill="x", pady=(0, 8))

        path_row = ttk.Frame(container)
        path_row.pack(fill="x", pady=(0, 8))
        ttk.Label(path_row, text="模式文件：").pack(side="left")
        self.path_entry = ttk.Entry(path_row, textvariable=self.path_var)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(8, 8))
        self.upload_btn = ttk.Button(path_row, text="上传当前模式文件", command=self.choose_file_for_current_mode)
        self.upload_btn.pack(side="left", padx=(0, 8))
        self.run_btn = ttk.Button(path_row, text="运行当前模式", command=self.run_mode)
        self.run_btn.pack(side="left")

        ttk.Label(container, text="运行日志：").pack(anchor="w")
        log_frame = ttk.Frame(container)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    def switch_mode(self, mode):
        self.current_mode = mode
        if mode == MODE_ONE:
            self.mode_label_var.set(f"当前模式：{MODE_ONE_TITLE}")
        else:
            self.mode_label_var.set(f"当前模式：{MODE_TWO_TITLE}")
        self.path_var.set(self.paths_by_mode.get(mode, ""))

    def choose_file_for_current_mode(self):
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
        self.save_config()

    def append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        text = message.strip()
        if text:
            self.status_var.set(f"运行状态：{text}")

    def set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.upload_btn.config(state=state)
        self.run_btn.config(state=state)
        self.mode_one_btn.config(state=state)
        self.mode_two_btn.config(state=state)

    def run_mode(self):
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
        self.save_config()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.append_log("------------------------------------------------------------")
        self.append_log("开始执行任务...")
        self.status_var.set("运行状态：任务启动中...")
        self.set_controls_enabled(False)

        self.worker = threading.Thread(
            target=self.run_worker,
            args=(self.current_mode, excel_path, username, password),
            daemon=True,
        )
        self.worker.start()

    def run_worker(self, mode, excel_path, username, password):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            mode_text = MODE_ONE_TITLE if mode == MODE_ONE else MODE_TWO_TITLE
            self.event_queue.put(("log", f"[启动] {mode_text}"))
            self.event_queue.put(("log", f"[参数] 账号: {username}"))
            self.event_queue.put(("log", f"[参数] 文件: {excel_path}"))

            config = {
                "page": ChromiumPage(),
                "mode": mode,
                "username": username,
                "password": password,
                "excel_path": excel_path,
                "base_dir": self.base_dir,
            }
            run = SaihuERP(config)
            run.main()

            self.flush()
            self.event_queue.put(("done", True, f"{mode_text} 运行完成"))
        except Exception as exc:
            traceback.print_exc()
            self.flush()
            self.event_queue.put(("done", False, str(exc)))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def on_done(self, success, message):
        self.set_controls_enabled(True)
        self.append_log("------------------------------------------------------------")
        self.append_log(message)
        self.status_var.set(f"运行状态：{'执行完成' if success else '执行失败'}")
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showwarning("失败", message)

    def poll_events(self):
        try:
            while True:
                event = self.event_queue.get_nowait()
                if not event:
                    continue
                if event[0] == "log":
                    self.append_log(event[1])
                elif event[0] == "done":
                    self.on_done(event[1], event[2])
        except queue.Empty:
            pass
        self.root.after(100, self.poll_events)

    def main(self):
        self.root.mainloop()


if __name__ == "__main__":
    config = {
        "base_dir": Path(__file__).resolve().parent,
        "config_path": Path(__file__).resolve().parent / "onlyrun_config.json",
    }
    run = SaihuERPRun(config)
    run.main()
