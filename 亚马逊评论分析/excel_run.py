"""
亚马逊评论 AI 分析 - GUI 入口
调用 analysis.CommentAnalyzer(path, api_key).run()，不修改 analysis.py 中的 CommentAnalyzer 类。
"""

import logging
import queue
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from analysis import CommentAnalyzer

# 与 analysis.py 中 if __name__ == "__main__" 内默认值保持一致（未改动 analysis.py，此处手工同步）
DEFAULT_CONFIG = {
    "excel_path": r"C:\RPA流程\亚马逊评论分析\flie\亚马逊评论.xlsx",
    "api_key": "sk-c6110db8ead745e5bf1078a63c80a427",
}


def get_app_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_log_file_path():
    log_dir = get_app_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "亚马逊评论分析.log"


class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class CommentAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("亚马逊评论 AI 分析工具")
        self.root.geometry("860x620")
        self.root.minsize(680, 480)

        self.is_running = False
        self.current_thread = None
        self.log_queue = queue.Queue()

        self.excel_path = tk.StringVar(value=DEFAULT_CONFIG["excel_path"])
        self.api_key = tk.StringVar(value=DEFAULT_CONFIG["api_key"])

        style = ttk.Style()
        style.theme_use("clam")

        self.setup_logging()
        self._build_ui()

        self.load_config()
        self.process_log_queue()
        logging.info("界面已就绪，填写配置后点击「开始 AI 分析」。")

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        form = ttk.Frame(outer, padding=8)
        form.pack(fill="both", expand=True)

        row = 0
        ttk.Label(form, text="评论 Excel 文件:").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.excel_path, width=72).grid(
            row=row, column=1, sticky="ew", padx=8, pady=6
        )
        ttk.Button(form, text="浏览", width=10, command=self.select_excel).grid(
            row=row, column=2, sticky="e", pady=6
        )
        row += 1

        ttk.Label(form, text="DeepSeek API Key:").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.api_key, width=72, show="*").grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=6
        )
        row += 1

        btn_row = ttk.Frame(form)
        btn_row.grid(row=row, column=0, columnspan=3, pady=12)
        self.run_btn = ttk.Button(
            btn_row, text="开始 AI 分析", command=self.run_analysis, width=18
        )
        self.run_btn.pack(side=tk.LEFT, padx=6)
        self.stop_btn = ttk.Button(
            btn_row, text="停止（仅标记）", command=self.stop_task, width=14, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=6)
        row += 1

        ttk.Label(
            form,
            text="说明：读取评论 Excel，调用 DeepSeek 生成好评卖点、差评痛点与改进建议，"
            "报告保存在 Excel 同目录下的「分析报告」文件夹。",
            foreground="gray",
            wraplength=760,
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=6)

        form.columnconfigure(1, weight=1)

        ttk.Separator(outer).pack(fill="x", pady=8)

        ctrl = ttk.Frame(outer)
        ctrl.pack(fill="x")
        self.status_label = ttk.Label(ctrl, text="就绪", foreground="green")
        self.status_label.pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(8, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=16, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="选择评论 Excel 文件",
            filetypes=[("Excel", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.excel_path.set(path)

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
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
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

    def _set_running(self, running):
        self.is_running = running
        if running:
            self.run_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.status_label.config(text="AI 分析运行中...", foreground="orange")
        else:
            self.run_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.status_label.config(text="就绪", foreground="green")

    def _build_params(self):
        excel_path = self.excel_path.get().strip()
        api_key = self.api_key.get().strip()

        if not excel_path:
            raise ValueError("请选择评论 Excel 文件。")
        if not excel_path.lower().endswith((".xlsx", ".xls")):
            raise ValueError("请选择 .xlsx 或 .xls 格式的 Excel 文件。")
        if not Path(excel_path).is_file():
            raise ValueError(f"Excel 文件不存在：\n{excel_path}")
        if not api_key:
            raise ValueError("请填写 DeepSeek API Key。")

        return excel_path, api_key

    def save_config(self):
        cfg = get_app_base_dir() / "comment_analyzer_gui_config.txt"
        try:
            with open(cfg, "w", encoding="utf-8") as f:
                f.write(f"excel_path={self.excel_path.get()}\n")
                f.write(f"api_key={self.api_key.get()}\n")
            logging.info("已保存界面配置")
        except OSError as e:
            logging.warning("保存配置失败: %s", e)

    def load_config(self):
        cfg = get_app_base_dir() / "comment_analyzer_gui_config.txt"
        if not cfg.exists():
            logging.info("未找到本地配置，使用 analysis.py 默认值。")
            return
        try:
            with open(cfg, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("excel_path="):
                        self.excel_path.set(line.split("=", 1)[1].strip())
                    elif line.startswith("api_key="):
                        self.api_key.set(line.split("=", 1)[1].strip())
            logging.info("已加载本地配置（覆盖默认值）")
        except OSError as e:
            logging.warning("加载配置失败，保留 analysis.py 默认值: %s", e)

    def stop_task(self):
        if not self.is_running:
            return
        self.is_running = False
        logging.info("已请求停止（AI 分析任务无法强制中断，仅作状态标记）。")
        self.status_label.config(text="已请求停止", foreground="orange")

    def run_analysis(self):
        if self.is_running:
            messagebox.showwarning("提示", "任务正在运行中，请稍候。")
            return

        try:
            excel_path, api_key = self._build_params()
        except ValueError as e:
            messagebox.showwarning("参数错误", str(e))
            return

        self.save_config()
        self.log_text.delete("1.0", tk.END)
        self._set_running(True)

        def target():
            try:
                logging.info("=" * 50)
                logging.info("开始分析：CommentAnalyzer(path, api_key).run()")
                logging.info("Excel: %s", excel_path)
                logging.info("=" * 50)

                analyzer = CommentAnalyzer(excel_path, api_key)
                analyzer.run()

                if self.is_running:
                    logging.info("AI 分析流程已结束。")
                    self.root.after(
                        0,
                        lambda: self.on_finish(
                            True,
                            "分析完成，报告已保存至 Excel 同目录下的「分析报告」文件夹。",
                        ),
                    )
                else:
                    self.root.after(0, lambda: self.on_finish(False, "任务已中止"))
            except Exception as e:
                logging.exception("分析任务出错: %s", e)
                err_msg = str(e)
                self.root.after(0, lambda msg=err_msg: self.on_finish(False, msg))

        self.current_thread = threading.Thread(target=target, daemon=True)
        self.current_thread.start()

    def on_finish(self, success, message):
        self._set_running(False)
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("错误", message)


def main():
    root = tk.Tk()
    CommentAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
