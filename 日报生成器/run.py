"""工作日报生成器 — Tkinter GUI 入口。"""

import json
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from main import (
    DEFAULT_NUMBER,
    DEFAULT_TITLE,
    DEFAULT_WEBHOOK,
    Comment,
)


def get_app_base_dir():
    """脚本目录；PyInstaller 打包时取 exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_FILE = get_app_base_dir() / "run_config.json"


class GuiComment(Comment):
    """GUI 录入的今日/明日事项作为数据源。"""

    def __init__(self, config, daily_items, tomorrow_items):
        super().__init__(config)
        self._daily_items = daily_items
        self.tomorrow_items = tomorrow_items

    def main(self):
        return self._daily_items


class DailyReportWindow:
    """主窗口：配置、今日事项、明日待办、预览与分条推送。"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("工作日报生成器")
        self.root.geometry("560x820")
        self.root.minsize(520, 760)

        self.daily_items = []
        self.tomorrow_items = []
        self.current_date = datetime.now().strftime("%Y-%m-%d")

        self.account_var = tk.StringVar(value=DEFAULT_NUMBER)
        self.webhook_var = tk.StringVar(value=DEFAULT_WEBHOOK)
        self.date_var = tk.StringVar(value=self.current_date)

        self.item_title_var = tk.StringVar()
        self.owner_var = tk.StringVar()
        self.status_var = tk.StringVar(value="开发中")

        self.tomorrow_title_var = tk.StringVar()
        self.tomorrow_owner_var = tk.StringVar()

        self._load_config()
        self.build_window()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_window(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(8, weight=1)

        row = 0
        ttk.Label(container, text="当前日期").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.date_var, state="readonly").grid(
            row=row, column=1, sticky="ew", pady=6
        )
        row += 1

        ttk.Label(container, text="企业微信账号").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.account_var).grid(
            row=row, column=1, sticky="ew", pady=6
        )
        row += 1

        ttk.Label(container, text="群 webhook").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.webhook_var).grid(
            row=row, column=1, sticky="ew", pady=6
        )
        row += 1

        webhook_hint = (
            "请确认群聊已添加消息推送机器人，并将 webhook 填入上方；\n"
            "今日日报与明日待办将分两条 Markdown 消息发送。"
        )
        ttk.Label(container, text=webhook_hint, foreground="#6b7280", wraplength=440).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )
        row += 1

        today_box = ttk.LabelFrame(container, text="今日工作日报", padding=10)
        today_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=6)
        today_box.columnconfigure(1, weight=1)
        self._build_item_form(
            today_box,
            self.item_title_var,
            self.owner_var,
            self.status_var,
            include_status=True,
        )
        self.detail_text = self._last_detail_widget
        ttk.Button(today_box, text="添加今日事项", command=self.add_daily_item).grid(
            row=4, column=1, sticky="e", pady=(8, 0)
        )
        row += 1

        tomorrow_box = ttk.LabelFrame(container, text="明日待办", padding=10)
        tomorrow_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=6)
        tomorrow_box.columnconfigure(1, weight=1)
        self._build_item_form(
            tomorrow_box,
            self.tomorrow_title_var,
            self.tomorrow_owner_var,
            None,
            include_status=False,
        )
        self.tomorrow_detail_text = self._last_detail_widget
        ttk.Button(tomorrow_box, text="添加明日待办", command=self.add_tomorrow_item).grid(
            row=3, column=1, sticky="e", pady=(8, 0)
        )
        row += 1

        button_box = ttk.Frame(container)
        button_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=4)
        ttk.Button(button_box, text="预览", command=self.preview_items).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(button_box, text="发送企业微信", command=self.send_message).pack(side=tk.LEFT)
        row += 1

        ttk.Label(container, text="推送预览").grid(row=row, column=0, sticky="w", pady=(8, 4))
        row += 1
        self.preview_text = tk.Text(container, height=12, wrap="word")
        self.preview_text.grid(row=row, column=0, columnspan=2, sticky="nsew")
        row += 1

        self.status_label = ttk.Label(
            container,
            text=self._status_summary(),
            foreground="#166534",
        )
        self.status_label.grid(row=row, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _build_item_form(self, parent, title_var, owner_var, status_var, include_status):
        ttk.Label(parent, text="标题").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=title_var).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(parent, text="负责人").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=owner_var).grid(row=1, column=1, sticky="ew", pady=6)

        detail_row = 2
        if include_status:
            ttk.Label(parent, text="状态").grid(row=2, column=0, sticky="w", pady=6)
            ttk.Combobox(
                parent,
                textvariable=status_var,
                values=("开发中", "已完成", "测试中"),
                state="readonly",
            ).grid(row=2, column=1, sticky="ew", pady=6)
            detail_row = 3

        ttk.Label(parent, text="详情").grid(row=detail_row, column=0, sticky="nw", pady=6)
        detail_widget = tk.Text(parent, height=4, wrap="word")
        detail_widget.grid(row=detail_row, column=1, sticky="ew", pady=6)
        self._last_detail_widget = detail_widget

    def _status_summary(self):
        return (
            f"今日 {len(self.daily_items)} 条，明日 {len(self.tomorrow_items)} 条，待发送"
        )

    def _get_text(self, widget):
        return widget.get("1.0", "end").strip()

    def _clear_daily_form(self):
        self.item_title_var.set("")
        self.owner_var.set("")
        self.status_var.set("开发中")
        self.detail_text.delete("1.0", "end")

    def _clear_tomorrow_form(self):
        self.tomorrow_title_var.set("")
        self.tomorrow_owner_var.set("")
        self.tomorrow_detail_text.delete("1.0", "end")

    def _validate_item(self, title, owner, detail, section_name):
        if not title:
            messagebox.showwarning("提示", f"请填写{section_name}标题")
            return False
        if not owner:
            messagebox.showwarning("提示", f"请填写{section_name}负责人")
            return False
        if not detail:
            messagebox.showwarning("提示", f"请填写{section_name}详情")
            return False
        return True

    def add_daily_item(self):
        title = self.item_title_var.get().strip()
        owner = self.owner_var.get().strip()
        status = self.status_var.get().strip()
        detail = self._get_text(self.detail_text)
        if not self._validate_item(title, owner, detail, "今日事项"):
            return

        comment = Comment({})
        self.daily_items.append(
            {
                "事项": title,
                "负责人": owner,
                "状态": status,
                "详情": comment.format_detail(detail),
            }
        )
        self._clear_daily_form()
        self._save_config()
        self.status_label.config(text=self._status_summary())
        messagebox.showinfo("提示", "已添加今日事项")

    def add_tomorrow_item(self):
        title = self.tomorrow_title_var.get().strip()
        owner = self.tomorrow_owner_var.get().strip()
        detail = self._get_text(self.tomorrow_detail_text)
        if not self._validate_item(title, owner, detail, "明日待办"):
            return

        comment = Comment({})
        self.tomorrow_items.append(
            {
                "事项": title,
                "负责人": owner,
                "详情": comment.format_detail(detail),
            }
        )
        self._clear_tomorrow_form()
        self._save_config()
        self.status_label.config(text=self._status_summary())
        messagebox.showinfo("提示", "已添加明日待办")

    def build_comment(self):
        return GuiComment(
            {
                "title": DEFAULT_TITLE,
                "report_date": self.current_date,
                "number": self.account_var.get().strip(),
                "wechat_webhook": self.webhook_var.get().strip(),
                "send_wechat": True,
            },
            self.daily_items,
            self.tomorrow_items,
        )

    def preview_items(self):
        if not self.daily_items and not self.tomorrow_items:
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", "暂无已保存的今日事项或明日待办")
            self.status_label.config(text="暂无已保存内容")
            return
        self.refresh_preview()
        self.status_label.config(text="已展示推送预览（今日与明日分两条发送）")

    def refresh_preview(self):
        comment = self.build_comment()
        blocks = []
        if self.daily_items:
            blocks.append(
                comment.build_items_markdown(
                    f"{DEFAULT_TITLE} {self.current_date}",
                    self.daily_items,
                    include_status=True,
                )
            )
        if self.tomorrow_items:
            blocks.append(
                comment.build_items_markdown(
                    f"明日待办 {comment.tomorrow_date()}",
                    self.tomorrow_items,
                    include_status=False,
                )
            )
        content = "\n\n━━━━━━━━━━━━━━━━\n【将分两条消息发送】\n━━━━━━━━━━━━━━━━\n\n".join(
            blocks
        )
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", content)

    def send_message(self):
        if not self.daily_items and not self.tomorrow_items:
            messagebox.showwarning("提示", "请至少添加一条今日事项或明日待办")
            return
        if not self.webhook_var.get().strip():
            messagebox.showwarning("提示", "请填写消息推送群的 webhook 值")
            return

        self._save_config()
        self.status_label.config(text="正在发送企业微信消息...")
        thread = threading.Thread(target=self.send_message_worker, daemon=True)
        thread.start()

    def send_message_worker(self):
        try:
            comment = self.build_comment()
            comment.message_send(
                daily_items=self.daily_items,
                tomorrow_items=self.tomorrow_items,
            )
            self.root.after(0, self.send_success)
        except Exception as exc:
            self.root.after(0, lambda error=exc: self.send_failed(error))

    def send_success(self):
        self.status_label.config(text="企业微信发送流程已执行，请查看控制台发送结果")
        messagebox.showinfo("完成", "今日日报与明日待办已分条发送，请查看群内消息")

    def send_failed(self, error):
        self.status_label.config(text=f"发送失败: {error}")
        messagebox.showerror("发送失败", str(error))

    def _load_config(self):
        if not CONFIG_FILE.is_file():
            return
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return

        number = str(data.get("number") or "").strip()
        webhook = str(data.get("wechat_webhook") or "").strip()
        if number:
            self.account_var.set(number)
        if webhook:
            self.webhook_var.set(webhook)

        self.daily_items = list(data.get("daily_items") or [])
        self.tomorrow_items = list(data.get("tomorrow_items") or [])

    def _save_config(self):
        data = {
            "number": self.account_var.get().strip() or DEFAULT_NUMBER,
            "wechat_webhook": self.webhook_var.get().strip() or DEFAULT_WEBHOOK,
            "daily_items": self.daily_items,
            "tomorrow_items": self.tomorrow_items,
        }
        try:
            CONFIG_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass

    def on_close(self):
        self._save_config()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DailyReportWindow()
    app.run()
