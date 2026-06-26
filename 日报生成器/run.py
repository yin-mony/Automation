import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from datetime import datetime

from main import Comment


class GuiComment(Comment):
    def __init__(self, config, daily_items):
        super().__init__(config)
        self.daily_items = daily_items

    # GUI 输入的数据作为自动化流程结果
    def main(self):
        return self.daily_items


class DailyReportWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("工作日报生成器")
        self.root.geometry("520x540")
        self.root.minsize(500, 520)

        self.daily_items = []
        self.current_date = datetime.now().strftime("%Y-%m-%d")

        self.account_var = tk.StringVar()
        self.webhook_var = tk.StringVar()
        self.item_title_var = tk.StringVar()
        self.owner_var = tk.StringVar()
        self.status_var = tk.StringVar(value="开发中")
        self.date_var = tk.StringVar(value=self.current_date)

        self.build_window()

    def build_window(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        frame = ttk.Frame(self.root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(7, weight=1)

        ttk.Label(frame, text="当前日期").grid(row=0, column=0, sticky="w", pady=6)
        date_entry = ttk.Entry(frame, textvariable=self.date_var, state="readonly")
        date_entry.grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(frame, text="企业微信账号").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.account_var).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(frame, text="群 webhook").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.webhook_var).grid(row=2, column=1, sticky="ew", pady=6)

        webhook_hint = (
            "请检查需要推送的群聊是否添加消息推送，\n如已添加请将对应的webhook值填入；\n"
            "如果未开启添加，请先开启添加"
        )
        ttk.Label(frame, text=webhook_hint, foreground="#6b7280", wraplength=420).grid(
            row=3,
            column=1,
            sticky="w",
            pady=(0, 10),
        )

        content_box = ttk.LabelFrame(frame, text="工作日报内容", padding=10)
        content_box.grid(row=4, column=0, columnspan=2, sticky="ew", pady=6)
        content_box.columnconfigure(1, weight=1)

        ttk.Label(content_box, text="标题").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(content_box, textvariable=self.item_title_var).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(content_box, text="负责人").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(content_box, textvariable=self.owner_var).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(content_box, text="状态").grid(row=2, column=0, sticky="w", pady=6)
        status_box = ttk.Combobox(
            content_box,
            textvariable=self.status_var,
            values=("开发中", "已完成", "测试中"),
            state="readonly",
        )
        status_box.grid(row=2, column=1, sticky="ew", pady=6)

        ttk.Label(content_box, text="详情").grid(row=3, column=0, sticky="nw", pady=6)
        self.detail_text = tk.Text(content_box, height=4, wrap="word")
        self.detail_text.grid(row=3, column=1, sticky="ew", pady=6)

        button_box = ttk.Frame(frame)
        button_box.grid(row=5, column=0, columnspan=2, sticky="ew", pady=4)
        button_box.columnconfigure(2, weight=1)

        ttk.Button(button_box, text="添加事项", command=self.add_item).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(button_box, text="预览", command=self.preview_items).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(button_box, text="发送企业微信", command=self.send_message).grid(row=0, column=3)

        ttk.Label(frame, text="日报预览").grid(row=6, column=0, sticky="w", pady=(8, 4))
        self.preview_text = tk.Text(frame, height=7, wrap="word")
        self.preview_text.grid(row=7, column=0, columnspan=2, sticky="nsew")

        self.status_label = ttk.Label(frame, text="等待输入日报内容", foreground="#166534")
        self.status_label.grid(row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def get_detail(self):
        return self.detail_text.get("1.0", "end").strip()

    def add_item(self):
        title = self.item_title_var.get().strip()
        owner = self.owner_var.get().strip()
        status = self.status_var.get().strip()
        detail = self.get_detail()

        if not title:
            messagebox.showwarning("提示", "请填写工作日报内容标题")
            return
        if not owner:
            messagebox.showwarning("提示", "请填写负责人")
            return
        if not detail:
            messagebox.showwarning("提示", "请填写详情")
            return

        self.daily_items.append(
            {
                "事项": title,
                "负责人": owner,
                "状态": status,
                "详情": detail,
            }
        )
        self.clear_form()
        self.status_label.config(text=f"已保存 {len(self.daily_items)} 条日报内容，待提交发送")
        messagebox.showinfo("提示", "已保存当前日报内容，待提交发送")

    def clear_form(self):
        self.item_title_var.set("")
        self.owner_var.set("")
        self.status_var.set("开发中")
        self.detail_text.delete("1.0", "end")

    def build_comment(self):
        return GuiComment(
            {
                "title": "工作日报",
                "number": self.account_var.get().strip(),
                "wechat_webhook": self.webhook_var.get().strip(),
                "send_wechat": True,
            },
            self.daily_items,
        )

    def preview_items(self):
        if not self.daily_items:
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", "暂无已保存的日报内容")
            self.status_label.config(text="暂无已保存的日报内容")
            return
        self.refresh_preview()
        self.status_label.config(text="已展示保存的日报内容")

    def refresh_preview(self):
        self.preview_text.delete("1.0", "end")

        report_name = f"工作日报-{self.current_date}"
        lines = [report_name]
        for index, item in enumerate(self.daily_items, 1):
            lines.extend(
                [
                    f"{index}.{item.get('事项', '')}",
                    f"负责人：{item.get('负责人', '')}",
                    f"状态：{item.get('状态', '')}",
                    f"详情：{item.get('详情', '')}",
                    "",
                ]
            )
        content = "\n".join(lines).rstrip()

        self.preview_text.insert("1.0", content)

    def send_message(self):
        if not self.daily_items:
            messagebox.showwarning("提示", "请先添加至少一条日报内容")
            return
        if not self.webhook_var.get().strip():
            messagebox.showwarning("提示", "请填写消息推送群的 webhook 值")
            return

        self.status_label.config(text="正在发送企业微信消息...")
        thread = threading.Thread(target=self.send_message_worker, daemon=True)
        thread.start()

    def send_message_worker(self):
        try:
            comment = self.build_comment()
            data = comment.excel()
            comment.message_send(data)
            self.root.after(0, self.send_success)
        except Exception as e:
            self.root.after(0, lambda error=e: self.send_failed(error))

    def send_success(self):
        self.status_label.config(text="企业微信发送流程已执行，请查看控制台发送结果")
        messagebox.showinfo("完成", "企业微信发送流程已执行，请查看控制台发送结果")

    def send_failed(self, error):
        self.status_label.config(text=f"发送失败: {error}")
        messagebox.showerror("发送失败", str(error))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DailyReportWindow()
    app.run()
