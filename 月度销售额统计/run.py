import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

from main import Automation


DEFAULT_EXPERTS = (
    "lydia_homegoods,carhack_ryan,k8paz0xqw4,chicpicksbylydia,c7crfmav15,"
    "dailyfindsbylydia,detailing_dave_,furfreeliving_,haley1110,lydiashomefinds,"
    "homewithcamila,kerryshares,cleanwithlydia18,pltejkffq9,shopwithlydia_,"
    "sneakerheadmax_,spicypotato571,gppzoa2o03,cleaningwithemma91,"
    "cleaningwithsofia_,hrmb03eak0"
)


class App:

    def __init__(self, root):
        self.root = root
        self.root.title("月度销售额统计 - 数据下载")
        self.root.geometry("620x520")
        self.root.resizable(False, False)

        self.username = tk.StringVar(value="18512836434")
        self.password = tk.StringVar(value="Gyh1185202898.")
        self.ips = tk.StringVar(value="35.85.87.195")
        self.ports = tk.StringVar(value="8945")
        self.file_path = tk.StringVar(value=r"C:\RPA流程\月度销售额统计\flie")
        self.running = False

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        title = ttk.Label(
            main_frame,
            text="月度销售额统计 - 数据下载",
            font=("微软雅黑", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, pady=(0, 16))

        self._add_row(main_frame, 1, "易得客账号:", self.username)
        self._add_row(main_frame, 2, "易得客密码:", self.password, show="*")

        ttk.Label(main_frame, text="店铺 IP:").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.ips, width=55).grid(row=3, column=1, padx=10)
        ttk.Label(main_frame, text="多个用逗号分隔", foreground="gray").grid(row=3, column=2)

        ttk.Label(main_frame, text="调试端口:").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.ports, width=55).grid(row=4, column=1, padx=10)
        ttk.Label(main_frame, text="与 IP 一一对应", foreground="gray").grid(row=4, column=2)

        ttk.Label(main_frame, text="达人账号:").grid(row=5, column=0, sticky="nw", pady=8)
        self.experts_text = scrolledtext.ScrolledText(main_frame, width=42, height=6)
        self.experts_text.grid(row=5, column=1, padx=10, pady=8, sticky="w")
        self.experts_text.insert("1.0", DEFAULT_EXPERTS)

        ttk.Label(main_frame, text="下载目录:").grid(row=6, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.file_path, width=55).grid(row=6, column=1, padx=10)
        ttk.Button(main_frame, text="浏览", width=10, command=self.select_folder).grid(row=6, column=2)

        ttk.Separator(main_frame).grid(row=7, column=0, columnspan=3, pady=16, sticky="ew")

        self.start_btn = ttk.Button(
            main_frame,
            text="开始下载",
            width=25,
            command=self.start,
        )
        self.start_btn.grid(row=8, column=0, columnspan=3)

    def _add_row(self, parent, row, label, variable, show=None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)
        options = {"textvariable": variable, "width": 55}
        if show:
            options["show"] = show
        ttk.Entry(parent, **options).grid(row=row, column=1, padx=10, columnspan=2, sticky="w")

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.file_path.set(path)

    @staticmethod
    def parse_list(text):
        items = []
        for part in text.replace("\n", ",").split(","):
            part = part.strip()
            if part:
                items.append(part)
        return items

    def build_config(self):
        username = self.username.get().strip()
        password = self.password.get()
        ips = self.parse_list(self.ips.get())
        ports_raw = self.parse_list(self.ports.get())
        experts = self.parse_list(self.experts_text.get("1.0", "end"))
        file_path = self.file_path.get().strip()

        if not username:
            raise ValueError("请填写易得客账号")
        if not password:
            raise ValueError("请填写易得客密码")
        if not ips:
            raise ValueError("请填写至少一个店铺 IP")
        if not ports_raw:
            raise ValueError("请填写调试端口")
        if not experts:
            raise ValueError("请填写至少一个达人账号")
        if not file_path:
            raise ValueError("请选择下载目录")

        try:
            ports = [int(p) for p in ports_raw]
        except ValueError as exc:
            raise ValueError("调试端口必须为整数") from exc

        if len(ips) != len(ports):
            raise ValueError("店铺 IP 数量与调试端口数量必须一一对应")

        return {
            "username": username,
            "password": password,
            "ip": ips,
            "port": ports,
            "experts": experts,
            "file_path": file_path,
        }

    def start(self):
        if self.running:
            return

        try:
            config = self.build_config()
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return

        self.running = True
        self.start_btn.config(state="disabled")

        def run_task():
            try:
                automation = Automation(config)
                automation.Start()
                self.root.after(0, lambda: messagebox.showinfo("完成", "数据下载完成"))
            except Exception as e:
                self.root.after(0, lambda err=e: messagebox.showerror("错误", str(err)))
            finally:
                self.root.after(0, self._on_finish)

        threading.Thread(target=run_task, daemon=True).start()

    def _on_finish(self):
        self.running = False
        self.start_btn.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()

    style = ttk.Style()
    style.theme_use("clam")

    app = App(root)
    root.mainloop()
