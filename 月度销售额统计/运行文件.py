import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from analysis import ExcelUtil


class App:

    def __init__(self, root):
        self.root = root
        self.root.title("月度销售额统计工具")
        self.root.geometry("620x260")
        self.root.resizable(False, False)

        # 默认路径（请按本机填写或通过界面选择）
        self.summary_path = tk.StringVar(value="")
        self.data_path = tk.StringVar(value="")

        self.create_widgets()

    def create_widgets(self):

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        # 标题
        title = ttk.Label(
            main_frame,
            text="月度销售额统计工具",
            font=("微软雅黑", 16, "bold")
        )
        title.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # 汇总表路径
        ttk.Label(main_frame, text="汇总表 Excel:").grid(
            row=1, column=0, sticky="w", pady=8
        )

        ttk.Entry(main_frame, textvariable=self.summary_path, width=55).grid(
            row=1, column=1, padx=10
        )

        ttk.Button(
            main_frame,
            text="浏览",
            width=10,
            command=self.select_summary
        ).grid(row=1, column=2)

        # 数据文件夹
        ttk.Label(main_frame, text="数据文件夹:").grid(
            row=2, column=0, sticky="w", pady=8
        )

        ttk.Entry(main_frame, textvariable=self.data_path, width=55).grid(
            row=2, column=1, padx=10
        )

        ttk.Button(
            main_frame,
            text="浏览",
            width=10,
            command=self.select_folder
        ).grid(row=2, column=2)

        # 分隔线
        ttk.Separator(main_frame).grid(
            row=3, column=0, columnspan=3, pady=20, sticky="ew"
        )

        # 开始按钮
        start_btn = ttk.Button(
            main_frame,
            text="开始执行",
            width=25,
            command=self.start
        )
        start_btn.grid(row=4, column=0, columnspan=3)

    def select_summary(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if path:
            self.summary_path.set(path)

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.data_path.set(path)

    def start(self):

        summary = self.summary_path.get()
        data = self.data_path.get()

        try:
            excel = ExcelUtil()   # 创建实例
            excel.MergeData(summary, data)

            messagebox.showinfo("完成", "数据处理完成")

        except Exception as e:
            messagebox.showerror("错误", str(e))


if __name__ == "__main__":

    root = tk.Tk()

    style = ttk.Style()
    style.theme_use("clam")

    app = App(root)
    root.mainloop()