"""TREC 推广邮件 GUI 窗口入口。

run.py 只负责界面、参数收集和日志展示；正式邮件读取与发送逻辑统一放在 main.py。
"""

from __future__ import annotations

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from main import Main


class RunGui:
    """TREC 推广邮件发送窗口。"""

    def __init__(self):
        """初始化窗口、默认配置、控件变量和界面布局。"""
        # 默认配置：直接读取 main.py，避免 GUI 和正式流程维护两份配置。
        self.defaultConfig = Main.defaultConfig()

        # 窗口配置：标题、尺寸和最小尺寸全部来自 main.py 默认配置。
        self.window = tk.Tk()
        self.window.title(self.defaultConfig["windowTitle"])
        self.window.geometry(self.defaultConfig["windowSize"])
        self.window.minsize(
            self.defaultConfig["windowMinWidth"],
            self.defaultConfig["windowMinHeight"],
        )

        # 运行环境配置：与其他子项目一致，允许选择本机或线上。
        self.envVar = tk.StringVar(value="online" if self.defaultConfig["isOnline"] else "offline")

        # 数据文件配置：公司/个人文件默认指向上游搜索匹配后的固定结果表。
        self.companyEnabled = tk.BooleanVar(value=self.defaultConfig["includeCompany"])
        self.personEnabled = tk.BooleanVar(value=self.defaultConfig["includePerson"])
        self.companyFile = tk.StringVar(value=self.defaultConfig["companyFile"])
        self.personFile = tk.StringVar(value=self.defaultConfig["personFile"])

        # 固定邮件配置展示：只读展示 main.py 中锁定的 SMTP 参数，GUI 不负责修改。
        self.senderEmail = tk.StringVar(value=self.defaultConfig["senderEmail"])
        self.smtpServer = tk.StringVar(value=self.defaultConfig["smtpServer"])
        self.smtpPort = tk.StringVar(value=str(self.defaultConfig["smtpPort"]))
        self.smtpUser = tk.StringVar(value=self.defaultConfig["smtpUser"])
        self.smtpPassword = tk.StringVar(value=self.defaultConfig["smtpPassword"])

        # 运行状态控件：用于发送中禁用按钮，避免重复点击。
        self.sendButton = None
        self.logText = None

        self.buildWindow()

    def buildWindow(self):
        """创建主窗口布局。"""
        mainFrame = ttk.Frame(self.window, padding=16)
        mainFrame.pack(fill="both", expand=True)
        mainFrame.columnconfigure(0, weight=1)
        mainFrame.rowconfigure(3, weight=1)

        self.buildRuntimeBox(mainFrame)
        self.buildDataBox(mainFrame)
        self.buildSmtpBox(mainFrame)
        self.buildLogBox(mainFrame)
        self.buildButtonBox(mainFrame)
        self.log("已加载默认推广文件路径。")

    def buildRuntimeBox(self, parent):
        """创建运行环境选择区域。"""
        runtimeBox = ttk.LabelFrame(parent, text="运行环境")
        runtimeBox.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        runtimeBox.columnconfigure(2, weight=1)

        ttk.Radiobutton(
            runtimeBox,
            text="本机",
            value="offline",
            variable=self.envVar,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=8)

        ttk.Radiobutton(
            runtimeBox,
            text="线上",
            value="online",
            variable=self.envVar,
        ).grid(row=0, column=1, sticky="w", padx=10, pady=8)

    def buildDataBox(self, parent):
        """创建推广数据文件选择区域。"""
        dataBox = ttk.LabelFrame(parent, text="推广数据文件")
        dataBox.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        dataBox.columnconfigure(1, weight=1)

        ttk.Checkbutton(dataBox, text="公司推广数据", variable=self.companyEnabled).grid(
            row=0,
            column=0,
            sticky="w",
            padx=10,
            pady=8,
        )
        ttk.Entry(dataBox, textvariable=self.companyFile).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(dataBox, text="选择", command=lambda: self.chooseFile(self.companyFile)).grid(
            row=0,
            column=2,
            padx=10,
        )

        ttk.Checkbutton(dataBox, text="个人推广数据", variable=self.personEnabled).grid(
            row=1,
            column=0,
            sticky="w",
            padx=10,
            pady=8,
        )
        ttk.Entry(dataBox, textvariable=self.personFile).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(dataBox, text="选择", command=lambda: self.chooseFile(self.personFile)).grid(
            row=1,
            column=2,
            padx=10,
        )

    def buildSmtpBox(self, parent):
        """创建只读邮件配置展示区域。"""
        smtpBox = ttk.LabelFrame(parent, text="邮件发送配置（只读）")
        smtpBox.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        smtpBox.columnconfigure(1, weight=1)
        smtpBox.columnconfigure(3, weight=1)

        ttk.Label(smtpBox, text="发件邮箱").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(smtpBox, textvariable=self.senderEmail, state="readonly").grid(
            row=0,
            column=1,
            sticky="ew",
            padx=8,
        )

        ttk.Label(smtpBox, text="SMTP").grid(row=0, column=2, sticky="w", padx=10)
        ttk.Entry(smtpBox, textvariable=self.smtpServer, state="readonly").grid(
            row=0,
            column=3,
            sticky="ew",
            padx=8,
        )

        ttk.Label(smtpBox, text="端口").grid(row=0, column=4, sticky="w", padx=10)
        ttk.Entry(smtpBox, textvariable=self.smtpPort, width=8, state="readonly").grid(
            row=0,
            column=5,
            sticky="w",
            padx=8,
        )

        ttk.Label(smtpBox, text="登录用户").grid(row=1, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(smtpBox, textvariable=self.smtpUser, state="readonly").grid(
            row=1,
            column=1,
            sticky="ew",
            padx=8,
        )

        ttk.Label(smtpBox, text="SMTP授权码").grid(row=1, column=2, sticky="w", padx=10)
        ttk.Entry(smtpBox, textvariable=self.smtpPassword, show="*", state="readonly").grid(
            row=1,
            column=3,
            columnspan=3,
            sticky="ew",
            padx=8,
        )

    def buildLogBox(self, parent):
        """创建运行日志区域。"""
        logBox = ttk.LabelFrame(parent, text="运行日志")
        logBox.grid(row=3, column=0, sticky="nsew")
        logBox.rowconfigure(0, weight=1)
        logBox.columnconfigure(0, weight=1)

        self.logText = tk.Text(logBox, height=16, wrap="word")
        self.logText.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(logBox, orient="vertical", command=self.logText.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.logText.configure(yscrollcommand=scrollbar.set)

    def buildButtonBox(self, parent):
        """创建底部发送按钮区域。"""
        buttonBox = ttk.Frame(parent)
        buttonBox.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        buttonBox.columnconfigure(0, weight=1)

        self.sendButton = ttk.Button(buttonBox, text="开始发送邮件", command=self.startTask)
        self.sendButton.grid(row=0, column=1)

    def chooseFile(self, targetVar):
        """选择 Excel 或 CSV 推广数据文件。"""
        filePath = filedialog.askopenfilename(
            title="选择推广数据文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if filePath:
            targetVar.set(filePath)

    def selectedSources(self):
        """返回当前 GUI 勾选的数据源。"""
        sources = []
        if self.companyEnabled.get():
            sources.append(("公司", self.companyFile.get()))
        if self.personEnabled.get():
            sources.append(("个人", self.personFile.get()))
        return sources

    def log(self, message):
        """向日志框追加一行文本。"""
        self.logText.insert("end", str(message) + "\n")
        self.logText.see("end")

    def threadLog(self, message):
        """从工作线程安全写入 GUI 日志。"""
        self.window.after(0, self.log, message)

    def setRunning(self, running):
        """切换发送按钮状态。"""
        state = "disabled" if running else "normal"
        self.sendButton.configure(state=state)

    def startTask(self):
        """校验配置并启动邮件发送线程。"""
        sources = self.selectedSources()
        if not sources:
            messagebox.showwarning("缺少数据文件", "请至少选择公司推广数据或个人推广数据。")
            return

        note = "即将真实发送邮件。\n\n没有邮箱的行会自动跳过；发送结果会保存到后台邮件发送记录。\n\n是否继续？"
        if not messagebox.askyesno("确认发送", note):
            return

        self.setRunning(True)
        self.log("开始发送邮件...")

        workerThread = threading.Thread(
            target=self.worker,
            kwargs={
                "sources": sources,
            },
            daemon=True,
        )
        workerThread.start()

    def worker(self, sources):
        """在线程中调用 main.py 正式发送流程。"""
        try:
            config = self.defaultConfig.copy()
            config.update({
                "isOnline": self.envVar.get() == "online",
                "includeCompany": False,
                "includePerson": False,
            })
            result = Main(config, logCallback=self.threadLog).run(
                executeSend=True,
                sourceFiles=sources,
            )
            summary = result["summary"]
            self.threadLog(
                "完成：邮件成功 {emailSent}/{emailTotal}，失败 {emailFailed}".format(**summary)
            )
        except Exception as error:
            self.threadLog(f"运行失败: {error}")
            self.window.after(0, messagebox.showerror, "运行失败", str(error))
        finally:
            self.window.after(0, self.setRunning, False)

    def run(self):
        """启动 GUI 主循环。"""
        self.window.mainloop()


if __name__ == "__main__":
    RunGui().run()
