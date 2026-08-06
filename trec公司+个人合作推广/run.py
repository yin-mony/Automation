"""TREC 公司+个人合作推广 GUI 窗口入口。"""

import json
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from main import Main


class RunGui:
    """TREC 公司+个人合作推广运行窗口。"""

    def __init__(self):
        """初始化窗口、默认配置、控件变量和日志轮询。"""
        self.baseDir = Path(__file__).resolve().parent
        self.configPath = self.baseDir / "run_config.json"
        self.baseMain = Main()
        self.baseConfig = self.baseMain.config
        self.helpText = {
            "isOnline": "运行环境标记，只影响日志和邮件正文。本机调试选本机，正式运行选线上。",
            "outputDir": "最终结果表、断点和缓存所在目录。默认 output，不能填 file。",
            "rawFileName": "file 内置目录中的未清洗全量底表文件名，仅用于保留和邮件附件。",
            "cleanFileName": "file 内置目录中的已清洗初始表文件名，当前搜索流程只读取它。",
            "companyResultFileName": "公司模式最终导出表文件名。",
            "personResultFileName": "个人模式最终导出表文件名。",
            "expireMonths": "个人模式到期月份。默认 6，表示未来六个月内到期。",
            "sendEmail": "邮件发送开关。开启后流程结束只发送固定数据表附件。",
            "email": "收件邮箱。开启邮件发送时必须填写，可用英文逗号或分号分隔多个邮箱。",
            "emailSubject": "邮件标题。默认即可。",
            "promotionExecuteSend": "推广邮件流程固定会生成后台发送记录；选择真实发送后才会连接 SMTP 发出推广邮件。",
            "promotionSenderEmail": "推广邮件固定发件邮箱，仅展示不可修改；SMTP 服务器和授权码不在界面展示。",
        }

        # 运行状态变量。
        self.workerThread = None
        self.isRunning = False
        self.logQueue = queue.Queue()

        # GUI 控件变量。
        self.configVars = {}
        self.modeVar = None
        self.envVar = None
        self.mailVar = None
        self.promotionMailVar = None
        self.emailEntry = None

        # 主窗口基础配置。
        self.window = tk.Tk()
        self.window.title("TREC 公司+个人合作推广")
        self.window.geometry("1120x760")
        self.window.minsize(980, 660)
        self.window.protocol("WM_DELETE_WINDOW", self.closeWindow)

        # 运行模式是窗口状态，不写入 main.py 默认配置，也不保存到 run_config.json。
        self.modeVar = tk.StringVar(value="company")
        self.envVar = tk.StringVar(value="online" if self.baseConfig.get("isOnline") else "offline")
        self.mailVar = tk.StringVar(value="yes" if self.baseConfig.get("sendEmail") else "no")
        self.promotionMailVar = tk.StringVar(value="send" if self.baseConfig.get("promotionExecuteSend") else "record")

        self.buildWindow()
        self.loadDefaultValues()
        self.loadConfig(silent=True)
        self.switchMode()
        self.window.after(200, self.pollLog)

    def buildWindow(self):
        """创建窗口布局。"""
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)

        self.buildTopBar()

        body = ttk.Frame(self.window, padding=(10, 0, 10, 10))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        self.buildModeBar(body)

        rightFrame = ttk.Frame(body)
        rightFrame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        rightFrame.columnconfigure(0, weight=1)
        rightFrame.rowconfigure(0, weight=4)
        rightFrame.rowconfigure(1, weight=1)

        self.contentFrame = ttk.Frame(rightFrame, padding=12)
        self.contentFrame.grid(row=0, column=0, sticky="nsew")
        self.contentFrame.columnconfigure(0, weight=1)

        logFrame = ttk.LabelFrame(rightFrame, text="运行日志", padding=8)
        logFrame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        logFrame.columnconfigure(0, weight=1)
        logFrame.rowconfigure(0, weight=1)
        self.logText = ScrolledText(logFrame, wrap=tk.WORD, height=10)
        self.logText.grid(row=0, column=0, sticky="nsew")

    def buildTopBar(self):
        """创建顶部运行按钮和状态提示。"""
        topFrame = ttk.Frame(self.window, padding=(10, 10, 10, 8))
        topFrame.grid(row=0, column=0, sticky="ew")
        topFrame.columnconfigure(1, weight=1)

        self.startButton = ttk.Button(topFrame, text="开始运行", command=self.startTask)
        self.startButton.grid(row=0, column=0, padx=(0, 8))

        self.statusVar = tk.StringVar(value="就绪")
        ttk.Label(topFrame, textvariable=self.statusVar, foreground="#0f6b7a").grid(row=0, column=1, sticky="e")

    def buildModeBar(self, parent):
        """创建左侧模式按钮。"""
        modeFrame = ttk.LabelFrame(parent, text="运行模式", padding=10)
        modeFrame.grid(row=0, column=0, sticky="ns")

        modes = [
            ("config", "配置"),
            ("company", "公司模式"),
            ("person", "个人模式"),
        ]
        for rowNo, (modeValue, modeText) in enumerate(modes):
            ttk.Radiobutton(
                modeFrame,
                text=modeText,
                value=modeValue,
                variable=self.modeVar,
                command=self.switchMode,
                width=18,
            ).grid(row=rowNo, column=0, sticky="ew", pady=4)

        noteText = (
            "搜索接口：SerpApi。\n\n"
            "搜索页：固定 Google 第一页。\n\n"
            "批量：每次固定 10 个对象。\n\n"
            "断点和缓存：后台自动保存。"
        )
        ttk.Label(modeFrame, text=noteText, foreground="#666666", wraplength=170, justify="left").grid(
            row=len(modes), column=0, sticky="w", pady=(16, 0)
        )

    def switchMode(self):
        """切换模式并刷新右侧配置。"""
        if not self.contentFrame:
            return
        for child in self.contentFrame.winfo_children():
            child.destroy()

        mode = self.modeVar.get()
        if mode == "config":
            self.buildConfigPanel()
        elif mode == "person":
            self.buildPersonPanel()
        else:
            self.buildCompanyPanel()
        self.refreshPlanInfo()

    def buildConfigPanel(self):
        """创建公共配置页。"""
        self.addTitle("配置", "file 是固定内置数据目录；输出目录只用于结果表、断点和缓存。")
        self.addConfigActionBox(row=1)
        self.addRuntimeBox(row=2)
        self.addPathBox(row=3)
        self.addRuleBox(row=4)
        self.addPlanBox(row=5)

    def buildCompanyPanel(self):
        """创建公司模式配置页。"""
        self.addTitle("公司模式", "读取清洗表中有挂靠公司的数据，按公司名去重后调用 SerpApi，并用浏览器打开返回 link 的二级页面。")
        self.addRuleBox(row=1)
        self.addPlanBox(row=2)

    def buildPersonPanel(self):
        """创建个人模式配置页。"""
        self.addTitle("个人模式", "读取清洗表中无挂靠公司、Active、指定月份内到期的个人，再调用 SerpApi 并用浏览器打开二级页面提取。")
        self.addPersonBox(row=1)
        self.addRuleBox(row=2)
        self.addPlanBox(row=3)

    def addTitle(self, title, description):
        """添加页面标题。"""
        titleFrame = ttk.Frame(self.contentFrame)
        titleFrame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        titleFrame.columnconfigure(0, weight=1)
        ttk.Label(titleFrame, text=title, font=("Microsoft YaHei UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(titleFrame, text=description, foreground="#666666", wraplength=820).grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

    def addConfigActionBox(self, row):
        """添加配置保存和恢复按钮。"""
        box = ttk.LabelFrame(self.contentFrame, text="配置操作", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(2, weight=1)
        ttk.Button(box, text="保存配置", command=self.saveConfig).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(box, text="恢复默认", command=lambda: self.loadDefaultValues(keepMode=True)).grid(
            row=0, column=1, padx=4
        )
        ttk.Label(
            box,
            text="保存后会写入 run_config.json；运行模式、SerpApi Key、固定发件邮箱、SMTP 服务器和授权码不会写入本地配置。",
            foreground="#666666",
            wraplength=700,
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))

    def addRuntimeBox(self, row):
        """添加运行环境和邮件配置。"""
        box = ttk.LabelFrame(self.contentFrame, text="运行环境和邮件", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(1, weight=1)
        box.columnconfigure(4, weight=1)

        ttk.Label(box, text="运行环境").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Radiobutton(box, text="本机", value="offline", variable=self.envVar).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(box, text="线上", value="online", variable=self.envVar).grid(row=0, column=2, sticky="w")
        self.addHelpLabel(box, "isOnline", row=0, column=3, columnspan=3, wraplength=520)

        ttk.Label(box, text="邮件发送").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Radiobutton(box, text="不发送", value="no", variable=self.mailVar, command=self.toggleEmail).grid(
            row=1, column=1, sticky="w"
        )
        ttk.Radiobutton(box, text="发送", value="yes", variable=self.mailVar, command=self.toggleEmail).grid(
            row=1, column=2, sticky="w"
        )
        self.addHelpLabel(box, "sendEmail", row=1, column=3, columnspan=3, wraplength=520)

        ttk.Label(box, text="收件邮箱").grid(row=2, column=0, sticky="w", pady=4)
        self.emailEntry = self.addEntry(box, "email", row=2, column=1, width=32)
        self.addHelpLabel(box, "email", row=2, column=2, columnspan=4, wraplength=620)

        ttk.Label(box, text="邮件标题").grid(row=3, column=0, sticky="w", pady=4)
        self.addEntry(box, "emailSubject", row=3, column=1, width=32)
        self.addHelpLabel(box, "emailSubject", row=3, column=2, columnspan=4, wraplength=620)

        ttk.Label(box, text="推广发件邮箱").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Label(
            box,
            text=self.baseConfig.get("promotionSenderEmail", ""),
            foreground="#666666",
        ).grid(row=4, column=1, sticky="w", padx=(4, 10), pady=4)
        self.addHelpLabel(box, "promotionSenderEmail", row=4, column=2, columnspan=4, wraplength=620)

        ttk.Label(box, text="推广邮件模式").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Radiobutton(box, text="只生成后台记录", value="record", variable=self.promotionMailVar).grid(
            row=5, column=1, sticky="w"
        )
        ttk.Radiobutton(box, text="真实发送推广邮件", value="send", variable=self.promotionMailVar).grid(
            row=5, column=2, sticky="w"
        )
        self.addHelpLabel(box, "promotionExecuteSend", row=5, column=3, columnspan=3, wraplength=520)
        self.toggleEmail()

    def addPathBox(self, row):
        """添加数据文件和输出配置。"""
        box = ttk.LabelFrame(self.contentFrame, text="内置数据和输出目录", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(1, weight=1)
        box.columnconfigure(4, weight=1)

        ttk.Label(box, text="输出目录").grid(row=0, column=0, sticky="w", pady=4)
        self.addEntry(box, "outputDir", row=0, column=1)
        ttk.Button(box, text="选择", command=lambda: self.browseDir("outputDir")).grid(row=0, column=2, padx=6)
        self.addHelpLabel(box, "outputDir", row=0, column=3, columnspan=3)

        ttk.Label(box, text="内置目录").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Label(box, text="file（固定不可改）", foreground="#666666").grid(row=1, column=1, sticky="w", padx=(4, 10), pady=4)
        ttk.Label(box, text="程序只从这里读取内置底表，不把结果写回 file。", foreground="#666666").grid(
            row=1, column=3, columnspan=3, sticky="w", pady=4
        )

        self.addLabeledEntry(box, "cleanFileName", "清洗初始表", row=2, column=0)
        self.addLabeledEntry(box, "rawFileName", "未清洗底表", row=2, column=3)
        self.addLabeledEntry(box, "companyResultFileName", "公司结果表", row=3, column=0)
        self.addLabeledEntry(box, "personResultFileName", "个人结果表", row=3, column=3)

    def addRuleBox(self, row):
        """添加固定搜索规则说明。"""
        box = ttk.LabelFrame(self.contentFrame, text="固定搜索规则", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(0, weight=1)
        text = (
            "1. SerpApi 固定请求 Google 第一页，不设置返回条数，按当前页实际返回的 organic_results 全部处理。\n"
            "2. 每次运行固定处理 10 个未完成对象，每个对象消耗 1 次 SerpApi 搜索额度。\n"
            "3. 程序会读取 organic_results 里的 link 字段，并用 DP 浏览器打开普通网页二级页面提取邮箱和电话。\n"
            "4. 断点文件和缓存文件由公司/个人逻辑后台自动保存，不需要在配置里填写。"
        )
        ttk.Label(box, text=text, foreground="#666666", justify="left", wraplength=840).grid(row=0, column=0, sticky="w")

    def addPersonBox(self, row):
        """添加个人模式配置。"""
        box = ttk.LabelFrame(self.contentFrame, text="个人模式配置", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(1, weight=1)
        self.addLabeledEntry(box, "expireMonths", "到期月份", row=0, column=0)
        ttk.Label(
            box,
            text="个人模式固定筛选 Active、无挂靠公司、未来到期月份范围内的数据；地区不限制。",
            foreground="#666666",
            wraplength=760,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def addPlanBox(self, row):
        """添加本次运行说明。"""
        box = ttk.LabelFrame(self.contentFrame, text="本次会怎么跑", padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=6)
        box.columnconfigure(0, weight=1)
        self.planVar = tk.StringVar(value="")
        ttk.Label(box, textvariable=self.planVar, justify="left", wraplength=840).grid(row=0, column=0, sticky="w")
        self.refreshPlanInfo()

    def addEntry(self, parent, key, row, column, width=18):
        """添加普通输入框。"""
        variable = self.configVars.get(key)
        if variable is None:
            variable = tk.StringVar()
            self.configVars[key] = variable
        entry = ttk.Entry(parent, textvariable=variable, width=width)
        entry.grid(row=row, column=column, sticky="ew", padx=(4, 10), pady=4)
        parent.columnconfigure(column, weight=1)
        return entry

    def addLabeledEntry(self, parent, key, label, row, column):
        """添加带标签的输入框。"""
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", pady=4)
        return self.addEntry(parent, key, row=row, column=column + 1)

    def addHelpLabel(self, parent, key, row, column, columnspan=1, wraplength=360):
        """添加灰色说明文字。"""
        text = self.helpText.get(key, "")
        if not text:
            return None
        label = ttk.Label(parent, text=text, foreground="#666666", justify="left", wraplength=wraplength)
        label.grid(row=row, column=column, columnspan=columnspan, sticky="w", padx=(6, 0), pady=4)
        return label

    def toggleEmail(self):
        """根据邮件开关启用或禁用收件邮箱。"""
        state = tk.NORMAL if self.mailVar and self.mailVar.get() == "yes" else tk.DISABLED
        if self.emailEntry:
            self.emailEntry.configure(state=state)

    def loadDefaultValues(self, keepMode=False):
        """把默认配置写入 GUI 控件。"""
        currentMode = self.modeVar.get() if self.modeVar else "company"
        hiddenKeys = {
            "isOnline", "sendEmail", "promotionExecuteSend", "promotionSenderEmail", "promotionRecordFileName",
            "serpapiUrl", "serpapiKey", "sender_email", "smtp_auth_code",
        }
        for key, value in self.baseConfig.items():
            if key in hiddenKeys:
                continue
            self.configVars.setdefault(key, tk.StringVar()).set("" if value is None else str(value))

        self.envVar.set("online" if self.baseConfig.get("isOnline") else "offline")
        self.mailVar.set("yes" if self.baseConfig.get("sendEmail") else "no")
        self.promotionMailVar.set("send" if self.baseConfig.get("promotionExecuteSend") else "record")
        self.modeVar.set(currentMode if keepMode else "company")
        self.toggleEmail()
        self.statusVar.set("已恢复默认配置")

    def loadConfig(self, silent=False):
        """从 run_config.json 加载本地配置。"""
        if not self.configPath.exists():
            return
        try:
            config = json.loads(self.configPath.read_text(encoding="utf-8"))
        except Exception as error:
            if not silent:
                messagebox.showerror("读取失败", str(error))
            return

        self.envVar.set("online" if config.get("isOnline") else "offline")
        self.mailVar.set("yes" if config.get("sendEmail") else "no")
        self.promotionMailVar.set("send" if config.get("promotionExecuteSend") else "record")

        hiddenKeys = {
            "isOnline", "sendEmail", "promotionExecuteSend", "promotionSenderEmail", "promotionRecordFileName", "runMode",
            "serpapiUrl", "serpapiKey", "sender_email", "smtp_auth_code",
        }
        for key, value in config.items():
            if key in hiddenKeys:
                continue
            if key in self.configVars:
                self.configVars[key].set("" if value is None else str(value))
        self.toggleEmail()
        if not silent:
            self.statusVar.set("配置已加载")

    def buildConfig(self, withCallback=True, showError=True):
        """读取 GUI 控件并转换为 Main 需要的配置。"""
        try:
            config = dict(self.baseConfig)
            config.update({
                "runMode": self.modeVar.get(),
                "isOnline": self.envVar.get() == "online",
                "sendEmail": self.mailVar.get() == "yes",
                "promotionExecuteSend": self.promotionMailVar.get() == "send",
                "email": self.getText("email"),
                "emailSubject": self.getText("emailSubject"),
                "outputDir": self.getText("outputDir") or "output",
                "rawFileName": self.getText("rawFileName"),
                "cleanFileName": self.getText("cleanFileName"),
                "companyResultFileName": self.getText("companyResultFileName"),
                "personResultFileName": self.getText("personResultFileName"),
                "expireMonths": self.getInt("expireMonths"),
            })
        except ValueError as error:
            if showError:
                messagebox.showerror("配置错误", str(error))
            return None

        if config["sendEmail"] and not config["email"]:
            if showError:
                messagebox.showerror("配置错误", "开启邮件发送时必须填写收件邮箱。")
            return None

        if Path(config["outputDir"]).name.lower() == "file":
            if showError:
                messagebox.showerror("配置错误", "输出目录不能设置为 file。file 是固定内置数据目录，只能读取底表。")
            return None

        if config["expireMonths"] <= 0:
            if showError:
                messagebox.showerror("配置错误", "到期月份必须大于 0。")
            return None

        return config

    def buildRunPlanText(self, config=None):
        """生成当前配置下的运行说明。"""
        if config is None:
            config = self.buildConfig(withCallback=False, showError=False)
        if config is None:
            return "当前配置暂时不完整，请检查数字输入框。"

        outputDir = Path(config["outputDir"] or "output")
        lines = [
            f"运行环境：{'线上' if config.get('isOnline') else '本机'}",
            "内置数据目录：file（固定不可改）",
            f"输出目录：{outputDir}",
            f"读取清洗初始表：file/{config['cleanFileName']}",
            "搜索页：固定 Google 第一页",
            "本次数量：固定 10 个对象",
            "自然结果：按 SerpApi 当前页实际返回的 link 全部处理",
            "断点缓存：后台自动保存",
        ]

        if config["sendEmail"]:
            lines.append(f"邮件发送：流程结束后发送到 {config['email']}")
        else:
            lines.append("邮件发送：不发送")

        if config["promotionExecuteSend"]:
            lines.append(f"推广邮件：真实发送，固定发件邮箱 {self.baseConfig.get('promotionSenderEmail', '')}")
        else:
            lines.append("推广邮件：只生成后台发送记录，不真实发送")

        if self.modeVar.get() == "config":
            lines.append("当前页面：配置页。请切换公司模式或个人模式后开始运行。")
        elif self.modeVar.get() == "person":
            lines.append(f"运行模式：个人模式，到期月份 {config['expireMonths']}。")
            lines.append("搜索对象：Active、无挂靠公司、未来指定月份内到期的个人。")
        else:
            lines.append("运行模式：公司模式。")
            lines.append("搜索对象：有挂靠公司信息的数据，按公司名去重。")

        lines.append("说明：每个搜索对象会调用 1 次 SerpApi，并用浏览器打开返回 link 做二级页面提取。")
        return "\n".join(lines)

    def refreshPlanInfo(self):
        """刷新运行说明。"""
        if hasattr(self, "planVar"):
            self.planVar.set(self.buildRunPlanText())

    def getText(self, key):
        """读取文本配置。"""
        return self.configVars[key].get().strip()

    def getInt(self, key):
        """读取整数配置。"""
        text = self.getText(key)
        try:
            return int(text)
        except ValueError as error:
            raise ValueError(f"{key} 必须是整数，当前值：{text}") from error

    def browseDir(self, key):
        """选择目录并写回输入框。"""
        currentText = self.getText(key)
        currentPath = Path(currentText) if currentText else self.baseDir
        initialDir = str(currentPath if currentPath.is_absolute() else self.baseDir / currentPath)
        selectedDir = filedialog.askdirectory(initialdir=initialDir)
        if selectedDir:
            self.configVars[key].set(selectedDir)
            self.refreshPlanInfo()

    def saveConfig(self):
        """保存 GUI 配置到 run_config.json。"""
        config = self.buildConfig(withCallback=False)
        if config is None:
            return

        # 不保存运行模式、固定密钥和固定发信凭据。
        for key in [
            "runMode", "serpapiUrl", "serpapiKey", "sender_email", "smtp_auth_code",
            "promotionSenderEmail", "promotionRecordFileName",
        ]:
            config.pop(key, None)

        self.configPath.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        self.statusVar.set("配置已保存")
        self.refreshPlanInfo()
        messagebox.showinfo("保存成功", f"配置已保存到：\n{self.configPath}")

    def startTask(self):
        """启动后台线程执行主流程。"""
        if self.isRunning:
            messagebox.showwarning("运行中", "当前任务仍在运行，请等待完成。")
            return
        if self.modeVar.get() == "config":
            messagebox.showwarning("请选择运行模式", "当前是配置页，请先切换公司模式或个人模式。")
            return

        config = self.buildConfig(withCallback=True)
        if config is None:
            return

        self.logText.delete("1.0", tk.END)
        self.logText.insert(tk.END, self.buildRunPlanText(config) + "\n\n")
        self.isRunning = True
        self.startButton.configure(state=tk.DISABLED)
        self.statusVar.set("运行中")

        self.workerThread = threading.Thread(target=self.runTask, args=(config,), daemon=True)
        self.workerThread.start()

    def runTask(self, config):
        """后台执行主流程，并把日志转发到窗口。"""
        oldStdout = sys.stdout
        oldStderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            print("开始运行 TREC SerpApi 搜索流程。")
            Main(config).run()
            print("主流程运行完成。")
            self.logQueue.put(("done", "完成"))
        except Exception:
            traceback.print_exc()
            self.logQueue.put(("done", "异常结束"))
        finally:
            sys.stdout = oldStdout
            sys.stderr = oldStderr

    def write(self, text):
        """作为 stdout/stderr 接收日志文本。"""
        if text:
            self.logQueue.put(("text", text))
        return len(text)

    def flush(self):
        """兼容 stdout/stderr flush 调用。"""
        return None

    def pollLog(self):
        """定时刷新后台日志。"""
        try:
            while True:
                messageType, messageText = self.logQueue.get_nowait()
                if messageType == "text":
                    self.logText.insert(tk.END, messageText)
                    self.logText.see(tk.END)
                elif messageType == "done":
                    self.isRunning = False
                    self.startButton.configure(state=tk.NORMAL)
                    self.statusVar.set(messageText)
        except queue.Empty:
            pass
        self.window.after(200, self.pollLog)

    def closeWindow(self):
        """关闭窗口前检查后台任务状态。"""
        if self.isRunning:
            allowed = messagebox.askyesno("任务仍在运行", "当前任务仍在运行，确定要关闭吗？", parent=self.window)
            if not allowed:
                return
        self.window.destroy()

    def run(self):
        """启动 GUI。"""
        self.window.mainloop()


if __name__ == "__main__":
    RunGui().run()
