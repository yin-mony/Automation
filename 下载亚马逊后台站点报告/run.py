"""下载亚马逊后台站点报告统一 GUI 入口。"""

import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from main import AmazonReport, reportTypes, siteMap, siteModes


class RunGui:
    """亚马逊后台站点报告请求 GUI"""

    class AutoPane:
        """易得客流程标签页"""

        def __init__(self, owner, parent):
            # 外层 GUI 引用
            self.owner = owner
            # 标签页容器
            self.parent = parent
            # 运行状态
            self.statusVar = tk.StringVar(value="待运行")
            # 创建易得客流程界面
            self.build()

        def build(self):
            """创建易得客流程标签页控件"""
            form = ttk.LabelFrame(self.parent, text="易得客流程配置（请求亚马逊后台站点报告）", padding=12)
            form.pack(fill=tk.X, padx=12, pady=12)

            ttk.Label(form, text="易得客账号").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.yidekeUsernameVar).grid(row=0, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="易得客密码").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            yidekePasswordFrame = ttk.Frame(form)
            yidekePasswordFrame.grid(row=1, column=1, sticky=tk.EW, pady=6)
            yidekePasswordEntry = ttk.Entry(yidekePasswordFrame, textvariable=self.owner.yidekePasswordVar, show="*")
            yidekePasswordEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            yidekePasswordButton = ttk.Button(yidekePasswordFrame, text="显示", width=6)
            yidekePasswordButton.config(
                command=lambda: self.owner.togglePassword(yidekePasswordEntry, yidekePasswordButton),
            )
            yidekePasswordButton.pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="店铺站点").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            siteCombo = ttk.Combobox(
                form,
                textvariable=self.owner.autoSiteNameVar,
                values=self.owner.autoSiteNames,
                state="readonly",
            )
            siteCombo.grid(row=2, column=1, sticky=tk.EW, pady=6)
            siteCombo.bind("<<ComboboxSelected>>", lambda event: self.owner.syncAmazonSiteFromAuto())

            ttk.Label(form, text="店铺 IP").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopIpVar).grid(row=3, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="店铺端口").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopPortVar).grid(row=4, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 邮箱").grid(row=5, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.amazonEmailVar).grid(row=5, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 密码").grid(row=6, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            amazonPasswordFrame = ttk.Frame(form)
            amazonPasswordFrame.grid(row=6, column=1, sticky=tk.EW, pady=6)
            amazonPasswordEntry = ttk.Entry(amazonPasswordFrame, textvariable=self.owner.amazonPasswordVar, show="*")
            amazonPasswordEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            amazonPasswordButton = ttk.Button(amazonPasswordFrame, text="显示", width=6)
            amazonPasswordButton.config(
                command=lambda: self.owner.togglePassword(amazonPasswordEntry, amazonPasswordButton),
            )
            amazonPasswordButton.pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="Amazon 后台站点").grid(row=7, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            self.owner.amazonSiteCombo = ttk.Combobox(
                form,
                textvariable=self.owner.amazonSiteNameVar,
                values=self.owner.autoSiteNames,
                state="readonly",
            )
            self.owner.amazonSiteCombo.grid(row=7, column=1, sticky=tk.EW, pady=6)
            form.columnconfigure(1, weight=1)

            reportForm = ttk.LabelFrame(self.parent, text="报告请求配置", padding=12)
            reportForm.pack(fill=tk.X, padx=12, pady=(0, 12))

            ttk.Label(reportForm, text="报告类型").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            reportCombo = ttk.Combobox(
                reportForm,
                textvariable=self.owner.reportTypeVar,
                values=self.owner.reportTypeLabels,
                state="readonly",
            )
            reportCombo.grid(row=0, column=1, sticky=tk.EW, pady=6)

            ttk.Label(reportForm, text="站点切换模式").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            siteModeFrame = ttk.Frame(reportForm)
            siteModeFrame.grid(row=1, column=1, sticky=tk.W, pady=6)
            ttk.Radiobutton(
                siteModeFrame,
                text=siteModes["single"],
                variable=self.owner.siteModeVar,
                value="single",
                command=self.owner.syncSiteMode,
            ).pack(side=tk.LEFT, padx=(0, 12))
            ttk.Radiobutton(
                siteModeFrame,
                text=siteModes["all"],
                variable=self.owner.siteModeVar,
                value="all",
                command=self.owner.syncSiteMode,
            ).pack(side=tk.LEFT)

            ttk.Label(reportForm, text="全站点范围").grid(row=2, column=0, sticky=tk.NW, padx=(0, 8), pady=6)
            allSiteBox = ttk.Frame(reportForm)
            allSiteBox.grid(row=2, column=1, sticky=tk.EW, pady=6)
            self.owner.buildAllSiteChecks(allSiteBox)
            reportForm.columnconfigure(1, weight=1)

            actions = ttk.Frame(self.parent)
            actions.pack(fill=tk.X, padx=12, pady=(0, 8))
            self.startButton = ttk.Button(actions, text="开始运行易得客流程", command=self.owner.startAuto)
            self.startButton.pack(side=tk.LEFT)
            ttk.Button(actions, text="清空日志", command=self.owner.clearAutoLog).pack(side=tk.LEFT, padx=(8, 0))
            ttk.Label(actions, textvariable=self.statusVar).pack(side=tk.RIGHT)

            logBox = ttk.LabelFrame(self.parent, text="易得客运行日志", padding=8)
            logBox.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
            self.logText = scrolledtext.ScrolledText(logBox, height=14, wrap=tk.WORD)
            self.logText.pack(fill=tk.BOTH, expand=True)
            # 易得客标签页控件已创建完成

    def __init__(self, root):
        # Tk 根窗口
        self.root = root
        self.root.title("下载亚马逊后台站点报告")
        self.root.geometry("980x860")
        self.root.minsize(900, 760)
        # 日志与线程状态
        self.logQueue = queue.Queue()
        self.worker = None
        self.currentLogText = None
        self.currentStartButton = None
        self.currentStatusVar = None
        # 易得客与 Amazon 后台站点选项映射，中文用于界面选择，英文用于 Seller Central 账号切换
        self.autoSiteMap = siteMap
        self.autoSiteNames = list(self.autoSiteMap.keys())
        self.reportTypeLabels = [reportTypes[key]["label"] for key in reportTypes]
        self.siteCheckVars = {}
        self.allSiteWidgets = []
        self.amazonSiteCombo = None
        # 公共表单变量
        self.isOnlineVar = tk.StringVar(value="offline")
        # 易得客流程表单变量
        self.yidekeUsernameVar = tk.StringVar(value="")
        self.yidekePasswordVar = tk.StringVar(value="")
        self.autoSiteNameVar = tk.StringVar(value="美国")
        self.amazonSiteNameVar = tk.StringVar(value="美国")
        self.shopIpVar = tk.StringVar(value="")
        self.shopPortVar = tk.StringVar(value="8888")
        self.amazonEmailVar = tk.StringVar(value="")
        self.amazonPasswordVar = tk.StringVar(value="")
        # 报告请求配置变量
        self.siteModeVar = tk.StringVar(value="single")
        self.reportTypeVar = tk.StringVar(value=reportTypes["summary"]["label"])

        self.buildUi()
        self.pollLog()

    def buildUi(self):
        """创建统一 GUI 界面"""
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        self.buildCommon(container)
        tabs = ttk.Notebook(container)
        tabs.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        autoFrame = ttk.Frame(tabs)
        tabs.add(autoFrame, text="易得客流程")

        self.autoPane = self.AutoPane(self, autoFrame)
        self.syncSiteMode()
        # 统一界面已创建完成

    def buildCommon(self, parent):
        """创建流程共用配置区"""
        common = ttk.LabelFrame(parent, text="公共配置", padding=12)
        common.pack(fill=tk.X)

        tipText = "本流程会先处理易得客进店与 Amazon 后台登录，再优先切换中文页面语言并请求所选报告。"
        ttk.Label(common, text=tipText, foreground="#6b4e00").grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        ttk.Label(common, text="运行环境").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        envFrame = ttk.Frame(common)
        envFrame.grid(row=1, column=1, sticky=tk.W, pady=6)
        ttk.Radiobutton(envFrame, text="线下", variable=self.isOnlineVar, value="offline").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(envFrame, text="线上", variable=self.isOnlineVar, value="online").pack(side=tk.LEFT)
        common.columnconfigure(1, weight=1)

    def togglePassword(self, entry, button):
        """切换密码输入框的显示与隐藏状态"""
        if entry.cget("show") == "*":
            entry.config(show="")
            button.config(text="隐藏")
        else:
            entry.config(show="*")
            button.config(text="显示")

    def buildAllSiteChecks(self, parent):
        """创建全站点自动切换范围多选区。"""
        buttonFrame = ttk.Frame(parent)
        buttonFrame.grid(row=0, column=0, columnspan=5, sticky=tk.W, pady=(0, 6))
        selectButton = ttk.Button(buttonFrame, text="全选", width=8, command=self.selectAllSites)
        selectButton.pack(side=tk.LEFT)
        clearButton = ttk.Button(buttonFrame, text="清空", width=8, command=self.clearAllSites)
        clearButton.pack(side=tk.LEFT, padx=(8, 0))
        self.allSiteWidgets.extend([selectButton, clearButton])
        for index, siteName in enumerate(self.autoSiteNames):
            var = tk.BooleanVar(value=True)
            self.siteCheckVars[siteName] = var
            check = tk.Checkbutton(
                parent,
                text=siteName,
                variable=var,
                anchor=tk.W,
                padx=0,
                pady=0,
                highlightthickness=0,
            )
            check.grid(row=1 + index // 5, column=index % 5, sticky=tk.W, padx=(0, 18), pady=2)
            self.allSiteWidgets.append(check)
        for column in range(5):
            parent.columnconfigure(column, weight=1)

    def syncAmazonSiteFromAuto(self):
        """按 FBA 界面习惯，把店铺站点同步到 Amazon 后台站点。"""
        if self.siteModeVar.get() == "single":
            self.amazonSiteNameVar.set(self.autoSiteNameVar.get())

    def syncSiteMode(self):
        """根据站点切换模式启用对应控件。"""
        isSingle = self.siteModeVar.get() == "single"
        if self.amazonSiteCombo:
            self.amazonSiteCombo.configure(state="readonly" if isSingle else tk.DISABLED)
        allSiteState = tk.DISABLED if isSingle else tk.NORMAL
        for widget in self.allSiteWidgets:
            widget.configure(state=allSiteState)

    def selectAllSites(self):
        """全选全站点自动切换范围。"""
        for var in self.siteCheckVars.values():
            var.set(True)

    def clearAllSites(self):
        """清空全站点自动切换范围。"""
        for var in self.siteCheckVars.values():
            var.set(False)

    def parseList(self, value):
        """将逗号/换行分隔的文本拆成列表"""
        items = []
        for item in value.replace("\n", ",").split(","):
            item = item.strip()
            if item:
                items.append(item)
        return items

    def buildAutoConfig(self):
        """校验易得客配置并组装 AmazonReport 入参"""
        yidekeUsername = self.yidekeUsernameVar.get().strip()
        yidekePassword = self.yidekePasswordVar.get()
        autoSiteName = self.autoSiteNameVar.get().strip()
        amazonSiteName = self.amazonSiteNameVar.get().strip()
        shopIps = self.parseList(self.shopIpVar.get())
        shopPorts = self.parseList(self.shopPortVar.get())
        amazonEmail = self.amazonEmailVar.get().strip()
        amazonPassword = self.amazonPasswordVar.get()
        reportType = self.reportTypeVar.get().strip()
        siteMode = self.siteModeVar.get()
        selectedSites = [siteName for siteName, var in self.siteCheckVars.items() if var.get()]

        if not autoSiteName:
            raise ValueError("请选择店铺站点")
        if autoSiteName not in self.autoSiteNames:
            raise ValueError(f"暂不支持该店铺站点: {autoSiteName}")
        if siteMode == "single":
            if not amazonSiteName:
                raise ValueError("请选择 Amazon 后台站点")
            if amazonSiteName not in self.autoSiteNames:
                raise ValueError(f"暂不支持该 Amazon 后台站点: {amazonSiteName}")
        if siteMode == "all" and not selectedSites:
            raise ValueError("全站点自动切换模式下至少选择一个站点")
        if not shopIps:
            raise ValueError("请填写店铺 IP")
        if not shopPorts:
            raise ValueError("请填写店铺端口")
        if not yidekeUsername:
            raise ValueError("请填写易得客账号")
        if not yidekePassword:
            raise ValueError("请填写易得客密码")
        if reportType not in self.reportTypeLabels:
            raise ValueError(f"暂不支持该报告类型: {reportType}")

        config = {
            # 易得客登录账号密码
            "username": yidekeUsername,
            "password": yidekePassword,
            "yidekeUsername": yidekeUsername,
            "yidekePassword": yidekePassword,
            # 店铺站点只影响易得客进店访问
            "autoSiteName": autoSiteName,
            # 店铺 IP 与调试端口用于接管易得客浏览器
            "ip": shopIps,
            "port": shopPorts,
            "shopIp": shopIps,
            "shopPort": shopPorts,
            # Amazon Seller Central 登录账号密码
            "amazonEmail": amazonEmail,
            "amazonPassword": amazonPassword,
            # Amazon 后台站点与报告请求配置
            "amazonSiteName": amazonSiteName,
            "amazonSiteNames": selectedSites if siteMode == "all" else [amazonSiteName],
            "siteMode": siteMode,
            "reportType": reportType,
            "isOnline": self.isOnlineVar.get() == "online",
        }
        return config

    def isRunning(self):
        """判断当前是否已有流程在运行"""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "已有流程正在运行中")
            return True
        return False

    def startAuto(self):
        """启动易得客流程后台线程"""
        if self.isRunning():
            return
        try:
            config = self.buildAutoConfig()
            # 通过主流程类做二次校验，保持 GUI 与 CLI 入口一致
            AmazonReport(config).validate()
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.currentLogText = self.autoPane.logText
        self.currentStartButton = self.autoPane.startButton
        self.currentStatusVar = self.autoPane.statusVar
        self.autoPane.startButton.config(state=tk.DISABLED)
        self.autoPane.statusVar.set("运行中")
        self.appendLog(self.autoPane.logText, "开始运行易得客流程...\n")
        self.appendLog(self.autoPane.logText, f"运行环境: {'线上' if config.get('isOnline') else '线下'}\n")
        self.appendLog(self.autoPane.logText, f"店铺站点: {config.get('autoSiteName')}\n")
        siteText = config.get("amazonSiteName")
        if config.get("siteMode") == "all":
            siteText = ", ".join(config.get("amazonSiteNames") or [])
        self.appendLog(self.autoPane.logText, f"Amazon后台站点: {siteText}\n")
        self.appendLog(self.autoPane.logText, f"站点切换模式: {siteModes.get(config.get('siteMode'))}\n")
        self.appendLog(self.autoPane.logText, f"报告类型: {config.get('reportType')}\n")
        self.appendLog(self.autoPane.logText, f"Amazon 邮箱: {config.get('amazonEmail') or '未填写'}\n")
        self.appendLog(self.autoPane.logText, f"店铺 IP: {', '.join(config.get('ip') or [])}\n")
        self.appendLog(self.autoPane.logText, f"店铺端口: {', '.join(str(port) for port in config.get('port') or [])}\n")
        # 后台线程执行易得客与 Amazon 页面流程，避免 GUI 窗口卡住
        self.worker = threading.Thread(target=self.runAutoTask, args=(config,), daemon=True)
        self.worker.start()
        # 易得客流程后台线程已启动

    def runAutoTask(self, config):
        """在线程中执行易得客流程并回写日志"""
        oldStdout = sys.stdout
        oldStderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            AmazonReport(config).run()
            self.logQueue.put((self.currentLogText, "\n易得客流程任务完成。\n"))
            self.root.after(0, lambda: self.currentStatusVar.set("已完成"))
        except Exception:
            self.logQueue.put((self.currentLogText, "\n易得客流程运行失败：\n"))
            self.logQueue.put((self.currentLogText, traceback.format_exc()))
            self.root.after(0, lambda: self.currentStatusVar.set("运行失败"))
        finally:
            sys.stdout = oldStdout
            sys.stderr = oldStderr
            self.root.after(0, lambda: self.currentStartButton.config(state=tk.NORMAL))
            # 标准输出已恢复，开始按钮已恢复可用

    def pollLog(self):
        """轮询日志队列并刷新到对应日志框"""
        while True:
            try:
                logText, text = self.logQueue.get_nowait()
            except queue.Empty:
                break
            self.appendLog(logText, text)
        self.root.after(100, self.pollLog)

    def appendLog(self, logText, text):
        """向指定日志框追加文本"""
        if not logText:
            return
        logText.insert(tk.END, text)
        logText.see(tk.END)

    def clearAutoLog(self):
        """清空易得客流程日志"""
        self.autoPane.logText.delete("1.0", tk.END)

    def write(self, text):
        """接收 print 输出并放入当前流程日志队列"""
        if text:
            self.logQueue.put((self.currentLogText, text))

    def flush(self):
        """兼容标准输出 flush 接口"""
        return


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = RunGui(root)
    root.mainloop()
