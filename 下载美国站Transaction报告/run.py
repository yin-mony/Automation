"""下载 Transaction 报告统一 GUI 入口。"""

import queue
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from main import TestPage


class RunGui:
    """易得客 Transaction 报告请求 GUI"""

    class TransactionPane:
        """Transaction 报告请求标签区"""

        def __init__(self, owner, parent):
            # 外层 GUI 引用
            self.owner = owner
            # 流程容器
            self.parent = parent
            # 当前流程运行状态
            self.statusVar = tk.StringVar(value="待运行")
            self.build()

        def build(self):
            """创建 Transaction 报告请求配置区"""
            form = ttk.LabelFrame(self.parent, text="易得客流程配置（请求 Transaction 报告）", padding=12)
            form.pack(fill=tk.X, padx=12, pady=12)

            ttk.Label(form, text="易得客账号").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.usernameVar).grid(row=0, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="易得客密码").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            passwordFrame = ttk.Frame(form)
            passwordFrame.grid(row=1, column=1, sticky=tk.EW, pady=6)
            passwordEntry = ttk.Entry(passwordFrame, textvariable=self.owner.passwordVar, show="*")
            passwordEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            passwordButton = ttk.Button(passwordFrame, text="显示", width=6)
            passwordButton.config(command=lambda: self.owner.togglePassword(passwordEntry, passwordButton))
            passwordButton.pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="店铺站点").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            shopSiteCombo = ttk.Combobox(
                form,
                textvariable=self.owner.autoSiteNameVar,
                values=self.owner.amazonSiteNames,
                state="readonly",
            )
            shopSiteCombo.grid(row=2, column=1, sticky=tk.EW, pady=6)
            shopSiteCombo.bind(
                "<<ComboboxSelected>>",
                self.owner.syncAmazonSiteSelection,
            )

            ttk.Label(form, text="店铺 IP").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopIpVar).grid(row=3, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="店铺端口").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopPortVar).grid(row=4, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 邮箱").grid(row=5, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.amazonEmailVar).grid(row=5, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 密码").grid(row=6, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.amazonPasswordVar, show="*").grid(
                row=6,
                column=1,
                sticky=tk.EW,
                pady=6,
            )

            ttk.Label(form, text="Amazon 后台站点（可多选）").grid(
                row=7,
                column=0,
                sticky=tk.NW,
                padx=(0, 8),
                pady=6,
            )
            amazonSiteFrame = ttk.Frame(form)
            amazonSiteFrame.grid(row=7, column=1, sticky=tk.EW, pady=6)
            for siteIndex, siteName in enumerate(self.owner.amazonSiteNames):
                selectedVar = tk.BooleanVar(value=siteName == "美国")
                displayVar = tk.StringVar(
                    value=f"{'☑' if selectedVar.get() else '☐'} {siteName}"
                )
                self.owner.amazonSiteVars[siteName] = selectedVar
                self.owner.amazonSiteTextVars[siteName] = displayVar
                ttk.Checkbutton(
                    amazonSiteFrame,
                    textvariable=displayVar,
                    variable=selectedVar,
                    style="Checkmark.TCheckbutton",
                    command=lambda name=siteName: self.owner.updateAmazonSiteCheckText(name),
                ).grid(
                    row=siteIndex // 4,
                    column=siteIndex % 4,
                    sticky=tk.W,
                    padx=(0, 18),
                    pady=3,
                )
            for columnIndex in range(4):
                amazonSiteFrame.columnconfigure(columnIndex, weight=1)

            helpText = "店铺站点用于易得客入口；Amazon 登录信息仅在出现登录页时使用；所选后台站点将按列表顺序逐个处理。"
            ttk.Label(
                form,
                text=helpText,
                foreground="#6b4e00",
                wraplength=720,
                justify=tk.LEFT,
            ).grid(
                row=8,
                column=0,
                columnspan=2,
                sticky=tk.W,
                pady=(4, 0),
            )

            form.columnconfigure(1, weight=1)

            actions = ttk.Frame(self.parent)
            actions.pack(fill=tk.X, padx=12, pady=(0, 8))
            self.startButton = ttk.Button(actions, text="开始请求 Transaction 报告", command=self.owner.startTransaction)
            self.startButton.pack(side=tk.LEFT)
            ttk.Button(actions, text="清空日志", command=self.owner.clearTransactionLog).pack(side=tk.LEFT, padx=(8, 0))
            ttk.Label(actions, textvariable=self.statusVar).pack(side=tk.RIGHT)

            logBox = ttk.LabelFrame(self.parent, text="Transaction 运行日志", padding=8)
            logBox.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
            self.logText = scrolledtext.ScrolledText(logBox, height=14, wrap=tk.WORD)
            self.logText.pack(fill=tk.BOTH, expand=True)

    def __init__(self, root):
        # Tk 根窗口
        self.root = root
        self.root.title("下载 Transaction 报告")
        self.root.geometry("900x820")
        self.root.minsize(780, 720)
        # 隐藏系统主题自带的复选框指示器，统一使用 ☐ / ☑ 显示选中状态。
        style = ttk.Style(self.root)
        style.layout(
            "Checkmark.TCheckbutton",
            [
                (
                    "Checkbutton.padding",
                    {
                        "sticky": tk.NSEW,
                        "children": [
                            (
                                "Checkbutton.focus",
                                {
                                    "sticky": tk.NSEW,
                                    "children": [
                                        ("Checkbutton.label", {"sticky": tk.NSEW})
                                    ],
                                },
                            )
                        ],
                    },
                )
            ],
        )
        style.configure("Checkmark.TCheckbutton", padding=(2, 2))
        # 日志与线程状态
        self.logQueue = queue.Queue()
        self.worker = None
        self.currentLogText = None
        self.currentStartButton = None
        self.currentStatusVar = None
        # 公共配置
        self.isOnlineVar = tk.StringVar(value="offline")
        self.sendEmailVar = tk.StringVar(value="no")
        self.emailVar = tk.StringVar(value="")
        self.emailEntry = None
        # Amazon 后台站点选项：中文用于界面选择，英文用于 Seller Central 账号切换
        self.amazonSiteMap = {
            "美国": "United States",
            "加拿大": "Canada",
            "墨西哥": "Mexico",
            "巴西": "Brazil",
            "英国": "United Kingdom",
            "法国": "France",
            "德国": "Germany",
            "意大利": "Italy",
            "西班牙": "Spain",
            "荷兰": "Netherlands",
            "瑞典": "Sweden",
            "波兰": "Poland",
            "比利时": "Belgium",
            "爱尔兰": "Ireland",
            "日本": "Japan",
            "新加坡": "Singapore",
            "澳大利亚": "Australia",
            "印度": "India",
            "阿联酋": "United Arab Emirates",
            "沙特阿拉伯": "Saudi Arabia",
            "土耳其": "Turkey",
            "埃及": "Egypt",
            "南非": "South Africa",
        }
        self.amazonSiteNames = list(self.amazonSiteMap.keys())
        # Transaction 流程配置
        self.usernameVar = tk.StringVar(
            value=os.getenv("YIDEKE_USERNAME", "13281439638")
        )
        self.passwordVar = tk.StringVar(
            value=os.getenv("YIDEKE_PASSWORD", "13281439638@MM")
        )
        self.shopIpVar = tk.StringVar(value="54.70.92.80")
        self.shopPortVar = tk.StringVar(value="9527")
        self.autoSiteNameVar = tk.StringVar(value="美国")
        self.amazonSiteVars = {}
        self.amazonSiteTextVars = {}
        self.amazonEmailVar = tk.StringVar(value="")
        self.amazonPasswordVar = tk.StringVar(value="")

        self.buildUi()
        self.pollLog()

    def buildUi(self):
        """创建统一 GUI 界面"""
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        self.buildCommon(container)

        flowFrame = ttk.Frame(container)
        flowFrame.pack(fill=tk.BOTH, expand=True, pady=(12, 0))
        self.transactionPane = self.TransactionPane(self, flowFrame)
        self.toggleEmail()

    def buildCommon(self, parent):
        """创建公共配置区"""
        common = ttk.LabelFrame(parent, text="公共配置", padding=12)
        common.pack(fill=tk.X)

        tipText = "本流程按顺序请求所选 Amazon 后台站点上一个自然月的 Transaction 报告；全部站点点击“请求报告”后完成。"
        ttk.Label(
            common,
            text=tipText,
            foreground="#6b4e00",
            wraplength=720,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        ttk.Label(common, text="运行环境").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        envFrame = ttk.Frame(common)
        envFrame.grid(row=1, column=1, sticky=tk.W, pady=6)
        ttk.Radiobutton(envFrame, text="线下", variable=self.isOnlineVar, value="offline").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(envFrame, text="线上", variable=self.isOnlineVar, value="online").pack(side=tk.LEFT)

        ttk.Label(common, text="邮件通知").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        mailFrame = ttk.Frame(common)
        mailFrame.grid(row=2, column=1, sticky=tk.W, pady=6)
        ttk.Radiobutton(mailFrame, text="不发送", variable=self.sendEmailVar, value="no", command=self.toggleEmail).pack(
            side=tk.LEFT,
            padx=(0, 12),
        )
        ttk.Radiobutton(mailFrame, text="发送", variable=self.sendEmailVar, value="yes", command=self.toggleEmail).pack(
            side=tk.LEFT,
        )

        ttk.Label(common, text="接收邮箱").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.emailEntry = ttk.Entry(common, textvariable=self.emailVar)
        self.emailEntry.grid(row=3, column=1, sticky=tk.EW, pady=6)
        common.columnconfigure(1, weight=1)

    def togglePassword(self, entry, button):
        """切换密码输入框的显示与隐藏状态"""
        if entry.cget("show") == "*":
            entry.config(show="")
            button.config(text="隐藏")
        else:
            entry.config(show="*")
            button.config(text="显示")

    def toggleEmail(self):
        """根据邮件开关启用或禁用邮箱输入框"""
        if not self.emailEntry:
            return
        state = tk.NORMAL if self.sendEmailVar.get() == "yes" else tk.DISABLED
        self.emailEntry.config(state=state)

    def syncAmazonSiteSelection(self, event=None):
        """店铺站点变化时，将 Amazon 复选框同步为对应的单个站点。"""
        if not self.amazonSiteVars:
            return
        siteName = self.autoSiteNameVar.get().strip()
        if siteName not in self.amazonSiteNames:
            return
        for optionSiteName, selectedVar in self.amazonSiteVars.items():
            selectedVar.set(optionSiteName == siteName)
            self.updateAmazonSiteCheckText(optionSiteName)

    def updateAmazonSiteCheckText(self, siteName):
        """根据站点选中状态显示空方框或勾选方框。"""
        selectedVar = self.amazonSiteVars[siteName]
        displayVar = self.amazonSiteTextVars[siteName]
        displayVar.set(f"{'☑' if selectedVar.get() else '☐'} {siteName}")

    def parseList(self, value):
        """将逗号/换行分隔的文本拆成列表"""
        items = []
        for item in value.replace("\n", ",").split(","):
            item = item.strip()
            if item:
                items.append(item)
        return items

    def buildCommonConfig(self):
        """读取公共配置并做基础校验"""
        sendEmail = self.sendEmailVar.get() == "yes"
        email = self.emailVar.get().strip()
        return {
            "isOnline": self.isOnlineVar.get() == "online",
            "sendEmail": sendEmail,
            "email": email,
        }

    def buildTransactionConfig(self):
        """校验 Transaction 流程配置并组装 TestPage 入参"""
        common = self.buildCommonConfig()
        username = self.usernameVar.get().strip()
        password = self.passwordVar.get().strip()
        autoSiteName = self.autoSiteNameVar.get().strip()
        ips = self.parseList(self.shopIpVar.get())
        portValues = self.parseList(self.shopPortVar.get())
        amazonSiteNames = [
            siteName for siteName in self.amazonSiteNames
            if self.amazonSiteVars[siteName].get()
        ]
        amazonEmail = self.amazonEmailVar.get().strip()
        amazonPassword = self.amazonPasswordVar.get()

        if not username:
            raise ValueError("请填写易得客账号")
        if not password:
            raise ValueError("请填写易得客密码")
        if not autoSiteName:
            raise ValueError("请选择店铺站点")
        if autoSiteName not in self.amazonSiteNames:
            raise ValueError(f"暂不支持该店铺站点: {autoSiteName}")
        if not ips:
            raise ValueError("请填写店铺 IP")
        if not portValues:
            raise ValueError("请填写店铺端口")
        if not amazonSiteNames:
            raise ValueError("请至少选择一个 Amazon 后台站点")
        unsupportedSiteNames = [
            siteName for siteName in amazonSiteNames
            if siteName not in self.amazonSiteNames
        ]
        if unsupportedSiteNames:
            raise ValueError(
                f"暂不支持以下 Amazon 后台站点: {', '.join(unsupportedSiteNames)}"
            )
        if common["sendEmail"] and not common["email"]:
            raise ValueError("选择发送邮件时必须填写接收邮箱")

        try:
            ports = [int(port) for port in portValues]
        except ValueError as exc:
            raise ValueError("店铺端口只能填写数字，多个端口用逗号或换行分隔") from exc
        if len(ports) == 1 and len(ips) > 1:
            ports = ports * len(ips)
        if len(ports) != len(ips):
            raise ValueError("端口数量需要和 IP 数量一致，或只填写一个端口")

        config = {
            "username": username,
            "password": password,
            "autoSiteName": autoSiteName,
            "ip": ips,
            "port": ports,
            "amazonSiteNames": amazonSiteNames,
            "amazonEmail": amazonEmail,
            "amazonPassword": amazonPassword,
        }
        config.update(common)
        return config

    def isRunning(self):
        """判断当前是否已有流程在运行"""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "已有流程正在运行中")
            return True
        return False

    def startTransaction(self):
        """启动 Transaction 报告请求后台线程"""
        if self.isRunning():
            return
        try:
            config = self.buildTransactionConfig()
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.currentLogText = self.transactionPane.logText
        self.currentStartButton = self.transactionPane.startButton
        self.currentStatusVar = self.transactionPane.statusVar
        self.transactionPane.startButton.config(state=tk.DISABLED)
        self.transactionPane.statusVar.set("运行中")
        self.appendLog(self.transactionPane.logText, "开始请求 Transaction 报告...\n")
        self.appendLog(self.transactionPane.logText, f"运行环境: {'线上' if config.get('isOnline') else '线下'}\n")
        self.appendLog(self.transactionPane.logText, f"店铺站点: {config.get('autoSiteName')}\n")
        amazonSiteNames = config.get("amazonSiteNames") or []
        amazonSiteText = "、".join(
            f"{siteName} / {self.amazonSiteMap.get(siteName, siteName)}"
            for siteName in amazonSiteNames
        )
        self.appendLog(self.transactionPane.logText, f"Amazon 后台站点: {amazonSiteText}\n")
        self.appendLog(self.transactionPane.logText, f"店铺 IP: {', '.join(config.get('ip') or [])}\n")
        self.appendLog(self.transactionPane.logText, f"店铺端口: {', '.join(str(port) for port in config.get('port') or [])}\n")
        self.appendLog(self.transactionPane.logText, f"邮件通知: {'发送' if config.get('sendEmail') else '不发送'}\n")

        self.worker = threading.Thread(target=self.runTransactionTask, args=(config,), daemon=True)
        self.worker.start()

    def runTransactionTask(self, config):
        """在线程中执行 Transaction 请求流程并回写日志"""
        oldStdout = sys.stdout
        oldStderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            TestPage(config).run()
            self.logQueue.put((self.currentLogText, "\nTransaction 报告请求任务完成。\n"))
            self.root.after(0, lambda: self.currentStatusVar.set("已完成"))
        except Exception:
            self.logQueue.put((self.currentLogText, "\nTransaction 报告请求运行失败：\n"))
            self.logQueue.put((self.currentLogText, traceback.format_exc()))
            self.root.after(0, lambda: self.currentStatusVar.set("运行失败"))
        finally:
            sys.stdout = oldStdout
            sys.stderr = oldStderr
            self.root.after(0, lambda: self.currentStartButton.config(state=tk.NORMAL))

    def pollLog(self):
        """轮询日志队列并刷新到日志框"""
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

    def clearTransactionLog(self):
        """清空 Transaction 流程日志"""
        self.transactionPane.logText.delete("1.0", tk.END)

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
