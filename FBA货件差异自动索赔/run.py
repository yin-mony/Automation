"""FBA 货件差异自动索赔统一 GUI 入口。"""

import json
import queue
import re
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from DrissionPage import ChromiumPage
from tkcalendar import DateEntry

from main import Main


class RunGui:
    """赛狐 POP 导出与易得客 CASE 提交统一 GUI"""

    class SaihuPane:
        """赛狐流程标签页"""

        def __init__(self, owner, parent):
            # 外层 GUI 引用
            self.owner = owner
            # 标签页容器
            self.parent = parent
            # 运行状态
            self.statusVar = tk.StringVar(value="待运行")
            # 创建赛狐流程界面
            self.build()

        def build(self):
            """创建赛狐流程标签页控件"""
            form = ttk.LabelFrame(self.parent, text="赛狐流程配置（生成 POP 发票文件）", padding=12)
            form.pack(fill=tk.X, padx=12, pady=12)

            ttk.Label(form, text="赛狐账号").grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.usernameVar).grid(row=0, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="赛狐密码").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            passwordFrame = ttk.Frame(form)
            passwordFrame.grid(row=1, column=1, sticky=tk.EW, pady=6)
            passwordEntry = ttk.Entry(passwordFrame, textvariable=self.owner.passwordVar, show="*")
            passwordEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            passwordButton = ttk.Button(passwordFrame, text="显示", width=6)
            passwordButton.config(command=lambda: self.owner.togglePassword(passwordEntry, passwordButton))
            passwordButton.pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="导出目录").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            exportFrame = ttk.Frame(form)
            exportFrame.grid(row=2, column=1, sticky=tk.EW, pady=6)
            ttk.Entry(exportFrame, textvariable=self.owner.exportDirVar).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(exportFrame, text="浏览", command=self.owner.selectExportDir).pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="筛选站点").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            siteCombo = ttk.Combobox(
                form,
                textvariable=self.owner.siteNameVar,
                values=self.owner.siteNames,
                state="readonly",
            )
            siteCombo.grid(row=3, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="筛选店铺").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            shopCombo = ttk.Combobox(
                form,
                textvariable=self.owner.shopNameVar,
                values=self.owner.shopNames,
                state="readonly",
            )
            shopCombo.grid(row=4, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="筛选时间").grid(row=5, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            dateFrame = ttk.Frame(form)
            dateFrame.grid(row=5, column=1, sticky=tk.EW, pady=6)
            startDateText = self.owner.startDateVar.get()
            startDateEntry = DateEntry(
                dateFrame,
                textvariable=self.owner.startDateVar,
                date_pattern="yyyy-mm-dd",
                width=16,
            )
            try:
                startDateEntry.set_date(datetime.strptime(startDateText, "%Y-%m-%d").date())
            except ValueError:
                startDateEntry.set_date(datetime.strptime(self.owner.defaultStartDate, "%Y-%m-%d").date())
            startDateEntry.pack(side=tk.LEFT)
            ttk.Label(dateFrame, text="至").pack(side=tk.LEFT, padx=8)
            endDateText = self.owner.endDateVar.get()
            endDateEntry = DateEntry(
                dateFrame,
                textvariable=self.owner.endDateVar,
                date_pattern="yyyy-mm-dd",
                width=16,
            )
            try:
                endDateEntry.set_date(datetime.strptime(endDateText, "%Y-%m-%d").date())
            except ValueError:
                endDateEntry.set_date(datetime.strptime(self.owner.defaultEndDate, "%Y-%m-%d").date())
            endDateEntry.pack(side=tk.LEFT)

            ttk.Label(form, text="模板文件").grid(row=6, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            templateFrame = ttk.Frame(form)
            templateFrame.grid(row=6, column=1, sticky=tk.EW, pady=6)
            ttk.Entry(templateFrame, textvariable=self.owner.templatePathVar).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(templateFrame, text="选择模板", command=self.owner.selectTemplateFile).pack(side=tk.LEFT, padx=(8, 0))
            form.columnconfigure(1, weight=1)

            actions = ttk.Frame(self.parent)
            actions.pack(fill=tk.X, padx=12, pady=(0, 8))
            self.startButton = ttk.Button(actions, text="开始运行赛狐流程", command=self.owner.startSaihu)
            self.startButton.pack(side=tk.LEFT)
            ttk.Button(actions, text="清空日志", command=self.owner.clearSaihuLog).pack(side=tk.LEFT, padx=(8, 0))
            ttk.Label(actions, textvariable=self.statusVar).pack(side=tk.RIGHT)

            logBox = ttk.LabelFrame(self.parent, text="赛狐运行日志", padding=8)
            logBox.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
            self.logText = scrolledtext.ScrolledText(logBox, height=14, wrap=tk.WORD)
            self.logText.pack(fill=tk.BOTH, expand=True)
            # 赛狐标签页控件已创建完成

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
            form = ttk.LabelFrame(self.parent, text="易得客流程配置（上传 POP/POD 并提交 CASE）", padding=12)
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
            siteCombo.bind("<<ComboboxSelected>>", lambda event: self.owner.amazonSiteNameVar.set(self.owner.autoSiteNameVar.get()))

            ttk.Label(form, text="店铺 IP").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopIpVar).grid(row=3, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="店铺端口").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.shopPortVar).grid(row=4, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 邮箱").grid(row=5, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.amazonEmailVar).grid(row=5, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 密码").grid(row=6, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            ttk.Entry(form, textvariable=self.owner.amazonPasswordVar, show="*").grid(row=6, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="Amazon 后台站点").grid(row=7, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            amazonSiteCombo = ttk.Combobox(
                form,
                textvariable=self.owner.amazonSiteNameVar,
                values=self.owner.autoSiteNames,
                state="readonly",
            )
            amazonSiteCombo.grid(row=7, column=1, sticky=tk.EW, pady=6)

            ttk.Label(form, text="已完成导出存放 POP 文件的目录").grid(row=8, column=0, sticky=tk.W, padx=(0, 8), pady=6)
            popFrame = ttk.Frame(form)
            popFrame.grid(row=8, column=1, sticky=tk.EW, pady=6)
            ttk.Entry(popFrame, textvariable=self.owner.popDirVar).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(popFrame, text="浏览", command=self.owner.selectPopDir).pack(side=tk.LEFT, padx=(8, 0))
            ttk.Button(popFrame, text="刷新编号", command=self.owner.refreshAutoShipmentIds).pack(side=tk.LEFT, padx=(8, 0))

            ttk.Label(form, text="待处理货件编号").grid(row=9, column=0, sticky=tk.NW, padx=(0, 8), pady=6)
            self.owner.autoShipmentBox = scrolledtext.ScrolledText(form, height=4, wrap=tk.WORD, state=tk.DISABLED)
            self.owner.autoShipmentBox.grid(row=9, column=1, sticky=tk.EW, pady=6)
            form.columnconfigure(1, weight=1)

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
            self.owner.refreshAutoShipmentIds()
            # 易得客标签页控件已创建完成

    def __init__(self, root):
        # 主配置入口，负责提供 GUI 默认值并调用两个业务流程
        self.main = Main()
        # 项目根目录，由 Main 统一维护
        self.baseDir = self.main.baseDir
        # GUI 本地缓存配置文件，由 Main 统一维护
        self.configFile = self.main.configFile
        # 默认 POP 输出目录，由 Main 统一维护
        self.defaultExportDir = self.main.defaultExportDir
        # 默认 POP 模板文件，由 Main 统一维护
        self.defaultTemplatePath = self.main.defaultTemplatePath
        # 默认邮件接收邮箱，由 Main 统一维护
        self.defaultEmail = self.main.defaultEmail
        # 默认邮件发件邮箱，由 Main 统一维护
        self.defaultSenderEmail = self.main.defaultSenderEmail
        # 默认 SMTP 授权码，由 Main 统一维护
        self.defaultSmtpAuthCode = self.main.defaultSmtpAuthCode
        # 默认 SMTP 服务地址，由 Main 统一维护
        self.defaultSmtpServer = self.main.defaultSmtpServer
        # 默认 SMTP SSL 端口，由 Main 统一维护
        self.defaultSmtpPort = self.main.defaultSmtpPort
        # 默认企业微信 Webhook，由 Main 统一维护
        self.defaultWechatWebhook = self.main.defaultWechatWebhook
        # 默认企业微信 @ 手机号，由 Main 统一维护
        self.defaultWechatMobile = self.main.defaultWechatMobile
        # 默认赛狐账号，由 Main 统一维护
        self.defaultSaihuUsername = self.main.defaultSaihuUsername
        # 默认赛狐密码，由 Main 统一维护
        self.defaultSaihuPassword = self.main.defaultSaihuPassword
        # 默认赛狐筛选站点，由 Main 统一维护
        self.defaultSiteName = self.main.defaultSiteName
        # 默认赛狐店铺主体名，由 Main 统一维护
        self.defaultShopBaseName = self.main.defaultShopBaseName
        # 默认赛狐筛选开始时间，由 Main 统一维护
        self.defaultStartDate = self.main.defaultStartDate
        # 默认赛狐筛选结束时间，由 Main 统一维护
        self.defaultEndDate = self.main.defaultEndDate
        # 默认易得客账号，由 Main 统一维护
        self.defaultYidekeUsername = self.main.defaultYidekeUsername
        # 默认易得客密码，由 Main 统一维护
        self.defaultYidekePassword = self.main.defaultYidekePassword
        # 默认易得客店铺站点，由 Main 统一维护
        self.defaultAutoSiteName = self.main.defaultAutoSiteName
        # 默认 Amazon 后台站点，由 Main 统一维护
        self.defaultAmazonSiteName = self.main.defaultAmazonSiteName
        # 默认店铺 IP，由 Main 统一维护
        self.defaultShopIp = self.main.defaultShopIp
        # 默认店铺端口，由 Main 统一维护
        self.defaultShopPort = self.main.defaultShopPort
        # 默认 Amazon 邮箱，由 Main 统一维护
        self.defaultAmazonEmail = self.main.defaultAmazonEmail
        # 默认 Amazon 密码，由 Main 统一维护
        self.defaultAmazonPassword = self.main.defaultAmazonPassword
        # Tk 根窗口
        self.root = root
        self.root.title("FBA 货件差异自动索赔")
        self.root.geometry("980x860")
        self.root.minsize(860, 760)
        # 日志与线程状态
        self.logQueue = queue.Queue()
        self.worker = None
        self.currentLogText = None
        self.currentStartButton = None
        self.currentStatusVar = None
        self.selectingExportDir = False
        self.selectingPopDir = False
        # 赛狐 FBA 站点选择项，由 Main 统一维护
        self.siteNames = list(self.main.siteNames)
        # 赛狐店铺主体名选项，由 Main 统一维护
        self.shopNames = list(self.main.shopNames)
        # 赛狐店铺后缀映射，由 Main 统一维护
        self.siteShopSuffixMap = dict(self.main.siteShopSuffixMap)
        # 易得客与 Amazon 后台站点映射，由 Main 统一维护
        self.autoSiteMap = dict(self.main.autoSiteMap)
        # 易得客流程 GUI 站点下拉项，由 Main 统一维护
        self.autoSiteNames = list(self.main.autoSiteNames)
        # 公共表单变量
        self.isOnlineVar = tk.StringVar(value="offline")
        self.sendEmailVar = tk.StringVar(value="no")
        self.emailVar = tk.StringVar(value=self.defaultEmail)
        self.senderEmailVar = tk.StringVar(value=self.defaultSenderEmail)
        self.smtpAuthCodeVar = tk.StringVar(value=self.defaultSmtpAuthCode)
        self.smtpServerVar = tk.StringVar(value=self.defaultSmtpServer)
        self.smtpPortVar = tk.StringVar(value=self.defaultSmtpPort)
        self.sendWechatVar = tk.StringVar(value="no")
        self.wechatWebhookVar = tk.StringVar(value=self.defaultWechatWebhook)
        self.wechatMobileVar = tk.StringVar(value=self.defaultWechatMobile)
        # 赛狐流程表单变量
        self.usernameVar = tk.StringVar(value=self.defaultSaihuUsername)
        self.passwordVar = tk.StringVar(value=self.defaultSaihuPassword)
        self.exportDirVar = tk.StringVar(value=self.defaultExportDir)
        self.siteNameVar = tk.StringVar(value=self.defaultSiteName)
        self.shopNameVar = tk.StringVar(value=self.defaultShopBaseName)
        self.startDateVar = tk.StringVar(value=self.defaultStartDate)
        self.endDateVar = tk.StringVar(value=self.defaultEndDate)
        self.templatePathVar = tk.StringVar(value=self.defaultTemplatePath)
        # 易得客流程表单变量
        self.yidekeUsernameVar = tk.StringVar(value=self.defaultYidekeUsername)
        self.yidekePasswordVar = tk.StringVar(value=self.defaultYidekePassword)
        self.autoSiteNameVar = tk.StringVar(value=self.defaultAutoSiteName)
        self.amazonSiteNameVar = tk.StringVar(value=self.defaultAmazonSiteName)
        self.shopIpVar = tk.StringVar(value=self.defaultShopIp)
        self.shopPortVar = tk.StringVar(value=self.defaultShopPort)
        self.amazonEmailVar = tk.StringVar(value=self.defaultAmazonEmail)
        self.amazonPasswordVar = tk.StringVar(value=self.defaultAmazonPassword)
        self.popDirVar = tk.StringVar(value="")
        self.autoShipmentBox = None
        self.exportDirVar.trace_add("write", self.syncPopDir)

        self.loadConfig()
        self.buildUi()
        self.pollLog()
        self.root.protocol("WM_DELETE_WINDOW", self.onClose)

    def buildUi(self):
        """创建统一 GUI 界面"""
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        self.buildCommon(container)
        # 两个流程分离在不同标签页中
        tabs = ttk.Notebook(container)
        tabs.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        saihuFrame = ttk.Frame(tabs)
        autoFrame = ttk.Frame(tabs)
        tabs.add(saihuFrame, text="赛狐流程")
        tabs.add(autoFrame, text="易得客流程")

        self.saihuPane = self.SaihuPane(self, saihuFrame)
        self.autoPane = self.AutoPane(self, autoFrame)
        self.toggleEmail()
        # 统一界面已创建完成

    def buildCommon(self, parent):
        """创建两个流程共用配置区"""
        common = ttk.LabelFrame(parent, text="公共配置", padding=12)
        common.pack(fill=tk.X)

        # 顶部提示语说明两个流程的执行顺序，避免用户未生成 POP 就直接运行易得客流程
        tipText = "两个流程是分开的：如果没有 POP 文件，请先运行赛狐流程完成 POP 文件生成，再运行易得客流程。"
        ttk.Label(common, text=tipText, foreground="#6b4e00").grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        # 运行环境用于传入业务流程，区分线上与线下数据环境
        ttk.Label(common, text="运行环境").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        envFrame = ttk.Frame(common)
        envFrame.grid(row=1, column=1, sticky=tk.W, pady=6)
        ttk.Radiobutton(envFrame, text="线下", variable=self.isOnlineVar, value="offline").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(envFrame, text="线上", variable=self.isOnlineVar, value="online").pack(side=tk.LEFT)

        # 邮件通知为公共开关，赛狐发送 POP，易得客发送 CASE 结果与 POP 附件
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

        # 接收邮箱随邮件开关启用，用于接收 POP 或 CASE 汇总邮件
        ttk.Label(common, text="接收邮箱").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        self.emailEntry = ttk.Entry(common, textvariable=self.emailVar)
        self.emailEntry.grid(row=3, column=1, sticky=tk.EW, pady=6)

        # 企业微信通知只发送最终汇总消息，Webhook 与手机号由公共区统一填写
        ttk.Label(common, text="企业微信通知").grid(row=4, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        wechatFrame = ttk.Frame(common)
        wechatFrame.grid(row=4, column=1, sticky=tk.W, pady=6)
        ttk.Radiobutton(wechatFrame, text="不发送", variable=self.sendWechatVar, value="no").pack(
            side=tk.LEFT,
            padx=(0, 12),
        )
        ttk.Radiobutton(wechatFrame, text="发送", variable=self.sendWechatVar, value="yes").pack(side=tk.LEFT)

        ttk.Label(common, text="企业微信 Webhook").grid(row=5, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        ttk.Entry(common, textvariable=self.wechatWebhookVar).grid(row=5, column=1, sticky=tk.EW, pady=6)

        ttk.Label(common, text="@ 手机号").grid(row=6, column=0, sticky=tk.W, padx=(0, 8), pady=6)
        ttk.Entry(common, textvariable=self.wechatMobileVar).grid(row=6, column=1, sticky=tk.EW, pady=6)
        # 公共配置输入框横向拉伸，保证长 Webhook 与邮箱显示完整
        common.columnconfigure(1, weight=1)
        # 公共配置区已创建完成

    def togglePassword(self, entry, button):
        """切换密码输入框的显示与隐藏状态"""
        if entry.cget("show") == "*":
            entry.config(show="")
            button.config(text="隐藏")
            # 当前密码框已切换为明文显示
        else:
            entry.config(show="*")
            button.config(text="显示")
            # 当前密码框已恢复为隐藏显示

    def toggleEmail(self):
        """根据邮件开关启用或禁用邮箱输入框"""
        state = tk.NORMAL if self.sendEmailVar.get() == "yes" else tk.DISABLED
        self.emailEntry.config(state=state)
        # 邮箱输入框状态已同步

    def selectExportDir(self):
        """选择 POP 文件导出目录"""
        if self.selectingExportDir:
            messagebox.showinfo("提示", "目录选择窗口已打开", parent=self.root)
            return
        initial = self.exportDirVar.get().strip() or self.defaultExportDir
        initialPath = Path(initial)
        # 初始目录不存在时回退到项目目录
        if not initialPath.is_dir():
            initialPath = self.baseDir
        self.selectingExportDir = True
        thread = threading.Thread(target=self.openExportDir, args=(str(initialPath.resolve()),), daemon=True)
        thread.start()
        # 已启动 Windows 原生导出目录选择线程

    def openExportDir(self, initialDir):
        """调用 Windows 目录选择窗口并回填赛狐导出目录"""
        safeDir = initialDir.replace("'", "''")
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.StartPosition = 'CenterScreen'
$form.Width = 1
$form.Height = 1
$form.Opacity = 0
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = '选择赛狐 POP 导出目录'
$dialog.SelectedPath = '{safeDir}'
$dialog.ShowNewFolderButton = $true
$form.Show()
$form.Activate()
$result = $dialog.ShowDialog($form)
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $dialog.SelectedPath
}}
$dialog.Dispose()
$form.Dispose()
"""
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-STA",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    script,
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode != 0:
                raise RuntimeError((result.stderr or "导出目录选择窗口打开失败").strip())
            path = (result.stdout or "").strip()
            if path:
                self.root.after(0, lambda: self.setExportDir(path))
                # 导出目录已回填到主线程
        except Exception as exc:
            msg = str(exc)
            self.root.after(0, lambda: messagebox.showerror("导出目录选择失败", msg, parent=self.root))
        finally:
            self.root.after(0, lambda: setattr(self, "selectingExportDir", False))

    def setExportDir(self, path):
        """写入赛狐导出目录并触发易得客 POP 目录同步"""
        self.exportDirVar.set(path)
        # 导出目录已写入表单，监听器会同步 POP 目录

    def syncPopDir(self, *args):
        """赛狐导出目录变化时同步易得客 POP 目录"""
        exportDir = self.exportDirVar.get().strip() or self.defaultExportDir
        self.popDirVar.set(exportDir)
        # 易得客 POP 目录已跟随赛狐导出目录更新
        self.refreshAutoShipmentIds()
        # 待处理货件编号展示已按新目录刷新

    def selectTemplateFile(self):
        """选择 POP 输出模板文件"""
        initial = self.templatePathVar.get().strip()
        initialDir = str(Path(initial).parent) if initial else str(self.baseDir)
        path = filedialog.askopenfilename(
            initialdir=initialDir,
            filetypes=[
                ("POP 模板文件", "*.docx;*.pdf"),
                ("Word 模板", "*.docx"),
                ("PDF 模板", "*.pdf"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.templatePathVar.set(path)
            # POP 模板路径已写入表单

    def selectPopDir(self):
        """选择已完成导出的 POP PDF 存放目录"""
        if self.selectingPopDir:
            messagebox.showinfo("提示", "目录选择窗口已打开", parent=self.root)
            return
        initial = self.popDirVar.get().strip() or self.exportDirVar.get().strip() or self.defaultExportDir
        initialPath = Path(initial)
        # 初始目录不存在时回退到项目目录
        if not initialPath.is_dir():
            initialPath = self.baseDir
        self.selectingPopDir = True
        thread = threading.Thread(target=self.openPopDir, args=(str(initialPath.resolve()),), daemon=True)
        thread.start()
        # 已启动 Windows 原生目录选择线程

    def openPopDir(self, initialDir):
        """调用 Windows 目录选择窗口并回填结果"""
        safeDir = initialDir.replace("'", "''")
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.StartPosition = 'CenterScreen'
$form.Width = 1
$form.Height = 1
$form.Opacity = 0
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = '选择已完成导出存放 POP 文件的目录'
$dialog.SelectedPath = '{safeDir}'
$dialog.ShowNewFolderButton = $false
$form.Show()
$form.Activate()
$result = $dialog.ShowDialog($form)
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $dialog.SelectedPath
}}
$dialog.Dispose()
$form.Dispose()
"""
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-STA",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    script,
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode != 0:
                raise RuntimeError((result.stderr or "目录选择窗口打开失败").strip())
            path = (result.stdout or "").strip()
            if path:
                self.root.after(0, lambda: self.setPopDir(path))
                # POP PDF 目录已写入表单
        except Exception as exc:
            msg = str(exc)
            self.root.after(0, lambda: messagebox.showerror("目录选择失败", msg, parent=self.root))
        finally:
            self.root.after(0, lambda: setattr(self, "selectingPopDir", False))

    def setPopDir(self, path):
        """写入 POP 目录并刷新待处理货件编号"""
        self.popDirVar.set(path)
        # POP 目录已写入界面变量
        self.refreshAutoShipmentIds()
        # 待处理货件编号展示已同步刷新

    def getAutoShipmentIds(self):
        """优先从本轮 shipment_ids.json 读取易得客待处理货件编号"""
        shipmentIds = []
        seen = set()
        popDir = self.popDirVar.get().strip()
        popPath = Path(popDir) if popDir else None

        jsonPaths = []
        if popPath:
            jsonPaths.append(popPath / "shipment_ids.json")
        exportDir = self.exportDirVar.get().strip()
        if exportDir:
            jsonPaths.append(Path(exportDir) / "shipment_ids.json")

        for jsonPath in jsonPaths:
            # 优先读取赛狐本轮覆盖生成的 JSON，避免 POP 目录旧文件混入
            if not jsonPath.is_file():
                continue
            try:
                data = json.loads(jsonPath.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return [], f"{jsonPath} 解析失败"
            if isinstance(data, list):
                rawIds = data
            elif isinstance(data, dict):
                if "shipmentIds" in data:
                    rawIds = data.get("shipmentIds") or []
                elif "allShipmentIds" in data:
                    rawIds = data.get("allShipmentIds") or []
                else:
                    rawIds = data.get("ids") or []
            else:
                rawIds = []
            for item in rawIds:
                shipmentId = str(item or "").strip().upper()
                if not re.fullmatch(r"FBA[A-Z0-9]{6,}", shipmentId):
                    continue
                if shipmentId in seen:
                    continue
                seen.add(shipmentId)
                shipmentIds.append(shipmentId)
            return shipmentIds, str(jsonPath)

        if popPath and popPath.is_dir():
            for pdfPath in sorted(popPath.iterdir()):
                # 没有本轮 JSON 时，才兜底从最终导出的 POP PDF 文件名提取货件编号
                if not pdfPath.is_file() or pdfPath.suffix.lower() != ".pdf":
                    continue
                match = re.search(r"(FBA[A-Z0-9]{6,})", pdfPath.stem, re.IGNORECASE)
                if not match:
                    continue
                shipmentId = match.group(1).upper()
                if shipmentId in seen:
                    continue
                seen.add(shipmentId)
                shipmentIds.append(shipmentId)
            if shipmentIds:
                return shipmentIds, "POP PDF 文件名（兜底）"

        return shipmentIds, ""

    def refreshAutoShipmentIds(self):
        """刷新易得客标签页里的待处理货件编号展示"""
        if not self.autoShipmentBox:
            return
        shipmentIds, source = self.getAutoShipmentIds()
        if shipmentIds:
            text = f"来源：{source}\n共 {len(shipmentIds)} 个\n" + "\n".join(shipmentIds)
        else:
            text = "未提取到货件编号。请选择 POP 目录，或先运行赛狐流程生成 POP 与 shipment_ids.json。"
        self.autoShipmentBox.config(state=tk.NORMAL)
        self.autoShipmentBox.delete("1.0", tk.END)
        self.autoShipmentBox.insert(tk.END, text)
        self.autoShipmentBox.config(state=tk.DISABLED)
        # 待处理货件编号展示已刷新

    def buildCommonConfig(self):
        """读取公共配置并做基础校验"""
        sendEmail = self.sendEmailVar.get() == "yes"
        sendWechat = self.sendWechatVar.get() == "yes"
        email = self.emailVar.get().strip()
        senderEmail = self.senderEmailVar.get().strip()
        smtpAuthCode = self.smtpAuthCodeVar.get().strip()
        smtpServer = self.smtpServerVar.get().strip()
        smtpPort = self.smtpPortVar.get().strip()
        wechatWebhook = self.wechatWebhookVar.get().strip()
        wechatMobile = self.wechatMobileVar.get().strip()
        # 邮件通知开启时必须具备完整 SMTP 发件配置
        if sendEmail:
            if not email:
                raise ValueError("选择发送邮件时必须填写接收邮箱")
            if not senderEmail:
                raise ValueError("选择发送邮件时必须配置发件邮箱")
            if not smtpAuthCode:
                raise ValueError("选择发送邮件时必须配置 SMTP 授权码")
            if not smtpServer:
                raise ValueError("选择发送邮件时必须配置 SMTP 服务地址")
            if not smtpPort.isdigit():
                raise ValueError("SMTP 端口只能填写数字")
        # 企业微信通知开启时必须提供 Webhook 与 @ 手机号
        if sendWechat:
            if not wechatWebhook:
                raise ValueError("请填写企业微信 Webhook")
            if not wechatMobile:
                raise ValueError("请填写企业微信 @ 手机号")
        # 公共配置会同时传给赛狐流程和易得客流程
        return {
            # 运行环境：True 为线上，False 为线下
            "isOnline": self.isOnlineVar.get() == "online",
            # 邮件通知：赛狐发送 POP，易得客发送 CASE 结果
            "sendEmail": sendEmail,
            # 邮件接收人
            "email": email,
            # 邮件发件邮箱
            "sender_email": senderEmail,
            # 邮件发件邮箱 SMTP 授权码
            "smtp_auth_code": smtpAuthCode,
            # SMTP 服务地址
            "smtp_server": smtpServer,
            # SMTP SSL 端口
            "smtp_port": smtpPort,
            # 企业微信通知开关
            "sendWechat": sendWechat,
            # 企业微信群机器人 Webhook
            "wechatWebhook": wechatWebhook,
            # 企业微信 @ 人手机号，可填写多个
            "wechatMobile": wechatMobile,
        }

    def buildSaihuConfig(self):
        """校验赛狐配置并组装 Saihu 入参"""
        common = self.buildCommonConfig()
        username = self.usernameVar.get().strip()
        password = self.passwordVar.get()
        exportDir = self.exportDirVar.get().strip()
        siteName = self.siteNameVar.get().strip()
        shopName = self.shopNameVar.get().strip()
        startDate = self.startDateVar.get().strip()
        endDate = self.endDateVar.get().strip()
        templatePath = self.templatePathVar.get().strip()

        # 校验赛狐账号密码
        if not username:
            raise ValueError("请填写赛狐账号")
        if not password:
            raise ValueError("请填写赛狐密码")
        # 校验导出目录
        if not exportDir:
            raise ValueError("请选择导出目录")
        exportPath = Path(exportDir)
        exportPath.mkdir(parents=True, exist_ok=True)
        # 校验站点选择
        if not siteName:
            raise ValueError("请选择筛选站点")
        if siteName not in self.siteNames:
            raise ValueError(f"暂不支持该筛选站点: {siteName}")
        # 校验赛狐店铺必须选择具体店铺
        if not shopName:
            raise ValueError("请选择筛选店铺")
        if shopName not in self.shopNames:
            raise ValueError(f"暂不支持该筛选店铺: {shopName}")
        if siteName not in self.siteShopSuffixMap:
            raise ValueError(f"暂未配置该站点的店铺后缀: {siteName}")
        fullShopName = f"{shopName}-{self.siteShopSuffixMap[siteName]}"
        # 校验赛狐筛选日期，格式固定为 YYYY-MM-DD
        if not startDate:
            raise ValueError("请填写开始时间")
        if not endDate:
            raise ValueError("请填写结束时间")
        try:
            startDateValue = datetime.strptime(startDate, "%Y-%m-%d").date()
            endDateValue = datetime.strptime(endDate, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError("开始时间和结束时间格式必须为 YYYY-MM-DD") from exc
        if startDateValue > endDateValue:
            raise ValueError("开始时间不能晚于结束时间")
        # 校验 POP 模板文件，支持 Word 与 PDF 两种模板
        if not templatePath:
            raise ValueError("请选择模板文件")
        templateFile = Path(templatePath)
        if not templateFile.is_file():
            raise ValueError(f"模板文件不存在: {templatePath}")
        if templateFile.suffix.lower() not in (".docx", ".pdf"):
            raise ValueError("模板文件仅支持 .docx 或 .pdf")
        # 赛狐流程实际发送 POP 邮件，开启邮件时必须填写邮箱
        if common["sendEmail"] and not common["email"]:
            raise ValueError("选择发送邮件时必须填写接收邮箱")

        page = ChromiumPage()
        # 赛狐流程入参：页面实例、账号、导出目录、站点与模板信息
        config = {
            # 当前 Chrome 页面实例，由 Saihu 接管赛狐页面
            "page": page,
            # 赛狐登录账号密码
            "username": username,
            "password": password,
            # POP 输出目录与项目资源目录
            "exportDir": str(exportPath.resolve()),
            "baseDir": str(self.baseDir),
            # FBA 货件筛选站点
            "siteName": siteName,
            # FBA 货件筛选店铺
            "shopName": fullShopName,
            # GUI 选择的店铺主体名
            "shopBaseName": shopName,
            # FBA 货件筛选更新时间范围
            "startDate": startDate,
            "endDate": endDate,
            # POP 输出模板，支持 Word 模板或 PDF 模板
            "templatePath": str(templateFile.resolve()),
        }
        config.update(common)
        return config

    def buildAutoConfig(self):
        """校验易得客配置并组装 Auto 入参"""
        common = self.buildCommonConfig()
        yidekeUsername = self.yidekeUsernameVar.get().strip()
        yidekePassword = self.yidekePasswordVar.get()
        autoSiteName = self.autoSiteNameVar.get().strip()
        amazonSiteName = self.amazonSiteNameVar.get().strip()
        shopIp = self.shopIpVar.get().strip()
        shopPort = self.shopPortVar.get().strip()
        amazonEmail = self.amazonEmailVar.get().strip()
        amazonPassword = self.amazonPasswordVar.get()
        popDir = self.popDirVar.get().strip()

        # 校验易得客店铺站点
        if not autoSiteName:
            raise ValueError("请选择易得客店铺站点")
        if autoSiteName not in self.autoSiteNames:
            raise ValueError(f"暂不支持该易得客店铺站点: {autoSiteName}")
        # 校验 Amazon 后台站点，该站点只用于 Seller Central 账号选择和站点切换
        if not amazonSiteName:
            raise ValueError("请选择 Amazon 后台站点")
        if amazonSiteName not in self.autoSiteNames:
            raise ValueError(f"暂不支持该 Amazon 后台站点: {amazonSiteName}")
        # 校验 POP PDF 目录
        if not popDir:
            raise ValueError("请选择已完成导出存放 POP 文件的目录")
        popPath = Path(popDir)
        if not popPath.is_dir():
            raise ValueError(f"POP 文件目录不存在: {popDir}")
        shipmentIds, source = self.getAutoShipmentIds()
        if not shipmentIds:
            raise ValueError("未从 POP 目录或 shipment_ids.json 中提取到货件编号")
        # 校验店铺端口
        if not shopPort:
            raise ValueError("请填写店铺端口")
        if not shopPort.isdigit():
            raise ValueError("店铺端口只能填写数字")
        # 正式模式必须提供易得客账号和店铺 IP
        if not yidekeUsername:
            raise ValueError("请填写易得客账号")
        if not yidekePassword:
            raise ValueError("请填写易得客密码")
        if not shopIp:
            raise ValueError("请填写店铺 IP")

        # 易得客流程入参：进店信息、Amazon 账号、POP 目录与货件来源
        config = {
            # 易得客登录账号密码
            "yidekeUsername": yidekeUsername,
            "yidekePassword": yidekePassword,
            # 店铺站点只影响易得客进店访问
            "autoSiteName": autoSiteName,
            # Amazon 后台站点只影响 Seller Central 账号选择和站点切换
            "amazonSiteName": amazonSiteName,
            # 店铺 IP 与调试端口用于接管易得客浏览器
            "shopIp": shopIp,
            "shopPort": int(shopPort),
            # Amazon Seller Central 登录账号密码
            "amazonEmail": amazonEmail,
            "amazonPassword": amazonPassword,
            # POP 目录用于读取待处理货件编号与上传对应 PDF
            "popDir": popDir,
            "shipmentSource": source,
            # 项目资源目录用于读取 POD 文件
            "baseDir": str(self.baseDir),
        }
        config.update(common)
        return config

    def isRunning(self):
        """判断当前是否已有流程在运行"""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "已有流程正在运行中")
            return True
        return False

    def startSaihu(self):
        """启动赛狐流程后台线程"""
        if self.isRunning():
            return
        try:
            config = self.buildSaihuConfig()
            # 表单配置已校验并生成赛狐运行参数
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return
        self.saveConfig()
        self.currentLogText = self.saihuPane.logText
        self.currentStartButton = self.saihuPane.startButton
        self.currentStatusVar = self.saihuPane.statusVar
        self.saihuPane.startButton.config(state=tk.DISABLED)
        self.saihuPane.statusVar.set("运行中")
        self.appendLog(self.saihuPane.logText, "开始运行赛狐流程...\n")
        self.appendLog(self.saihuPane.logText, f"运行环境: {'线上' if config.get('isOnline') else '线下'}\n")
        self.appendLog(self.saihuPane.logText, f"筛选站点: {config.get('siteName')}\n")
        self.appendLog(self.saihuPane.logText, f"筛选店铺: {config.get('shopName')}\n")
        self.appendLog(self.saihuPane.logText, f"筛选时间: {config.get('startDate')} 至 {config.get('endDate')}\n")
        self.appendLog(self.saihuPane.logText, f"导出目录: {config.get('exportDir')}\n")
        self.appendLog(self.saihuPane.logText, f"模板文件: {config.get('templatePath')}\n")
        self.appendLog(self.saihuPane.logText, f"邮件通知: {'发送' if config.get('sendEmail') else '不发送'}\n")
        # 后台线程执行耗时浏览器流程，避免 GUI 窗口卡住
        self.worker = threading.Thread(target=self.runSaihuTask, args=(config,), daemon=True)
        self.worker.start()
        # 赛狐流程后台线程已启动

    def startAuto(self):
        """启动易得客流程后台线程"""
        if self.isRunning():
            return
        try:
            config = self.buildAutoConfig()
            # 表单配置已校验并生成易得客运行参数
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return
        self.saveConfig()
        self.currentLogText = self.autoPane.logText
        self.currentStartButton = self.autoPane.startButton
        self.currentStatusVar = self.autoPane.statusVar
        self.autoPane.startButton.config(state=tk.DISABLED)
        self.autoPane.statusVar.set("运行中")
        self.appendLog(self.autoPane.logText, "开始运行易得客流程...\n")
        self.appendLog(self.autoPane.logText, f"店铺站点: {config.get('autoSiteName')}\n")
        self.appendLog(self.autoPane.logText, f"Amazon后台站点: {config.get('amazonSiteName')}\n")
        self.appendLog(self.autoPane.logText, f"POP目录: {config.get('popDir') or '未配置'}\n")
        self.appendLog(self.autoPane.logText, f"货件编号来源: {config.get('shipmentSource') or '未识别'}\n")
        self.appendLog(self.autoPane.logText, f"企业微信通知: {'发送' if config.get('sendWechat') else '不发送'}\n")
        # 后台线程执行易得客与 Amazon 页面流程，避免 GUI 窗口卡住
        self.worker = threading.Thread(target=self.runAutoTask, args=(config,), daemon=True)
        self.worker.start()
        # 易得客流程后台线程已启动

    def runSaihuTask(self, config):
        """在线程中执行赛狐主流程并回写日志"""
        oldStdout = sys.stdout
        oldStderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            self.main.runSaihu(config)
            self.logQueue.put((self.currentLogText, "\n赛狐流程任务完成。\n"))
            self.root.after(0, lambda: self.currentStatusVar.set("已完成"))
        except Exception:
            self.logQueue.put((self.currentLogText, "\n赛狐流程运行失败：\n"))
            self.logQueue.put((self.currentLogText, traceback.format_exc()))
            self.root.after(0, lambda: self.currentStatusVar.set("运行失败"))
        finally:
            sys.stdout = oldStdout
            sys.stderr = oldStderr
            self.root.after(0, lambda: self.currentStartButton.config(state=tk.NORMAL))
            # 标准输出已恢复，开始按钮已恢复可用

    def runAutoTask(self, config):
        """在线程中执行易得客流程并回写日志"""
        oldStdout = sys.stdout
        oldStderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        try:
            self.main.runAuto(config)
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
            # 已追加一段日志文本
        self.root.after(100, self.pollLog)

    def appendLog(self, logText, text):
        """向指定日志框追加文本"""
        if not logText:
            return
        logText.insert(tk.END, text)
        logText.see(tk.END)
        # 日志框已滚动到末尾

    def clearSaihuLog(self):
        """清空赛狐流程日志"""
        self.saihuPane.logText.delete("1.0", tk.END)
        # 赛狐日志已清空

    def clearAutoLog(self):
        """清空易得客流程日志"""
        self.autoPane.logText.delete("1.0", tk.END)
        # 易得客日志已清空

    def write(self, text):
        """接收 print 输出并放入当前流程日志队列"""
        if text:
            self.logQueue.put((self.currentLogText, text))
            # print 输出已进入当前流程日志队列

    def flush(self):
        """兼容标准输出 flush 接口"""
        return

    def loadConfig(self):
        """读取统一 GUI 持久化配置"""
        data = {}
        if self.configFile.is_file():
            try:
                data = json.loads(self.configFile.read_text(encoding="utf-8"))
                # 统一配置文件已读取并解析为 JSON
            except (json.JSONDecodeError, OSError):
                data = {}
        # 回填公共配置
        self.isOnlineVar.set("online" if data.get("isOnline") else "offline")
        self.sendEmailVar.set("yes" if data.get("sendEmail") else "no")
        self.emailVar.set(str(data.get("email") or self.defaultEmail).strip())
        self.senderEmailVar.set(str(data.get("sender_email") or self.defaultSenderEmail).strip())
        self.smtpAuthCodeVar.set(str(data.get("smtp_auth_code") or self.defaultSmtpAuthCode).strip())
        self.smtpServerVar.set(str(data.get("smtp_server") or self.defaultSmtpServer).strip())
        self.smtpPortVar.set(str(data.get("smtp_port") or self.defaultSmtpPort).strip())
        self.sendWechatVar.set("yes" if data.get("sendWechat") else "no")
        self.wechatWebhookVar.set(str(data.get("wechatWebhook") or self.defaultWechatWebhook).strip())
        self.wechatMobileVar.set(str(data.get("wechatMobile") or self.defaultWechatMobile).strip())

        # 回填赛狐配置
        self.usernameVar.set(str(data.get("username") or self.usernameVar.get()).strip())
        self.passwordVar.set(str(data.get("password") or self.defaultSaihuPassword))
        self.exportDirVar.set(str(data.get("exportDir") or self.defaultExportDir).strip())
        self.siteNameVar.set(str(data.get("siteName") or self.defaultSiteName).strip())
        savedShopName = str(data.get("shopBaseName") or data.get("shopName") or self.defaultShopBaseName).strip()
        for siteSuffix in self.siteShopSuffixMap.values():
            suffixText = f"-{siteSuffix}"
            if savedShopName.endswith(suffixText):
                savedShopName = savedShopName[:-len(suffixText)]
                break
        if savedShopName not in self.shopNames:
            savedShopName = self.defaultShopBaseName
        self.shopNameVar.set(savedShopName)
        self.startDateVar.set(str(data.get("startDate") or self.defaultStartDate).strip())
        self.endDateVar.set(str(data.get("endDate") or self.defaultEndDate).strip())
        savedTemplatePath = str(data.get("templatePath") or self.defaultTemplatePath).strip()
        if not Path(savedTemplatePath).is_file():
            savedTemplatePath = self.defaultTemplatePath
        self.templatePathVar.set(savedTemplatePath)

        # 回填易得客配置
        self.yidekeUsernameVar.set(str(data.get("yidekeUsername") or self.defaultYidekeUsername).strip())
        self.yidekePasswordVar.set(str(data.get("yidekePassword") or self.defaultYidekePassword))
        self.autoSiteNameVar.set(str(data.get("autoSiteName") or data.get("siteName") or self.defaultAutoSiteName).strip())
        self.amazonSiteNameVar.set(str(data.get("amazonSiteName") or data.get("autoSiteName") or data.get("siteName") or self.defaultAmazonSiteName).strip())
        self.shopIpVar.set(str(data.get("shopIp") or self.defaultShopIp).strip())
        self.shopPortVar.set(str(data.get("shopPort") or self.defaultShopPort).strip())
        self.amazonEmailVar.set(str(data.get("amazonEmail") or self.defaultAmazonEmail).strip())
        self.amazonPasswordVar.set(str(data.get("amazonPassword") or self.defaultAmazonPassword))
        self.popDirVar.set(self.exportDirVar.get().strip() or self.defaultExportDir)

    def saveConfig(self):
        """保存统一 GUI 配置到本地文件"""
        # 统一 GUI 仅保存一个 run_config.json，删除后下次会按默认值重建
        data = {
            # 公共配置
            "isOnline": self.isOnlineVar.get() == "online",
            "sendEmail": self.sendEmailVar.get() == "yes",
            "email": self.emailVar.get().strip(),
            "sender_email": self.senderEmailVar.get().strip() or self.defaultSenderEmail,
            "smtp_auth_code": self.smtpAuthCodeVar.get().strip() or self.defaultSmtpAuthCode,
            "smtp_server": self.smtpServerVar.get().strip() or self.defaultSmtpServer,
            "smtp_port": self.smtpPortVar.get().strip() or self.defaultSmtpPort,
            "sendWechat": self.sendWechatVar.get() == "yes",
            "wechatWebhook": self.wechatWebhookVar.get().strip(),
            "wechatMobile": self.wechatMobileVar.get().strip(),
            # 赛狐流程配置
            "username": self.usernameVar.get().strip(),
            "password": self.passwordVar.get(),
            "exportDir": self.exportDirVar.get().strip() or self.defaultExportDir,
            "siteName": self.siteNameVar.get().strip() or self.defaultSiteName,
            "shopName": self.shopNameVar.get().strip() or self.defaultShopBaseName,
            "shopBaseName": self.shopNameVar.get().strip() or self.defaultShopBaseName,
            "startDate": self.startDateVar.get().strip() or self.defaultStartDate,
            "endDate": self.endDateVar.get().strip() or self.defaultEndDate,
            "templatePath": self.templatePathVar.get().strip() or self.defaultTemplatePath,
            # 易得客流程配置
            "yidekeUsername": self.yidekeUsernameVar.get().strip(),
            "yidekePassword": self.yidekePasswordVar.get(),
            "autoSiteName": self.autoSiteNameVar.get().strip() or self.defaultAutoSiteName,
            "amazonSiteName": self.amazonSiteNameVar.get().strip() or self.autoSiteNameVar.get().strip() or self.defaultAmazonSiteName,
            "shopIp": self.shopIpVar.get().strip(),
            "shopPort": self.shopPortVar.get().strip() or self.defaultShopPort,
            "amazonEmail": self.amazonEmailVar.get().strip(),
            "amazonPassword": self.amazonPasswordVar.get(),
            # 易得客 POP 目录默认跟随赛狐导出目录，保证两段流程衔接
            "popDir": self.exportDirVar.get().strip() or self.defaultExportDir,
        }
        try:
            self.configFile.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            # 统一 GUI 配置已写入本地文件
        except OSError:
            # 配置保存失败不影响当前运行
            pass

    def onClose(self):
        """关闭窗口前保存配置"""
        self.saveConfig()
        self.root.destroy()
        # GUI 窗口已关闭


if __name__ == "__main__":
    # 本文件独立调试入口
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = RunGui(root)
    root.mainloop()
