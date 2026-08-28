"""TDI 公司与个人合作推广 PySide6 桌面入口。"""

import json
import os
import sys
import threading
import traceback
import ctypes
from pathlib import Path


def configureFrozenDllSearchPath():
    if not getattr(sys, "frozen", False):
        return
    candidates = [
        Path(str(getattr(sys, "_MEIPASS", ""))),
        Path(sys.executable).resolve().parent / "_internal",
    ]
    for candidate in candidates:
        if not candidate or not candidate.exists():
            continue
        try:
            os.add_dll_directory(str(candidate))
        except Exception:
            pass
        os.environ["PATH"] = f"{candidate}{os.pathsep}{os.environ.get('PATH', '')}"
        for name in ("libcrypto-3-x64.dll", "libssl-3-x64.dll"):
            dll = candidate / name
            if dll.exists():
                try:
                    ctypes.WinDLL(str(dll))
                except Exception:
                    pass


configureFrozenDllSearchPath()

try:
    from main import Main

    from PySide6.QtCore import QObject, QSize, Qt, QThread, QTimer, QUrl, Signal, Slot
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QSizePolicy,
        QSpinBox,
        QStackedWidget,
        QStyle,
        QTableWidget,
        QTableWidgetItem,
        QVBoxLayout,
        QWidget,
    )
except Exception:
    try:
        baseDir = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parent
        )
        path = baseDir / "output" / "import_error.txt"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(traceback.format_exc(), encoding="utf-8")
    except Exception:
        pass
    raise


def appBaseDir():
    return (
        Path(sys.executable).resolve().parent
        if getattr(sys, "frozen", False)
        else Path(__file__).resolve().parent
    )


def writeDiagnosticJson(baseDir, name, payload):
    path = Path(baseDir) / "output" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def runQuotaDiagnostic():
    baseDir = appBaseDir()
    writeDiagnosticJson(baseDir, "quota_diagnostic.json", {"stage": "start"})
    try:
        model = Main({"baseDir": str(baseDir)})
        writeDiagnosticJson(baseDir, "quota_diagnostic.json", {"stage": "main-created"})
        import time

        started = time.time()
        status = model.status(refreshQuota=True)
        writeDiagnosticJson(
            baseDir,
            "quota_diagnostic.json",
            {
                "stage": "complete",
                "elapsed": round(time.time() - started, 2),
                "quota": (status or {}).get("quota") or {},
                "checks": [
                    {
                        "name": item.get("name", ""),
                        "ok": bool(item.get("ok")),
                        "detail": item.get("detail", ""),
                    }
                    for item in list((status or {}).get("checks") or [])
                ],
            },
        )
        return 0
    except Exception:
        writeDiagnosticJson(
            baseDir,
            "quota_diagnostic.json",
            {"stage": "error", "traceback": traceback.format_exc()},
        )
        return 1


class RunGui(QMainWindow):
    """提供任务配置、运行监控、环境状态和人工审核界面。"""

    class Worker(QObject):
        """在独立线程执行完整流程、刷新或推广预览。"""

        log = Signal(str)
        progress = Signal(int, int, str)
        state = Signal(str, str)
        humanRequired = Signal(str, str)
        finished = Signal(str, dict)
        failed = Signal(str, str)

        def __init__(self, baseDir, operation, config):
            """保存线程操作、配置和协作控制事件。"""
            super().__init__()
            self.baseDir = baseDir
            self.operation = operation
            self.config = dict(config)
            self.pauseEvent = threading.Event()
            self.stopEvent = threading.Event()

        def humanPause(self, reason, url):
            """暂停后台流程并通知主窗口人工处理。"""
            self.pauseEvent.set()
            self.humanRequired.emit(reason, url)

        @Slot()
        def run(self):
            """按操作类型调用 Main 的明确入口。"""
            runtime = dict(self.config)
            runtime.update({
                "baseDir": str(self.baseDir),
                "pauseEvent": self.pauseEvent,
                "stopEvent": self.stopEvent,
                "logCallback": self.log.emit,
                "progressCallback": self.progress.emit,
                "stateCallback": self.state.emit,
                "humanCallback": self.humanPause,
            })
            try:
                main = Main(runtime)
                if self.operation == "run":
                    result = main.run()
                elif self.operation == "refresh":
                    result = main.status(refreshQuota=True)
                elif self.operation == "preview":
                    result = main.previewPromotion()
                elif self.operation == "mail":
                    result = main.processPromotion(str(runtime.get("promotionMode") or "draft"))
                else:
                    raise ValueError(f"未知操作：{self.operation}")
                self.finished.emit(self.operation, dict(result or {}))
            except Exception:
                self.failed.emit(self.operation, traceback.format_exc())

        def pause(self):
            self.pauseEvent.set()
            self.state.emit("paused", "任务已暂停")

        def resume(self):
            self.pauseEvent.clear()
            self.state.emit("running", "任务继续运行")

        def stop(self):
            self.stopEvent.set()
            self.pauseEvent.clear()
            self.state.emit("stopping", "正在安全停止")

    def __init__(self):
        """初始化项目路径、状态模型和全部页面。"""
        super().__init__()
        self.baseDir = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parent
        )
        self.model = Main({"baseDir": str(self.baseDir)})
        self.thread = None
        self.worker = None
        self.metricValues = {}
        self.metricNotes = {}
        self.flowStates = {}
        self.officialQuota = None
        self.todayUsed = 0
        self.pageNames = [
            "工作台",
            "运行设置",
            "任务监控",
            "邮件设置",
            "结果审核",
            "结果数据",
            "连接状态",
            "操作说明",
        ]
        self.setWindowTitle("TDI 公司与个人合作推广")
        self.resize(1380, 860)
        self.setMinimumSize(1120, 720)
        self.buildUi()
        self.applyTheme()
        self.refreshLocal()
        self.markQuotaRefreshing()
        QTimer.singleShot(0, self.refreshRemote)
        self.refreshReview()
        self.refreshResults()

    def styleIcon(self, icon):
        return self.style().standardIcon(icon)

    def buildUi(self):
        """构建侧栏、标题栏和八个工作页面。"""
        root = QWidget()
        rootLayout = QHBoxLayout(root)
        rootLayout.setContentsMargins(0, 0, 0, 0)
        rootLayout.setSpacing(0)
        self.setCentralWidget(root)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(228)
        sideLayout = QVBoxLayout(sidebar)
        sideLayout.setContentsMargins(16, 22, 16, 18)
        sideLayout.setSpacing(12)
        brand = QLabel("TDI\n推广工作台")
        brand.setObjectName("brand")
        sideLayout.addWidget(brand)
        edition = QLabel("SERPAPI FREE · 250 / MONTH")
        edition.setObjectName("edition")
        sideLayout.addWidget(edition)

        self.navigation = QListWidget()
        self.navigation.setObjectName("navigation")
        self.navigation.setSpacing(4)
        self.navigation.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        icons = [
            QStyle.SP_ComputerIcon,
            QStyle.SP_FileDialogDetailedView,
            QStyle.SP_FileDialogContentsView,
            QStyle.SP_FileDialogNewFolder,
            QStyle.SP_DialogApplyButton,
            QStyle.SP_FileDialogListView,
            QStyle.SP_DriveNetIcon,
            QStyle.SP_MessageBoxQuestion,
        ]
        for label, icon in zip(self.pageNames, icons):
            item = QListWidgetItem(self.styleIcon(icon), label)
            item.setSizeHint(QSize(176, 42))
            self.navigation.addItem(item)
        self.navigation.currentRowChanged.connect(self.switchPage)
        sideLayout.addWidget(self.navigation, 1)

        safety = QLabel("普通网页：直连访问\nFacebook：仅保存链接")
        safety.setObjectName("safety")
        sideLayout.addWidget(safety)
        rootLayout.addWidget(sidebar)

        content = QWidget()
        contentLayout = QVBoxLayout(content)
        contentLayout.setContentsMargins(24, 18, 24, 20)
        contentLayout.setSpacing(16)
        contentLayout.addWidget(self.buildHeader())
        self.stack = QStackedWidget()
        self.stack.addWidget(self.buildOverview())
        self.stack.addWidget(self.buildTask())
        self.stack.addWidget(self.buildMonitor())
        self.stack.addWidget(self.buildMail())
        self.stack.addWidget(self.buildReview())
        self.stack.addWidget(self.buildResults())
        self.stack.addWidget(self.buildEnvironment())
        self.stack.addWidget(self.buildHelp())
        contentLayout.addWidget(self.stack, 1)
        rootLayout.addWidget(content, 1)
        self.navigation.setCurrentRow(0)

    def buildHeader(self):
        header = QFrame()
        header.setObjectName("header")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(0, 0, 0, 0)
        titleLayout = QVBoxLayout()
        titleLayout.setSpacing(1)
        self.pageTitle = QLabel(self.pageNames[0])
        self.pageTitle.setObjectName("pageTitle")
        self.pageDetail = QLabel("流程状态、免费额度与启动条件")
        self.pageDetail.setObjectName("pageDetail")
        titleLayout.addWidget(self.pageTitle)
        titleLayout.addWidget(self.pageDetail)
        layout.addLayout(titleLayout)
        layout.addStretch()
        self.stateBadge = QLabel("就绪")
        self.stateBadge.setObjectName("stateBadge")
        self.stateBadge.setAlignment(Qt.AlignCenter)
        self.stateBadge.setMinimumWidth(86)
        layout.addWidget(self.stateBadge)
        self.headerStart = QPushButton("开始运行")
        self.headerStart.setObjectName("primaryButton")
        self.headerStart.setIcon(self.styleIcon(QStyle.SP_MediaPlay))
        self.headerStart.clicked.connect(self.startRun)
        layout.addWidget(self.headerStart)
        return header

    def metricCard(self, key, title, tone):
        card = QFrame()
        card.setObjectName("metricCard")
        card.setProperty("tone", tone)
        card.setMinimumHeight(108)
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(15, 13, 15, 12)
        layout.setSpacing(3)
        titleLabel = QLabel(title)
        titleLabel.setObjectName("metricTitle")
        valueLabel = QLabel("--")
        valueLabel.setObjectName("metricValue")
        noteLabel = QLabel("")
        noteLabel.setObjectName("metricNote")
        noteLabel.setWordWrap(True)
        layout.addWidget(titleLabel)
        layout.addWidget(valueLabel)
        layout.addWidget(noteLabel)
        self.metricValues[key] = valueLabel
        self.metricNotes[key] = noteLabel
        return card

    def workflowStep(self, number, title):
        step = QFrame()
        step.setObjectName("workflowStep")
        step.setMinimumHeight(72)
        layout = QHBoxLayout(step)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(10)
        indexLabel = QLabel(str(number))
        indexLabel.setObjectName("flowIndex")
        indexLabel.setAlignment(Qt.AlignCenter)
        indexLabel.setFixedSize(28, 28)
        textLayout = QVBoxLayout()
        textLayout.setSpacing(2)
        titleLabel = QLabel(title)
        titleLabel.setObjectName("flowTitle")
        stateLabel = QLabel("待检查")
        stateLabel.setObjectName("flowState")
        textLayout.addWidget(titleLabel)
        textLayout.addWidget(stateLabel)
        layout.addWidget(indexLabel)
        layout.addLayout(textLayout, 1)
        self.flowStates[number] = stateLabel
        return step

    def buildOverview(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)
        flowTitle = QLabel("流程状态")
        flowTitle.setObjectName("sectionTitle")
        layout.addWidget(flowTitle)
        flowLayout = QHBoxLayout()
        flowLayout.setSpacing(10)
        for number, title in enumerate(
            ("检查连接", "读取名单", "设置任务", "执行搜索", "审核结果"),
            start=1,
        ):
            flowLayout.addWidget(self.workflowStep(number, title), 1)
        layout.addLayout(flowLayout)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(12)
        metrics.setVerticalSpacing(12)
        cards = [
            self.metricCard("quota", "SerpApi 剩余额度", "green"),
            self.metricCard("today", "今日新搜索", "blue"),
            self.metricCard("company", "公司结果", "orange"),
            self.metricCard("review", "待人工审核", "red"),
        ]
        for column, card in enumerate(cards):
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        toolbar = QHBoxLayout()
        section = QLabel("启动条件")
        section.setObjectName("sectionTitle")
        toolbar.addWidget(section)
        toolbar.addStretch()
        self.refreshButton = QPushButton("刷新状态")
        self.refreshButton.setIcon(self.styleIcon(QStyle.SP_BrowserReload))
        self.refreshButton.clicked.connect(self.refreshRemote)
        toolbar.addWidget(self.refreshButton)
        openButton = QPushButton("打开输出目录")
        openButton.setIcon(self.styleIcon(QStyle.SP_DirOpenIcon))
        openButton.clicked.connect(self.openOutput)
        toolbar.addWidget(openButton)
        layout.addLayout(toolbar)

        self.checkTable = QTableWidget(0, 3)
        self.checkTable.setHorizontalHeaderLabels(["启动条件", "结论", "当前信息"])
        self.checkTable.verticalHeader().setVisible(False)
        self.checkTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.checkTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.checkTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.checkTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.checkTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.checkTable, 1)
        return page

    def buildTask(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(18)
        modeTitle = QLabel("处理模式")
        modeTitle.setObjectName("sectionTitle")
        layout.addWidget(modeTitle)
        modeRow = QHBoxLayout()
        modeRow.setSpacing(0)
        self.modeGroup = QButtonGroup(self)
        self.modeGroup.setExclusive(True)
        self.modeButtons = {}
        for index, (key, label) in enumerate(
            (("both", "公司 + 个人"), ("company", "仅公司"), ("person", "仅个人"))
        ):
            button = QPushButton(label)
            button.setObjectName("modeButton")
            button.setCheckable(True)
            button.setProperty("mode", key)
            button.setMinimumWidth(140)
            self.modeGroup.addButton(button, index)
            self.modeButtons[key] = button
            button.clicked.connect(self.updateModeView)
            modeRow.addWidget(button)
        modeRow.addStretch()
        self.modeButtons[str(self.model.config.get("runMode") or "both")].setChecked(True)
        layout.addLayout(modeRow)
        modeBand = QFrame()
        modeBand.setObjectName("modeSummaryBand")
        modeBandLayout = QGridLayout(modeBand)
        modeBandLayout.setContentsMargins(16, 12, 16, 12)
        modeBandLayout.setHorizontalSpacing(24)
        self.modeSummary = {}
        for column, key in enumerate(("target", "rule", "quota", "output")):
            titleLabel = QLabel("")
            titleLabel.setObjectName("policyTitle")
            valueLabel = QLabel("")
            valueLabel.setObjectName("modeSummaryValue")
            valueLabel.setWordWrap(True)
            modeBandLayout.addWidget(titleLabel, 0, column)
            modeBandLayout.addWidget(valueLabel, 1, column)
            self.modeSummary[key] = (titleLabel, valueLabel)
        layout.addWidget(modeBand)

        optionsTitle = QLabel("本次运行")
        optionsTitle.setObjectName("sectionTitle")
        layout.addWidget(optionsTitle)
        optionRow = QHBoxLayout()
        batchLabel = QLabel("每类数量")
        batchLabel.setObjectName("policyTitle")
        optionRow.addWidget(batchLabel)
        self.companyBatchInput = QSpinBox()
        self.companyBatchInput.setRange(1, 5000)
        self.companyBatchInput.setSuffix(" 家")
        self.companyBatchInput.setValue(int(self.model.config.get("companyBatch", 5)))
        self.companyBatchInput.setMinimumWidth(120)
        optionRow.addWidget(self.companyBatchInput)
        dailyCapLabel = QLabel("每日新搜索上限")
        dailyCapLabel.setObjectName("policyTitle")
        optionRow.addWidget(dailyCapLabel)
        self.dailyCapInput = QSpinBox()
        self.dailyCapInput.setRange(1, 5000)
        self.dailyCapInput.setSuffix(" 次 / 天")
        self.dailyCapInput.setValue(int(self.model.config.get("dailySerpCap", 6)))
        self.dailyCapInput.setMinimumWidth(130)
        self.dailyCapInput.valueChanged.connect(self.updateModeView)
        optionRow.addWidget(self.dailyCapInput)
        optionRow.addStretch()
        self.runSaveButton = QPushButton("保存运行设置")
        self.runSaveButton.setIcon(self.styleIcon(QStyle.SP_DialogSaveButton))
        self.runSaveButton.clicked.connect(self.saveRunSettings)
        optionRow.addWidget(self.runSaveButton)
        layout.addLayout(optionRow)

        scopeTitle = QLabel("数据处理范围")
        scopeTitle.setObjectName("sectionTitle")
        layout.addWidget(scopeTitle)
        self.scopeTable = QTableWidget(3, 3)
        self.scopeTable.setObjectName("scopeTable")
        self.scopeTable.setHorizontalHeaderLabels(["数据源", "处理方式", "保存内容"])
        self.scopeTable.verticalHeader().setVisible(False)
        self.scopeTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.scopeTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.scopeTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.scopeTable.setFixedHeight(150)
        layout.addWidget(self.scopeTable)

        policyTitle = QLabel("搜索额度策略")
        policyTitle.setObjectName("sectionTitle")
        layout.addWidget(policyTitle)
        policy = QFrame()
        policy.setObjectName("policyBand")
        policyLayout = QGridLayout(policy)
        policyLayout.setContentsMargins(16, 14, 16, 14)
        policyLayout.setHorizontalSpacing(30)
        labels = [
            ("公司目标", "5 / 天"),
            ("个人目标", "5 / 天"),
            ("新搜索上限", "6 / 天"),
            ("额度分配", "公司 6"),
            ("月末预留", "20 次"),
        ]
        self.policyLabels = []
        for column, (title, value) in enumerate(labels):
            titleLabel = QLabel(title)
            titleLabel.setObjectName("policyTitle")
            valueLabel = QLabel(value)
            valueLabel.setObjectName("policyValue")
            policyLayout.addWidget(titleLabel, 0, column)
            policyLayout.addWidget(valueLabel, 1, column)
            self.policyLabels.append((titleLabel, valueLabel))
        layout.addWidget(policy)
        layout.addStretch()
        startRow = QHBoxLayout()
        startRow.addStretch()
        startButton = QPushButton("开始完整流程")
        startButton.setObjectName("primaryButton")
        startButton.setIcon(self.styleIcon(QStyle.SP_MediaPlay))
        startButton.setMinimumWidth(172)
        startButton.clicked.connect(self.startRun)
        startRow.addWidget(startButton)
        layout.addLayout(startRow)
        self.updateModeView()
        return page

    def buildMonitor(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        top = QHBoxLayout()
        title = QLabel("任务进度")
        title.setObjectName("sectionTitle")
        top.addWidget(title)
        top.addStretch()
        self.pauseButton = QPushButton("暂停")
        self.pauseButton.setIcon(self.styleIcon(QStyle.SP_MediaPause))
        self.pauseButton.setEnabled(False)
        self.pauseButton.clicked.connect(self.togglePause)
        top.addWidget(self.pauseButton)
        self.stopButton = QPushButton("安全停止")
        self.stopButton.setObjectName("dangerButton")
        self.stopButton.setIcon(self.styleIcon(QStyle.SP_MediaStop))
        self.stopButton.setEnabled(False)
        self.stopButton.clicked.connect(self.stopRun)
        top.addWidget(self.stopButton)
        clearButton = QPushButton("清空日志")
        clearButton.clicked.connect(lambda: self.logBox.clear())
        top.addWidget(clearButton)
        layout.addLayout(top)
        self.progressLabel = QLabel("等待任务")
        self.progressLabel.setObjectName("progressLabel")
        layout.addWidget(self.progressLabel)
        self.progressBar = QProgressBar()
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(0)
        self.progressBar.setTextVisible(False)
        layout.addWidget(self.progressBar)
        self.logBox = QPlainTextEdit()
        self.logBox.setObjectName("logBox")
        self.logBox.setReadOnly(True)
        self.logBox.setPlaceholderText("运行日志将在这里显示")
        layout.addWidget(self.logBox, 1)
        return page

    def buildMail(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        modeTitle = QLabel("邮件处理模式")
        modeTitle.setObjectName("sectionTitle")
        layout.addWidget(modeTitle)
        modeRow = QHBoxLayout()
        modeRow.setSpacing(0)
        self.mailModeGroup = QButtonGroup(self)
        self.mailModeGroup.setExclusive(True)
        self.mailModeButtons = {}
        for index, (key, label) in enumerate(
            (("draft", "生成邮箱草稿（默认）"), ("send", "真实发送"))
        ):
            button = QPushButton(label)
            button.setObjectName("mailModeButton")
            button.setCheckable(True)
            button.setMinimumWidth(190)
            self.mailModeGroup.addButton(button, index)
            self.mailModeButtons[key] = button
            button.clicked.connect(self.updateMailMode)
            modeRow.addWidget(button)
        modeRow.addStretch()
        selectedMode = str(self.model.config.get("promotionMode") or "draft")
        self.mailModeButtons[selectedMode if selectedMode in self.mailModeButtons else "draft"].setChecked(True)
        layout.addLayout(modeRow)

        configTitle = QLabel("阿里邮箱配置")
        configTitle.setObjectName("sectionTitle")
        layout.addWidget(configTitle)
        configBand = QFrame()
        configBand.setObjectName("mailConfigBand")
        configLayout = QGridLayout(configBand)
        configLayout.setContentsMargins(16, 14, 16, 14)
        configLayout.setHorizontalSpacing(14)
        configLayout.setVerticalSpacing(10)
        accountLabel = QLabel("邮箱账号")
        codeLabel = QLabel("第三方客户端安全密码")
        subjectLabel = QLabel("邮件主题")
        self.mailAccountInput = QLineEdit(str(self.model.config.get("promotionSenderEmail") or ""))
        self.mailCodeInput = QLineEdit(str(self.model.config.get("promotionSmtpAuthCode") or ""))
        self.mailCodeInput.setEchoMode(QLineEdit.Password)
        self.mailSubjectInput = QLineEdit(str(self.model.config.get("promotionSubject") or ""))
        configLayout.addWidget(accountLabel, 0, 0)
        configLayout.addWidget(self.mailAccountInput, 0, 1)
        configLayout.addWidget(codeLabel, 0, 2)
        configLayout.addWidget(self.mailCodeInput, 0, 3)
        configLayout.addWidget(subjectLabel, 1, 0)
        configLayout.addWidget(self.mailSubjectInput, 1, 1, 1, 3)
        layout.addWidget(configBand)

        bodyTitle = QLabel("邮件正文")
        bodyTitle.setObjectName("sectionTitle")
        layout.addWidget(bodyTitle)
        self.mailBodyInput = QPlainTextEdit()
        self.mailBodyInput.setObjectName("mailBody")
        self.mailBodyInput.setPlainText(
            str(self.model.config.get("promotionBody") or self.model.mail.defaultPromotionBody)
        )
        self.mailBodyInput.setMinimumHeight(155)
        layout.addWidget(self.mailBodyInput, 1)
        self.mailLogoLabel = QLabel("正文 Logo：正在检查 file/time2renew-logo.png")
        self.mailLogoLabel.setObjectName("filterNote")
        layout.addWidget(self.mailLogoLabel)

        summaryBand = QFrame()
        summaryBand.setObjectName("policyBand")
        summaryLayout = QGridLayout(summaryBand)
        summaryLayout.setContentsMargins(16, 12, 16, 12)
        self.mailSummaryValues = {}
        for column, (key, title) in enumerate(
            (
                ("eligible", "本次可处理"),
                ("skipped", "重复跳过"),
                ("drafts", "历史草稿"),
                ("sent", "历史发送"),
                ("unique", "唯一联系人"),
            )
        ):
            titleLabel = QLabel(title)
            titleLabel.setObjectName("policyTitle")
            valueLabel = QLabel("0")
            valueLabel.setObjectName("policyValue")
            summaryLayout.addWidget(titleLabel, 0, column)
            summaryLayout.addWidget(valueLabel, 1, column)
            self.mailSummaryValues[key] = valueLabel
        layout.addWidget(summaryBand)

        actionRow = QHBoxLayout()
        self.mailConnectionLabel = QLabel("IMAP 草稿 / SMTP 发送")
        self.mailConnectionLabel.setObjectName("filterNote")
        actionRow.addWidget(self.mailConnectionLabel)
        actionRow.addStretch()
        self.mailSaveButton = QPushButton("保存邮件配置")
        self.mailSaveButton.clicked.connect(self.saveMailSettings)
        actionRow.addWidget(self.mailSaveButton)
        refreshButton = QPushButton("刷新去重统计")
        refreshButton.setIcon(self.styleIcon(QStyle.SP_BrowserReload))
        refreshButton.clicked.connect(self.refreshMail)
        actionRow.addWidget(refreshButton)
        self.mailActionButton = QPushButton("生成邮箱草稿")
        self.mailActionButton.setObjectName("primaryButton")
        self.mailActionButton.clicked.connect(self.startMail)
        actionRow.addWidget(self.mailActionButton)
        layout.addLayout(actionRow)
        self.refreshMail()
        return page

    def buildReview(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        toolbar = QHBoxLayout()
        title = QLabel("联系方式审核")
        title.setObjectName("sectionTitle")
        toolbar.addWidget(title)
        self.reviewFilter = QComboBox()
        self.reviewFilter.addItem("待审核", "pending")
        self.reviewFilter.addItem("已通过", "approved")
        self.reviewFilter.addItem("已拒绝", "rejected")
        self.reviewFilter.addItem("全部", "all")
        self.reviewFilter.currentIndexChanged.connect(self.refreshReview)
        toolbar.addWidget(self.reviewFilter)
        self.reviewSelectAll = QCheckBox("全选当前列表")
        self.reviewSelectAll.stateChanged.connect(self.toggleReviewSelection)
        toolbar.addWidget(self.reviewSelectAll)
        self.reviewSelectedLabel = QLabel("已选择 0 条")
        self.reviewSelectedLabel.setObjectName("filterNote")
        toolbar.addWidget(self.reviewSelectedLabel)
        toolbar.addStretch()
        refreshButton = QPushButton("刷新")
        refreshButton.setIcon(self.styleIcon(QStyle.SP_BrowserReload))
        refreshButton.clicked.connect(self.refreshReview)
        toolbar.addWidget(refreshButton)
        previewButton = QPushButton("进入邮件设置")
        previewButton.setIcon(self.styleIcon(QStyle.SP_FileDialogNewFolder))
        previewButton.clicked.connect(self.previewPromotion)
        toolbar.addWidget(previewButton)
        layout.addLayout(toolbar)

        self.reviewTable = QTableWidget(0, 7)
        self.reviewTable.setHorizontalHeaderLabels(
            ["ID", "类型", "对象", "邮箱 / 电话", "来源", "置信度", "状态"]
        )
        self.reviewTable.verticalHeader().setVisible(False)
        self.reviewTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.reviewTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.reviewTable.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.reviewTable.itemSelectionChanged.connect(self.updateReviewSelection)
        self.reviewTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.reviewTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.reviewTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.reviewTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.reviewTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.reviewTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.reviewTable.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        layout.addWidget(self.reviewTable, 1)

        actions = QHBoxLayout()
        actions.addStretch()
        pendingButton = QPushButton("设为待审核")
        pendingButton.clicked.connect(lambda: self.setReview("pending"))
        actions.addWidget(pendingButton)
        rejectButton = QPushButton("拒绝")
        rejectButton.setObjectName("dangerButton")
        rejectButton.clicked.connect(lambda: self.setReview("rejected"))
        actions.addWidget(rejectButton)
        approveButton = QPushButton("通过")
        approveButton.setObjectName("primaryButton")
        approveButton.setIcon(self.styleIcon(QStyle.SP_DialogApplyButton))
        approveButton.clicked.connect(lambda: self.setReview("approved"))
        actions.addWidget(approveButton)
        layout.addLayout(actions)
        return page

    def buildHelp(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        title = QLabel("页面操作说明")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        order = QLabel(
            "推荐顺序：连接状态 → 运行设置 → 开始运行 → 任务监控 → "
            "结果审核 → 邮件设置 → 结果数据"
        )
        order.setObjectName("helpOrder")
        order.setWordWrap(True)
        layout.addWidget(order)
        rows = [
            ("工作台", "确认能否开始任务", "查看额度、结果数和启动条件；需要时点击刷新状态。", "额度以 SerpApi 官方账号为准；必需检查未通过时不要直接运行。"),
            ("运行设置", "设置本次处理范围", "确认公司数量和每日新搜索上限；默认生成草稿。", "公司从 TDI 名单按代理机构类牌照过滤并去重。"),
            ("任务监控", "观察完整流程", "运行后查看进度和日志；需要时暂停、继续或安全停止。", "出现验证码时按提示人工处理；任务运行中不要强制关闭程序。"),
            ("邮件设置", "创建服务器草稿或发信", "确认账号、客户端安全密码、主题和正文；默认生成阿里邮箱草稿，真实发送需二次确认。", "只有已通过且未重复处理的邮箱会进入本次操作；真实发送会立即联系外部收件人。"),
            ("结果审核", "人工确认联系方式", "按状态筛选，单选、多选或全选当前列表，再统一设为通过、拒绝或待审核。", "全选只作用于当前筛选结果；只有通过的邮箱才能进入邮件处理。"),
            ("结果数据", "查看和导出采集结果", "检查邮箱、电话、来源链接和审核状态。", "Facebook 只保存公开主页链接，不登录、不打开页面。"),
            ("连接状态", "检查外部服务路径", "确认 SerpApi、直连抓取和阿里邮箱配置状态。", "普通网页、SerpApi 和邮件服务均直连；凭据不要提交或分享。"),
            ("操作说明", "随时核对操作方法", "按页面名称查找用途、操作步骤和注意事项。", "本页只提供说明，不会启动搜索、修改审核状态或发送邮件。"),
        ]
        self.helpTable = QTableWidget(len(rows), 4)
        self.helpTable.setHorizontalHeaderLabels(["页面", "主要用途", "怎么操作", "需要注意"])
        self.helpTable.verticalHeader().setVisible(False)
        self.helpTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.helpTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.helpTable.setWordWrap(True)
        self.helpTable.setTextElideMode(Qt.ElideNone)
        self.helpTable.verticalHeader().setMinimumSectionSize(48)
        self.helpTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.helpTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.helpTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.helpTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        for row, values in enumerate(rows):
            for column, value in enumerate(values):
                self.helpTable.setItem(row, column, QTableWidgetItem(value))
        self.helpTable.resizeRowsToContents()
        layout.addWidget(self.helpTable, 1)
        return page

    def buildResults(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        toolbar = QHBoxLayout()
        title = QLabel("已保存结果")
        title.setObjectName("sectionTitle")
        toolbar.addWidget(title)
        self.resultCountLabel = QLabel("0 条记录")
        self.resultCountLabel.setObjectName("filterNote")
        toolbar.addWidget(self.resultCountLabel)
        toolbar.addStretch()
        refreshButton = QPushButton("刷新")
        refreshButton.setIcon(self.styleIcon(QStyle.SP_BrowserReload))
        refreshButton.clicked.connect(self.refreshResults)
        toolbar.addWidget(refreshButton)
        openButton = QPushButton("打开输出目录")
        openButton.setIcon(self.styleIcon(QStyle.SP_DirOpenIcon))
        openButton.clicked.connect(self.openOutput)
        toolbar.addWidget(openButton)
        layout.addLayout(toolbar)

        self.resultTable = QTableWidget(0, 7)
        self.resultTable.setHorizontalHeaderLabels(
            ["类型", "对象", "邮箱", "电话", "来源链接", "采集状态", "更新时间"]
        )
        self.resultTable.verticalHeader().setVisible(False)
        self.resultTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.resultTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.resultTable.setSelectionMode(QAbstractItemView.SingleSelection)
        self.resultTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.resultTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.resultTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.resultTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.resultTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.resultTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.resultTable.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        layout.addWidget(self.resultTable, 1)
        return page

    def buildEnvironment(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        toolbar = QHBoxLayout()
        title = QLabel("网络与配置")
        title.setObjectName("sectionTitle")
        toolbar.addWidget(title)
        toolbar.addStretch()
        configButton = QPushButton("打开配置目录")
        configButton.setIcon(self.styleIcon(QStyle.SP_DirOpenIcon))
        configButton.clicked.connect(self.openConfigDir)
        toolbar.addWidget(configButton)
        layout.addLayout(toolbar)
        self.environmentTable = QTableWidget(0, 3)
        self.environmentTable.setHorizontalHeaderLabels(["服务", "网络路径", "当前状态"])
        self.environmentTable.verticalHeader().setVisible(False)
        self.environmentTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.environmentTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.environmentTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.environmentTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.environmentTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.environmentTable, 1)
        return page

    def applyTheme(self):
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background:#F4F6F5; color:#1D2926; font-family:'Microsoft YaHei UI'; font-size:13px; }
            #sidebar { background:#182421; }
            #sidebar QLabel { background:transparent; }
            #brand { color:#FFFFFF; font-size:22px; font-weight:700; line-height:1.25; }
            #edition { color:#8FC9BA; font-size:10px; font-weight:600; }
            #safety { color:#9FB0AA; font-size:11px; line-height:1.5; }
            #navigation { background:transparent; border:0; outline:0; color:#C8D2CE; }
            #navigation::item { border-radius:5px; padding:8px 10px; margin:1px 0; }
            #navigation::item:selected { background:#285C50; color:#FFFFFF; }
            #navigation::item:hover:!selected { background:#22332E; color:#FFFFFF; }
            #pageTitle { font-size:22px; font-weight:700; }
            #pageDetail { color:#66736F; font-size:12px; }
            #sectionTitle { font-size:15px; font-weight:700; }
            #stateBadge { background:#E7F5F0; color:#0F6B59; border:1px solid #B9DED3; border-radius:6px; padding:7px 12px; font-weight:600; }
            #workflowStep { background:#FFFFFF; border:1px solid #DCE3E0; border-radius:6px; }
            #workflowStep QLabel { background:transparent; }
            #flowIndex { background:#E8F1EE; color:#285C50; border-radius:6px; font-weight:700; }
            #flowTitle { font-size:12px; font-weight:700; }
            #flowState { color:#687771; font-size:11px; }
            #filterNote { color:#687771; font-size:11px; }
            #metricCard { background:#FFFFFF; border:1px solid #DCE3E0; border-radius:6px; }
            #metricCard QLabel, #policyBand QLabel { background:transparent; }
            #metricCard[tone='green'] { border-top:3px solid #0F766E; }
            #metricCard[tone='blue'] { border-top:3px solid #2563EB; }
            #metricCard[tone='orange'] { border-top:3px solid #B45309; }
            #metricCard[tone='red'] { border-top:3px solid #B42318; }
            #metricTitle, #policyTitle { color:#66736F; font-size:11px; font-weight:600; }
            #metricValue { font-size:25px; font-weight:700; }
            #metricNote { color:#77837F; font-size:11px; }
            #policyBand, #modeSummaryBand, #mailConfigBand { background:#FFFFFF; border:1px solid #DCE3E0; border-radius:6px; }
            #modeSummaryBand QLabel, #mailConfigBand QLabel { background:transparent; }
            #policyValue { font-size:17px; font-weight:700; }
            #modeSummaryValue { color:#273632; font-size:12px; font-weight:600; }
            QPushButton { min-height:34px; padding:0 13px; background:#FFFFFF; border:1px solid #C9D2CE; border-radius:5px; }
            QPushButton:hover { background:#EEF3F1; border-color:#9BAAA4; }
            QPushButton:focus { border:2px solid #2E7D6F; }
            QPushButton:disabled { color:#9FA9A5; background:#EDF0EF; border-color:#DDE2E0; }
            #primaryButton { background:#0F766E; color:#FFFFFF; border-color:#0F766E; font-weight:600; }
            #primaryButton:hover { background:#0B655E; }
            #dangerButton { color:#A62B23; border-color:#D8AAA5; }
            #dangerButton:hover { background:#FFF0EE; border-color:#C8756D; }
            #mailModeButton { border-radius:0; min-height:38px; }
            #mailModeButton:checked { background:#285C50; color:#FFFFFF; border-color:#285C50; }
            QCheckBox { spacing:9px; min-height:28px; }
            QComboBox, QSpinBox { min-height:34px; padding:0 10px; background:#FFFFFF; border:1px solid #C9D2CE; border-radius:5px; }
            QLineEdit { min-height:34px; padding:0 10px; background:#FFFFFF; border:1px solid #C9D2CE; border-radius:5px; }
            QTableWidget { background:#FFFFFF; alternate-background-color:#F7F9F8; border:1px solid #DCE3E0; border-radius:5px; gridline-color:#E7ECEA; selection-background-color:#DDEFEA; selection-color:#1D2926; }
            QHeaderView::section { background:#EEF2F0; color:#43504C; border:0; border-bottom:1px solid #D5DDDA; padding:9px 8px; font-weight:600; }
            QTableWidget::item { padding:7px; }
            QProgressBar { min-height:8px; max-height:8px; border:0; background:#DDE4E1; border-radius:4px; }
            QProgressBar::chunk { background:#0F766E; border-radius:4px; }
            #progressLabel { color:#52615C; }
            #logBox { background:#17211E; color:#D8E7E1; border:1px solid #2C3A35; border-radius:5px; padding:10px; font-family:Consolas; font-size:12px; }
            #mailBody { background:#FFFFFF; color:#273632; border:1px solid #C9D2CE; border-radius:5px; padding:10px; }
            #helpOrder { background:#E8F3EF; color:#234B41; border-left:4px solid #0F766E; padding:10px 12px; }
            """
        )
        self.checkTable.setAlternatingRowColors(True)
        self.scopeTable.setAlternatingRowColors(True)
        self.reviewTable.setAlternatingRowColors(True)
        self.resultTable.setAlternatingRowColors(True)
        self.environmentTable.setAlternatingRowColors(True)
        self.helpTable.setAlternatingRowColors(True)

    def switchPage(self, index):
        if index < 0:
            return
        self.stack.setCurrentIndex(index)
        self.pageTitle.setText(self.pageNames[index])
        if index == 7:
            self.helpTable.resizeRowsToContents()
        details = [
            "流程状态、免费额度与启动条件",
            "处理模式、数据范围与额度策略",
            "实时进度、运行日志和安全控制",
            "阿里邮箱草稿、真实发送和去重台账",
            "公开联系方式人工确认",
            "联系方式、来源链接与导出状态",
            "凭据、直连抓取和外部服务连接状态",
            "每个页面的操作步骤与使用注意事项",
        ]
        self.pageDetail.setText(details[index])

    def selectedMode(self):
        for mode, button in self.modeButtons.items():
            if button.isChecked():
                return mode
        return "both"

    def updateModeView(self):
        mode = self.selectedMode()
        dailyCap = self.dailyCapInput.value()
        batch = self.companyBatchInput.value()
        companyCap = (dailyCap + 1) // 2
        personCap = dailyCap // 2
        views = {
            "company": {
                "summary": [
                    ("每日目标", f"公司 {batch}"),
                    ("候选规则", "公司名去重（所有牌照类型）"),
                    ("搜索额度", f"公司最多 {dailyCap} 次"),
                    ("结果文件", "TDI 公司联系信息表"),
                ],
                "scope": [
                    ("TDI 名单", "公司名去重（所有牌照类型）", "公司、牌照类型、城市、州"),
                    ("Google 与普通网站", "公司名 + 城市州搜索并核对", "公司邮箱、电话、来源链接"),
                    ("Facebook", "仅识别主页链接", "公司 Facebook 链接"),
                ],
                "policy": [
                    ("公司目标", f"{batch} / 天"),
                    ("个人目标", "关闭"),
                    ("新搜索上限", f"{dailyCap} / 天"),
                    ("额度分配", f"公司 {dailyCap}"),
                    ("月末预留", "20 次"),
                ],
            },
            "person": {
                "summary": [
                    ("每日目标", f"个人 {batch}"),
                    ("候选规则", "Expiration date 为 2026"),
                    ("搜索额度", f"个人最多 {dailyCap} 次"),
                    ("结果文件", "TDI 个人联系信息表"),
                ],
                "scope": [
                    ("证书名单", "Expiration date 2026 筛选 + NPN 去重", "姓名、牌照类型、城市、州、到期日"),
                    ("Google 与普通网站", "人名 + insurance 搜索并核对", "个人邮箱、电话、来源链接"),
                    ("Facebook", "仅识别主页链接", "个人 Facebook 链接"),
                ],
                "policy": [
                    ("公司目标", "关闭"),
                    ("个人目标", f"{batch} / 天"),
                    ("新搜索上限", f"{dailyCap} / 天"),
                    ("额度分配", f"个人 {dailyCap}"),
                    ("月末预留", "20 次"),
                ],
            },
            "both": {
                "summary": [
                    ("每日目标", f"公司 {batch} + 个人 {batch}"),
                    ("候选规则", "公司去重 + 个人 2026 到期"),
                    ("搜索额度", f"公司 {companyCap} + 个人 {personCap}"),
                    ("结果文件", "公司表 + 个人表"),
                ],
                "scope": [
                    ("TDI 名单", "公司名去重（所有牌照类型）", "公司候选"),
                    ("证书名单", "Expiration date 2026 + NPN 去重", "个人候选"),
                    ("Google 与普通网站", "两类对象交替搜索", "邮箱、电话、来源链接"),
                ],
                "policy": [
                    ("公司目标", f"{batch} / 天"),
                    ("个人目标", f"{batch} / 天"),
                    ("新搜索上限", f"{dailyCap} / 天"),
                    ("额度分配", f"公司 {companyCap} + 个人 {personCap}"),
                    ("月末预留", "20 次"),
                ],
            },
        }
        view = views.get(mode, views["company"])
        for key, pair in zip(("target", "rule", "quota", "output"), view["summary"]):
            self.modeSummary[key][0].setText(pair[0])
            self.modeSummary[key][1].setText(pair[1])
        self.scopeTable.setRowCount(len(view["scope"]))
        for row, values in enumerate(view["scope"]):
            for column, value in enumerate(values):
                self.scopeTable.setItem(row, column, QTableWidgetItem(value))
        for labels, values in zip(self.policyLabels, view["policy"]):
            labels[0].setText(values[0])
            labels[1].setText(values[1])
        self.updateDailyMetric()

    def updateDailyMetric(self):
        if "today" not in self.metricValues:
            return
        dailyCap = self.dailyCapInput.value()
        mode = self.selectedMode()
        self.metricValues["today"].setText(f"{self.todayUsed} / {dailyCap}")
        if mode == "both":
            companyCap = (dailyCap + 1) // 2
            personCap = dailyCap // 2
            note = f"合并模式公司 {companyCap}、个人 {personCap}"
        elif mode == "company":
            note = f"仅公司模式最多 {dailyCap} 次"
        else:
            note = f"仅个人模式最多 {dailyCap} 次"
        self.metricNotes["today"].setText(note)

    def runRuntimeConfig(self):
        return {
            "runMode": self.selectedMode(),
            "companyBatch": self.companyBatchInput.value(),
            "personBatch": self.companyBatchInput.value(),
            "dailySerpCap": self.dailyCapInput.value(),
            "proxyRequired": False,
            "useBrowser": False,
            "useDirectFallback": True,
        }

    def saveRunSettings(self):
        values = self.runRuntimeConfig()
        path = self.baseDir / "config.local.json"
        try:
            current = {}
            if path.exists():
                current = json.loads(path.read_text(encoding="utf-8-sig"))
                if not isinstance(current, dict):
                    current = {}
            current.update(values)
            path.write_text(json.dumps(current, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            self.model = Main({"baseDir": str(self.baseDir), **values})
        except Exception as error:
            QMessageBox.critical(self, "保存失败", str(error))
            return
        self.updateState("ready", f"运行设置已保存（公司 {values['companyBatch']} 家，日上限 {values['dailySerpCap']} 次）")
        self.updateModeView()

    def selectedMailMode(self):
        for mode, button in self.mailModeButtons.items():
            if button.isChecked():
                return mode
        return "draft"

    def mailRuntimeConfig(self):
        return {
            "promotionMode": self.selectedMailMode(),
            "promotionSenderEmail": self.mailAccountInput.text().strip(),
            "promotionSmtpAuthCode": self.mailCodeInput.text(),
            "promotionSubject": self.mailSubjectInput.text().strip(),
            "promotionBody": self.mailBodyInput.toPlainText(),
        }

    def updateMailMode(self):
        sending = self.selectedMailMode() == "send"
        self.mailActionButton.setText("确认真实发送" if sending else "生成邮箱草稿")
        self.mailActionButton.setObjectName("dangerButton" if sending else "primaryButton")
        self.mailActionButton.style().unpolish(self.mailActionButton)
        self.mailActionButton.style().polish(self.mailActionButton)
        self.refreshMail()

    def saveMailSettings(self):
        values = self.mailRuntimeConfig()
        path = self.baseDir / "config.local.json"
        try:
            current = {}
            if path.exists():
                current = json.loads(path.read_text(encoding="utf-8-sig"))
                if not isinstance(current, dict):
                    current = {}
            current.update(values)
            path.write_text(json.dumps(current, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            self.model = Main({"baseDir": str(self.baseDir), **values})
        except Exception as error:
            QMessageBox.critical(self, "保存失败", str(error))
            return
        self.updateState("ready", "邮件配置已保存到本机")
        self.populateEnvironment(self.model.status(refreshQuota=False))

    def refreshMail(self):
        if not hasattr(self, "mailSummaryValues"):
            return
        try:
            runtime = self.mailRuntimeConfig()
            main = Main({"baseDir": str(self.baseDir), **runtime})
            records = main.mail.approvedRecords(self.selectedMailMode())
            summary = main.data.mailActionSummary()
            skipped = len(main.mail.skippedRecords)
        except Exception as error:
            self.mailConnectionLabel.setText(f"配置读取失败：{str(error)[:100]}")
            return
        self.mailSummaryValues["eligible"].setText(str(len(records)))
        self.mailSummaryValues["skipped"].setText(str(skipped))
        self.mailSummaryValues["drafts"].setText(str(summary.get("drafts", 0)))
        self.mailSummaryValues["sent"].setText(str(summary.get("sent", 0)))
        self.mailSummaryValues["unique"].setText(str(summary.get("uniqueRecipients", 0)))
        logoPath = main.mail.logoPath()
        self.mailLogoLabel.setText(
            f"正文 Logo：已加载 {logoPath.name}，草稿和真实发送都会内嵌显示"
            if logoPath.is_file()
            else f"正文 Logo：未找到 {logoPath}"
        )

    def startMail(self):
        config = self.mailRuntimeConfig()
        try:
            main = Main({"baseDir": str(self.baseDir), **config})
            records = main.mail.approvedRecords(self.selectedMailMode())
            skipped = len(main.mail.skippedRecords)
        except Exception as error:
            QMessageBox.critical(self, "邮件配置无效", str(error))
            return
        if not records:
            QMessageBox.information(self, "没有待处理邮件", f"没有新的审核通过邮箱；去重跳过 {skipped} 个。")
            self.refreshMail()
            return
        if self.selectedMailMode() == "send":
            answer = QMessageBox.warning(
                self,
                "确认真实发送",
                f"将立即向 {len(records)} 个唯一邮箱真实发送推广邮件。此操作不会自动撤回。",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
        self.navigation.setCurrentRow(2)
        self.startWorker("mail", config)

    def runtimeConfig(self):
        return {**self.runRuntimeConfig(), **self.mailRuntimeConfig()}

    def populateChecks(self, checks):
        self.checkTable.setRowCount(len(checks))
        for row, check in enumerate(checks):
            status = "通过" if check.get("ok") else "需处理"
            values = [check.get("name", ""), status, check.get("detail", "")]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 1:
                    item.setForeground(Qt.darkGreen if check.get("ok") else Qt.darkRed)
                self.checkTable.setItem(row, column, item)
        self.checkTable.resizeRowsToContents()

    def populateEnvironment(self, status):
        config = self.model.config
        rows = [
            ("TDI 公司名单", "本地 xlsx", f"公司结果 {int(status.get('companyResults') or 0):,} 条"),
            ("证书名单", "本地 xlsx", f"个人结果 {int(status.get('personResults') or 0):,} 条"),
            ("SerpApi", "直连", "Key 已配置" if config.get("serpapiKey") else "Key 未配置"),
            ("普通二级页面", "直连 HTTP 抓取", "不启动代理桥，不读取 ipfiy.py"),
            ("Facebook 主页", "不发起页面请求", "仅记录搜索结果和普通网页发现的链接"),
            ("阿里邮箱", "IMAP / SMTP SSL 直连", "账号和授权码已配置" if config.get("promotionSenderEmail") and config.get("promotionSmtpAuthCode") else "账号或授权码未配置"),
            ("SQLite", "本地", str(status.get("database") or "")),
            ("初始化配置", "Main.__init__ 默认值", "已加载"),
        ]
        self.environmentTable.setRowCount(len(rows))
        for row, values in enumerate(rows):
            for column, value in enumerate(values):
                self.environmentTable.setItem(row, column, QTableWidgetItem(str(value)))
        self.environmentTable.resizeRowsToContents()

    def applyStatus(self, status):
        status = self.mergeOfficialQuota(status)
        self.rememberOfficialQuota(status)
        quota = status.get("quota") or {}
        remaining = int(quota.get("remaining") or 0)
        allowance = int(quota.get("allowance") or 250)
        today = int(quota.get("todayUsed") or 0)
        source = str(quota.get("source") or "本地预算")
        if source.startswith("本地预算（官方刷新失败）"):
            self.metricValues["quota"].setText("官方失败")
            note = "SerpApi 官方额度刷新失败"
            if quota.get("remoteError"):
                note += f"：{str(quota['remoteError'])[:80]}"
            self.metricNotes["quota"].setText(note)
        else:
            self.metricValues["quota"].setText(f"{remaining} / {allowance}")
            self.metricNotes["quota"].setText(f"已用 {int(quota.get('used') or 0)} · 来源：{source}")
        self.todayUsed = today
        self.updateDailyMetric()
        self.metricValues["company"].setText(f"{int(status.get('companyResults') or 0):,}")
        personCount = int(status.get("personResults") or 0)
        self.metricNotes["company"].setText(
            f"公司 {int(status.get('companyResults') or 0):,} · 个人 {personCount:,} · 已通过 {int(status.get('approvedResults') or 0)}"
        )
        self.metricValues["review"].setText(str(int(status.get("reviewPending") or 0)))
        self.metricNotes["review"].setText("通过后才能生成推广记录")
        requiredChecks = [item for item in status.get("checks") or [] if item.get("level") == "required"]
        ready = all(item.get("ok") for item in requiredChecks)
        self.flowStates[1].setText("已通过" if ready else "需处理")
        self.flowStates[2].setText("已就绪" if self.model.data.xlsxPath.exists() else "缺少名单")
        modeNames = {"both": "公司 + 个人", "company": "仅公司", "person": "仅个人"}
        self.flowStates[3].setText(modeNames.get(self.selectedMode(), "仅公司"))
        if self.stateBadge.text() not in {"运行中", "已暂停", "停止中"}:
            self.flowStates[4].setText("待启动")
        self.flowStates[5].setText(f"待审核 {int(status.get('reviewPending') or 0)}")
        self.populateChecks(list(status.get("checks") or []))
        self.populateEnvironment(status)

    def rememberOfficialQuota(self, status):
        """缓存最近一次 SerpApi 官方额度，避免后续本地刷新覆盖首屏。"""
        quota = (status or {}).get("quota") or {}
        if str(quota.get("source") or "") == "SerpApi":
            self.officialQuota = dict(quota)

    def mergeOfficialQuota(self, status):
        """本地状态刷新只更新业务数据，额度继续使用已获取的官方值。"""
        status = dict(status or {})
        quota = status.get("quota") or {}
        if self.officialQuota and str(quota.get("source") or "").startswith("本地预算"):
            status["quota"] = dict(self.officialQuota)
            checks = []
            for item in list(status.get("checks") or []):
                value = dict(item)
                if value.get("name") == "SerpApi 免费额度":
                    value["detail"] = (
                        f"剩余 {self.officialQuota['remaining']} / "
                        f"{self.officialQuota['allowance']}，来源：{self.officialQuota['source']}"
                    )
                checks.append(value)
            status["checks"] = checks
        return status

    def refreshOfficialStartup(self):
        """首屏直接读取 SerpApi 官方额度，不展示本地预算作为最终值。"""
        try:
            self.model = Main({"baseDir": str(self.baseDir)})
            status = self.model.status(refreshQuota=True)
            self.applyStatus(status)
            self.writeStartupDiagnostic(status)
            self.appendLog("SerpApi 官方额度已读取。")
        except Exception as error:
            self.appendLog(f"SerpApi 官方额度读取失败：{error}")
            self.refreshLocal()

    def writeStartupDiagnostic(self, status):
        """测试打包 exe 时可选写出首屏状态；普通用户运行不写。"""
        if os.environ.get("TDI_WRITE_STARTUP_STATUS") != "1":
            return
        try:
            path = self.baseDir / "output" / "startup_status.json"
            path.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "quota": (status or {}).get("quota") or {},
                "checks": [
                    {
                        "name": item.get("name", ""),
                        "ok": bool(item.get("ok")),
                        "detail": item.get("detail", ""),
                    }
                    for item in list((status or {}).get("checks") or [])
                ],
            }
            path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as error:
            self.appendLog(f"启动状态诊断写入失败：{error}")

    def refreshLocal(self):
        try:
            self.model = Main({"baseDir": str(self.baseDir)})
            self.applyStatus(self.model.status(refreshQuota=False))
        except Exception as error:
            self.appendLog(f"本地状态读取失败：{error}")

    def markQuotaRefreshing(self):
        """启动后官方额度未返回前，避免把本地预算显示成最终结果。"""
        if "quota" in self.metricValues:
            self.metricValues["quota"].setText("读取中")
        if "quota" in self.metricNotes:
            self.metricNotes["quota"].setText("正在获取 SerpApi 官方额度...")
        if hasattr(self, "checkTable"):
            for row in range(self.checkTable.rowCount()):
                nameItem = self.checkTable.item(row, 0)
                if nameItem and nameItem.text() == "SerpApi 免费额度":
                    infoItem = self.checkTable.item(row, 2)
                    if infoItem:
                        infoItem.setText("正在获取 SerpApi 官方额度...")
                    break
        self.appendLog("正在获取 SerpApi 官方额度...")

    def refreshRemote(self):
        self.startWorker("refresh", {})

    def startRun(self):
        config = self.runtimeConfig()
        try:
            main = Main({"baseDir": str(self.baseDir), **config})
            checks = main.preflight()
        except Exception as error:
            QMessageBox.critical(self, "无法启动", str(error))
            return
        blockers = [item for item in checks if item.get("level") == "required" and not item.get("ok")]
        if blockers:
            detail = "\n".join(f"• {item['name']}：{item['detail']}" for item in blockers)
            QMessageBox.warning(self, "运行条件不完整", detail)
            self.navigation.setCurrentRow(6)
            return
        self.navigation.setCurrentRow(2)
        self.startWorker("run", config)

    def previewPromotion(self):
        self.navigation.setCurrentRow(3)
        self.refreshMail()

    def startWorker(self, operation, config):
        if self.thread and self.thread.isRunning():
            QMessageBox.information(self, "任务运行中", "请先等待当前任务结束或安全停止。")
            return
        self.thread = QThread(self)
        self.worker = self.Worker(self.baseDir, operation, config)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self.appendLog)
        self.worker.progress.connect(self.updateProgress)
        self.worker.state.connect(self.updateState)
        self.worker.humanRequired.connect(self.handleHuman)
        self.worker.finished.connect(self.workerFinished)
        self.worker.failed.connect(self.workerFailed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.clearWorker)
        self.setBusy(True, operation)
        self.thread.start()

    def setBusy(self, busy, operation=""):
        self.headerStart.setEnabled(not busy)
        self.refreshButton.setEnabled(not busy)
        running = busy and operation == "run"
        self.pauseButton.setEnabled(running)
        self.stopButton.setEnabled(running)
        self.dailyCapInput.setEnabled(not busy)
        self.companyBatchInput.setEnabled(not busy)
        self.runSaveButton.setEnabled(not busy)
        self.mailSaveButton.setEnabled(not busy)
        self.mailActionButton.setEnabled(not busy)
        for button in self.mailModeButtons.values():
            button.setEnabled(not busy)
        if busy:
            self.updateState("running", "正在执行" if operation != "refresh" else "正在刷新")

    def togglePause(self):
        if not self.worker:
            return
        if self.worker.pauseEvent.is_set():
            self.worker.resume()
            self.pauseButton.setText("暂停")
            self.pauseButton.setIcon(self.styleIcon(QStyle.SP_MediaPause))
        else:
            self.worker.pause()
            self.pauseButton.setText("继续")
            self.pauseButton.setIcon(self.styleIcon(QStyle.SP_MediaPlay))

    def stopRun(self):
        if not self.worker:
            return
        answer = QMessageBox.question(
            self,
            "安全停止",
            "当前对象完成或到达安全点后停止，已保存数据不会丢失。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.worker.stop()
            self.stopButton.setEnabled(False)

    def handleHuman(self, reason, url):
        message = reason
        if url:
            message += f"\n\n当前页面：{url}"
        dialog = QMessageBox(self)
        dialog.setWindowTitle("需要人工处理")
        dialog.setText(message)
        dialog.setIcon(QMessageBox.Information)
        continueButton = dialog.addButton("已完成，继续", QMessageBox.AcceptRole)
        dialog.addButton("保持暂停", QMessageBox.RejectRole)
        dialog.exec()
        if dialog.clickedButton() == continueButton and self.worker:
            self.worker.resume()
            self.pauseButton.setText("暂停")

    def appendLog(self, message):
        self.logBox.appendPlainText(str(message))
        bar = self.logBox.verticalScrollBar()
        bar.setValue(bar.maximum())

    def updateProgress(self, current, total, message):
        self.progressLabel.setText(message)
        if total > 0:
            self.progressBar.setRange(0, 100)
            self.progressBar.setValue(min(100, int(current * 100 / total)))
        else:
            self.progressBar.setRange(0, 0)

    def updateState(self, state, message):
        labels = {
            "running": "运行中",
            "paused": "已暂停",
            "stopping": "停止中",
            "stopped": "已停止",
            "ready": "就绪",
        }
        self.stateBadge.setText(labels.get(state, state or "就绪"))
        self.progressLabel.setText(message)
        if hasattr(self, "flowStates") and 4 in self.flowStates:
            self.flowStates[4].setText(labels.get(state, state or "待启动"))

    def workerFinished(self, operation, result):
        if operation == "refresh":
            self.applyStatus(result)
            self.writeStartupDiagnostic(result)
            self.appendLog("SerpApi 官方额度和环境状态已刷新。")
        elif operation == "preview":
            self.appendLog(f"推广预览已生成：{int(result.get('total') or 0)} 个审核通过邮箱。")
            QMessageBox.information(self, "推广预览已生成", f"记录数：{int(result.get('total') or 0)}\n文件：{result.get('recordFile') or ''}")
        elif operation == "mail":
            if result.get("mode") == "draft":
                message = f"邮箱草稿已创建 {int(result.get('drafted') or 0)} 封，重复跳过 {int(result.get('skipped') or 0)} 封。"
                title = "阿里邮箱草稿完成"
            else:
                message = f"真实发送成功 {int(result.get('sent') or 0)} 封，重复跳过 {int(result.get('skipped') or 0)} 封。"
                title = "真实发送完成"
            if result.get("failed"):
                message += f" 失败 {int(result.get('failed') or 0)} 封。"
            self.appendLog(message)
            QMessageBox.information(self, title, message)
        else:
            self.appendLog(f"流程结束：公司完成 {int(result.get('completed') or 0)}，待审核 {int(result.get('reviewPending') or 0)}。")
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(100 if operation == "run" else 0)
        self.updateState("ready", "任务完成")
        if operation != "refresh":
            self.refreshLocal()
        self.refreshReview()
        self.refreshResults()
        self.refreshMail()

    def workerFailed(self, operation, details):
        self.appendLog(details)
        self.updateState("stopped", "任务失败")
        summary = details.strip().splitlines()[-1] if details.strip() else "未知错误"
        QMessageBox.critical(self, "任务失败", summary)

    def clearWorker(self):
        self.setBusy(False)
        self.pauseButton.setText("暂停")
        self.pauseButton.setIcon(self.styleIcon(QStyle.SP_MediaPause))
        if self.worker:
            self.worker.deleteLater()
        if self.thread:
            self.thread.deleteLater()
        self.worker = None
        self.thread = None

    def refreshReview(self):
        if not hasattr(self, "reviewTable"):
            return
        status = str(self.reviewFilter.currentData() or "pending")
        try:
            data = Main({"baseDir": str(self.baseDir)}).data
            rows = data.reviewItems(status)
        except Exception as error:
            self.appendLog(f"审核数据读取失败：{error}")
            return
        self.reviewTable.setRowCount(len(rows))
        statusNames = {"pending": "待审核", "approved": "已通过", "rejected": "已拒绝"}
        for row, item in enumerate(rows):
            contact = "\n".join(value for value in (item.get("emailText"), item.get("phoneText")) if value)
            values = [
                item.get("id", ""),
                "公司" if item.get("mode") == "company" else "个人",
                item.get("object_name", ""),
                contact,
                item.get("source_url", ""),
                item.get("confidence", 0),
                statusNames.get(str(item.get("status") or ""), item.get("status", "")),
            ]
            for column, value in enumerate(values):
                self.reviewTable.setItem(row, column, QTableWidgetItem(str(value)))
        self.reviewTable.resizeRowsToContents()
        self.reviewSelectAll.blockSignals(True)
        self.reviewSelectAll.setChecked(False)
        self.reviewSelectAll.blockSignals(False)
        self.updateReviewSelection()

    def toggleReviewSelection(self, state):
        if state:
            self.reviewTable.selectAll()
        else:
            self.reviewTable.clearSelection()

    def updateReviewSelection(self):
        selected = self.reviewTable.selectionModel().selectedRows()
        count = len(selected)
        total = self.reviewTable.rowCount()
        self.reviewSelectedLabel.setText(f"已选择 {count} 条")
        self.reviewSelectAll.blockSignals(True)
        self.reviewSelectAll.setChecked(total > 0 and count == total)
        self.reviewSelectAll.blockSignals(False)

    def refreshResults(self):
        if not hasattr(self, "resultTable"):
            return
        try:
            data = Main({"baseDir": str(self.baseDir)}).data
            rows = data.contactResults()
        except Exception as error:
            self.appendLog(f"结果数据读取失败：{error}")
            return
        self.resultCountLabel.setText(f"{len(rows)} 条记录")
        self.resultTable.setRowCount(len(rows))
        for row, item in enumerate(rows):
            values = [
                "公司" if item.get("mode") == "company" else "个人",
                item.get("objectName", ""),
                "\n".join(item.get("emails") or []),
                "\n".join(item.get("phones") or []),
                "\n".join(item.get("verifiedUrls") or item.get("sourceUrls") or []),
                item.get("contactStatus", ""),
                item.get("updatedAt", ""),
            ]
            for column, value in enumerate(values):
                self.resultTable.setItem(row, column, QTableWidgetItem(str(value)))
        self.resultTable.resizeRowsToContents()

    def setReview(self, status):
        rows = sorted({index.row() for index in self.reviewTable.selectionModel().selectedRows()})
        if not rows:
            QMessageBox.information(self, "请选择记录", "请先选择一条或多条联系方式记录。")
            return
        itemIds = [int(self.reviewTable.item(row, 0).text()) for row in rows if self.reviewTable.item(row, 0)]
        if not itemIds:
            return
        try:
            data = Main({"baseDir": str(self.baseDir)}).data
            updated = data.setReviewStatuses(itemIds, status)
        except Exception as error:
            QMessageBox.critical(self, "审核失败", str(error))
            return
        statusNames = {"pending": "待审核", "approved": "通过", "rejected": "拒绝"}
        self.updateState("ready", f"已将 {updated} 条记录设为{statusNames.get(status, status)}")
        self.refreshReview()
        self.refreshLocal()

    def openOutput(self):
        self.model.outputDir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.model.outputDir.resolve())))

    def openConfigDir(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.baseDir.resolve())))

    def closeEvent(self, event):
        if self.thread and self.thread.isRunning() and self.worker:
            answer = QMessageBox.question(
                self,
                "退出程序",
                "任务仍在运行，是否请求安全停止并退出？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self.worker.stop()
            self.thread.quit()
            self.thread.wait(3000)
        event.accept()


if __name__ == "__main__":
    try:
        if "--diagnose-quota" in sys.argv:
            raise SystemExit(runQuotaDiagnostic())
        application = QApplication(sys.argv)
        application.setApplicationName("TDI 公司与个人合作推广")
        window = RunGui()
        window.show()
        raise SystemExit(application.exec())
    except Exception:
        if os.environ.get("TDI_WRITE_STARTUP_STATUS") == "1":
            try:
                baseDir = (
                    Path(sys.executable).resolve().parent
                    if getattr(sys, "frozen", False)
                    else Path(__file__).resolve().parent
                )
                path = baseDir / "output" / "fatal_error.txt"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(traceback.format_exc(), encoding="utf-8")
            except Exception:
                pass
        raise
