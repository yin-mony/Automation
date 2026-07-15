import queue
import re
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from analysis import CommentAnalyzer
from main import Auto


class RunGui:
    """亚马逊评论工具统一窗口"""

    class LogStream:
        """线程日志流，负责把任务输出写入界面队列和日志文件"""

        def __init__(self, logName="亚马逊评论工具.log", fallback=None):
            """初始化日志流"""
            # 程序目录，打包后使用 exe 所在目录
            if getattr(sys, "frozen", False):
                self.baseDir = Path(sys.executable).resolve().parent
            else:
                self.baseDir = Path(__file__).resolve().parent

            # 日志目录，不存在时自动创建
            self.logDir = self.baseDir / "logs"
            self.logDir.mkdir(parents=True, exist_ok=True)

            # 日志文件路径
            self.logPath = self.logDir / logName
            # 界面消费队列
            self.queue = queue.Queue()
            # 线程输出路由表
            self.routes = {}
            # 路由锁，避免多线程同时改写
            self.lock = threading.RLock()
            # 未命中路由时回落到原始输出
            self.fallback = fallback

        def routeThread(self, writer):
            """把当前线程的 print 输出路由到指定 writer"""
            # 登记当前线程对应的 writer
            with self.lock:
                self.routes[threading.get_ident()] = writer

        def clearThread(self):
            """清除当前线程的输出路由"""
            # 删除当前线程路由，避免任务结束后继续接管输出
            with self.lock:
                self.routes.pop(threading.get_ident(), None)

        def write(self, text):
            """兼容 sys.stdout.write，把文本写入当前线程日志"""
            # 空文本直接忽略
            if not text:
                return

            # 查找当前线程的日志 writer
            with self.lock:
                writer = self.routes.get(threading.get_ident())

            # 当前线程有路由时写入对应日志
            if writer:
                if text.strip():
                    writer(text.rstrip())
                return

            # 没有路由时交给原始输出
            if self.fallback:
                self.fallback.write(text)

        def flush(self):
            """兼容 stdout flush 接口"""
            # 原始输出存在时同步刷新
            if self.fallback:
                self.fallback.flush()

        def info(self, msg):
            """写入普通日志"""
            # 写入 INFO 级别日志
            self.push("INFO", msg)

        def warn(self, msg):
            """写入警告日志"""
            # 写入 WARN 级别日志
            self.push("WARN", msg)

        def error(self, msg):
            """写入错误日志"""
            # 写入 ERROR 级别日志
            self.push("ERROR", msg)

        def exception(self, msg):
            """写入异常日志和 traceback"""
            # 先写入错误摘要
            self.error(msg)
            # 再逐行写入异常堆栈
            for line in traceback.format_exc().splitlines():
                self.error(line)

        def push(self, level, msg):
            """写入界面队列和日志文件"""
            # 生成日志时间
            now = datetime.now().strftime("%H:%M:%S")
            # 组装界面日志行
            line = f"{now} - {level} - {msg}"
            # 推入界面队列
            self.queue.put(line)

            # 生成文件日志时间
            fileNow = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # 追加写入日志文件
            with open(self.logPath, "a", encoding="utf-8") as file:
                file.write(f"{fileNow} - {level} - {msg}\n")

        def poll(self):
            """取出当前队列内的全部日志"""
            # 收集本轮日志
            messages = []
            # 持续取出队列中已有日志
            while True:
                try:
                    messages.append(self.queue.get_nowait())
                except queue.Empty:
                    break
            return messages

    class DownloadPage(ttk.Frame):
        """评论下载配置与运行页面"""

        def __init__(self, parent, outputLog=None):
            """初始化评论下载页面"""
            # 初始化 Frame
            super().__init__(parent, padding=12)
            # 统一输出路由
            self.outputLog = outputLog
            # 当前页日志流
            self.log = RunGui.LogStream("亚马逊评论工具-下载.log")
            # 运行状态
            self.isRunning = False
            # 当前后台线程
            self.currentThread = None
            # 当前自动化实例
            self.autoTask = None

            # 默认配置值放在实例属性中
            self.defaultFilePath = r"C:\RPA流程\亚马逊评论分析\flie"
            self.ipPattern = re.compile(r"^(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}$")
            self.ipChars = re.compile(r"^[\d., ]*$")
            self.siteNames = [
                "美国", "加拿大", "墨西哥", "巴西",
                "英国", "法国", "德国", "意大利", "西班牙", "荷兰", "瑞典", "波兰", "比利时", "爱尔兰",
                "日本", "新加坡", "澳大利亚", "印度", "阿联酋", "沙特阿拉伯", "土耳其", "埃及", "南非",
            ]

            # 界面变量
            self.username = tk.StringVar(value="")
            self.password = tk.StringVar(value="")
            self.siteName = tk.StringVar(value="美国")
            self.shopIp = tk.StringVar(value="")
            self.shopPort = tk.StringVar(value="8945")
            self.amazonEmail = tk.StringVar(value="")
            self.amazonPassword = tk.StringVar(value="")
            self.filePath = tk.StringVar(value=self.defaultFilePath)

            # 构建界面并加载配置
            self.buildUi()
            self.loadConfig()
            self.normalizeIp()
            self.processLog()
            self.log.info(f"下载页已就绪。日志文件: {self.log.logPath}")

        def buildUi(self):
            """构建下载页界面"""
            # 创建表单容器
            form = ttk.Frame(self, padding=8)
            form.pack(fill="both", expand=True)

            # 易得客账号输入
            row = 0
            ttk.Label(form, text="易得客账号:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.username, width=50).grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # 易得客密码输入
            row += 1
            ttk.Label(form, text="易得客密码:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.password, width=50, show="*").grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # 店铺 IP 输入
            row += 1
            ttk.Label(form, text="店铺 IP（英文逗号分隔）:").grid(row=row, column=0, sticky="w", pady=4)
            self.shopIpEntry = ttk.Entry(form, textvariable=self.shopIp, width=72)
            self.shopIpEntry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)
            validCmd = (self.register(self.validateIpTyping), "%P")
            self.shopIpEntry.configure(validate="key", validatecommand=validCmd)
            self.shopIpEntry.bind("<FocusOut>", self.normalizeIp)
            self.shopIpEntry.bind("<<Paste>>", self.pasteIp)

            # 调试端口输入
            row += 1
            ttk.Label(form, text="调试端口（与 IP 个数一致）:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.shopPort, width=72).grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # 店铺站点选择
            row += 1
            ttk.Label(form, text="店铺站点:").grid(row=row, column=0, sticky="w", pady=4)
            siteCombo = ttk.Combobox(form, textvariable=self.siteName, values=self.siteNames, state="readonly", width=48)
            siteCombo.grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # Amazon 后台账号输入
            row += 1
            ttk.Label(form, text="Amazon 邮箱:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.amazonEmail, width=72).grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # Amazon 后台密码输入
            row += 1
            ttk.Label(form, text="Amazon 密码:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.amazonPassword, width=72, show="*").grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            # 保存目录输入
            row += 1
            ttk.Label(form, text="评论保存目录:").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.filePath, width=62).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
            ttk.Button(form, text="浏览", width=10, command=self.selectDir).grid(row=row, column=2, sticky="e", pady=4)

            # ASIN 输入框
            row += 1
            asinRow = row
            ttk.Label(form, text="商品 ASIN（逗号或换行）:").grid(row=row, column=0, sticky="nw", pady=4)
            asinFrame = ttk.Frame(form)
            asinFrame.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=8, pady=4)
            self.asinText = tk.Text(asinFrame, height=8, width=70, font=("Consolas", 9), wrap=tk.WORD)
            asinScroll = ttk.Scrollbar(asinFrame, orient=tk.VERTICAL, command=self.asinText.yview)
            self.asinText.configure(yscrollcommand=asinScroll.set)
            self.asinText.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            asinScroll.pack(side=tk.RIGHT, fill=tk.Y)

            # 操作按钮
            row += 1
            btnRow = ttk.Frame(form)
            btnRow.grid(row=row, column=0, columnspan=3, pady=12)
            self.runBtn = ttk.Button(btnRow, text="开始下载评论", command=self.runDownload, width=18)
            self.runBtn.pack(side=tk.LEFT, padx=6)
            self.stopBtn = ttk.Button(btnRow, text="强制停止", command=self.forceStop, width=12, state=tk.DISABLED)
            self.stopBtn.pack(side=tk.LEFT, padx=6)

            # 设置表单自适应
            form.columnconfigure(1, weight=1)
            form.rowconfigure(asinRow, weight=1)

            # 状态栏与日志区
            ttk.Separator(self).pack(fill="x", pady=8)
            ctrl = ttk.Frame(self)
            ctrl.pack(fill="x")
            self.statusLabel = ttk.Label(ctrl, text="就绪", foreground="green")
            self.statusLabel.pack(side=tk.LEFT)
            logFrame = ttk.LabelFrame(self, text="下载日志", padding=8)
            logFrame.pack(fill="both", expand=True, pady=(8, 0))
            self.logText = scrolledtext.ScrolledText(logFrame, height=14, wrap=tk.WORD, font=("Consolas", 9))
            self.logText.pack(fill=tk.BOTH, expand=True)

        def normalizeIp(self, event=None):
            """规范化店铺 IP 标点"""
            # 将常见中文句号替换为英文点
            text = self.shopIp.get()
            for full in ("\uff0e", "\u3002", "\uff61"):
                text = text.replace(full, ".")
            self.shopIp.set(text)

        def pasteIp(self, event):
            """粘贴 IP 时过滤非法字符"""
            try:
                # 读取剪贴板文本
                clip = self.clipboard_get()
            except tk.TclError:
                return "break"

            # 只保留 IP 输入允许字符
            for full in ("\uff0e", "\u3002", "\uff61"):
                clip = clip.replace(full, ".")
            text = "".join(ch for ch in clip if ord(ch) < 128 and ch in "0123456789., ")
            text = text.replace("\n", ",").replace("\r", "").strip()
            if not text:
                return "break"

            # 替换当前选区或插入到光标位置
            widget = event.widget
            try:
                widget.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            widget.insert("insert", text)
            return "break"

        def selectDir(self):
            """选择评论 Excel 保存目录"""
            # 打开目录选择窗口
            path = filedialog.askdirectory(title="选择评论 Excel 保存目录")
            if path:
                self.filePath.set(path)

        def getBaseDir(self):
            """获取配置文件所在目录"""
            # 打包后配置写到 exe 所在目录
            if getattr(sys, "frozen", False):
                return Path(sys.executable).resolve().parent
            # 源码运行时配置写到当前脚本目录
            return Path(__file__).resolve().parent

        def processLog(self):
            """刷新下载日志文本框"""
            # 取出本轮日志并写入文本框
            for msg in self.log.poll():
                self.logText.insert(tk.END, msg + "\n")
                self.logText.see(tk.END)
            # 定时继续刷新
            self.after(100, self.processLog)

        def setRunning(self, running):
            """切换运行中状态"""
            # 保存当前运行状态
            self.isRunning = running
            if running:
                self.runBtn.config(state=tk.DISABLED)
                self.stopBtn.config(state=tk.NORMAL)
                self.statusLabel.config(text="下载运行中...", foreground="orange")
            else:
                self.runBtn.config(state=tk.NORMAL)
                self.stopBtn.config(state=tk.DISABLED)
                self.statusLabel.config(text="就绪", foreground="green")
                self.autoTask = None

        def parseList(self, text):
            """按逗号或换行解析列表"""
            # 统一换行和逗号后拆分
            return [item.strip() for item in text.replace("\n", ",").split(",") if item.strip()]

        def parseAsin(self):
            """解析 ASIN 多行输入"""
            # 按行和逗号拆分 ASIN
            asinList = []
            for line in self.asinText.get("1.0", tk.END).splitlines():
                for part in line.split(","):
                    value = part.strip()
                    if value:
                        asinList.append(value)
            return asinList

        def validIp(self, token):
            """校验 IPv4 字符串"""
            # 使用正则校验完整 IPv4
            return bool(self.ipPattern.fullmatch((token or "").strip()))

        def validateIpTyping(self, proposed):
            """限制 IP 输入框只允许数字、点、逗号与空格"""
            # 空文本允许
            if proposed == "":
                return True
            # 非 ASCII 字符不允许直接输入
            if any(ord(ch) > 127 for ch in proposed):
                return False
            return self.ipChars.fullmatch(proposed) is not None

        def buildConfig(self):
            """校验界面输入并组装下载配置"""
            # 规范化 IP 文本后解析配置项
            self.normalizeIp()
            ipList = self.parseList(self.shopIp.get())
            portRaw = self.parseList(self.shopPort.get())
            asinList = self.parseAsin()
            filePath = self.filePath.get().strip()
            username = self.username.get().strip()
            password = self.password.get()
            siteName = self.siteName.get().strip() or "美国"
            amazonEmail = self.amazonEmail.get().strip()
            amazonPassword = self.amazonPassword.get()

            # 校验账号密码
            if not username or not password:
                raise ValueError("请填写易得客账号与密码。")
            # 校验 IP 列表
            if not ipList:
                raise ValueError("请填写至少一个店铺 IP。")
            badIp = [ip for ip in ipList if not self.validIp(ip)]
            if badIp:
                raise ValueError("以下为无效 IPv4：\n" + "\n".join(badIp[:8]))

            # 校验端口列表
            try:
                portList = [int(item) for item in portRaw]
            except ValueError as exc:
                raise ValueError("端口必须为整数，多个时用英文逗号分隔。") from exc
            if len(portList) != len(ipList):
                raise ValueError(f"IP 数量 ({len(ipList)}) 与端口数量 ({len(portList)}) 不一致。")

            # 校验保存目录和 ASIN
            if not filePath:
                raise ValueError("请选择评论保存目录。")
            Path(filePath).mkdir(parents=True, exist_ok=True)
            if not asinList:
                raise ValueError("请填写至少一个商品 ASIN。")

            return {
                "yidekeUsername": username,
                "yidekePassword": password,
                "username": username,
                "password": password,
                "autoSiteName": siteName,
                "siteName": siteName,
                "amazonEmail": amazonEmail,
                "amazonPassword": amazonPassword,
                "ip": ipList,
                "port": portList,
                "shopIp": ipList,
                "shopPort": portList,
                "experts": asinList,
                "asinList": asinList,
                "file_path": filePath,
                "filePath": filePath,
            }

        def saveConfig(self):
            """保存下载页配置"""
            # 读取 ASIN 输入并转义换行
            cfg = self.getBaseDir() / "comment_download_gui_config.txt"
            asinText = self.asinText.get("1.0", tk.END).strip().replace("\n", "\\n")
            try:
                with open(cfg, "w", encoding="utf-8") as file:
                    file.write(f"username={self.username.get()}\n")
                    file.write(f"password={self.password.get()}\n")
                    file.write(f"auto_site_name={self.siteName.get()}\n")
                    file.write(f"shop_ip={self.shopIp.get()}\n")
                    file.write(f"shop_port={self.shopPort.get()}\n")
                    file.write(f"amazon_email={self.amazonEmail.get()}\n")
                    file.write(f"amazon_password={self.amazonPassword.get()}\n")
                    file.write(f"file_path={self.filePath.get()}\n")
                    file.write(f"experts={asinText}\n")
                self.log.info("已保存下载配置")
            except OSError as exc:
                self.log.warn(f"保存下载配置失败: {exc}")

        def loadConfig(self):
            """加载下载页配置"""
            # 配置文件不存在时使用默认值
            cfg = self.getBaseDir() / "comment_download_gui_config.txt"
            if not cfg.exists():
                self.log.info("未找到下载配置，使用默认值。")
                return

            try:
                asinText = None
                with open(cfg, "r", encoding="utf-8") as file:
                    for line in file:
                        if line.startswith("username="):
                            self.username.set(line.split("=", 1)[1].strip())
                        elif line.startswith("password="):
                            self.password.set(line.split("=", 1)[1].strip())
                        elif line.startswith("auto_site_name="):
                            self.siteName.set(line.split("=", 1)[1].strip() or "美国")
                        elif line.startswith("shop_ip="):
                            self.shopIp.set(line.split("=", 1)[1].strip())
                        elif line.startswith("shop_port="):
                            self.shopPort.set(line.split("=", 1)[1].strip())
                        elif line.startswith("amazon_email="):
                            self.amazonEmail.set(line.split("=", 1)[1].strip())
                        elif line.startswith("amazon_password="):
                            self.amazonPassword.set(line.split("=", 1)[1].strip())
                        elif line.startswith("file_path="):
                            self.filePath.set(line.split("=", 1)[1].strip())
                        elif line.startswith("experts="):
                            asinText = line.split("=", 1)[1].strip().replace("\\n", "\n")
                if asinText is not None:
                    self.asinText.delete("1.0", tk.END)
                    self.asinText.insert("1.0", asinText)
                self.log.info("已加载下载配置")
            except OSError as exc:
                self.log.warn(f"加载下载配置失败: {exc}")

        def runDownload(self):
            """启动后台线程执行评论下载"""
            # 避免重复启动下载任务
            if self.isRunning:
                messagebox.showwarning("提示", "评论下载任务正在运行中，请稍候。")
                return

            try:
                config = self.buildConfig()
            except ValueError as exc:
                messagebox.showwarning("参数错误", str(exc))
                return

            # 保存配置并切换运行状态
            self.saveConfig()
            self.logText.delete("1.0", tk.END)
            self.setRunning(True)

            def target():
                """后台线程执行下载任务"""
                # 将当前线程 print 输出路由到下载日志
                router = self.outputLog or self.log
                router.routeThread(self.log.info)
                try:
                    self.log.info("=" * 50)
                    self.log.info("开始下载：Auto(config).run()")
                    self.log.info(f"店铺站点: {config['autoSiteName']}")
                    self.log.info(f"店铺 IP: {config['ip']} | 端口: {config['port']}")
                    self.log.info(f"Amazon 邮箱: {'已填写' if config['amazonEmail'] else '未填写'}")
                    self.log.info(f"保存目录: {config['file_path']}")
                    self.log.info(f"ASIN 数量: {len(config['experts'])}")
                    self.log.info("=" * 50)

                    # 创建并运行下载任务
                    self.autoTask = Auto(config)
                    self.autoTask.run()

                    # 根据当前状态提示完成或中止
                    if self.isRunning:
                        self.log.info("评论下载流程已结束。")
                        self.after(0, lambda: self.finishTask(True, "评论下载完成，已导出 Excel。"))
                    else:
                        self.after(0, lambda: self.finishTask(False, "任务已中止"))
                except Exception as exc:
                    self.log.exception(f"下载任务出错: {exc}")
                    self.after(0, lambda msg=str(exc): self.finishTask(False, msg))
                finally:
                    router.clearThread()

            # 启动后台线程
            self.currentThread = threading.Thread(target=target, daemon=True)
            self.currentThread.start()

        def forceStop(self):
            """强制停止评论下载任务"""
            # 未运行时无需处理
            if not self.isRunning:
                return
            if not messagebox.askyesno("确认停止", "将终止 Chrome/eDecker 进程并强制退出程序，确定继续？"):
                return

            # 调用业务对象强制结束进程
            self.isRunning = False
            self.log.warn("用户请求强制停止...")
            if self.autoTask is not None:
                self.autoTask.stopProgram()
            else:
                self.log.warn("任务实例尚未创建，仅标记停止。")

        def finishTask(self, success, message):
            """任务结束后恢复界面状态并弹窗提示"""
            # 恢复按钮和状态栏
            self.setRunning(False)
            if success:
                messagebox.showinfo("完成", message)
            else:
                messagebox.showerror("错误", message)

    class AnalysisPage(ttk.Frame):
        """AI 分析配置与运行页面"""

        def __init__(self, parent, outputLog=None):
            """初始化 AI 分析页面"""
            # 初始化 Frame
            super().__init__(parent, padding=12)
            # 统一输出路由
            self.outputLog = outputLog
            # 当前页日志流
            self.log = RunGui.LogStream("亚马逊评论工具-分析.log")
            # 运行状态
            self.isRunning = False
            # 当前后台线程
            self.currentThread = None
            # 默认 Excel 路径
            self.defaultExcelPath = r"C:\RPA流程\亚马逊评论分析\flie\亚马逊评论.xlsx"

            # 界面变量
            self.excelPath = tk.StringVar(value=self.defaultExcelPath)
            self.apiKey = tk.StringVar(value="")

            # 构建界面并加载配置
            self.buildUi()
            self.loadConfig()
            self.processLog()
            self.log.info(f"分析页已就绪。日志文件: {self.log.logPath}")

        def buildUi(self):
            """构建 AI 分析页界面"""
            # 创建表单容器
            form = ttk.Frame(self, padding=8)
            form.pack(fill="both", expand=True)

            # 评论 Excel 文件选择
            row = 0
            ttk.Label(form, text="评论 Excel 文件:").grid(row=row, column=0, sticky="w", pady=6)
            ttk.Entry(form, textvariable=self.excelPath, width=72).grid(row=row, column=1, sticky="ew", padx=8, pady=6)
            ttk.Button(form, text="浏览", width=10, command=self.selectExcel).grid(row=row, column=2, sticky="e", pady=6)

            # API Key 输入
            row += 1
            ttk.Label(form, text="OpenAI API Key:").grid(row=row, column=0, sticky="w", pady=6)
            ttk.Entry(form, textvariable=self.apiKey, width=72, show="*").grid(row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=6)

            # 操作按钮
            row += 1
            btnRow = ttk.Frame(form)
            btnRow.grid(row=row, column=0, columnspan=3, pady=12)
            self.runBtn = ttk.Button(btnRow, text="开始 AI 分析", command=self.runAnalysis, width=18)
            self.runBtn.pack(side=tk.LEFT, padx=6)
            self.stopBtn = ttk.Button(btnRow, text="停止（仅标记）", command=self.stopTask, width=14, state=tk.DISABLED)
            self.stopBtn.pack(side=tk.LEFT, padx=6)

            # 设置表单自适应
            form.columnconfigure(1, weight=1)

            # 状态栏与日志区
            ttk.Separator(self).pack(fill="x", pady=8)
            ctrl = ttk.Frame(self)
            ctrl.pack(fill="x")
            self.statusLabel = ttk.Label(ctrl, text="就绪", foreground="green")
            self.statusLabel.pack(side=tk.LEFT)
            logFrame = ttk.LabelFrame(self, text="分析日志", padding=8)
            logFrame.pack(fill="both", expand=True, pady=(8, 0))
            self.logText = scrolledtext.ScrolledText(logFrame, height=18, wrap=tk.WORD, font=("Consolas", 9))
            self.logText.pack(fill=tk.BOTH, expand=True)

        def selectExcel(self):
            """选择评论 Excel 文件"""
            # 打开 Excel 文件选择窗口
            path = filedialog.askopenfilename(
                title="选择评论 Excel 文件",
                filetypes=[("Excel", "*.xlsx *.xls"), ("所有文件", "*.*")],
            )
            if path:
                self.excelPath.set(path)

        def getBaseDir(self):
            """获取配置文件所在目录"""
            # 打包后配置写到 exe 所在目录
            if getattr(sys, "frozen", False):
                return Path(sys.executable).resolve().parent
            # 源码运行时配置写到当前脚本目录
            return Path(__file__).resolve().parent

        def processLog(self):
            """刷新 AI 分析日志文本框"""
            # 取出本轮日志并写入文本框
            for msg in self.log.poll():
                self.logText.insert(tk.END, msg + "\n")
                self.logText.see(tk.END)
            # 定时继续刷新
            self.after(100, self.processLog)

        def setRunning(self, running):
            """切换 AI 分析运行状态"""
            # 保存当前运行状态
            self.isRunning = running
            if running:
                self.runBtn.config(state=tk.DISABLED)
                self.stopBtn.config(state=tk.NORMAL)
                self.statusLabel.config(text="AI 分析运行中...", foreground="orange")
            else:
                self.runBtn.config(state=tk.NORMAL)
                self.stopBtn.config(state=tk.DISABLED)
                self.statusLabel.config(text="就绪", foreground="green")

        def buildParams(self):
            """校验界面输入并返回分析参数"""
            # 读取界面输入
            excelPath = self.excelPath.get().strip()
            apiKey = self.apiKey.get().strip()

            # 校验 Excel 路径
            if not excelPath:
                raise ValueError("请选择评论 Excel 文件。")
            if not excelPath.lower().endswith((".xlsx", ".xls")):
                raise ValueError("请选择 .xlsx 或 .xls 格式的 Excel 文件。")
            if not Path(excelPath).is_file():
                raise ValueError(f"Excel 文件不存在：\n{excelPath}")

            # 校验 API Key
            if not apiKey:
                raise ValueError("请填写 OpenAI API Key。")

            return excelPath, apiKey

        def saveConfig(self):
            """保存 AI 分析页配置"""
            # 写入本地配置文件
            cfg = self.getBaseDir() / "comment_analyzer_gui_config.txt"
            try:
                with open(cfg, "w", encoding="utf-8") as file:
                    file.write(f"excel_path={self.excelPath.get()}\n")
                    file.write(f"api_key={self.apiKey.get()}\n")
                self.log.info("已保存分析配置")
            except OSError as exc:
                self.log.warn(f"保存分析配置失败: {exc}")

        def loadConfig(self):
            """加载 AI 分析页配置"""
            # 配置文件不存在时使用默认值
            cfg = self.getBaseDir() / "comment_analyzer_gui_config.txt"
            if not cfg.exists():
                self.log.info("未找到分析配置，使用默认值。")
                return

            try:
                with open(cfg, "r", encoding="utf-8") as file:
                    for line in file:
                        if line.startswith("excel_path="):
                            self.excelPath.set(line.split("=", 1)[1].strip())
                        elif line.startswith("api_key="):
                            self.apiKey.set(line.split("=", 1)[1].strip())
                self.log.info("已加载分析配置")
            except OSError as exc:
                self.log.warn(f"加载分析配置失败: {exc}")

        def stopTask(self):
            """标记停止 AI 分析任务"""
            # 未运行时无需处理
            if not self.isRunning:
                return

            # AI 请求无法强杀，只标记状态
            self.isRunning = False
            self.log.info("已请求停止（AI 分析任务无法强制中断，仅作状态标记）。")
            self.statusLabel.config(text="已请求停止", foreground="orange")

        def runAnalysis(self):
            """启动后台线程执行 AI 分析"""
            # 避免重复启动分析任务
            if self.isRunning:
                messagebox.showwarning("提示", "AI 分析任务正在运行中，请稍候。")
                return

            try:
                excelPath, apiKey = self.buildParams()
            except ValueError as exc:
                messagebox.showwarning("参数错误", str(exc))
                return

            # 保存配置并切换运行状态
            self.saveConfig()
            self.logText.delete("1.0", tk.END)
            self.setRunning(True)

            def target():
                """后台线程执行分析任务"""
                # 将当前线程 print 输出路由到分析日志
                router = self.outputLog or self.log
                router.routeThread(self.log.info)
                try:
                    self.log.info("=" * 50)
                    self.log.info("开始分析：CommentAnalyzer(path, apiKey).run()")
                    self.log.info(f"Excel: {excelPath}")
                    self.log.info("=" * 50)

                    # 创建并运行分析任务
                    analyzer = CommentAnalyzer(excelPath, apiKey)
                    analyzer.run()

                    # 根据当前状态提示完成或中止
                    if self.isRunning:
                        self.log.info("AI 分析流程已结束。")
                        self.after(0, lambda: self.finishTask(True, "分析完成，报告已保存至 Excel 同目录下的「分析报告」文件夹。"))
                    else:
                        self.after(0, lambda: self.finishTask(False, "任务已中止"))
                except Exception as exc:
                    self.log.exception(f"分析任务出错: {exc}")
                    self.after(0, lambda msg=str(exc): self.finishTask(False, msg))
                finally:
                    router.clearThread()

            # 启动后台线程
            self.currentThread = threading.Thread(target=target, daemon=True)
            self.currentThread.start()

        def finishTask(self, success, message):
            """任务结束后恢复界面状态并弹窗提示"""
            # 恢复按钮和状态栏
            self.setRunning(False)
            if success:
                messagebox.showinfo("完成", message)
            else:
                messagebox.showerror("错误", message)

    def __init__(self, root=None, startTab="download"):
        """初始化统一窗口"""
        # 创建或复用 Tk 根窗口
        self.root = root or tk.Tk()
        # 默认启动页
        self.startTab = startTab
        # 原始 stdout 和 stderr
        self.stdout = sys.stdout
        self.stderr = sys.stderr
        # 统一输出路由，后台线程 print 会分发到对应页日志
        self.outputLog = RunGui.LogStream("亚马逊评论工具.log", fallback=self.stdout)
        # Notebook 控件
        self.notebook = None
        # 下载页
        self.downloadPage = None
        # 分析页
        self.analysisPage = None

    def buildUi(self):
        """构建统一窗口界面"""
        # 设置窗口基础属性
        self.root.title("亚马逊评论工具")
        self.root.geometry("960x740")
        self.root.minsize(780, 580)

        # 使用 clam 主题，保持 Windows 下控件一致
        style = ttk.Style()
        style.theme_use("clam")

        # 接管标准输出，后台任务 print 可进入界面日志
        sys.stdout = self.outputLog
        sys.stderr = self.outputLog

        # 创建双功能页 Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # 创建评论下载页
        self.downloadPage = self.DownloadPage(self.notebook, outputLog=self.outputLog)
        self.notebook.add(self.downloadPage, text="评论下载")

        # 创建 AI 分析页
        self.analysisPage = self.AnalysisPage(self.notebook, outputLog=self.outputLog)
        self.notebook.add(self.analysisPage, text="AI 分析")

        # 根据启动参数选择默认页
        if self.startTab == "analysis":
            self.notebook.select(self.analysisPage)
        else:
            self.notebook.select(self.downloadPage)

    def run(self):
        """启动统一 GUI 主循环"""
        # 构建界面
        self.buildUi()
        # 进入 Tk 主循环
        self.root.mainloop()


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "startTab": "download",
    }

    # 创建并启动统一窗口
    app = RunGui(startTab=config["startTab"])
    app.run()
