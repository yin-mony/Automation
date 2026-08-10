import queue
import threading
import tkinter as tk
from queue import Queue
from tkinter import messagebox, scrolledtext, ttk
from boss_web.auto import BossAuto
from boss_web.db import BossDb
from boss_web.login import BossLogin
from boss_web.job import BossJob
from boss_web.template import BossTemplate
from boss_web.report import BossReport
from boss_web.reply import BossReply

class RunGui:
    """BOSS GUI：扫码登录 + 推荐牛人与简历任务"""

    def __init__(self):
        """初始化 GUI 配置、运行状态与控件引用"""
        # 窗口标题与尺寸
        self.title = 'BOSS 直聘 · 招聘自动化'
        self.winWidth = 560
        self.winHeight = 1010
        # 分发标准目录
        self.appRootDir = 'D:\\boss_zhaopin_筛选简历'
        self.defaultUserDataPath = self.appRootDir + '\\boss_chrome_profile'
        # 登录与自动化实例
        self.login = BossLogin()
        self.reply = BossReply()
        self.bossAuto = None
        self.db = None
        # 工作线程
        self.workerThread = None
        # 日志队列（子线程 -> GUI 线程）
        self.logQueue = queue.Queue()
        self.autoTaskAfterLogin = True
        self.testCandidateName = ''
        self.root = None
        self.pathVar = None
        self.autoTaskVar = None
        self.logText = None
        self.loginBtn = None
        self.taskBtn = None
        self.stopBtn = None
        self.formalDoneBtn = None
        self.manualHandoffBtn = None
        self.manualReplyDoneBtn = None
        self.manualResumeBtn = None
        self.manualUnsuitableBtn = None
        self.manualReplyMode = ''
        self.recommendLimitVar = None
        self.recommendCountVar = None
        self.recommendSaveBtn = None
        self.unsuitableBtn = None
        self.unsuitableWin = None
        self.unsuitableListbox = None
        self.unsuitableRows = []
        self.wecomPushBtn = None
        self.wecomSaveBtn = None
        self.wecomPushThread = None
        self.wecomWebhookVar = None
        self.wecomMobileVar = None
        self.statusVar = None
        self.templateBtn = None
        self.templateEditorWin = None
        self.templateTypeVar = None
        self.templateListbox = None
        self.templateWordText = None
        self.templateEnabledVar = None
        self.selectedTemplateId = None
        self.templateRows = []
        self.templateListHintLabel = None
        self.tplPhCompanyVar = None
        self.tplPhJobVar = None
        self.tplPhNameVar = None
        self.tplPhAddressVar = None
        self.tplPhDurationVar = None
        self.tplPhDayOffsetVar = None
        self.tplPhTimeSlotsVar = None
        self.jobBtn = None
        self.jobEditorWin = None
        self.jobListbox = None
        self.jobListHintLabel = None
        self.selectedJobId = None
        self.jobRows = []
        self.jobNameVar = None
        self.jobMatchKeysVar = None
        self.jobIntroText = None
        self.jobAgeMinVar = None
        self.jobAgeMaxVar = None
        self.jobWorkYearsVar = None
        self.jobEducationVar = None
        self.jobMustVar = None
        self.jobPreferVar = None
        self.jobRejectVar = None
        self.jobAnyText = None
        self.jobEnabledVar = None
        self.replyConfigBtn = None
        self.replyConfigWin = None
        self.replyEnabledVar = None
        self.replyBaseUrlVar = None
        self.replyModelVar = None
        self.replyTimeoutVar = None
        self.replyConfigStatusVar = None
        self.replySkillListbox = None
        self.replySkillRows = []
        self.selectedReplySkillId = None
        self.replySkillNameVar = None
        self.replySkillEnabledVar = None
        self.replyInstructionText = None
        self.replyExamplesText = None
        self.replyTestBtn = None
        self.replyAssistBtn = None
        self.replyAssistWin = None
        self.replyAssistText = None
        self.replyAssistStatusVar = None
        self.replyAssistRecommendVar = None
        self.replyGenerateBtn = None
        self.replyFillBtn = None
        self.replyGenerateThread = None
        self.replyFillThread = None
        self.manualReplyInfo = {}
        self.replyResult = None

    def templateTypeLabels(self):
        """话术类型中文说明（下拉展示用）"""
        return {
            'greeting': '岗位介绍 greeting',
            'followup': '求简历跟进 followup',
            'interview_pre': '面试预邀请 interview_pre',
            'interview_ask_time': '追问面试时间 interview_ask_time',
            'interview_remind': '面试提醒 interview_remind',
            'interview_reschedule': '改期确认 interview_reschedule',
            'interview_cancel': '取消面试 interview_cancel',
            'reject': '审核不通过 reject',
            'reply_interest': '智能回复-感兴趣 reply_interest',
            'reply_resume': '智能回复-发简历 reply_resume',
            'reply_learn': '智能回复-了解岗位 reply_learn',
        }

    def templatePlaceholderHint(self):
        """占位符说明"""
        return '话术中可写占位符；下方「占位符取值」配置发送时的实际替换内容（{name}/{job} 优先用候选人信息，为空时用默认值）'

    def parsePlaceholderTimeSlots(self, text):
        """解析占位符配置中的面试时段整点列表"""
        parts = str(text or '').replace('，', ',').split(',')
        slots = []
        for part in parts:
            part = part.strip()
            if not part:
                continue
            try:
                slots.append(int(part))
            except ValueError:
                continue
        return slots or [14]

    def loadPlaceholderFormToGui(self):
        """从数据库加载占位符配置到话术模板窗口"""
        db = self.ensureDb()
        gs = BossJob().globalSettings
        cfg = db.getPlaceholderConfig(gs)
        if self.tplPhCompanyVar:
            self.tplPhCompanyVar.set(str(cfg.get('company') or ''))
        if self.tplPhJobVar:
            self.tplPhJobVar.set(str(cfg.get('jobDefault') or ''))
        if self.tplPhNameVar:
            self.tplPhNameVar.set(str(cfg.get('nameDefault') or ''))
        if self.tplPhAddressVar:
            self.tplPhAddressVar.set(str(cfg.get('address') or ''))
        if self.tplPhDurationVar:
            self.tplPhDurationVar.set(str(cfg.get('duration') or ''))
        if self.tplPhDayOffsetVar:
            self.tplPhDayOffsetVar.set(str(cfg.get('dayOffset') if cfg.get('dayOffset') is not None else 1))
        if self.tplPhTimeSlotsVar:
            self.tplPhTimeSlotsVar.set(str(cfg.get('timeSlots') or ''))

    def collectPlaceholderFromGui(self):
        """从话术模板窗口收集占位符配置"""
        try:
            dayOffset = int(str(self.tplPhDayOffsetVar.get() if self.tplPhDayOffsetVar else '1').strip() or 1)
        except ValueError:
            raise ValueError('面试日期偏移须为整数（0=今天，1=明天，2=后天）')
        if dayOffset < 0:
            raise ValueError('面试日期偏移不能为负数')
        slotsText = self.tplPhTimeSlotsVar.get().strip() if self.tplPhTimeSlotsVar else ''
        slots = self.parsePlaceholderTimeSlots(slotsText)
        return {
            'company': self.tplPhCompanyVar.get().strip() if self.tplPhCompanyVar else '',
            'jobDefault': self.tplPhJobVar.get().strip() if self.tplPhJobVar else '',
            'nameDefault': self.tplPhNameVar.get().strip() if self.tplPhNameVar else '',
            'address': self.tplPhAddressVar.get().strip() if self.tplPhAddressVar else '',
            'duration': self.tplPhDurationVar.get().strip() if self.tplPhDurationVar else '',
            'dayOffset': dayOffset,
            'timeSlots': ','.join(str(x) for x in slots),
        }

    def savePlaceholderSettings(self):
        """保存占位符配置到数据库"""
        try:
            data = self.collectPlaceholderFromGui()
        except ValueError as exc:
            messagebox.showwarning('提示', str(exc))
            return
        if not data.get('company'):
            messagebox.showwarning('提示', '公司名称不能为空')
            return
        db = self.ensureDb()
        db.savePlaceholderConfig(data)
        self.appendLog('已保存话术占位符配置')
        messagebox.showinfo('已保存', '占位符配置已保存，下次开始任务时生效')

    def restorePlaceholderDefaults(self):
        """恢复占位符为 job.py 内置默认值"""
        if not messagebox.askyesno('确认恢复', '确定将占位符配置恢复为代码默认值吗？\n当前自定义内容将被覆盖。'):
            return
        gs = BossJob().globalSettings
        defaults = BossDb().defaultPlaceholderConfig(gs)
        db = self.ensureDb()
        db.savePlaceholderConfig(defaults)
        self.loadPlaceholderFormToGui()
        self.appendLog('已恢复话术占位符默认配置')
        messagebox.showinfo('已恢复', '占位符配置已恢复默认')

    def openTemplateEditor(self):
        """打开话术模板管理窗口"""
        if self.templateEditorWin and self.templateEditorWin.winfo_exists():
            self.templateEditorWin.lift()
            self.templateEditorWin.focus_force()
            self.refreshTemplateList()
            return
        win = tk.Toplevel(self.root)
        win.title('话术模板管理')
        win.geometry('680x780')
        win.minsize(600, 640)
        self.templateEditorWin = win
        pad = {'padx': 10, 'pady': 6}
        top = ttk.Frame(win, padding=10)
        top.pack(fill=tk.BOTH, expand=True)
        ttk.Label(top, text='话术类型:').pack(anchor=tk.W)
        labels = self.templateTypeLabels()
        typeKeys = list(BossTemplate().types())
        displayValues = [labels.get(key, key) for key in typeKeys]
        self.templateTypeVar = tk.StringVar(value=displayValues[0] if displayValues else '')
        typeCombo = ttk.Combobox(top, textvariable=self.templateTypeVar, values=displayValues, state='readonly')
        typeCombo.pack(fill=tk.X, pady=(2, 0))
        typeCombo.bind('<<ComboboxSelected>>', lambda e: self.refreshTemplateList())
        self.templateListHintLabel = ttk.Label(top, text='当前类型全部话术（共 0 条，发送时随机选用已启用条目）')
        self.templateListHintLabel.pack(anchor=tk.W, pady=(8, 0))
        listFrame = ttk.Frame(top)
        listFrame.pack(fill=tk.X, pady=(2, 0))
        self.templateListbox = tk.Listbox(listFrame, height=4, exportselection=False)
        self.templateListbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        listScroll = ttk.Scrollbar(listFrame, orient=tk.VERTICAL, command=self.templateListbox.yview)
        listScroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.templateListbox.configure(yscrollcommand=listScroll.set)
        self.templateListbox.bind('<<ListboxSelect>>', self.onTemplateListSelect)
        ttk.Label(top, text='话术内容:').pack(anchor=tk.W, pady=(8, 0))
        self.templateWordText = scrolledtext.ScrolledText(top, height=5, font=('Microsoft YaHei UI', 9))
        self.templateWordText.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(top, text=self.templatePlaceholderHint(), foreground='#666', wraplength=640).pack(anchor=tk.W, pady=(4, 0))
        self.templateEnabledVar = tk.IntVar(value=1)
        ttk.Checkbutton(top, text='启用此条话术', variable=self.templateEnabledVar).pack(anchor=tk.W, pady=(4, 0))
        phGroup = ttk.LabelFrame(top, text='占位符取值配置')
        phGroup.pack(fill=tk.X, pady=(8, 0))
        phForm = ttk.Frame(phGroup, padding=8)
        phForm.pack(fill=tk.X)
        phForm.columnconfigure(1, weight=1)
        self.tplPhCompanyVar = tk.StringVar(value='')
        self.tplPhJobVar = tk.StringVar(value='')
        self.tplPhNameVar = tk.StringVar(value='')
        self.tplPhAddressVar = tk.StringVar(value='')
        self.tplPhDurationVar = tk.StringVar(value='')
        self.tplPhDayOffsetVar = tk.StringVar(value='1')
        self.tplPhTimeSlotsVar = tk.StringVar(value='')
        phFields = [
            ('{company} 公司名称', self.tplPhCompanyVar),
            ('{job} 默认岗位名（岗位为空时）', self.tplPhJobVar),
            ('{name} 默认称呼（姓名为空时）', self.tplPhNameVar),
            ('{address} 面试地址', self.tplPhAddressVar),
            ('{duration} 面试时长', self.tplPhDurationVar),
            ('{date} 日期偏移（0今天/1明天/2后天）', self.tplPhDayOffsetVar),
            ('{time} 可选整点时段（逗号分隔）', self.tplPhTimeSlotsVar),
        ]
        for rowIdx, (labelText, var) in enumerate(phFields):
            ttk.Label(phForm, text=labelText + ':').grid(row=rowIdx, column=0, sticky=tk.W, pady=2, padx=(0, 8))
            ttk.Entry(phForm, textvariable=var).grid(row=rowIdx, column=1, sticky=tk.EW, pady=2)
        phBtnRow = ttk.Frame(phGroup, padding=(8, 0, 8, 8))
        phBtnRow.pack(fill=tk.X)
        ttk.Button(phBtnRow, text='保存占位符配置', command=self.savePlaceholderSettings).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(phBtnRow, text='恢复占位符默认', command=self.restorePlaceholderDefaults).pack(side=tk.LEFT)
        btnRow = ttk.Frame(top)
        btnRow.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btnRow, text='新增', command=self.addTemplateRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='保存', command=self.saveTemplateRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='删除', command=self.deleteTemplateRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='恢复当前类型默认', command=self.restoreTemplateType).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='恢复全部默认', command=self.restoreAllTemplates).pack(side=tk.LEFT)
        win.protocol('WM_DELETE_WINDOW', self.closeTemplateEditor)
        self.selectedTemplateId = None
        self.refreshTemplateList()
        self.loadPlaceholderFormToGui()

    def closeTemplateEditor(self):
        """关闭话术模板管理窗口"""
        if self.templateEditorWin and self.templateEditorWin.winfo_exists():
            self.templateEditorWin.destroy()
        self.templateEditorWin = None
        self.templateListbox = None
        self.templateWordText = None
        self.selectedTemplateId = None
        self.templateRows = []
        self.templateListHintLabel = None
        self.tplPhCompanyVar = None
        self.tplPhJobVar = None
        self.tplPhNameVar = None
        self.tplPhAddressVar = None
        self.tplPhDurationVar = None
        self.tplPhDayOffsetVar = None
        self.tplPhTimeSlotsVar = None

    def currentTemplateTypeKey(self):
        """从下拉中文标签反查 template_type"""
        selected = self.templateTypeVar.get().strip() if self.templateTypeVar else ''
        labels = self.templateTypeLabels()
        for key, label in labels.items():
            if label == selected:
                return key
        for key in BossTemplate().types():
            if key == selected:
                return key
        return BossTemplate().types()[0] if BossTemplate().types() else 'greeting'

    def refreshTemplateList(self):
        """刷新当前类型的话术列表"""
        if not self.templateListbox:
            return
        db = self.ensureDb()
        templateType = self.currentTemplateTypeKey()
        self.templateRows = db.getTemplates(templateType)
        count = len(self.templateRows)
        # 更新列表标题，标明当前类型全部条数
        if self.templateListHintLabel:
            self.templateListHintLabel.configure(
                text=f'当前类型全部话术（共 {count} 条，发送时随机选用已启用条目）'
            )
        self.templateListbox.delete(0, tk.END)
        if count == 0:
            # 空列表占位提示
            self.templateListbox.insert(tk.END, '暂无，可点新增')
        else:
            for row in self.templateRows:
                enabledMark = '✓' if int(row.get('enabled') or 0) else '×'
                preview = str(row.get('word') or '').replace('\n', ' ')
                if len(preview) > 36:
                    preview = preview[:36] + '...'
                self.templateListbox.insert(tk.END, f'[{enabledMark}] #{row.get("id")} {preview}')
        self.selectedTemplateId = None
        if self.templateWordText:
            self.templateWordText.delete('1.0', tk.END)
        if self.templateEnabledVar:
            self.templateEnabledVar.set(1)

    def onTemplateListSelect(self, event=None):
        """选中列表项时加载到编辑框"""
        if not self.templateListbox:
            return
        # 空列表占位项不可选
        if not self.templateRows:
            return
        sel = self.templateListbox.curselection()
        if not sel:
            return
        idx = int(sel[0])
        if idx < 0 or idx >= len(self.templateRows):
            return
        row = self.templateRows[idx]
        self.selectedTemplateId = int(row.get('id') or 0)
        if self.templateWordText:
            self.templateWordText.delete('1.0', tk.END)
            self.templateWordText.insert(tk.END, str(row.get('word') or ''))
        if self.templateEnabledVar:
            self.templateEnabledVar.set(1 if int(row.get('enabled') or 0) else 0)

    def addTemplateRow(self):
        """新增一条空话术，等待用户编辑后保存"""
        self.selectedTemplateId = None
        if self.templateListbox:
            self.templateListbox.selection_clear(0, tk.END)
        if self.templateWordText:
            self.templateWordText.delete('1.0', tk.END)
            self.templateWordText.focus_set()
        if self.templateEnabledVar:
            self.templateEnabledVar.set(1)

    def saveTemplateRow(self):
        """保存新增或修改的话术"""
        if not self.templateWordText:
            return
        word = self.templateWordText.get('1.0', tk.END).strip()
        if not word:
            messagebox.showwarning('提示', '话术内容不能为空')
            return
        templateType = self.currentTemplateTypeKey()
        enabled = bool(self.templateEnabledVar.get()) if self.templateEnabledVar else True
        db = self.ensureDb()
        if self.selectedTemplateId:
            db.updateTemplate(self.selectedTemplateId, templateType, word, enabled)
            self.appendLog(f'已更新话术 #{self.selectedTemplateId}（{templateType}）')
        else:
            newId = db.createTemplate(templateType, word, enabled)
            self.selectedTemplateId = newId
            self.appendLog(f'已新增话术 #{newId}（{templateType}）')
        self.refreshTemplateList()
        messagebox.showinfo('已保存', '话术已保存，下次开始任务时生效')

    def deleteTemplateRow(self):
        """删除选中的话术"""
        if not self.selectedTemplateId:
            messagebox.showwarning('提示', '请先在列表中选择要删除的话术')
            return
        if not messagebox.askyesno('确认删除', f'确定删除话术 #{self.selectedTemplateId} 吗？'):
            return
        db = self.ensureDb()
        db.deleteTemplate(self.selectedTemplateId)
        self.appendLog(f'已删除话术 #{self.selectedTemplateId}')
        self.selectedTemplateId = None
        self.refreshTemplateList()

    def restoreTemplateType(self):
        """恢复当前类型为 template.py 默认值"""
        templateType = self.currentTemplateTypeKey()
        if not messagebox.askyesno('确认恢复', f'确定将「{templateType}」恢复为代码默认话术吗？\n当前类型的自定义内容将被覆盖。'):
            return
        tpl = BossTemplate()
        db = self.ensureDb()
        db.reloadTemplatesOfType(templateType, tpl.wordsOf(templateType))
        self.appendLog(f'已恢复默认话术类型：{templateType}')
        self.refreshTemplateList()
        messagebox.showinfo('已恢复', f'类型 {templateType} 已恢复默认')

    def restoreAllTemplates(self):
        """恢复全部话术为 template.py 默认值"""
        if not messagebox.askyesno('确认恢复', '确定将全部话术恢复为代码默认值吗？\n所有自定义话术将被覆盖。'):
            return
        tpl = BossTemplate()
        db = self.ensureDb()
        db.reloadTemplatesFromConfig(tpl.bundle())
        self.appendLog('已恢复全部默认话术')
        self.refreshTemplateList()
        messagebox.showinfo('已恢复', '全部话术已恢复默认')

    def openReplyEditor(self):
        """打开本地模型与回复 Skill 管理窗口"""
        if self.replyConfigWin and self.replyConfigWin.winfo_exists():
            self.replyConfigWin.lift()
            self.replyConfigWin.focus_force()
            self.loadReplyConfig()
            return
        win = tk.Toplevel(self.root)
        win.title('本地模型与回复 Skill')
        win.geometry('760x800')
        win.minsize(680, 680)
        self.replyConfigWin = win
        top = ttk.Frame(win, padding=10)
        top.pack(fill=tk.BOTH, expand=True)

        modelGroup = ttk.LabelFrame(top, text='本地模型服务')
        modelGroup.pack(fill=tk.X)
        modelForm = ttk.Frame(modelGroup, padding=8)
        modelForm.pack(fill=tk.X)
        modelForm.columnconfigure(1, weight=1)
        self.replyEnabledVar = tk.IntVar(value=1)
        self.replyBaseUrlVar = tk.StringVar(value=self.reply.baseUrl)
        self.replyModelVar = tk.StringVar(value=self.reply.modelName)
        self.replyTimeoutVar = tk.StringVar(value='90')
        ttk.Checkbutton(modelForm, text='启用本地回复建议', variable=self.replyEnabledVar).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(modelForm, text='API 地址:').grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=2)
        ttk.Entry(modelForm, textvariable=self.replyBaseUrlVar).grid(row=1, column=1, sticky=tk.EW, pady=2)
        ttk.Label(modelForm, text='模型名称:').grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=2)
        ttk.Entry(modelForm, textvariable=self.replyModelVar).grid(row=2, column=1, sticky=tk.EW, pady=2)
        ttk.Label(modelForm, text='超时秒数:').grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=2)
        ttk.Spinbox(modelForm, from_=5, to=600, width=8, textvariable=self.replyTimeoutVar).grid(row=3, column=1, sticky=tk.W, pady=2)
        modelBtns = ttk.Frame(modelGroup, padding=(8, 0, 8, 8))
        modelBtns.pack(fill=tk.X)
        ttk.Button(modelBtns, text='保存模型配置', command=self.saveReplyConfig).pack(side=tk.LEFT, padx=(0, 8))
        self.replyTestBtn = ttk.Button(modelBtns, text='测试本地服务', command=self.testReplyServer)
        self.replyTestBtn.pack(side=tk.LEFT, padx=(0, 8))
        self.replyConfigStatusVar = tk.StringVar(value='')
        ttk.Label(modelBtns, textvariable=self.replyConfigStatusVar, foreground='#0066cc').pack(side=tk.LEFT)

        skillGroup = ttk.LabelFrame(top, text='可切换回复 Skill')
        skillGroup.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        skillBody = ttk.Frame(skillGroup, padding=8)
        skillBody.pack(fill=tk.BOTH, expand=True)
        listFrame = ttk.Frame(skillBody)
        listFrame.pack(fill=tk.X)
        self.replySkillListbox = tk.Listbox(listFrame, height=4, exportselection=False)
        self.replySkillListbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        listScroll = ttk.Scrollbar(listFrame, orient=tk.VERTICAL, command=self.replySkillListbox.yview)
        listScroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.replySkillListbox.configure(yscrollcommand=listScroll.set)
        self.replySkillListbox.bind('<<ListboxSelect>>', self.onReplySkillSelect)
        skillBtns = ttk.Frame(skillBody)
        skillBtns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(skillBtns, text='新建', command=self.newReplySkill).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(skillBtns, text='保存 Skill', command=self.saveReplySkill).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(skillBtns, text='设为当前', command=self.setActiveReplySkill).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(skillBtns, text='删除', command=self.deleteReplySkill).pack(side=tk.LEFT)

        nameRow = ttk.Frame(skillBody)
        nameRow.pack(fill=tk.X, pady=(8, 0))
        self.replySkillNameVar = tk.StringVar(value='')
        self.replySkillEnabledVar = tk.IntVar(value=1)
        ttk.Label(nameRow, text='Skill 名称:').pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(nameRow, textvariable=self.replySkillNameVar).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Checkbutton(nameRow, text='启用', variable=self.replySkillEnabledVar).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(skillBody, text='回复规则:').pack(anchor=tk.W, pady=(8, 0))
        self.replyInstructionText = scrolledtext.ScrolledText(skillBody, height=10, font=('Microsoft YaHei UI', 9))
        self.replyInstructionText.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        ttk.Label(skillBody, text='参考案例:').pack(anchor=tk.W, pady=(8, 0))
        self.replyExamplesText = scrolledtext.ScrolledText(skillBody, height=8, font=('Microsoft YaHei UI', 9))
        self.replyExamplesText.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        ttk.Label(
            skillBody,
            text='Skill 修改保存后立即生效；模型建议始终需要人工检查，程序不会自动发送。',
            foreground='#666',
        ).pack(anchor=tk.W, pady=(6, 0))
        win.protocol('WM_DELETE_WINDOW', self.closeReplyEditor)
        self.loadReplyConfig()

    def closeReplyEditor(self):
        """关闭回复配置窗口并清理控件引用"""
        if self.replyConfigWin and self.replyConfigWin.winfo_exists():
            self.replyConfigWin.destroy()
        self.replyConfigWin = None
        self.replySkillListbox = None
        self.replyInstructionText = None
        self.replyExamplesText = None
        self.selectedReplySkillId = None

    def collectReplyConfig(self):
        """从 GUI 读取本地模型配置"""
        try:
            timeoutSec = int(self.replyTimeoutVar.get().strip())
        except (TypeError, ValueError):
            raise ValueError('模型超时秒数必须是整数')
        baseUrl = self.replyBaseUrlVar.get().strip().rstrip('/')
        if not self.reply.isLocalUrl(baseUrl):
            raise ValueError('只允许填写 127.0.0.1、localhost 或 ::1 的本机地址')
        return {
            'enabled': bool(self.replyEnabledVar.get()),
            'baseUrl': baseUrl,
            'modelName': self.replyModelVar.get().strip(),
            'timeoutSec': timeoutSec,
            'activeSkillId': self.ensureDb().getReplySettings().get('activeSkillId'),
        }

    def saveReplyConfig(self):
        """保存本地模型连接配置"""
        try:
            config = self.collectReplyConfig()
            self.ensureDb().saveReplySettings(config)
        except Exception as exc:
            messagebox.showwarning('保存失败', str(exc), parent=self.replyConfigWin)
            return False
        if self.replyConfigStatusVar:
            self.replyConfigStatusVar.set('已保存')
        self.appendLog('已保存本地回复模型配置')
        return True

    def loadReplyConfig(self):
        """从数据库加载模型配置与活动 Skill"""
        settings = self.ensureDb().getReplySettings()
        if self.replyEnabledVar:
            self.replyEnabledVar.set(1 if settings.get('enabled') else 0)
        if self.replyBaseUrlVar:
            self.replyBaseUrlVar.set(str(settings.get('baseUrl') or self.reply.baseUrl))
        if self.replyModelVar:
            self.replyModelVar.set(str(settings.get('modelName') or self.reply.modelName))
        if self.replyTimeoutVar:
            self.replyTimeoutVar.set(str(settings.get('timeoutSec') or 90))
        self.refreshReplySkills(settings.get('activeSkillId'))

    def testReplyServer(self):
        """后台测试本地 llama-server，避免阻塞 GUI"""
        try:
            config = self.collectReplyConfig()
        except Exception as exc:
            messagebox.showwarning('配置错误', str(exc), parent=self.replyConfigWin)
            return
        if self.replyTestBtn:
            self.replyTestBtn.configure(state=tk.DISABLED)
        if self.replyConfigStatusVar:
            self.replyConfigStatusVar.set('正在连接...')

        def worker():
            """调用模型列表接口并回到 GUI 显示结果"""
            try:
                names = self.reply.testConnection(config)
                text = '连接成功' + (f"：{', '.join(names[:2])}" if names else '')
                error = ''
            except Exception as exc:
                text = ''
                error = str(exc)
            if self.root:
                self.root.after(0, lambda: self.showReplyTest(text, error))

        threading.Thread(target=worker, daemon=True).start()

    def showReplyTest(self, text, error):
        """显示本地模型连接测试结果"""
        if self.replyTestBtn and self.replyTestBtn.winfo_exists():
            self.replyTestBtn.configure(state=tk.NORMAL)
        if self.replyConfigStatusVar:
            self.replyConfigStatusVar.set(text or '连接失败')
        if error:
            messagebox.showwarning('连接失败', error, parent=self.replyConfigWin)

    def refreshReplySkills(self, selectId=None):
        """刷新 Skill 列表并标记当前活动项"""
        if not self.replySkillListbox:
            return
        db = self.ensureDb()
        self.replySkillRows = db.getReplySkills()
        activeId = int(db.getReplySettings().get('activeSkillId') or 0)
        self.replySkillListbox.delete(0, tk.END)
        selectedIndex = None
        for index, row in enumerate(self.replySkillRows):
            prefix = '【当前】' if int(row['id']) == activeId else ''
            suffix = '' if int(row.get('enabled') or 0) else '（停用）'
            self.replySkillListbox.insert(tk.END, prefix + str(row.get('skill_name') or '') + suffix)
            if int(row['id']) == int(selectId or activeId):
                selectedIndex = index
        if selectedIndex is not None:
            self.replySkillListbox.selection_set(selectedIndex)
            self.replySkillListbox.see(selectedIndex)
            self.onReplySkillSelect()

    def onReplySkillSelect(self, event=None):
        """将列表中选中的 Skill 加载到编辑表单"""
        if not self.replySkillListbox:
            return
        selected = self.replySkillListbox.curselection()
        if not selected:
            return
        row = self.replySkillRows[selected[0]]
        self.selectedReplySkillId = int(row['id'])
        self.replySkillNameVar.set(str(row.get('skill_name') or ''))
        self.replySkillEnabledVar.set(1 if row.get('enabled') else 0)
        self.replyInstructionText.delete('1.0', tk.END)
        self.replyInstructionText.insert('1.0', str(row.get('instruction') or ''))
        self.replyExamplesText.delete('1.0', tk.END)
        self.replyExamplesText.insert('1.0', str(row.get('examples') or ''))

    def newReplySkill(self):
        """清空表单以创建新的回复 Skill"""
        self.selectedReplySkillId = None
        self.replySkillNameVar.set('')
        self.replySkillEnabledVar.set(1)
        self.replyInstructionText.delete('1.0', tk.END)
        self.replyInstructionText.insert('1.0', self.reply.defaultInstruction)
        self.replyExamplesText.delete('1.0', tk.END)

    def collectReplySkill(self):
        """从编辑表单读取回复 Skill 内容"""
        return {
            'skillName': self.replySkillNameVar.get().strip(),
            'instruction': self.replyInstructionText.get('1.0', tk.END).strip(),
            'examples': self.replyExamplesText.get('1.0', tk.END).strip(),
            'enabled': bool(self.replySkillEnabledVar.get()),
        }

    def saveReplySkill(self):
        """新增或更新回复 Skill"""
        try:
            data = self.collectReplySkill()
            db = self.ensureDb()
            if self.selectedReplySkillId:
                db.updateReplySkill(self.selectedReplySkillId, data)
                skillId = self.selectedReplySkillId
            else:
                skillId = db.createReplySkill(data)
        except Exception as exc:
            messagebox.showwarning('保存失败', str(exc), parent=self.replyConfigWin)
            return
        self.refreshReplySkills(skillId)
        self.appendLog(f"已保存回复 Skill：{data.get('skillName')}")

    def setActiveReplySkill(self):
        """将当前选中的 Skill 设为模型生成时使用的规则"""
        if not self.selectedReplySkillId:
            messagebox.showwarning('提示', '请先选择一条 Skill', parent=self.replyConfigWin)
            return
        row = self.ensureDb().getReplySkill(self.selectedReplySkillId)
        if not row or not row.get('enabled'):
            messagebox.showwarning('提示', '停用的 Skill 不能设为当前', parent=self.replyConfigWin)
            return
        try:
            config = self.collectReplyConfig()
            config['activeSkillId'] = self.selectedReplySkillId
            self.ensureDb().saveReplySettings(config)
        except Exception as exc:
            messagebox.showwarning('切换失败', str(exc), parent=self.replyConfigWin)
            return
        self.refreshReplySkills(self.selectedReplySkillId)
        self.appendLog(f"已切换回复 Skill：{row.get('skill_name')}")

    def deleteReplySkill(self):
        """删除当前选中的回复 Skill"""
        if not self.selectedReplySkillId:
            messagebox.showwarning('提示', '请先选择一条 Skill', parent=self.replyConfigWin)
            return
        if not messagebox.askyesno('确认删除', '确定删除当前回复 Skill 吗？', parent=self.replyConfigWin):
            return
        try:
            self.ensureDb().deleteReplySkill(self.selectedReplySkillId)
        except Exception as exc:
            messagebox.showwarning('删除失败', str(exc), parent=self.replyConfigWin)
            return
        self.selectedReplySkillId = None
        self.refreshReplySkills()

    def openReplyAssist(self):
        """打开当前候选人的本地模型回复建议窗口"""
        if not self.bossAuto or not self.bossAuto.manualReplyWaitActive:
            messagebox.showwarning('提示', '当前没有等待人工回复的候选人')
            return
        info = dict(self.manualReplyInfo or {})
        if not str(info.get('friendText') or '').strip():
            messagebox.showwarning('提示', '当前人工步骤没有候选人文字消息可供生成')
            return
        settings = self.ensureDb().getReplySettings()
        if not settings.get('enabled'):
            messagebox.showwarning('提示', '请先在“本地回复模型”中启用回复建议')
            return
        if self.replyAssistWin and self.replyAssistWin.winfo_exists():
            self.replyAssistWin.destroy()
        win = tk.Toplevel(self.root)
        win.title('本地模型回复助手')
        win.geometry('700x620')
        win.minsize(600, 520)
        self.replyAssistWin = win
        body = ttk.Frame(win, padding=10)
        body.pack(fill=tk.BOTH, expand=True)
        name = str(info.get('candidateName') or '')
        job = str(info.get('jobName') or '')
        ttk.Label(body, text=f'候选人：{name}    岗位：{job}').pack(anchor=tk.W)
        ttk.Label(body, text='最近对话:').pack(anchor=tk.W, pady=(8, 0))
        friendBox = scrolledtext.ScrolledText(body, height=7, font=('Microsoft YaHei UI', 9))
        friendBox.pack(fill=tk.X, pady=(2, 0))
        contextText = str(info.get('conversationText') or info.get('friendText') or '')
        friendBox.insert('1.0', contextText)
        friendBox.configure(state=tk.DISABLED)
        self.replyAssistRecommendVar = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.replyAssistRecommendVar, foreground='#7a4d00').pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(body, text='建议回复（可直接修改）:').pack(anchor=tk.W, pady=(6, 0))
        self.replyAssistText = scrolledtext.ScrolledText(body, height=10, font=('Microsoft YaHei UI', 10))
        self.replyAssistText.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        btnRow = ttk.Frame(body)
        btnRow.pack(fill=tk.X, pady=(8, 0))
        self.replyGenerateBtn = ttk.Button(btnRow, text='重新生成', command=self.generateReply)
        self.replyGenerateBtn.pack(side=tk.LEFT, padx=(0, 8))
        self.replyFillBtn = ttk.Button(btnRow, text='填入 BOSS 聊天框', command=self.fillReplyAssist, state=tk.DISABLED)
        self.replyFillBtn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='关闭', command=self.closeReplyAssist).pack(side=tk.LEFT)
        self.replyAssistStatusVar = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.replyAssistStatusVar, foreground='#0066cc').pack(anchor=tk.W, pady=(6, 0))
        ttk.Label(body, text='填入后仍需在 BOSS 页面人工检查并点击发送。', foreground='#666').pack(anchor=tk.W)
        win.protocol('WM_DELETE_WINDOW', self.closeReplyAssist)
        self.generateReply()

    def closeReplyAssist(self):
        """关闭回复建议窗口并清理当前建议"""
        if self.replyAssistWin and self.replyAssistWin.winfo_exists():
            self.replyAssistWin.destroy()
        self.replyAssistWin = None
        self.replyAssistText = None
        self.replyResult = None

    def generateReply(self):
        """后台调用本地模型生成当前候选人的回复建议"""
        if self.replyGenerateThread and self.replyGenerateThread.is_alive():
            return
        settings = self.ensureDb().getReplySettings()
        skill = settings.get('skill')
        if not skill:
            messagebox.showwarning('提示', '没有可用的回复 Skill', parent=self.replyAssistWin)
            return
        info = dict(self.manualReplyInfo or {})
        jobName = str(info.get('jobName') or '')
        jobRules = self.ensureDb().matchJobRules(jobName) or {}
        jobInfo = {'jobName': jobName, 'jobIntro': str(jobRules.get('jobIntro') or '')}
        if self.replyGenerateBtn:
            self.replyGenerateBtn.configure(state=tk.DISABLED)
        if self.replyFillBtn:
            self.replyFillBtn.configure(state=tk.DISABLED)
        if self.replyAssistStatusVar:
            self.replyAssistStatusVar.set('正在由本地模型生成...')

        def worker():
            """调用本地模型并回传结构化建议"""
            try:
                result = self.reply.generate(info, settings, skill, jobInfo)
                error = ''
            except Exception as exc:
                result = None
                error = str(exc)
            if self.root:
                self.root.after(0, lambda: self.showReplyResult(result, error))

        self.replyGenerateThread = threading.Thread(target=worker, daemon=True)
        self.replyGenerateThread.start()

    def showReplyResult(self, result, error):
        """在回复助手窗口展示模型结果或错误"""
        if not self.replyAssistWin or not self.replyAssistWin.winfo_exists():
            return
        if self.replyGenerateBtn:
            self.replyGenerateBtn.configure(state=tk.NORMAL)
        if error:
            self.replyAssistStatusVar.set('生成失败')
            messagebox.showwarning('生成失败', error, parent=self.replyAssistWin)
            return
        self.replyResult = dict(result or {})
        self.replyAssistText.delete('1.0', tk.END)
        self.replyAssistText.insert('1.0', str(self.replyResult.get('reply') or ''))
        labels = {
            'reply_only': '建议：仅回复，后续由人工判断',
            'consider_resume': '建议：可考虑继续沟通，是否索要简历由人工决定',
            'unsuitable': '建议：对方表达无意向或不合适，请人工确认',
        }
        recommendation = str(self.replyResult.get('recommendation') or 'reply_only')
        risk = str(self.replyResult.get('risk') or '')
        text = labels.get(recommendation, labels['reply_only'])
        if risk:
            text += f'；需核实：{risk}'
        self.replyAssistRecommendVar.set(text)
        self.replyAssistStatusVar.set('生成完成，请人工审核和修改')
        self.replyFillBtn.configure(state=tk.NORMAL)

    def fillReplyAssist(self):
        """将人工确认后的建议填入 BOSS 输入框并记录反馈"""
        if self.replyFillThread and self.replyFillThread.is_alive():
            return
        if not self.bossAuto or not self.bossAuto.manualReplyWaitActive:
            messagebox.showwarning('提示', '人工回复等待已经结束', parent=self.replyAssistWin)
            return
        finalReply = self.replyAssistText.get('1.0', tk.END).strip()
        if not finalReply:
            messagebox.showwarning('提示', '回复内容不能为空', parent=self.replyAssistWin)
            return
        info = dict(self.manualReplyInfo or {})
        result = dict(self.replyResult or {})
        settings = self.ensureDb().getReplySettings()
        self.replyFillBtn.configure(state=tk.DISABLED)
        self.replyAssistStatusVar.set('正在填入聊天框...')

        def worker():
            """填入 BOSS 聊天框并保存人工采用结果"""
            try:
                filled = self.bossAuto.fillManualReply(finalReply)
                if filled:
                    self.ensureDb().saveReplyFeedback({
                        'candidateKey': info.get('candidateKey'),
                        'skillId': settings.get('activeSkillId'),
                        'jobName': info.get('jobName'),
                        'friendText': info.get('friendText'),
                        'suggestedReply': result.get('reply'),
                        'finalReply': finalReply,
                        'recommendation': result.get('recommendation'),
                        'accepted': True,
                    })
                error = '' if filled else '未找到当前 BOSS 聊天输入框'
            except Exception as exc:
                filled = False
                error = str(exc)
            if self.root:
                self.root.after(0, lambda: self.showReplyFill(filled, error))

        self.replyFillThread = threading.Thread(target=worker, daemon=True)
        self.replyFillThread.start()

    def showReplyFill(self, filled, error):
        """显示回复填入结果并继续保留人工发送边界"""
        if not self.replyAssistWin or not self.replyAssistWin.winfo_exists():
            return
        self.replyFillBtn.configure(state=tk.NORMAL)
        if not filled:
            self.replyAssistStatusVar.set('填入失败')
            messagebox.showwarning('填入失败', error, parent=self.replyAssistWin)
            return
        self.replyAssistStatusVar.set('已填入，请到 BOSS 页面检查并人工发送')
        self.appendLog('本地模型建议已由人工确认并填入 BOSS 聊天框')

    def parseCommaList(self, text):
        """逗号分隔文本转关键词列表"""
        parts = str(text or '').replace('，', ',').split(',')
        return [part.strip() for part in parts if part.strip()]

    def formatCommaList(self, items):
        """关键词列表转逗号分隔文本"""
        return ', '.join(str(item or '').strip() for item in (items or []) if str(item or '').strip())

    def parseAnyKeywordGroups(self, text):
        """分组关键词文本转 anyKeywords（每组一行，组内逗号分隔）"""
        groups = []
        for line in str(text or '').splitlines():
            words = self.parseCommaList(line)
            if words:
                groups.append(words)
        return groups

    def formatAnyKeywordGroups(self, groups):
        """anyKeywords 转分组文本"""
        lines = []
        for group in groups or []:
            line = self.formatCommaList(group)
            if line:
                lines.append(line)
        return '\n'.join(lines)

    def openJobEditor(self):
        """打开岗位规则管理窗口"""
        if self.jobEditorWin and self.jobEditorWin.winfo_exists():
            self.jobEditorWin.lift()
            self.jobEditorWin.focus_force()
            self.refreshJobList()
            return
        win = tk.Toplevel(self.root)
        win.title('岗位规则管理')
        win.geometry('680x640')
        win.minsize(560, 520)
        self.jobEditorWin = win
        top = ttk.Frame(win, padding=10)
        top.pack(fill=tk.BOTH, expand=True)
        self.jobListHintLabel = ttk.Label(top, text='全部岗位规则（共 0 条，按岗位名与别名匹配沟通列表）')
        self.jobListHintLabel.pack(anchor=tk.W)
        listFrame = ttk.Frame(top)
        listFrame.pack(fill=tk.X, pady=(4, 0))
        self.jobListbox = tk.Listbox(listFrame, height=5, exportselection=False)
        self.jobListbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        listScroll = ttk.Scrollbar(listFrame, orient=tk.VERTICAL, command=self.jobListbox.yview)
        listScroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.jobListbox.configure(yscrollcommand=listScroll.set)
        self.jobListbox.bind('<<ListboxSelect>>', self.onJobListSelect)
        form = ttk.Frame(top)
        form.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        ttk.Label(form, text='岗位名称:').grid(row=0, column=0, sticky=tk.W, pady=2)
        self.jobNameVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobNameVar).grid(row=0, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='匹配别名（逗号分隔）:').grid(row=1, column=0, sticky=tk.W, pady=2)
        self.jobMatchKeysVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobMatchKeysVar).grid(row=1, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='岗位介绍 intro（可用 {job}）:').grid(row=2, column=0, sticky=tk.NW, pady=2)
        self.jobIntroText = scrolledtext.ScrolledText(form, height=4, font=('Microsoft YaHei UI', 9))
        self.jobIntroText.grid(row=2, column=1, sticky=tk.EW, pady=2)
        ageRow = ttk.Frame(form)
        ageRow.grid(row=3, column=1, sticky=tk.W, pady=2)
        ttk.Label(form, text='年龄 / 年限:').grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Label(ageRow, text='最小年龄').pack(side=tk.LEFT)
        self.jobAgeMinVar = tk.StringVar(value='18')
        ttk.Entry(ageRow, textvariable=self.jobAgeMinVar, width=6).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(ageRow, text='最大年龄').pack(side=tk.LEFT)
        self.jobAgeMaxVar = tk.StringVar(value='45')
        ttk.Entry(ageRow, textvariable=self.jobAgeMaxVar, width=6).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(ageRow, text='最低工作年限').pack(side=tk.LEFT)
        self.jobWorkYearsVar = tk.StringVar(value='0')
        ttk.Entry(ageRow, textvariable=self.jobWorkYearsVar, width=6).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(form, text='学历（逗号，留空不限）:').grid(row=4, column=0, sticky=tk.W, pady=2)
        self.jobEducationVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobEducationVar).grid(row=4, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='必须关键词（逗号，全部命中）:').grid(row=5, column=0, sticky=tk.W, pady=2)
        self.jobMustVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobMustVar).grid(row=5, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='分组关键词（每组一行，组内逗号）:').grid(row=6, column=0, sticky=tk.NW, pady=2)
        self.jobAnyText = scrolledtext.ScrolledText(form, height=3, font=('Microsoft YaHei UI', 9))
        self.jobAnyText.grid(row=6, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='优先关键词（逗号）:').grid(row=7, column=0, sticky=tk.W, pady=2)
        self.jobPreferVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobPreferVar).grid(row=7, column=1, sticky=tk.EW, pady=2)
        ttk.Label(form, text='排除关键词（逗号，命中即拒）:').grid(row=8, column=0, sticky=tk.W, pady=2)
        self.jobRejectVar = tk.StringVar(value='')
        ttk.Entry(form, textvariable=self.jobRejectVar).grid(row=8, column=1, sticky=tk.EW, pady=2)
        form.columnconfigure(1, weight=1)
        ttk.Label(
            form,
            text='分组说明：每组至少命中一个词；多组之间为「且」关系。岗位规则优先于全局默认筛选。',
            foreground='#666',
            wraplength=520,
        ).grid(row=9, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))
        self.jobEnabledVar = tk.IntVar(value=1)
        ttk.Checkbutton(form, text='启用此岗位规则', variable=self.jobEnabledVar).grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))
        btnRow = ttk.Frame(top)
        btnRow.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btnRow, text='新增', command=self.addJobRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='保存', command=self.saveJobRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='删除', command=self.deleteJobRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='恢复当前岗位默认', command=self.restoreJobRow).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnRow, text='恢复全部默认', command=self.restoreAllJobs).pack(side=tk.LEFT)
        win.protocol('WM_DELETE_WINDOW', self.closeJobEditor)
        self.selectedJobId = None
        self.refreshJobList()
        self.clearJobForm()

    def closeJobEditor(self):
        """关闭岗位规则管理窗口"""
        if self.jobEditorWin and self.jobEditorWin.winfo_exists():
            self.jobEditorWin.destroy()
        self.jobEditorWin = None
        self.jobListbox = None
        self.jobListHintLabel = None
        self.selectedJobId = None
        self.jobRows = []
        self.jobNameVar = None
        self.jobMatchKeysVar = None
        self.jobIntroText = None
        self.jobAgeMinVar = None
        self.jobAgeMaxVar = None
        self.jobWorkYearsVar = None
        self.jobEducationVar = None
        self.jobMustVar = None
        self.jobPreferVar = None
        self.jobRejectVar = None
        self.jobAnyText = None
        self.jobEnabledVar = None

    def clearJobForm(self):
        """清空岗位编辑表单"""
        self.selectedJobId = None
        if self.jobListbox:
            self.jobListbox.selection_clear(0, tk.END)
        if self.jobNameVar:
            self.jobNameVar.set('')
        if self.jobMatchKeysVar:
            self.jobMatchKeysVar.set('')
        if self.jobIntroText:
            self.jobIntroText.delete('1.0', tk.END)
        if self.jobAgeMinVar:
            self.jobAgeMinVar.set('18')
        if self.jobAgeMaxVar:
            self.jobAgeMaxVar.set('45')
        if self.jobWorkYearsVar:
            self.jobWorkYearsVar.set('0')
        if self.jobEducationVar:
            self.jobEducationVar.set('')
        if self.jobMustVar:
            self.jobMustVar.set('')
        if self.jobPreferVar:
            self.jobPreferVar.set('')
        if self.jobRejectVar:
            self.jobRejectVar.set('')
        if self.jobAnyText:
            self.jobAnyText.delete('1.0', tk.END)
        if self.jobEnabledVar:
            self.jobEnabledVar.set(1)

    def loadJobForm(self, row):
        """将岗位规则加载到编辑表单"""
        if self.jobNameVar:
            self.jobNameVar.set(str(row.get('jobName') or ''))
        if self.jobMatchKeysVar:
            self.jobMatchKeysVar.set(self.formatCommaList(row.get('matchKeys')))
        if self.jobIntroText:
            self.jobIntroText.delete('1.0', tk.END)
            self.jobIntroText.insert(tk.END, str(row.get('jobIntro') or ''))
        if self.jobAgeMinVar:
            self.jobAgeMinVar.set(str(row.get('ageMin') or 18))
        if self.jobAgeMaxVar:
            self.jobAgeMaxVar.set(str(row.get('ageMax') or 45))
        if self.jobWorkYearsVar:
            self.jobWorkYearsVar.set(str(row.get('workYearsMin') or 0))
        if self.jobEducationVar:
            self.jobEducationVar.set(self.formatCommaList(row.get('educationList')))
        if self.jobMustVar:
            self.jobMustVar.set(self.formatCommaList(row.get('mustKeywords')))
        if self.jobPreferVar:
            self.jobPreferVar.set(self.formatCommaList(row.get('preferKeywords')))
        if self.jobRejectVar:
            self.jobRejectVar.set(self.formatCommaList(row.get('rejectKeywords')))
        if self.jobAnyText:
            self.jobAnyText.delete('1.0', tk.END)
            self.jobAnyText.insert(tk.END, self.formatAnyKeywordGroups(row.get('anyKeywords')))
        if self.jobEnabledVar:
            self.jobEnabledVar.set(1 if int(row.get('enabled') or 0) else 0)

    def collectJobForm(self):
        """从编辑表单收集岗位规则数据"""
        try:
            ageMin = int(str(self.jobAgeMinVar.get() if self.jobAgeMinVar else '18').strip() or 18)
            ageMax = int(str(self.jobAgeMaxVar.get() if self.jobAgeMaxVar else '45').strip() or 45)
            workYearsMin = int(str(self.jobWorkYearsVar.get() if self.jobWorkYearsVar else '0').strip() or 0)
        except ValueError:
            raise ValueError('年龄与年限须为整数')
        intro = self.jobIntroText.get('1.0', tk.END).strip() if self.jobIntroText else ''
        anyText = self.jobAnyText.get('1.0', tk.END) if self.jobAnyText else ''
        return {
            'jobName': self.jobNameVar.get().strip() if self.jobNameVar else '',
            'matchKeys': self.parseCommaList(self.jobMatchKeysVar.get() if self.jobMatchKeysVar else ''),
            'jobIntro': intro,
            'ageMin': ageMin,
            'ageMax': ageMax,
            'workYearsMin': workYearsMin,
            'educationList': self.parseCommaList(self.jobEducationVar.get() if self.jobEducationVar else ''),
            'mustKeywords': self.parseCommaList(self.jobMustVar.get() if self.jobMustVar else ''),
            'anyKeywords': self.parseAnyKeywordGroups(anyText),
            'preferKeywords': self.parseCommaList(self.jobPreferVar.get() if self.jobPreferVar else ''),
            'rejectKeywords': self.parseCommaList(self.jobRejectVar.get() if self.jobRejectVar else ''),
            'enabled': bool(self.jobEnabledVar.get()) if self.jobEnabledVar else True,
        }

    def refreshJobList(self):
        """刷新岗位规则列表"""
        if not self.jobListbox:
            return
        db = self.ensureDb()
        self.jobRows = db.getJobRulesList()
        count = len(self.jobRows)
        if self.jobListHintLabel:
            self.jobListHintLabel.configure(text=f'全部岗位规则（共 {count} 条，按岗位名与别名匹配沟通列表）')
        self.jobListbox.delete(0, tk.END)
        if count == 0:
            self.jobListbox.insert(tk.END, '暂无，可点新增')
        else:
            for row in self.jobRows:
                enabledMark = '✓' if int(row.get('enabled') or 0) else '×'
                name = str(row.get('jobName') or '')
                self.jobListbox.insert(tk.END, f'[{enabledMark}] #{row.get("id")} {name}')

    def onJobListSelect(self, event=None):
        """选中岗位时加载到表单"""
        if not self.jobListbox or not self.jobRows:
            return
        sel = self.jobListbox.curselection()
        if not sel:
            return
        idx = int(sel[0])
        if idx < 0 or idx >= len(self.jobRows):
            return
        row = self.jobRows[idx]
        self.selectedJobId = int(row.get('id') or 0)
        self.loadJobForm(row)

    def addJobRow(self):
        """新增岗位，清空表单等待编辑"""
        self.clearJobForm()
        if self.jobNameVar:
            self.jobNameVar.set('')
            self.root.focus_set()
        if self.jobIntroText:
            self.jobIntroText.focus_set()

    def saveJobRow(self):
        """保存岗位规则"""
        try:
            data = self.collectJobForm()
        except ValueError as exc:
            messagebox.showwarning('提示', str(exc))
            return
        if not data['jobName']:
            messagebox.showwarning('提示', '岗位名称不能为空')
            return
        db = self.ensureDb()
        try:
            if self.selectedJobId:
                db.updateJobRule(self.selectedJobId, data)
                self.appendLog(f'已更新岗位规则 #{self.selectedJobId}：{data["jobName"]}')
            else:
                newId = db.createJobRule(data)
                self.selectedJobId = newId
                self.appendLog(f'已新增岗位规则 #{newId}：{data["jobName"]}')
        except Exception as exc:
            messagebox.showerror('保存失败', str(exc))
            return
        self.refreshJobList()
        # 保存后重新选中刚保存的项
        if self.selectedJobId and self.jobListbox:
            for idx, row in enumerate(self.jobRows):
                if int(row.get('id') or 0) == self.selectedJobId:
                    self.jobListbox.selection_set(idx)
                    self.loadJobForm(row)
                    break
        messagebox.showinfo('已保存', '岗位规则已保存，下次开始任务时生效')

    def deleteJobRow(self):
        """删除选中的岗位规则"""
        if not self.selectedJobId:
            messagebox.showwarning('提示', '请先在列表中选择要删除的岗位')
            return
        if not messagebox.askyesno('确认删除', f'确定删除岗位规则 #{self.selectedJobId} 吗？'):
            return
        db = self.ensureDb()
        db.deleteJobRule(self.selectedJobId)
        self.appendLog(f'已删除岗位规则 #{self.selectedJobId}')
        self.selectedJobId = None
        self.refreshJobList()
        self.clearJobForm()

    def findDefaultJobProfile(self, jobName):
        """按岗位名查找 job.py 内置 profile"""
        name = str(jobName or '').strip()
        for profile in BossJob().profiles:
            if str(profile.get('jobName') or '').strip() == name:
                return profile
        return None

    def restoreJobRow(self):
        """恢复当前岗位为 job.py 默认值"""
        jobName = self.jobNameVar.get().strip() if self.jobNameVar else ''
        if not jobName:
            messagebox.showwarning('提示', '请先填写或选择岗位名称')
            return
        profile = self.findDefaultJobProfile(jobName)
        if not profile:
            messagebox.showwarning('提示', f'代码默认配置中未找到岗位「{jobName}」\n仅可恢复 job.py 内置的 4 个岗位')
            return
        if not messagebox.askyesno('确认恢复', f'确定将「{jobName}」恢复为代码默认规则吗？\n当前岗位的自定义内容将被覆盖。'):
            return
        db = self.ensureDb()
        db.reloadJobProfileFromConfig(profile)
        self.appendLog(f'已恢复默认岗位规则：{jobName}')
        self.refreshJobList()
        for row in self.jobRows:
            if str(row.get('jobName') or '').strip() == jobName:
                self.selectedJobId = int(row.get('id') or 0)
                self.loadJobForm(row)
                if self.jobListbox:
                    idx = self.jobRows.index(row)
                    self.jobListbox.selection_clear(0, tk.END)
                    self.jobListbox.selection_set(idx)
                break
        messagebox.showinfo('已恢复', f'岗位 {jobName} 已恢复默认')

    def restoreAllJobs(self):
        """恢复全部岗位为 job.py 默认值"""
        if not messagebox.askyesno('确认恢复', '确定将全部岗位规则恢复为代码默认值吗？\n所有自定义岗位规则将被覆盖。'):
            return
        job = BossJob()
        db = self.ensureDb()
        db.reloadJobProfilesFromConfig(job.profiles)
        self.appendLog('已恢复全部默认岗位规则')
        self.refreshJobList()
        messagebox.showinfo('已恢复', '全部岗位规则已恢复默认')

    def appendLog(self, text):
        """向日志框追加一行（仅主线程调用）"""
        # 日志控件尚未创建则跳过
        if not self.logText:
            return
        # 临时解锁文本框以便写入
        self.logText.configure(state=tk.NORMAL)
        # 追加一行并自动滚到底部
        self.logText.insert(tk.END, text + '\n')
        self.logText.see(tk.END)
        # 恢复只读，防止用户误改
        self.logText.configure(state=tk.DISABLED)

    def pollLogQueue(self):
        """从队列读取子线程日志并刷新界面"""
        # 非阻塞取出队列中全部待显示日志
        while True:
            try:
                msg = self.logQueue.get_nowait()
            except queue.Empty:
                break
            # 每条日志写入界面
            self.appendLog(msg)
        # 200ms 后再次轮询，形成定时刷新
        if self.root:
            self.root.after(200, self.pollLogQueue)

    def logFromWorker(self, text):
        """子线程安全写日志"""
        # 放入队列，由主线程 pollLogQueue 统一刷新
        self.logQueue.put(text)

    def setUiBusy(self, busy, status='就绪'):
        """切换按钮可用状态"""
        # 忙碌时禁用登录/任务按钮，启用停止按钮
        state = tk.DISABLED if busy else tk.NORMAL
        self.loginBtn.configure(state=state)
        self.taskBtn.configure(state=state)
        self.stopBtn.configure(state=tk.NORMAL if busy else tk.DISABLED)
        # 正式面试确认按钮仅在等待人工发 BOSS 邀约时可点
        if self.formalDoneBtn and not (self.bossAuto and self.bossAuto.formalInterviewWaitActive):
            self.formalDoneBtn.configure(state=tk.DISABLED)
        # 人工回复确认按钮仅在等待人工聊天回复时可点
        if self.manualReplyDoneBtn and not (self.bossAuto and self.bossAuto.manualReplyWaitActive):
            self.manualReplyDoneBtn.configure(state=tk.DISABLED)
        if self.manualResumeBtn and not (self.bossAuto and self.bossAuto.manualReplyWaitActive):
            self.manualResumeBtn.configure(state=tk.DISABLED)
        if self.manualUnsuitableBtn and not (self.bossAuto and self.bossAuto.manualReplyWaitActive):
            self.manualUnsuitableBtn.configure(state=tk.DISABLED)
        if self.replyAssistBtn and not (self.bossAuto and self.bossAuto.manualReplyWaitActive):
            self.replyAssistBtn.configure(state=tk.DISABLED)
        # 切换人工回复按钮：任务运行中可点
        if self.manualHandoffBtn:
            self.manualHandoffBtn.configure(state=tk.NORMAL if busy else tk.DISABLED)
        # 更新底部状态栏文字
        self.statusVar.set(status)

    def onFormalInterviewWait(self, info):
        """子线程回调：进入人工正式面试邀约等待"""
        if self.root:
            self.root.after(0, lambda: self.showFormalInterviewWait(info))

    def onFormalInterviewWaitDone(self):
        """子线程回调：结束人工正式面试邀约等待"""
        if self.root:
            self.root.after(0, self.hideFormalInterviewWait)

    def showFormalInterviewWait(self, info):
        """主线程：弹窗通知并启用确认按钮"""
        name = str(info.get('candidateName') or '')
        job = str(info.get('jobName') or '')
        dateText = str(info.get('agreedDate') or '')
        timeText = str(info.get('agreedTime') or '')
        address = str(info.get('address') or '')
        self.formalDoneBtn.configure(state=tk.NORMAL)
        self.statusVar.set(f'等待人工发正式面试邀约：{name}')
        self.appendLog('—— 等待人工发送 BOSS 正式面试邀约 ——')
        messagebox.showinfo(
            '需人工操作',
            f'候选人：{name}\n岗位：{job}\n时间：{dateText} {timeText}\n地址：{address}\n\n'
            '请在 BOSS 聊天页点击「发送面试」完成正式邀约。\n'
            '完成后请点击「已完成正式面试发送」按钮，程序才会继续下一位候选人。',
        )

    def hideFormalInterviewWait(self):
        """主线程：禁用确认按钮并恢复状态栏"""
        self.formalDoneBtn.configure(state=tk.DISABLED)
        if self.workerThread and self.workerThread.is_alive():
            self.statusVar.set('运行中：招聘任务...')
        else:
            self.statusVar.set('就绪')

    def confirmFormalInterviewDone(self):
        """用户确认已在 BOSS 完成正式面试邀约"""
        if not self.bossAuto or not self.bossAuto.formalInterviewWaitActive:
            messagebox.showwarning('提示', '当前不在等待正式面试邀约阶段')
            return
        self.bossAuto.signalFormalInterviewDone()
        self.formalDoneBtn.configure(state=tk.DISABLED)
        self.appendLog('已收到确认：正式面试邀约已完成')

    def onManualReplyWait(self, info):
        """子线程回调：进入人工聊天回复等待"""
        if self.root:
            self.root.after(0, lambda: self.showManualReplyWait(info))

    def onManualReplyWaitDone(self):
        """子线程回调：结束人工聊天回复等待"""
        if self.root:
            self.root.after(0, self.hideManualReplyWait)

    def showManualReplyWait(self, info):
        """主线程：弹窗通知并启用人工回复确认按钮"""
        name = str(info.get('candidateName') or '')
        job = str(info.get('jobName') or '')
        reason = str(info.get('reason') or '')
        friendText = str(info.get('friendText') or '')
        mode = str(info.get('mode') or 'reply_only')
        self.manualReplyMode = mode
        self.manualReplyInfo = dict(info or {})
        # 有候选人文字消息时才开放本地模型回复建议
        if self.replyAssistBtn:
            state = tk.NORMAL if friendText else tk.DISABLED
            self.replyAssistBtn.configure(state=state)
        # 候选人回复场景仅开放继续求简历或不合适两个明确决定
        if mode == 'candidate_reply':
            self.manualReplyDoneBtn.configure(state=tk.DISABLED)
            self.manualResumeBtn.configure(state=tk.NORMAL)
            self.manualUnsuitableBtn.configure(state=tk.NORMAL)
        else:
            self.manualReplyDoneBtn.configure(state=tk.NORMAL)
            self.manualResumeBtn.configure(state=tk.DISABLED)
            self.manualUnsuitableBtn.configure(state=tk.DISABLED)
        self.statusVar.set(f'等待人工聊天回复：{name}')
        self.appendLog('—— 等待人工在 BOSS 聊天框自由回复 ——')
        detail = f'候选人：{name}\n岗位：{job}\n原因：{reason}'
        if friendText:
            detail += f'\n对方消息：{friendText[:200]}'
        if mode == 'candidate_reply':
            detail += '\n\n请先在 BOSS 聊天框人工回复。\n回复后点击「合适，继续索要简历」；表达不合适则点击「标记不合适」。'
        else:
            detail += '\n\n请在 BOSS 聊天框输入并发送回复。\n完成后请点击「已完成人工回复」按钮，程序才会继续当前候选人流程。'
        messagebox.showinfo('需人工操作', detail)

    def hideManualReplyWait(self):
        """主线程：禁用人工回复确认按钮并恢复状态栏"""
        self.manualReplyDoneBtn.configure(state=tk.DISABLED)
        self.manualResumeBtn.configure(state=tk.DISABLED)
        self.manualUnsuitableBtn.configure(state=tk.DISABLED)
        if self.replyAssistBtn:
            self.replyAssistBtn.configure(state=tk.DISABLED)
        self.manualReplyMode = ''
        self.manualReplyInfo = {}
        self.closeReplyAssist()
        if self.workerThread and self.workerThread.is_alive():
            self.statusVar.set('运行中：招聘任务...')
        else:
            self.statusVar.set('就绪')

    def requestManualReply(self):
        """用户手动切换当前候选人为人工回复"""
        if not self.bossAuto:
            messagebox.showwarning('提示', '请先开始招聘任务')
            return
        if not self.bossAuto.currentCandidateKey:
            messagebox.showwarning('提示', '当前没有正在处理的候选人')
            return
        if self.bossAuto.formalInterviewWaitActive:
            messagebox.showwarning('提示', '当前正在等待正式面试邀约，请先完成该步骤')
            return
        if self.bossAuto.manualReplyWaitActive:
            messagebox.showinfo('提示', '已在等待人工聊天回复')
            return
        if self.bossAuto.requestManualReply():
            self.appendLog(f'已请求切换人工回复：{self.bossAuto.currentCandidateName}')
        else:
            messagebox.showwarning('提示', '无法切换人工回复')

    def confirmManualReplyDone(self):
        """用户确认已在 BOSS 完成人工聊天回复"""
        if not self.bossAuto or not self.bossAuto.manualReplyWaitActive:
            messagebox.showwarning('提示', '当前不在等待人工聊天回复阶段')
            return
        if self.manualReplyMode == 'candidate_reply':
            messagebox.showwarning('提示', '请使用「合适，继续索要简历」或「标记不合适」')
            return
        self.bossAuto.signalManualReplyDone('done')
        self.appendLog('已收到确认：人工聊天回复已完成')

    def confirmResumeContinue(self):
        """人工确认已回复且候选人合适，可继续索要简历"""
        if not self.bossAuto or not self.bossAuto.manualReplyWaitActive:
            messagebox.showwarning('提示', '当前不在等待候选人回复处理阶段')
            return
        if self.manualReplyMode != 'candidate_reply':
            messagebox.showwarning('提示', '当前人工步骤不能发起求简历')
            return
        # 自动化线程会再次读取聊天时间线，未检测到人工回复时继续等待
        self.bossAuto.signalManualReplyDone('continue')
        self.appendLog('已提交：候选人合适，等待校验人工回复')

    def confirmManualUnsuitable(self):
        """人工确认候选人表达不合适并永久停止跟进"""
        if not self.bossAuto or not self.bossAuto.manualReplyWaitActive:
            messagebox.showwarning('提示', '当前不在等待候选人回复处理阶段')
            return
        if self.manualReplyMode != 'candidate_reply':
            messagebox.showwarning('提示', '当前人工步骤不能标记候选人不合适')
            return
        if not messagebox.askyesno('确认不合适', '确认将当前候选人永久标记为不合适吗？'):
            return
        self.bossAuto.signalManualReplyDone('unsuitable')
        self.manualResumeBtn.configure(state=tk.DISABLED)
        self.manualUnsuitableBtn.configure(state=tk.DISABLED)
        self.appendLog('已提交：标记当前候选人不合适')

    def pushTodayWecom(self):
        """推送今日日报到企业微信群（不导出本地文件）"""
        webhook = self.wecomWebhookVar.get().strip() if self.wecomWebhookVar else ''
        if not webhook:
            messagebox.showwarning('提示', '请先填写企业微信 Webhook URL')
            return
        if self.wecomPushThread and self.wecomPushThread.is_alive():
            messagebox.showinfo('提示', '日报推送进行中，请稍候')
            return
        self.saveWecomToDb()
        reportSettings = self.buildWecomReportSettings()

        def worker():
            """后台线程执行企微推送"""
            try:
                db = self.ensureDb()
                report = BossReport(database=db)
                report.loadSettings(reportSettings)
                self.logFromWorker('—— 开始推送今日日报到企业微信 ——')
                result = report.pushToday()
                sent = int(result.get('sent') or 0)
                interviewTotal = int(result.get('interviewTotal') or 0)
                batchCount = int(result.get('batchCount') or 0)
                self.logFromWorker(f'推送完成：共 {sent} 条消息，面试邀约 {interviewTotal} 人（{batchCount} 批）')
                if self.root:
                    self.root.after(0, lambda: messagebox.showinfo('推送成功', f'已发送 {sent} 条企微消息\n面试邀约 {interviewTotal} 人'))
            except Exception as exc:
                self.logFromWorker(f'企微推送失败: {exc}')
                if self.root:
                    self.root.after(0, lambda: messagebox.showerror('推送失败', str(exc)))
            finally:
                if self.root and self.wecomPushBtn:
                    self.root.after(0, lambda: self.wecomPushBtn.configure(state=tk.NORMAL))

        self.wecomPushBtn.configure(state=tk.DISABLED)
        self.appendLog('正在推送今日日报到企业微信...')
        self.wecomPushThread = threading.Thread(target=worker, name='WecomPush', daemon=True)
        self.wecomPushThread.start()

    def buildWecomReportSettings(self):
        """组装企微日报推送配置（缓存 Webhook/手机号 + job 默认批次参数）"""
        gs = BossJob().globalSettings
        webhook = self.wecomWebhookVar.get().strip() if self.wecomWebhookVar else ''
        mobile = self.wecomMobileVar.get().strip() if self.wecomMobileVar else ''
        return {
            'wecomWebhookUrl': webhook,
            'wecomMentionMobile': mobile,
            'wecomInterviewBatchSize': int(gs.get('wecomInterviewBatchSize') or 10),
            'wecomPushGapSec': float(gs.get('wecomPushGapSec') or 1.5),
        }

    def saveWecomToDb(self):
        """保存企微 Webhook 与 @ 手机号到本地缓存"""
        db = self.ensureDb()
        webhook = self.wecomWebhookVar.get().strip() if self.wecomWebhookVar else ''
        mobile = self.wecomMobileVar.get().strip() if self.wecomMobileVar else ''
        db.saveWecomSettings(webhook, mobile)

    def loadWecomToGui(self):
        """从本地缓存加载企微配置到界面"""
        db = self.ensureDb()
        gs = BossJob().globalSettings
        webhook = db.getWecomWebhook(str(gs.get('wecomWebhookUrl') or ''))
        mobile = db.getWecomMentionMobile(str(gs.get('wecomMentionMobile') or ''))
        if self.wecomWebhookVar:
            self.wecomWebhookVar.set(webhook)
        if self.wecomMobileVar:
            self.wecomMobileVar.set(mobile)

    def saveWecomSettings(self):
        """用户点击保存企微配置"""
        self.saveWecomToDb()
        self.appendLog('已保存企业微信 Webhook 与 @ 手机号')

    def loadRecommendSettings(self):
        """从数据库加载推荐牛人主动联系配置与今日进度"""
        db = self.ensureDb()
        if self.recommendLimitVar:
            self.recommendLimitVar.set(str(db.getRecommendLimit()))
        self.refreshRecommendCount()

    def refreshRecommendCount(self):
        """刷新推荐牛人今日主动联系次数显示"""
        if not self.recommendCountVar:
            return
        db = self.ensureDb()
        used = db.countTodayRecommend()
        dailyLimit = db.getRecommendLimit()
        # 仅展示推荐牛人首次招呼额度，不混入其他聊天动作
        self.recommendCountVar.set(f'今日已主动联系 {used}/{dailyLimit} 人')

    def saveRecommendSettings(self):
        """保存推荐牛人每日主动联系上限"""
        raw = self.recommendLimitVar.get().strip() if self.recommendLimitVar else ''
        try:
            dailyLimit = int(raw)
        except ValueError:
            messagebox.showwarning('提示', '每日主动联系上限必须是整数')
            return False
        if dailyLimit < 1:
            messagebox.showwarning('提示', '每日主动联系上限必须大于 0')
            return False
        db = self.ensureDb()
        saved = db.saveRecommendLimit(dailyLimit)
        self.recommendLimitVar.set(str(saved))
        self.refreshRecommendCount()
        self.appendLog(f'已保存推荐牛人每日主动联系上限：{saved}')
        return True

    def openUnsuitableList(self):
        """打开不合适候选人名单供人工解除"""
        if self.unsuitableWin and self.unsuitableWin.winfo_exists():
            self.unsuitableWin.lift()
            return
        db = self.ensureDb()
        self.unsuitableRows = db.getUnsuitableList()
        self.unsuitableWin = tk.Toplevel(self.root)
        self.unsuitableWin.title('不合适候选人')
        self.unsuitableWin.geometry('620x420')
        frame = ttk.Frame(self.unsuitableWin, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text='解除前请先确认已在 BOSS 人工回复；未回复的会话下次仍会被重新标记。').pack(anchor=tk.W, pady=(0, 8))
        self.unsuitableListbox = tk.Listbox(frame)
        self.unsuitableListbox.pack(fill=tk.BOTH, expand=True)
        # 显示候选人姓名、原因与最近更新时间
        for row in self.unsuitableRows:
            name = str(row.get('candidate_name') or '')
            reason = str(row.get('stop_reason') or '')
            updatedAt = str(row.get('updated_at') or '')
            self.unsuitableListbox.insert(tk.END, f'{name} | {reason} | {updatedAt}')
        btnRow = ttk.Frame(frame)
        btnRow.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btnRow, text='解除选中候选人', command=self.restoreUnsuitable).pack(side=tk.LEFT)
        ttk.Button(btnRow, text='关闭', command=self.unsuitableWin.destroy).pack(side=tk.RIGHT)

    def restoreUnsuitable(self):
        """解除名单中选中候选人的不合适状态"""
        if not self.unsuitableListbox:
            return
        selected = self.unsuitableListbox.curselection()
        if not selected:
            messagebox.showwarning('提示', '请先选择一位候选人')
            return
        index = int(selected[0])
        row = self.unsuitableRows[index]
        name = str(row.get('candidate_name') or '')
        if not messagebox.askyesno('确认解除', f'确认解除 {name} 的不合适状态吗？'):
            return
        db = self.ensureDb()
        if not db.restoreUnsuitable(str(row.get('candidate_key') or '')):
            messagebox.showwarning('提示', '候选人状态已变化，请重新打开名单')
            return
        # 数据库恢复成功后同步移除当前列表项
        self.unsuitableRows.pop(index)
        self.unsuitableListbox.delete(index)
        self.appendLog(f'已人工解除不合适状态：{name}')

    def ensureDb(self):
        """初始化数据库；话术/岗位首次 seed，全局简历规则仍从 job.py 灌库"""
        # 首次使用时创建数据库连接
        if not self.db:
            self.db = BossDb()
        # 补全表结构
        self.db.migrateSchema()
        tpl = BossTemplate()
        job = BossJob()
        gs = job.globalSettings
        # 首次运行写入企微默认 Webhook / @ 手机号（已有缓存不覆盖）
        self.db.seedWecomSettings(gs.get('wecomWebhookUrl'), gs.get('wecomMentionMobile'))
        # 话术仅首次从 template.py 灌入，GUI 保存后不再覆盖
        self.db.seedTemplatesIfEmpty(tpl.bundle())
        # 岗位规则仅首次从 job.py 灌入，GUI 保存后不再覆盖
        self.db.seedJobRulesIfEmpty(job.profiles)
        # 话术占位符仅首次写入默认配置
        self.db.seedPlaceholderConfigIfEmpty(gs)
        # 推荐牛人每日主动联系额度仅首次写入默认值 15
        self.db.seedRecommendLimit()
        # 本地回复模型与默认 Skill 仅首次写入
        self.db.seedReply(self.reply.defaultSkill())
        # 全局简历规则仍全量覆盖（阶段 2 再 GUI 化）
        self.db.reloadResumeRulesFromConfig(gs)
        return self.db

    def buildTaskParams(self):
        """从数据库组装简历任务参数"""
        db = self.ensureDb()
        # 读取全局简历筛选规则
        rules = db.getResumeRules()

        def pickWords(templateType):
            """从 message_templates 表读取某类已启用话术"""
            rows = db.getTemplates(templateType, enabledOnly=True)
            # 过滤空话术，返回纯文本列表
            return [str(row.get('word') or '').strip() for row in rows if str(row.get('word') or '').strip()]

        gs = BossJob().globalSettings
        placeholderCfg = db.getPlaceholderConfig(gs)
        timeSlots = self.parsePlaceholderTimeSlots(placeholderCfg.get('timeSlots'))
        try:
            dayOffset = int(placeholderCfg.get('dayOffset') if placeholderCfg.get('dayOffset') is not None else 1)
        except (TypeError, ValueError):
            dayOffset = 1
        rateLimits = {
            'todayOnly': bool(gs.get('todayOnly', True)),
            'maxCandidatesPerRun': int(gs.get('maxCandidatesPerRun') or 5),
            'maxMessagesPerDay': int(gs.get('maxMessagesPerDay') or 50),
            'maxPerActionType': int(gs.get('maxPerActionType') or 10),
            'chatIntervalMin': int(gs.get('chatIntervalMin') or 25),
            'chatIntervalMax': int(gs.get('chatIntervalMax') or 50),
            'skipIntervalMin': float(gs.get('skipIntervalMin') or 3),
            'skipIntervalMax': float(gs.get('skipIntervalMax') or 6),
            'minActionGapMin': float(gs.get('minActionGapMin') or 2),
            'minActionGapMax': float(gs.get('minActionGapMax') or 5),
            'maxVerifyPerDay': int(gs.get('maxVerifyPerDay') or 2),
            'workWindows': list(gs.get('workWindows') or [['10:30', '11:30'], ['14:00', '17:00']]),
            'riskKeywords': list(gs.get('riskKeywords') or []),
            'handoffWhenLimit': bool(gs.get('handoffWhenLimit', True)),
        }
        interviewConfig = {
            'dayOffset': dayOffset,
            'timeSlots': timeSlots,
            'timeSpread': bool(gs.get('interviewTimeSpread', True)),
            'address': str(placeholderCfg.get('address') or gs.get('interviewAddress') or ''),
            'duration': str(placeholderCfg.get('duration') or gs.get('interviewDuration') or '40-60'),
            'noReplyHours': float(gs.get('interviewNoReplyHours') or 1),
            'cancelAfterRemind': bool(gs.get('interviewCancelAfterRemind', True)),
        }
        manualHandoff = {
            'keywords': list(gs.get('manualHandoffKeywords') or []),
            'whenUnknownIntent': bool(gs.get('manualWhenUnknownIntent', True)),
            'whenNoTemplate': bool(gs.get('manualWhenNoTemplate', True)),
        }
        # 话术来自数据库（GUI 可编辑保存），规则来自 job.py globalSettings
        return {
            'greetingWords': pickWords('greeting'),
            'followupWords': pickWords('followup'),
            'interviewPreWords': pickWords('interview_pre'),
            'interviewAskWords': pickWords('interview_ask_time'),
            'interviewRemindWords': pickWords('interview_remind'),
            'interviewRescheduleWords': pickWords('interview_reschedule'),
            'interviewCancelWords': pickWords('interview_cancel'),
            'rejectWords': pickWords('reject'),
            'replyInterestWords': pickWords('reply_interest'),
            'replyResumeWords': pickWords('reply_resume'),
            'replyLearnWords': pickWords('reply_learn'),
            'interviewTime': rules.get('interviewTime') or '明天下午14:00',
            'chatInterval': int(gs.get('chatIntervalMin') or rules.get('chatInterval') or 25),
            'maxFollowDays': int(rules.get('maxFollowDays') or 7),
            'noForeigner': True,
            'testCandidateName': self.testCandidateName,
            'recommend': {
                'dailyLimit': db.getRecommendLimit(),
            },
            'rateLimits': rateLimits,
            'interviewConfig': interviewConfig,
            'placeholderConfig': placeholderCfg,
            'manualHandoff': manualHandoff,
            'resumeRules': {
                'ageMin': int(rules.get('ageMin') or 18),
                'ageMax': int(rules.get('ageMax') or 45),
                'educationList': list(rules.get('educationList') or ['本科', '大专']),
                'workYearsMin': int(rules.get('workYearsMin') or 0),
                'mustKeywords': list(rules.get('mustKeywords') or []),
                'rejectKeywords': list(rules.get('rejectKeywords') or []),
                'interviewTime': rules.get('interviewTime') or '明天下午14:00',
                'maxFollowDays': int(rules.get('maxFollowDays') or 7),
                'chatInterval': int(gs.get('chatIntervalMin') or rules.get('chatInterval') or 25),
            },
        }

    def runResumeTask(self, reusePage=None):
        """执行消息列表简历自动化（可复用登录阶段的浏览器）"""
        # 读取用户指定的 Chrome 配置目录
        userDataPath = self.pathVar.get().strip()
        db = self.ensureDb()
        # 创建自动化实例，日志回调走队列
        self.bossAuto = BossAuto(
            browserId='gui-browser',
            userDataPath=userDataPath,
            database=db,
            logCallback=self.logFromWorker,
            onFormalInterviewWait=self.onFormalInterviewWait,
            onFormalInterviewWaitDone=self.onFormalInterviewWaitDone,
            onManualReplyWait=self.onManualReplyWait,
            onManualReplyWaitDone=self.onManualReplyWaitDone,
        )
        self.bossAuto.connectMode = 'local'
        # 清除上次可能残留的停止标志
        self.bossAuto.stopFlag.clear()
        # 登录阶段已打开浏览器则直接复用 page
        if reusePage:
            self.bossAuto.page = reusePage
        # 构造任务队列：一条 resume 任务 + 一条 stop 结束信号
        taskQueue = Queue()
        taskQueue.put({'taskType': 'resume', 'params': self.buildTaskParams()})
        taskQueue.put({'taskType': 'stop'})
        # 进入自动化主循环
        self.bossAuto.main(taskQueue)

    def startWorker(self, mode):
        """后台线程：login / task"""
        # 已有线程在跑则提示并退出
        if self.workerThread and self.workerThread.is_alive():
            messagebox.showinfo('提示', '已有任务在运行中')
            return
        willRunTask = mode == 'task' or (mode == 'login' and bool(self.autoTaskVar.get()))
        # 开始招聘任务前保存 GUI 当前额度，避免使用旧配置
        if willRunTask and not self.saveRecommendSettings():
            return
        # 同步 Chrome 目录与登录模式到 login 实例
        userDataPath = self.pathVar.get().strip()
        self.login.userDataPath = userDataPath
        self.login.loginMode = 'scan'
        self.login.stopFlag.clear()
        # 读取是否登录后自动跑简历任务
        self.autoTaskAfterLogin = bool(self.autoTaskVar.get())
        if mode == 'login':
            self.appendLog('—— 开始登录流程 ——')
            self.setUiBusy(True, '运行中：登录...')
        elif mode == 'task':
            self.appendLog('—— 开始推荐牛人与沟通处理任务 ——')
            self.setUiBusy(True, '运行中：招聘任务...')
        else:
            self.setUiBusy(True, '运行中...')

        def worker():
            """后台线程实际执行登录与/或简历任务"""
            try:
                reusePage = None
                if mode == 'login':
                    # 执行扫码登录流程
                    self.login.run(logCallback=self.logFromWorker)
                    # 用户中途停止则直接结束
                    if self.login.stopFlag.is_set():
                        self.logFromWorker('登录已停止')
                        return
                    self.logFromWorker('—— 登录流程完成 ——')
                    # 勾选自动任务时接管浏览器 page，避免重复启动
                    if self.autoTaskAfterLogin:
                        reusePage = self.login.page
                        self.login.page = None
                # 纯任务模式，或登录后自动任务
                shouldRunTask = mode == 'task' or (mode == 'login' and self.autoTaskAfterLogin)
                if shouldRunTask:
                    if reusePage:
                        self.logFromWorker('—— 复用登录浏览器，开始扫描沟通列表 ——')
                    else:
                        self.logFromWorker('—— 开始扫描沟通列表并按状态处理 ——')
                    # 执行消息列表简历自动化
                    self.runResumeTask(reusePage=reusePage)
                    self.logFromWorker('—— 招聘任务结束 ——')
            except Exception as exc:
                # 异常写入日志供界面显示
                self.logFromWorker(f'错误: {exc}')
            finally:
                # 回到主线程恢复按钮可用
                if self.root:
                    self.root.after(0, lambda: self.setUiBusy(False))
                    self.root.after(0, self.refreshRecommendCount)
        # 启动守护线程，不阻塞 GUI
        self.workerThread = threading.Thread(target=worker, name='BossGuiWorker', daemon=True)
        self.workerThread.start()

    def startLogin(self):
        """开始扫码登录（可选登录后自动任务）"""
        self.startWorker('login')

    def startTaskOnly(self):
        """跳过登录，直接开始简历任务（需 Chrome 目录已有登录态）"""
        self.startWorker('task')

    def stopAll(self):
        """停止登录或自动化"""
        # 通知登录模块停止
        self.login.requestStop()
        # 通知自动化模块停止
        if self.bossAuto:
            self.bossAuto.requestStop()
        self.appendLog('已发送停止请求...')
        self.statusVar.set('正在停止...')

    def buildUi(self):
        """构建界面"""
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title(self.title)
        self.root.geometry(f'{self.winWidth}x{self.winHeight}')
        self.root.minsize(440, 400)
        pad = {'padx': 10, 'pady': 6}
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        # Chrome 用户数据目录输入
        ttk.Label(frame, text='Chrome 用户数据目录（保存登录态）:').pack(anchor=tk.W, **pad)
        self.pathVar = tk.StringVar(value=self.defaultUserDataPath)
        ttk.Entry(frame, textvariable=self.pathVar).pack(fill=tk.X, **pad)
        # 登录后是否自动跑简历任务
        self.autoTaskVar = tk.IntVar(value=1)
        ttk.Checkbutton(frame, text='登录成功后自动开始招聘任务', variable=self.autoTaskVar).pack(anchor=tk.W, **pad)
        # 操作说明
        hint = '1. 登录成功后先浏览推荐牛人，仅联系刚刚/在线、今日或 3 天内活跃且岗位匹配的人\n2. 候选人回复后必须人工处理，可用本地模型生成并填入建议，但发送仍由人工完成\n3. 人工回复后再选择继续索要简历或标记不合适；未回复时不追问、不索要简历\n4. 历史已读未回自动标记不合适；收到简历后继续按原岗位规则审核\n5. 若出现 BOSS 安全验证，请在浏览器中手动完成'
        ttk.Label(frame, text=hint, wraplength=520, foreground='#555').pack(anchor=tk.W, **pad)
        groupPad = {'padx': 10, 'pady': 4}
        btnInnerPad = 6
        btnGap = {'padx': (0, 8), 'pady': 2}
        # 任务控制
        taskGroup = ttk.LabelFrame(frame, text='任务控制')
        taskGroup.pack(fill=tk.X, **groupPad)
        taskRow = ttk.Frame(taskGroup, padding=btnInnerPad)
        taskRow.pack(fill=tk.X)
        self.loginBtn = ttk.Button(taskRow, text='开始登录', command=self.startLogin)
        self.loginBtn.pack(side=tk.LEFT, **btnGap)
        self.taskBtn = ttk.Button(taskRow, text='开始招聘任务', command=self.startTaskOnly)
        self.taskBtn.pack(side=tk.LEFT, **btnGap)
        self.stopBtn = ttk.Button(taskRow, text='停止', command=self.stopAll, state=tk.DISABLED)
        self.stopBtn.pack(side=tk.LEFT, **btnGap)
        # 推荐牛人主动联系额度
        recommendGroup = ttk.LabelFrame(frame, text='推荐牛人主动联系')
        recommendGroup.pack(fill=tk.X, **groupPad)
        recommendRow = ttk.Frame(recommendGroup, padding=btnInnerPad)
        recommendRow.pack(fill=tk.X)
        ttk.Label(recommendRow, text='每日上限:').pack(side=tk.LEFT, padx=(0, 6))
        self.recommendLimitVar = tk.StringVar(value='15')
        ttk.Spinbox(recommendRow, from_=1, to=999, width=6, textvariable=self.recommendLimitVar).pack(side=tk.LEFT, padx=(0, 8))
        self.recommendSaveBtn = ttk.Button(recommendRow, text='保存上限', command=self.saveRecommendSettings)
        self.recommendSaveBtn.pack(side=tk.LEFT, **btnGap)
        self.recommendCountVar = tk.StringVar(value='今日已主动联系 0/15 人')
        ttk.Label(recommendRow, textvariable=self.recommendCountVar).pack(side=tk.LEFT, padx=(6, 0))
        # 人工介入
        manualGroup = ttk.LabelFrame(frame, text='人工介入（任务运行中）')
        manualGroup.pack(fill=tk.X, **groupPad)
        manualRow1 = ttk.Frame(manualGroup, padding=btnInnerPad)
        manualRow1.pack(fill=tk.X)
        self.formalDoneBtn = ttk.Button(manualRow1, text='已完成正式面试发送', command=self.confirmFormalInterviewDone, state=tk.DISABLED)
        self.formalDoneBtn.pack(side=tk.LEFT, **btnGap)
        manualRow2 = ttk.Frame(manualGroup, padding=btnInnerPad)
        manualRow2.pack(fill=tk.X)
        self.manualHandoffBtn = ttk.Button(manualRow2, text='切换人工回复', command=self.requestManualReply, state=tk.DISABLED)
        self.manualHandoffBtn.pack(side=tk.LEFT, **btnGap)
        self.manualReplyDoneBtn = ttk.Button(manualRow2, text='已完成人工回复', command=self.confirmManualReplyDone, state=tk.DISABLED)
        self.manualReplyDoneBtn.pack(side=tk.LEFT, **btnGap)
        manualRow3 = ttk.Frame(manualGroup, padding=btnInnerPad)
        manualRow3.pack(fill=tk.X)
        self.manualResumeBtn = ttk.Button(manualRow3, text='合适，继续索要简历', command=self.confirmResumeContinue, state=tk.DISABLED)
        self.manualResumeBtn.pack(side=tk.LEFT, **btnGap)
        self.manualUnsuitableBtn = ttk.Button(manualRow3, text='标记不合适', command=self.confirmManualUnsuitable, state=tk.DISABLED)
        self.manualUnsuitableBtn.pack(side=tk.LEFT, **btnGap)
        manualRow4 = ttk.Frame(manualGroup, padding=(btnInnerPad, 0, btnInnerPad, btnInnerPad))
        manualRow4.pack(fill=tk.X)
        self.replyAssistBtn = ttk.Button(manualRow4, text='本地模型回复建议', command=self.openReplyAssist, state=tk.DISABLED)
        self.replyAssistBtn.pack(side=tk.LEFT, **btnGap)
        # 配置管理
        configGroup = ttk.LabelFrame(frame, text='配置管理')
        configGroup.pack(fill=tk.X, **groupPad)
        configRow = ttk.Frame(configGroup, padding=btnInnerPad)
        configRow.pack(fill=tk.X)
        self.templateBtn = ttk.Button(configRow, text='话术模板管理', command=self.openTemplateEditor)
        self.templateBtn.pack(side=tk.LEFT, **btnGap)
        self.jobBtn = ttk.Button(configRow, text='岗位规则管理', command=self.openJobEditor)
        self.jobBtn.pack(side=tk.LEFT, **btnGap)
        self.unsuitableBtn = ttk.Button(configRow, text='不合适名单', command=self.openUnsuitableList)
        self.unsuitableBtn.pack(side=tk.LEFT, **btnGap)
        configRow2 = ttk.Frame(configGroup, padding=(btnInnerPad, 0, btnInnerPad, btnInnerPad))
        configRow2.pack(fill=tk.X)
        self.replyConfigBtn = ttk.Button(configRow2, text='本地回复模型与 Skill', command=self.openReplyEditor)
        self.replyConfigBtn.pack(side=tk.LEFT, **btnGap)
        # 企业微信配置
        wecomGroup = ttk.LabelFrame(frame, text='企业微信配置')
        wecomGroup.pack(fill=tk.X, **groupPad)
        wecomRow = ttk.Frame(wecomGroup, padding=btnInnerPad)
        wecomRow.pack(fill=tk.X)
        ttk.Label(wecomRow, text='Webhook:').pack(anchor=tk.W)
        self.wecomWebhookVar = tk.StringVar(value='')
        ttk.Entry(wecomRow, textvariable=self.wecomWebhookVar).pack(fill=tk.X, pady=(2, 0))
        ttk.Label(wecomRow, text='@ 通知手机号（企微绑定手机号，多个用逗号分隔）:').pack(anchor=tk.W, pady=(4, 0))
        self.wecomMobileVar = tk.StringVar(value='')
        ttk.Entry(wecomRow, textvariable=self.wecomMobileVar).pack(fill=tk.X, pady=(2, 0))
        wecomBtnGroup = ttk.LabelFrame(frame, text='企微日报')
        wecomBtnGroup.pack(fill=tk.X, **groupPad)
        wecomBtnRow = ttk.Frame(wecomBtnGroup, padding=btnInnerPad)
        wecomBtnRow.pack(fill=tk.X)
        self.wecomSaveBtn = ttk.Button(wecomBtnRow, text='保存企微配置', command=self.saveWecomSettings)
        self.wecomSaveBtn.pack(side=tk.LEFT, **btnGap)
        self.wecomPushBtn = ttk.Button(wecomBtnRow, text='推送今日日报到企业微信', command=self.pushTodayWecom)
        self.wecomPushBtn.pack(side=tk.LEFT, **btnGap)
        # 状态栏
        self.statusVar = tk.StringVar(value='就绪')
        ttk.Label(frame, textvariable=self.statusVar, foreground='#0066cc').pack(anchor=tk.W, **pad)
        # 只读日志区域
        ttk.Label(frame, text='运行日志:').pack(anchor=tk.W)
        self.logText = scrolledtext.ScrolledText(frame, height=8, state=tk.DISABLED, font=('Consolas', 9))
        self.logText.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        # 关闭窗口时先停任务
        self.root.protocol('WM_DELETE_WINDOW', self.onClose)

    def onClose(self):
        """关闭窗口前停止后台任务"""
        self.stopAll()
        # 关闭数据库连接
        if self.db:
            self.db.close()
        self.root.destroy()

    def run(self):
        """启动 GUI"""
        self.buildUi()
        self.loadWecomToGui()
        self.loadRecommendSettings()
        # 启动日志队列轮询
        self.pollLogQueue()
        self.root.mainloop()

if __name__ == '__main__':
    config = {'userDataPath': 'D:\\boss_zhaopin_筛选简历\\boss_chrome_profile', 'testCandidateName': ''}
    app = RunGui()
    app.defaultUserDataPath = config['userDataPath']
    app.testCandidateName = config.get('testCandidateName') or ''
    app.run()
