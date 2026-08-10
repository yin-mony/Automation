import contextlib
import datetime as dt
import json
import random
import re
import threading
import time
import traceback
from pathlib import Path
from queue import Empty, Queue
from urllib.parse import urljoin
import requests
from .db import BossDb
from .match import ResumeMatch
from .parse import ResumeParse
try:
    from DrissionPage import ChromiumOptions, ChromiumPage
except ImportError:
    ChromiumOptions = None
    ChromiumPage = None

class StopRequested(RuntimeError):
    """用户请求立即停止自动化"""

class BossAuto:
    """BOSS 直聘 DrissionPage 自动化：话术、求简历、审核、面试邀请"""

    def __init__(self, browserId, userDataPath='', database=None, logCallback=None, pauseEvent=None, taskDoneCallback=None, onFormalInterviewWait=None, onFormalInterviewWaitDone=None, onManualReplyWait=None, onManualReplyWaitDone=None):
        """初始化浏览器自动化配置、状态与人工回调"""
        self.browserId = browserId  # 浏览器实例 ID（ADS profile 或本地标识）
        self.userDataPath = userDataPath  # Chrome 用户数据目录，本地模式持久化登录
        self.db = database  # 数据库访问对象，记录候选人状态与操作日志
        self.logCallback = logCallback  # 日志回调，推送到 Web 前端
        self.taskDoneCallback = taskDoneCallback  # 任务完成回调，通知上层调度
        self.currentTaskRunId = None  # 当前任务运行记录 ID，写入 action 日志
        self.page = None  # DrissionPage 页面对象
        self.stopFlag = threading.Event()  # 停止信号，用户请求立即停止
        self.pauseEvent = pauseEvent or threading.Event()  # 暂停信号，clear 表示暂停中
        self.pauseLogged = False  # 是否已输出「任务已暂停」日志，避免重复刷屏
        self.mousePauseFlag = threading.Event()  # 鼠标线程暂停标志（登录/验证/暂停时）
        self.actionLock = threading.RLock()  # 页面操作与鼠标移动的互斥锁
        self.mouseThread = None  # 后台随机移动鼠标的守护线程
        self.mousePos = [300, 300]  # 鼠标线程最近一次移动到的坐标
        self.parser = ResumeParse()  # 简历解析器，从预览弹窗或页面提取结构化信息
        self.matcher = ResumeMatch()  # 简历匹配器，按岗位要求审核通过/不通过
        self.adsUrl = 'http://127.0.0.1:50325'  # ADS 指纹浏览器本地 API 地址
        self.bitUrl = 'http://127.0.0.1:54345'  # BitBrowser 本地 API 地址（预留）
        self.connectMode = 'local'  # 浏览器连接模式：local 本地启动 / ads 远程调试
        self.minActionGap = 2.0  # 两次页面操作之间的最小间隔秒数
        self.loginPollSec = 2.0  # 等待登录时的轮询间隔
        self.loginRemindSec = 30.0  # 等待登录时的提醒间隔
        self.verifyPollSec = 2.0  # 等待安全验证时的轮询间隔
        self.verifyRemindSec = 30.0  # 等待安全验证时的提醒间隔
        # BOSS 安全验证页面关键词，用于检测滑块/点选等人机验证
        self.verifyKeywords = ['安全验证', '滑动验证', '人机验证', '请完成验证', '请完成安全验证', '向右滑动', '拖动滑块', '滑动完成验证', '点击完成验证', '异常访问', '访问验证', '行为验证', '完成拼图']
        self.verifyWaiting = False  # 当前是否处于等待用户完成安全验证的状态
        # 智能回复意图规则：关键词 → intent，供 trySendSmartReply 选用话术
        self.replyIntentRules = [{'intent': 'reply_resume', 'keywords': ['发简历', '发一份简历', '发个简历', '发份简历', '发您简历', '看下我的简历', '简历看看', '我的简历', '发我的简历']}, {'intent': 'reply_learn', 'keywords': ['想了解', '了解一下', '介绍一下', '介绍介绍', '这一职位', '这个岗位', '这个职位', '职位详情', '占用您一点时间', '了解下']}, {'intent': 'reply_interest', 'keywords': ['感兴趣', '很感兴趣', '特别感兴趣', '进一步沟通', '可以沟通', '聊一聊', '聊一下', '想要加入', '很想', '非常想要', '很匹配', '岗位觉得', '觉得匹配', '方便沟通']}]
        # 面试预邀请回复意图关键词
        self.interviewDeclineKeywords = ['不方便', '没空', '无法', '不能', '去不了', '有事', '冲突', '不行', '不太方便', '没时间', '顾不上']
        self.interviewConfirmKeywords = ['可以', '好的', '好', '没问题', '方便', '确认', 'ok', 'OK', '行', '准时', '参加', '没问题']
        self.onFormalInterviewWait = onFormalInterviewWait
        self.onFormalInterviewWaitDone = onFormalInterviewWaitDone
        self.formalInterviewDoneEvent = threading.Event()
        self.formalInterviewWaitActive = False
        self.formalInterviewRemindSec = 30.0
        self.onManualReplyWait = onManualReplyWait
        self.onManualReplyWaitDone = onManualReplyWaitDone
        self.manualReplyDoneEvent = threading.Event()
        self.manualReplyWaitActive = False
        self.manualReplyRemindSec = 30.0
        self.manualReplyRequested = False
        self.manualReplyDecision = ''
        self.manualReplyMode = ''
        self.currentCandidateKey = ''
        self.currentCandidateName = ''
        self.currentJobName = ''
        self.currentTask = None
        self.taskRateLimits = {}
        self.processedThisRun = 0
        self.verifyCountedEpisode = False
        # 推荐牛人页面最多连续空滚动次数
        self.recommendMaxIdle = 3

    def applyRateLimits(self, task):
        """从任务参数加载风控与时段配置"""
        limits = dict(task.get('rateLimits') or {})
        self.taskRateLimits = limits
        self.processedThisRun = 0
        self.verifyCountedEpisode = False

    def rateLimits(self):
        """当前任务风控配置"""
        return self.taskRateLimits or {}

    def randomActionGap(self):
        """随机页面操作间隔"""
        limits = self.rateLimits()
        lo = float(limits.get('minActionGapMin') or 2)
        hi = float(limits.get('minActionGapMax') or 5)
        if hi < lo:
            hi = lo
        return random.uniform(lo, hi)

    def gapWait(self):
        """按随机间隔等待"""
        self.pauseWait(self.randomActionGap())

    def randomChatWait(self, task):
        """有效处理完候选人后的随机等待"""
        limits = dict(task.get('rateLimits') or {})
        lo = float(limits.get('chatIntervalMin') or task.get('chatInterval') or 25)
        hi = float(limits.get('chatIntervalMax') or max(lo + 20, 50))
        if hi < lo:
            hi = lo
        self.pauseWait(random.uniform(lo, hi))

    def skipChatWait(self, task):
        """跳过或无有效操作时的短等待"""
        limits = dict(task.get('rateLimits') or {})
        lo = float(limits.get('skipIntervalMin') or 3)
        hi = float(limits.get('skipIntervalMax') or 6)
        if hi < lo:
            hi = lo
        self.pauseWait(random.uniform(lo, hi))

    def parseHm(self, text):
        """解析 HH:MM 为 time"""
        parts = str(text or '').strip().split(':')
        if len(parts) != 2:
            return None
        try:
            hour = int(parts[0])
            minute = int(parts[1])
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                return dt.time(hour, minute)
        except ValueError:
            return None
        return None

    def getCurrentWorkWindow(self):
        """返回当前时刻所在工作时段 (start, end)，不在时段内返回 None"""
        windows = self.rateLimits().get('workWindows') or []
        if not windows:
            return None
        now = dt.datetime.now().time()
        for item in windows:
            if not item or len(item) < 2:
                continue
            start = self.parseHm(item[0])
            end = self.parseHm(item[1])
            if start and end and start <= now <= end:
                return (start, end)
        return None

    def ensureWorkWindow(self):
        """不在工作时段则停止任务（仅校验当前时钟，不筛选列表消息时间）"""
        if not self.rateLimits().get('workWindows'):
            return True
        if self.getCurrentWorkWindow():
            return True
        self.log('当前不在工作时段内，任务停止')
        self.requestStop(closeSession=False)
        raise StopRequested('不在工作时段')

    def isTodayTimeLabel(self, label):
        """判断 BOSS 时间标签是否表示「当天」"""
        text = str(label or '').strip()
        if not text:
            return True
        oldMarks = ['昨天', '前天', '上周', '星期', '周', '日前']
        for mark in oldMarks:
            if mark in text:
                return False
        if re.search(r'\d{4}[-/年]', text):
            return False
        if re.search(r'\d{1,2}[-/]\d{1,2}', text):
            return False
        today = dt.date.today()
        if f'{today.month}月{today.day}日' in text:
            return True
        if '刚刚' in text or '分钟' in text or '小时' in text:
            return True
        if re.search(r'\d{1,2}:\d{2}', text):
            return True
        return False

    def isPastTodayListBoundary(self, listLabel):
        """列表时间标签是否已越过当天（如昨天及更早）"""
        text = str(listLabel or '').strip()
        if not text:
            return False
        return not self.isTodayTimeLabel(text)

    def readListItemTimeLabel(self, listItem):
        """读取沟通列表项上的时间标签"""
        xpaths = [
            'xpath:.//span[contains(@class,"time")]',
            'xpath:.//div[contains(@class,"time")]',
            'xpath:.//*[contains(@class,"last-msg-time")]',
        ]
        for xpath in xpaths:
            ele = listItem.ele(xpath, timeout=0.2)
            if ele and str(ele.text or '').strip():
                return str(ele.text).strip()
        return ''

    def isListItemUnread(self, listItem):
        """在打开聊天前判断沟通列表项是否带未读标记"""
        selectors = [
            'xpath:.//*[contains(@class,"unread")]',
            'xpath:.//*[contains(@class,"badge") and not(contains(@class,"job"))]',
            'xpath:.//*[contains(@class,"notice-dot")]',
            'xpath:.//*[contains(@class,"message-count")]',
        ]
        # 未读快照必须在点击列表项前完成，避免程序打开聊天后误判
        for selector in selectors:
            element = listItem.ele(selector, timeout=0.1)
            if element:
                return True
        itemClass = str(listItem.attr('class') or '').lower()
        return 'unread' in itemClass

    def readMessageTimeLabel(self, msgEle):
        """读取单条聊天消息旁的时间标签"""
        xpaths = [
            'xpath:.//span[contains(@class,"time")]',
            'xpath:.//div[contains(@class,"time")]',
            'xpath:.//*[contains(@class,"msg-time")]',
        ]
        for xpath in xpaths:
            ele = msgEle.ele(xpath, timeout=0.1)
            if ele and str(ele.text or '').strip():
                return str(ele.text).strip()
        return ''

    def hasTodayFriendInTimeline(self, timeline):
        """聊天时间线里是否有对方「当天」的消息"""
        for item in timeline:
            if item.get('sender') != 'friend':
                continue
            if not str(item.get('text') or '').strip():
                continue
            if item.get('isToday', True):
                return True
        return False

    def resumeReceivedToday(self, candidateKey):
        """简历是否于今日收到"""
        if not self.db:
            return False
        row = self.db.getCandidate(candidateKey) or {}
        received = str(row.get('resume_received_at') or '')
        return received.startswith(self.db.todayText())

    def shouldProcessToday(self, task, candidateKey, candidateName, listItem, timeline, status, resumeCard):
        """是否满足「仅处理当天候选人消息」"""
        limits = task.get('rateLimits') or {}
        if not limits.get('todayOnly', True):
            return True
        if str(task.get('testCandidateName') or '').strip():
            return True
        # 人工刚确认继续时立即执行求简历，不再受历史消息日期限制
        if status == 'manual_approved':
            return True
        listLabel = self.readListItemTimeLabel(listItem)
        if listLabel and not self.isTodayTimeLabel(listLabel):
            # 预邀请已提醒、待自动取消的候选人允许继续处理
            if status != 'interview_pre_sent' or not self.shouldProcessForRemindCancel(candidateKey, task):
                self.log(f'跳过 {candidateName}: 列表时间非今日（{listLabel}）')
                return False
        if resumeCard in ('pending_accept', 'accepted'):
            if resumeCard == 'pending_accept' and (self.isTodayTimeLabel(listLabel) or self.hasTodayFriendInTimeline(timeline)):
                return True
            if self.resumeReceivedToday(candidateKey):
                return True
            if self.hasTodayFriendInTimeline(timeline):
                return True
            self.log(f'跳过 {candidateName}: 简历非今日收到')
            return False
        # 面试预邀请阶段：允许处理追问/改期/无回复提醒
        if status in ('interview_pre_sent', 'interview_awaiting_time', 'interview_reschedule_pending', 'interview_formal_pending'):
            if self.hasTodayFriendInTimeline(timeline):
                return True
            pendingToday = self.getPendingFriendTexts(timeline, todayOnly=True)
            if pendingToday:
                return True
            if self.needInterviewNoReplyRemind(candidateKey, task):
                return True
            if self.needInterviewAutoCancelAfterRemind(candidateKey, task, timeline, todayOnly=True):
                return True
            if status == 'interview_formal_pending':
                return True
            self.log(f'跳过 {candidateName}: 面试流程无当天待处理消息')
            return False
        if self.hasTodayFriendInTimeline(timeline):
            return True
        pendingToday = self.getPendingFriendTexts(timeline, todayOnly=True)
        if pendingToday:
            return True
        self.log(f'跳过 {candidateName}: 无当天待处理消息')
        return False

    def shouldPickListItem(self, task, candidateKey, candidateName, listLabel):
        """列表扫描阶段判断候选人是否可能有当天待办（无需点进聊天）"""
        if str(task.get('testCandidateName') or '').strip():
            return True
        if not self.db:
            return True
        maxFollowDays = int(task.get('maxFollowDays') or (task.get('resumeRules') or {}).get('maxFollowDays') or 7)
        canFollow, stopReason = self.db.canFollowCandidate(candidateKey, maxFollowDays)
        if not canFollow:
            return False
        status = self.db.getResumeStatus(candidateKey)
        if status in ('interview_sent', 'interview_cancelled', 'resume_rejected', 'unsuitable'):
            return False
        if self.alreadySentInterview(candidateKey):
            return False
        if status in ('interview_pre_sent', 'interview_awaiting_time', 'interview_reschedule_pending', 'interview_formal_pending'):
            return True
        if status in ('resume_received', 'resume_requested', 'resume_passed'):
            return True
        if self.db.pendingResumeReview(candidateKey):
            return True
        if status in ('new', 'greeted'):
            return True
        return True

    def scrollChatListDown(self):
        """沟通列表向下滚动，加载更多联系人"""
        result = self.page.run_js(
            """
            const box = document.querySelector('.user-container')
                || document.querySelector('div.user-container');
            if (!box) return [0, 0, false];
            const before = box.scrollTop;
            const step = Math.max(200, Math.floor(box.clientHeight * 0.8));
            box.scrollTop = before + step;
            return [before, box.scrollTop, box.scrollTop > before];
            """
        )
        if result and len(result) >= 3 and result[2]:
            self.pauseWait(0.8)
            return True
        return False

    def handoffEnabled(self):
        """话术配额用尽时是否改由人工在输入框确认发送"""
        return bool(self.rateLimits().get('handoffWhenLimit', True))

    def canAutoSend(self, actionType):
        """检查是否还能自动发送（仅判断，不停任务）"""
        if not self.db:
            return True
        limits = self.rateLimits()
        ok, reason = self.db.canSendMessage(
            actionType,
            int(limits.get('maxMessagesPerDay') or 50),
            int(limits.get('maxPerActionType') or 10),
        )
        if not ok:
            self.log(f'消息限流: {reason}')
        return ok

    def fillMsgOnly(self, text):
        """将话术填入聊天输入框，不点击发送"""
        self.page.run_js(f"\n            const el = document.querySelector('#boss-chat-editor-input');\n            if (el) {{\n                el.focus();\n                el.innerHTML = '';\n                document.execCommand('insertText', false, {text!r});\n                el.dispatchEvent(new Event('input', {{ bubbles: true }}));\n            }}\n            ")
        self.gapWait()
        return bool(self.page.ele('#boss-chat-editor-input', timeout=2))

    def fillManualReply(self, text):
        """将人工审核后的模型建议填入当前聊天框，但不点击发送"""
        replyText = str(text or '').strip()
        if not replyText:
            return False
        if not self.manualReplyWaitActive or not self.page:
            return False
        with self.actionLock:
            # 只填入输入框，发送动作必须由人工在 BOSS 页面完成
            filled = self.fillMsgOnly(replyText)
        if filled:
            self.log('本地模型建议已填入聊天框，请人工检查、修改并发送')
        return filled

    def sendOrHandoff(self, text, actionType, candidateKey=None):
        """自动发送；配额用尽则填入输入框供人工发送，返回 sent / handoff / handoff_done / 空"""
        if self.canAutoSend(actionType):
            if self.sendMsg(text):
                return 'sent'
            return ''
        if not self.handoffEnabled():
            self.log('消息限流且未启用手动接管，跳过发送')
            return ''
        if candidateKey and self.db and self.db.hasHandoffToday(candidateKey, actionType):
            return 'handoff_done'
        if self.fillMsgOnly(text):
            self.log('话术已达上限，已填入输入框，请人工确认发送')
            if candidateKey and self.db:
                self.db.recordHandoff(candidateKey, actionType, word=text, taskRunId=self.currentTaskRunId)
            return 'handoff'
        return ''

    def checkRiskPage(self):
        """检测页面风控文案并停跑"""
        keywords = self.rateLimits().get('riskKeywords') or []
        if not keywords:
            return
        text = self.pageText()
        for keyword in keywords:
            if keyword and keyword in text:
                self.log(f'检测到风控提示「{keyword}」，停止任务')
                self.requestStop(closeSession=False)
                raise StopRequested(f'风控: {keyword}')

    def onVerifyDetected(self):
        """记录安全验证次数，超限则停跑"""
        if self.verifyCountedEpisode:
            return
        self.verifyCountedEpisode = True
        if not self.db:
            return
        limits = self.rateLimits()
        maxCount = int(limits.get('maxVerifyPerDay') or 2)
        count = self.db.recordRiskEvent('verify')
        self.log(f'今日安全验证次数: {count}/{maxCount}')
        if count >= maxCount:
            self.log('安全验证过于频繁，今日停止自动任务')
            self.requestStop(closeSession=False)
            raise StopRequested('验证次数超限')

    def log(self, message):
        """输出日志到控制台和回调"""
        text = str(message)
        print(text)
        # 有回调则同步推送到 Web 前端
        if self.logCallback:
            self.logCallback(text)

    def requestStop(self, closeSession=True):
        """请求停止任务"""
        self.stopFlag.set()
        self.pauseEvent.clear()  # 解除暂停，让阻塞中的 waitIfPaused 尽快退出
        self.mousePauseFlag.set()  # 暂停鼠标线程，避免停止后仍操作页面
        # 可选关闭浏览器会话引用
        if closeSession and self.page:
            self.page = None

    def ensureNotStopped(self):
        """未停止才继续"""
        # 已收到停止请求则抛出 StopRequested
        if self.stopFlag.is_set():
            raise StopRequested('用户请求停止任务')

    def pollManualReplyRequest(self):
        """等待期间响应 GUI 手动切换人工回复"""
        if not self.manualReplyRequested or not self.currentCandidateKey:
            return
        if self.manualReplyWaitActive or self.formalInterviewWaitActive:
            return
        task = self.currentTask or {}
        self.ensureManualReply(task, self.currentCandidateKey, self.currentCandidateName, self.currentJobName, '')

    def waitIfPaused(self):
        """暂停时阻塞等待"""
        self.ensureNotStopped()
        # 未处于暂停状态则直接返回
        if not self.pauseEvent.is_set():
            return
        # 首次进入暂停时输出提示
        if not self.pauseLogged:
            self.log('任务已暂停，等待继续...')
            self.pauseLogged = True
        self.mousePauseFlag.set()
        # 循环等待直到恢复或停止
        while self.pauseEvent.is_set() and (not self.stopFlag.is_set()):
            time.sleep(0.2)
        if self.pauseLogged:
            self.log('任务已继续')
            self.pauseLogged = False
        self.mousePauseFlag.clear()
        self.ensureNotStopped()

    def pauseWait(self, seconds, step=0.2):
        """可暂停的等待"""
        waited = 0.0
        while waited < seconds:
            self.waitIfPaused()  # 每小段 sleep 前检查暂停/停止
            self.pollManualReplyRequest()
            sleepTime = min(step, seconds - waited)
            time.sleep(sleepTime)
            waited += sleepTime
            self.ensureNotStopped()

    def createPageLocal(self):
        """创建本地 Chrome/Edge 页面"""
        if ChromiumOptions is None:
            raise RuntimeError('缺少 DrissionPage，请执行: pip install DrissionPage')
        co = ChromiumOptions()
        co.set_argument('--disable-blink-features=AutomationControlled')
        # 指定用户数据目录以复用登录态
        if self.userDataPath:
            co.set_paths(user_data_path=self.userDataPath)
        return ChromiumPage(co)

    def reqAdsBrowser(self):
        """通过 ADS 启动浏览器并返回调试地址"""
        data = {'profile_id': self.browserId}
        res = requests.post(f'{self.adsUrl}/api/v2/browser-profile/start', json=data, timeout=30)
        res.raise_for_status()
        body = res.json()
        # ADS 返回非 success 则启动失败
        if body.get('msg') != 'success':
            raise RuntimeError(f'打开 ADS 浏览器失败: {self.browserId}')
        return body['data']['ws']['selenium']

    def initBrowser(self):
        """初始化浏览器连接"""
        if self.connectMode == 'ads':
            debugAddr = self.reqAdsBrowser()
            self.pauseWait(3)  # 等待 ADS 浏览器进程就绪
            if ChromiumOptions is None:
                raise RuntimeError('缺少 DrissionPage')
            co = ChromiumOptions()
            co.set_paths(address=debugAddr)
            self.page = ChromiumPage(co)
        else:
            self.page = self.createPageLocal()
        self.pauseWait(2)
        self.log('浏览器已连接')

    def selectChatTab(self):
        """进入 BOSS 沟通页"""
        chatUrl = 'https://www.zhipin.com/web/chat/index'
        # 不在沟通页则导航过去
        if chatUrl not in (self.page.url or ''):
            self.page.get(chatUrl)
        self.page.wait.doc_loaded()
        self.pauseWait(2)
        # 关闭可能出现的引导弹窗
        closeBtn = self.page.ele('xpath://i[@class="icon-close"]', timeout=1)
        if closeBtn:
            closeBtn.click()
        self.ensureVerifyClear()

    def selectRecommendTab(self):
        """进入 BOSS 推荐牛人页"""
        selectors = [
            'xpath://a[contains(normalize-space(.),"推荐牛人")]',
            'xpath://dl[contains(@class,"menu-recommend")]//a',
            'xpath://*[contains(@class,"menu")]//a[contains(@href,"recommend")]',
        ]
        clicked = False
        # 优先点击当前招聘端菜单，兼容站点保留登录上下文
        for selector in selectors:
            link = self.page.ele(selector, timeout=1)
            if not link:
                continue
            link.click()
            clicked = True
            break
        # 菜单未识别时直接进入招聘端推荐页
        if not clicked:
            self.page.get('https://www.zhipin.com/web/boss/recommend')
        self.page.wait.doc_loaded()
        self.pauseWait(2)
        self.ensureVerifyClear()
        cards = self.readRecommendCards()
        if not cards:
            self.log('推荐牛人页面未识别到候选人卡片，本轮跳过主动联系')
            return False
        return True

    def readRecommendCards(self):
        """读取推荐牛人页面当前已加载的候选人卡片"""
        selectors = [
            'xpath://div[contains(@class,"recommend-list")]//div[contains(@class,"geek-card")]',
            'xpath://div[contains(@class,"recommend-list")]//li[contains(@class,"card")]',
            'xpath://div[contains(@class,"geek-list")]//li',
            'xpath://div[contains(@class,"recommend")]//div[contains(@class,"candidate-card")]',
            'xpath://div[contains(@class,"recommend")]//div[contains(@class,"card-item")]',
        ]
        # 使用第一组有效选择器，避免同一卡片被嵌套节点重复返回
        for selector in selectors:
            cards = self.page.eles(selector, timeout=1) or []
            if cards:
                return cards
        return []

    def readRecommendPageJob(self):
        """读取推荐页当前选择的招聘岗位"""
        if not self.db:
            return ''
        selectors = [
            'xpath://*[contains(@class,"job-select")]//*[contains(@class,"selected")]',
            'xpath://*[contains(@class,"position-select")]//*[contains(@class,"selected")]',
            'xpath://*[contains(@class,"current-job")]',
            'xpath://*[contains(@class,"job-name") and contains(@class,"active")]',
        ]
        # 仅接受能命中已启用岗位规则的文本，避免误读候选人期望职位
        for selector in selectors:
            elements = self.page.eles(selector, timeout=0.5) or []
            for element in elements:
                text = str(element.text or '').strip()
                rules = self.db.matchJobRules(text)
                if rules:
                    return str(rules.get('jobName') or text)
        return ''

    def parseActiveRank(self, activeText):
        """将活跃标签转换为优先级，超过三天或未知返回 -1"""
        text = re.sub(r'\s+', '', str(activeText or ''))
        if not text:
            return -1
        # 刚刚活跃与在线优先级最高
        if '刚刚活跃' in text or text in ('刚刚', '在线') or '当前在线' in text:
            return 0
        # 今日活跃排在第二档
        if '今日活跃' in text or '今天活跃' in text or '今日在线' in text:
            return 1
        # 一至三天内活跃排在第三档
        if any(mark in text for mark in ['1天内活跃', '2天内活跃', '3天内活跃', '1日内活跃', '2日内活跃', '3日内活跃', '一天内活跃', '两天内活跃', '三天内活跃']):
            return 2
        match = re.search(r'([1-3])(?:天|日)(?:内|前)?活跃', text)
        if match:
            return 2
        return -1

    def readRecommendActive(self, card):
        """读取推荐牛人卡片的活跃标签"""
        selectors = [
            'xpath:.//*[contains(@class,"active-time")]',
            'xpath:.//*[contains(@class,"activity")]',
            'xpath:.//*[contains(@class,"online")]',
            'xpath:.//*[contains(text(),"活跃") or normalize-space(text())="在线"]',
        ]
        # 优先读取带语义 class 的活跃标签
        for selector in selectors:
            element = card.ele(selector, timeout=0.2)
            text = str(element.text or '').strip() if element else ''
            if text and self.parseActiveRank(text) >= 0:
                return text
        # 页面改版时从卡片逐行文本中兜底识别
        for line in str(card.text or '').splitlines():
            text = line.strip()
            if self.parseActiveRank(text) >= 0:
                return text
        return ''

    def readRecommendName(self, card):
        """读取推荐牛人卡片的候选人姓名"""
        selectors = [
            'xpath:.//*[contains(@class,"geek-name")]',
            'xpath:.//*[contains(@class,"candidate-name")]',
            'xpath:.//*[contains(@class,"name") and not(contains(@class,"job")) and not(contains(@class,"position")) and not(contains(@class,"company"))][self::span or self::div or self::h3]',
        ]
        # 姓名必须来自明确节点，无法识别时不执行联系
        for selector in selectors:
            element = card.ele(selector, timeout=0.2)
            text = str(element.text or '').strip() if element else ''
            if text and len(text) <= 30:
                return text
        return ''

    def recommendWords(self, rules):
        """汇总岗位规则中可用于推荐卡片匹配的关键词"""
        words = [str(rules.get('jobName') or '').strip()]
        for key in ('matchKeys', 'mustKeywords', 'preferKeywords'):
            words.extend(str(word or '').strip() for word in rules.get(key) or [])
        for group in rules.get('anyKeywords') or []:
            if isinstance(group, (list, tuple)):
                words.extend(str(word or '').strip() for word in group)
            else:
                words.append(str(group or '').strip())
        # 去除空词并保持原顺序，避免重复关键词放大分数
        return list(dict.fromkeys(word for word in words if word))

    def matchRecommendRules(self, cardText, pageJobName, jobRules):
        """按当前岗位与卡片可见信息选择最匹配的启用岗位规则"""
        text = re.sub(r'\s+', '', str(cardText or ''))
        pageRules = self.db.matchJobRules(pageJobName) if self.db and pageJobName else None
        bestRules = None
        bestScore = -1
        for rules in jobRules:
            # 已识别当前招聘岗位时，仅在对应岗位规则内比较候选人
            if pageRules and str(pageRules.get('jobName') or '') != str(rules.get('jobName') or ''):
                continue
            score = 0
            for word in self.recommendWords(rules):
                wordNorm = re.sub(r'\s+', '', word)
                if wordNorm and wordNorm in text:
                    score += 1
            # 推荐页当前岗位命中规则时，以 BOSS 的岗位推荐上下文为基础匹配
            if pageRules and str(pageRules.get('jobName') or '') == str(rules.get('jobName') or ''):
                score += 10
            if score > bestScore:
                bestRules = rules
                bestScore = score
        # 未识别当前岗位时，卡片至少需要命中一个岗位关键词
        if not pageRules and bestScore <= 0:
            return (None, -1)
        return (bestRules, bestScore)

    def getRecommendButton(self, card):
        """读取推荐卡片中的首次沟通按钮"""
        selectors = [
            'xpath:.//button[normalize-space(.)="打招呼"]',
            'xpath:.//button[normalize-space(.)="立即沟通"]',
            'xpath:.//span[normalize-space(.)="打招呼"]/ancestor::*[self::button or @role="button"][1]',
            'xpath:.//span[normalize-space(.)="立即沟通"]/ancestor::*[self::button or @role="button"][1]',
        ]
        # 只识别首次沟通文案，不点击继续沟通或已沟通
        for selector in selectors:
            button = card.ele(selector, timeout=0.2)
            if button:
                return button
        return None

    def recommendCardKey(self, card, candidateName, jobName):
        """生成推荐卡片稳定标识供数据库去重"""
        for attrName in ('data-geek-id', 'data-geekid', 'data-id', 'key', 'ka'):
            value = str(card.attr(attrName) or '').strip()
            if value:
                return f'recommend:{jobName}:{value}'
        cardText = re.sub(r'\s+', '', str(card.text or ''))[:100]
        # 无页面 ID 时使用岗位、姓名与卡片摘要组成兜底标识
        return f'recommend:{jobName}:{candidateName}:{cardText}'

    def resetRecommendScroll(self):
        """将推荐牛人列表恢复到顶部"""
        self.page.run_js(
            """
            const box = document.querySelector('.recommend-list')
                || document.querySelector('.geek-list')
                || document.scrollingElement;
            if (box) box.scrollTop = 0;
            window.scrollTo(0, 0);
            """
        )
        self.pauseWait(1)

    def scrollRecommendDown(self):
        """向下滚动推荐牛人列表并返回是否发生位移"""
        result = self.page.run_js(
            """
            const box = document.querySelector('.recommend-list')
                || document.querySelector('.geek-list')
                || document.scrollingElement;
            if (!box) return [0, 0, false];
            const before = box.scrollTop;
            const step = Math.max(400, Math.floor((box.clientHeight || window.innerHeight) * 0.85));
            box.scrollTop = before + step;
            if (box === document.scrollingElement) window.scrollTo(0, box.scrollTop);
            return [before, box.scrollTop, box.scrollTop > before];
            """
        )
        if result and len(result) >= 3 and result[2]:
            self.pauseWait(1)
            return True
        return False

    def pickRecommendIntro(self, task, rules, candidateName):
        """为推荐牛人首次沟通选择岗位介绍话术"""
        intro = str((rules or {}).get('jobIntro') or '').strip()
        if not intro:
            intro = self.pickTemplateWord('greeting', task.get('greetingWords') or [])
        jobName = str((rules or {}).get('jobName') or '')
        # 统一替换姓名、岗位和公司等占位符
        return self.applyPlaceholders(intro, task, candidateName=candidateName, jobName=jobName) if intro else ''

    def contactRecommend(self, task, card, candidateName, rules, button):
        """联系一位推荐牛人，成功后记录独立的首次招呼动作"""
        jobName = str((rules or {}).get('jobName') or '')
        listItemKey = self.recommendCardKey(card, candidateName, jobName)
        candidateKey = BossDb.buildCandidateKey(candidateName, listItemKey)
        if self.db:
            self.db.getOrCreateCandidate(candidateName, listItemKey)
            if self.db.hasRecommendContact(candidateKey):
                return False
        # 点击卡片上的首次沟通按钮
        button.click()
        self.gapWait()
        self.ensureVerifyClear()
        # 部分页面会再弹出一次立即沟通确认
        confirm = self.page.ele(
            'xpath://div[contains(@class,"dialog") or contains(@class,"modal")]//button[normalize-space(.)="立即沟通" or normalize-space(.)="打招呼"]',
            timeout=1,
        )
        if confirm:
            confirm.click()
            self.gapWait()
            self.ensureVerifyClear()
        # 精确命中首次沟通按钮且点击无异常，按已消耗一次推荐牛人联系额度处理
        editor = self.page.ele('#boss-chat-editor-input', timeout=1)
        # 按钮仅打开聊天框但没有自动发话术时，发送岗位介绍作为首次招呼
        if editor and not self.hasSelfChatMsg():
            intro = self.pickRecommendIntro(task, rules, candidateName)
            if intro and not self.sendMsg(intro):
                self.log(f'{candidateName} 已触发首次沟通，但岗位介绍发送失败')
        if self.db:
            intro = self.pickRecommendIntro(task, rules, candidateName)
            self.db.recordAction(
                candidateKey,
                'recommend_greeting',
                word=intro,
                success=True,
                taskRunId=self.currentTaskRunId,
                extra={'jobName': jobName},
            )
            self.db.setResumeStatus(candidateKey, 'greeted')
            self.db.touchCandidate(candidateKey)
        self.log(f'已从推荐牛人主动联系 {candidateName}（{jobName or "岗位待确认"}）')
        return True

    def runRecommendRank(self, task, activeRank, dailyLimit, pageJobName, jobRules):
        """扫描一个活跃优先级并联系符合岗位的推荐牛人"""
        self.resetRecommendScroll()
        visited = set()
        idleCount = 0
        while self.db.countTodayRecommend() < dailyLimit:
            self.waitIfPaused()
            self.ensureVerifyClear()
            self.ensureWorkWindow()
            candidates = []
            cards = self.readRecommendCards()
            for card in cards:
                cardText = str(card.text or '')
                activeText = self.readRecommendActive(card)
                candidateName = self.readRecommendName(card)
                rawKey = self.recommendCardKey(card, candidateName, pageJobName)
                if rawKey in visited:
                    continue
                visited.add(rawKey)
                # 当前轮次只处理对应活跃等级，其他等级留给后续轮次
                if self.parseActiveRank(activeText) != activeRank or not candidateName:
                    continue
                rules, score = self.matchRecommendRules(cardText, pageJobName, jobRules)
                button = self.getRecommendButton(card)
                if not rules or not button:
                    continue
                candidates.append((score, candidateName, rules, card, button))
            # 同活跃等级内优先联系岗位关键词命中更多的人
            candidates.sort(key=lambda item: item[0], reverse=True)
            if candidates:
                score, candidateName, rules, card, button = candidates[0]
                if self.contactRecommend(task, card, candidateName, rules, button):
                    used = self.db.countTodayRecommend()
                    self.log(f'推荐牛人今日主动联系进度：{used}/{dailyLimit}')
                    self.randomChatWait(task)
                    # 页面跳转到聊天后重新进入推荐页继续扫描
                    if 'recommend' not in str(self.page.url or '').lower():
                        if not self.selectRecommendTab():
                            return
                    pageJobName = self.readRecommendPageJob() or pageJobName
                else:
                    self.skipChatWait(task)
                idleCount = 0
                continue
            if self.scrollRecommendDown():
                idleCount = 0
                continue
            idleCount += 1
            if idleCount >= self.recommendMaxIdle:
                return

    def runRecommend(self, task):
        """按活跃优先级执行推荐牛人主动联系"""
        if not self.db:
            self.log('数据库未初始化，跳过推荐牛人主动联系')
            return
        config = dict(task.get('recommend') or {})
        dailyLimit = max(1, int(config.get('dailyLimit') or 15))
        used = self.db.countTodayRecommend()
        if used >= dailyLimit:
            self.log(f'推荐牛人今日主动联系已达上限 {dailyLimit} 人')
            return
        if not self.selectRecommendTab():
            return
        pageJobName = self.readRecommendPageJob()
        jobRules = [row for row in self.db.getJobRulesList() if int(row.get('enabled') or 0) == 1]
        if not jobRules:
            self.log('没有启用的岗位规则，跳过推荐牛人主动联系')
            return
        if pageJobName:
            self.log(f'推荐牛人当前招聘岗位：{pageJobName}')
        else:
            self.log('未识别推荐页当前岗位，将按卡片可见关键词匹配已启用岗位')
        # 严格按刚刚/在线、今日、三天内三个等级依次扫描
        for activeRank in (0, 1, 2):
            if self.db.countTodayRecommend() >= dailyLimit:
                break
            self.runRecommendRank(task, activeRank, dailyLimit, pageJobName, jobRules)

    def checkLogin(self, timeout=5):
        """检查是否已登录"""
        nameEle = self.page.ele('xpath://span[@class="user-name"]', timeout=timeout)
        # 找到用户名元素说明已登录
        if nameEle:
            self.log(f'当前 BOSS 账号: {nameEle.text}')
            return True
        return False

    def waitForLogin(self):
        """未登录时等待用户在浏览器窗口中手动登录"""
        # 已登录则无需等待
        if self.checkLogin():
            return
        self.log('未检测到 BOSS 登录，请在已打开的浏览器窗口中登录招聘端账号')
        self.log('登录完成后程序将自动继续，无需重启任务')
        lastRemind = time.time()
        while True:
            self.waitIfPaused()
            # 轮询检测是否已登录成功
            if self.checkLogin(timeout=2):
                self.log('登录成功，继续执行任务')
                return
            now = time.time()
            # 定期提醒用户并尝试回到沟通页
            if now - lastRemind >= self.loginRemindSec:
                self.log('仍在等待登录，请在浏览器中完成 BOSS 登录...')
                lastRemind = now
                try:
                    self.selectChatTab()
                except Exception:
                    pass  # 页面未就绪时忽略，下一轮继续轮询
            self.pauseWait(self.loginPollSec)

    def pageText(self):
        """读取页面可见文本"""
        try:
            return self.page.run_js("return document.body ? document.body.innerText : ''") or ''
        except Exception:
            return ''  # 页面未就绪或 JS 执行失败时返回空串

    def isVerifyVisible(self):
        """检测页面是否出现 BOSS 安全验证（滑块/点选等）"""
        if not self.page:
            return False
        text = self.pageText()
        html = (self.page.html or '')[:80000]
        combined = f'{text}\n{html}'
        # 按关键词匹配验证文案
        for keyword in self.verifyKeywords:
            if keyword in combined:
                return True
        verifySelectors = ['xpath://iframe[contains(@src,"geetest")]', 'xpath://iframe[contains(@src,"captcha")]', 'xpath://*[contains(@class,"geetest")]', 'xpath://*[contains(@class,"nc-container")]', 'xpath://*[contains(@class,"verify-wrap")]', 'xpath://*[contains(@class,"captcha")]', 'xpath://*[contains(text(),"滑动完成验证")]', 'xpath://*[contains(text(),"请完成安全验证")]']
        # 按 DOM 选择器检测验证组件
        for selector in verifySelectors:
            if self.page.ele(selector, timeout=0.3):
                return True
        return False

    def waitForVerify(self):
        """出现安全验证时暂停，等待用户在浏览器中手动完成"""
        # 无验证则直接继续
        if not self.isVerifyVisible():
            return
        self.onVerifyDetected()
        self.verifyWaiting = True
        self.mousePauseFlag.set()  # 验证期间暂停鼠标随机移动
        self.log('检测到 BOSS 安全验证，请在浏览器中手动完成（滑块/点选等）')
        self.log('验证完成后程序将自动继续简历任务')
        lastRemind = time.time()
        while True:
            self.waitIfPaused()
            # 验证消失则恢复自动化
            if not self.isVerifyVisible():
                self.log('安全验证已通过，继续简历任务')
                self.verifyWaiting = False
                self.verifyCountedEpisode = False
                self.mousePauseFlag.clear()
                return
            now = time.time()
            # 定期提醒用户完成验证
            if now - lastRemind >= self.verifyRemindSec:
                self.log('仍在等待安全验证，请在浏览器窗口中完成验证...')
                lastRemind = now
            self.pauseWait(self.verifyPollSec)

    def ensureVerifyClear(self):
        """简历相关操作前确保页面未被安全验证阻挡"""
        self.waitForVerify()  # 有验证则阻塞至用户手动完成

    def sendMsg(self, text):
        """在聊天框发送消息"""
        # 通过 JS 写入聊天编辑器并触发 input 事件
        self.page.run_js(f"\n            const el = document.querySelector('#boss-chat-editor-input');\n            if (el) {{\n                el.focus();\n                el.innerHTML = '';\n                document.execCommand('insertText', false, {text!r});\n                el.dispatchEvent(new Event('input', {{ bubbles: true }}));\n            }}\n            ")
        self.gapWait()
        submit = self.page.ele('xpath://div[@class="submit-content"]', timeout=3)
        if not submit:
            return False
        submit.click()
        self.gapWait()
        self.ensureVerifyClear()
        return True

    def requestResume(self):
        """点击求简历"""
        btn = self.page.ele('xpath://span[@class="operate-btn" and text()="求简历"]', timeout=2)
        if not btn:
            return False
        btn.click()
        self.gapWait()
        # 确认弹窗中点击「确定」
        confirm = self.page.ele('xpath://span[contains(text(),"确定向牛人索取简历")]/..//span[contains(@class,"boss-btn-primary") and text()="确定"]', timeout=3)
        if confirm:
            confirm.click()
            self.gapWait()
            self.ensureVerifyClear()
            return True
        return False

    def listChatMessages(self):
        """读取当前聊天窗口的消息列表"""
        # 超时 2 秒查找 message-item 节点，无则返回空列表
        return self.page.eles('xpath://div[@class="message-item"]', timeout=2) or []

    def hasSelfChatMsg(self):
        """聊天里是否已有我方文字消息（含推荐牛人页手动打招呼）"""
        for msg in self.listChatMessages():
            # 跳过非我方消息
            if not msg.ele('xpath:.//div[contains(@class,"item-myself")]', timeout=0.1):
                continue
            textEle = msg.ele('xpath:.//span[@class="text-content"]', timeout=0.1)
            if textEle and str(textEle.text or '').strip():
                return True
        return False

    def hasFriendChatMsg(self):
        """聊天里是否已有对方文字消息"""
        for msg in self.listChatMessages():
            # 跳过非对方消息
            if not msg.ele('xpath:.//div[contains(@class,"item-friend")]', timeout=0.1):
                continue
            textEle = msg.ele('xpath:.//span[@class="text-content"]', timeout=0.1)
            if textEle and str(textEle.text or '').strip():
                return True
        return False

    def readChatTimeline(self):
        """读取聊天时间线，区分我方与对方文字消息"""
        timeline = []
        for msg in self.listChatMessages():
            isSelf = bool(msg.ele('xpath:.//div[contains(@class,"item-myself")]', timeout=0.1))
            isFriend = bool(msg.ele('xpath:.//div[contains(@class,"item-friend")]', timeout=0.1))
            # 忽略系统消息等非对话条目
            if not isSelf and (not isFriend):
                continue
            textEle = msg.ele('xpath:.//span[@class="text-content"]', timeout=0.1)
            text = str(textEle.text or '').strip() if textEle else ''
            if not text:
                continue
            timeLabel = self.readMessageTimeLabel(msg)
            timeline.append({
                'sender': 'self' if isSelf else 'friend',
                'text': text,
                'timeLabel': timeLabel,
                'isToday': self.isTodayTimeLabel(timeLabel),
            })
        return timeline

    def readReplyContext(self, limit=12):
        """读取最近对话并整理为本地模型可理解的角色文本"""
        timeline = self.readChatTimeline()
        rows = []
        # 仅保留最近有限条文字消息，控制本地模型上下文长度
        for item in timeline[-max(1, int(limit)):]:
            text = str(item.get('text') or '').strip()
            if not text:
                continue
            role = '招聘方' if item.get('sender') == 'self' else '候选人'
            rows.append(f'{role}：{text[:300]}')
        return '\n'.join(rows)

    def getPendingFriendTexts(self, timeline, todayOnly=False):
        """获取我方最后一条消息之后对方发来的未回复文字"""
        lastSelfIdx = -1
        # 找到我方最后一条消息的索引
        for idx, item in enumerate(timeline):
            if item['sender'] == 'self':
                lastSelfIdx = idx
        items = [item for item in timeline[lastSelfIdx + 1:] if item['sender'] == 'friend']
        # 仅当天模式：过滤非今日对方消息
        if todayOnly:
            items = [item for item in items if item.get('isToday', True)]
        return [item['text'] for item in items if str(item.get('text') or '').strip()]

    def detectReplyIntent(self, text):
        """根据对方消息关键词识别回复意图类型"""
        content = str(text or '').strip()
        if not content:
            return ''
        # 按规则表顺序匹配第一个命中的意图
        for rule in self.replyIntentRules:
            for keyword in rule['keywords']:
                if keyword in content:
                    return str(rule['intent'])
        return ''

    def pickSmartReplyWords(self, task, intent):
        """按意图从任务参数中取对应话术列表"""
        wordMap = {'reply_resume': task.get('replyResumeWords') or [], 'reply_learn': task.get('replyLearnWords') or [], 'reply_interest': task.get('replyInterestWords') or []}
        # 过滤空字符串，只返回有效话术
        return [w for w in wordMap.get(intent) or [] if w]

    def manualHandoffCfg(self, task):
        """读取人工切入聊天配置"""
        return dict(task.get('manualHandoff') or {})

    def needManualHandoff(self, task, pendingTexts, intent='', templateEmpty=False):
        """判断是否需人工自由回复"""
        if self.manualReplyRequested:
            return True, '手动切换人工回复'
        merged = '\n'.join(pendingTexts) if isinstance(pendingTexts, list) else str(pendingTexts or '')
        cfg = self.manualHandoffCfg(task)
        for keyword in cfg.get('keywords') or []:
            if keyword and keyword in merged:
                return True, f'命中关键词「{keyword}」'
        if templateEmpty and cfg.get('whenNoTemplate', True):
            return True, '无可用模板话术'
        if (not intent) and merged and cfg.get('whenUnknownIntent', True):
            return True, '无法识别回复意图'
        return False, ''

    def signalManualReplyDone(self, decision='done'):
        """GUI 提交人工回复结果或候选人处理决定"""
        self.manualReplyDecision = str(decision or 'done')
        # 唤醒等待线程，由自动化线程复核聊天时间线
        self.manualReplyDoneEvent.set()

    def requestManualReply(self):
        """GUI 请求对当前候选人切换人工回复"""
        if not self.currentCandidateKey:
            return False
        if self.formalInterviewWaitActive or self.manualReplyWaitActive:
            return False
        self.manualReplyRequested = True
        return True

    def hasSelfReplyAfterFriend(self):
        """判断我方是否已回复对方最新一条文字消息"""
        timeline = self.readChatTimeline()
        lastSelfIdx = -1
        lastFriendIdx = -1
        # 分别记录双方最后一条文字消息位置
        for idx, item in enumerate(timeline):
            if item.get('sender') == 'self':
                lastSelfIdx = idx
            elif item.get('sender') == 'friend':
                lastFriendIdx = idx
        # 必须存在对方消息，且我方最后消息位于其后
        return lastFriendIdx >= 0 and lastSelfIdx > lastFriendIdx

    def waitForManualReply(self, task, candidateKey, candidateName, jobName, reason, friendText='', mode='reply_only'):
        """阻塞等待 HR 完成人工回复或作出候选人决定"""
        if self.formalInterviewWaitActive:
            self.log('当前正在等待正式面试邀约，无法切换人工回复')
            return ''
        info = {
            'candidateKey': candidateKey,
            'candidateName': candidateName,
            'jobName': jobName or self.currentJobName or '',
            'reason': reason,
            'friendText': str(friendText or '')[:300],
            'conversationText': self.readReplyContext() if friendText else '',
            'mode': mode,
        }
        self.manualReplyDoneEvent.clear()
        self.manualReplyDecision = ''
        self.manualReplyMode = mode
        self.manualReplyWaitActive = True
        self.mousePauseFlag.set()
        if mode == 'candidate_reply':
            self.log('【需人工】请回复候选人，再选择继续索要简历或标记不合适')
        else:
            self.log('【需人工】请在 BOSS 聊天框自由回复并发送')
        self.log(f"  候选人: {info['candidateName']} | 岗位: {info['jobName']}")
        self.log(f"  原因: {reason}")
        if friendText:
            self.log(f"  对方消息: {str(friendText)[:120]}")
        if mode == 'candidate_reply':
            self.log('回复后点击「合适，继续索要简历」；不合适则点击「标记不合适」')
        else:
            self.log('完成后请点击界面「已完成人工回复」按钮')
        if self.onManualReplyWait:
            self.onManualReplyWait(info)
        lastRemind = time.time()
        decision = ''
        try:
            while True:
                self.waitIfPaused()
                if self.manualReplyDoneEvent.is_set():
                    choice = str(self.manualReplyDecision or 'done')
                    # 继续索要简历前，必须确认聊天中已有我方最新回复
                    if mode == 'candidate_reply' and choice == 'continue':
                        if not self.hasSelfReplyAfterFriend():
                            self.log('尚未检测到人工回复，请先在 BOSS 发送回复后再点击继续')
                            self.manualReplyDoneEvent.clear()
                            self.manualReplyDecision = ''
                            continue
                    decision = choice
                    break
                now = time.time()
                if now - lastRemind >= self.manualReplyRemindSec:
                    self.log('仍在等待人工聊天回复...')
                    lastRemind = now
                self.pauseWait(1.0)
        finally:
            self.manualReplyWaitActive = False
            self.manualReplyRequested = False
            self.manualReplyMode = ''
            self.mousePauseFlag.clear()
            if self.onManualReplyWaitDone:
                self.onManualReplyWaitDone()
        if decision:
            self.log(f'已收到 {candidateName} 的人工处理结果：{decision}')
        return decision

    def ensureManualReply(self, task, candidateKey, candidateName, jobName, pendingTexts='', intent='', templateEmpty=False):
        """需人工则等待 HR 回复，返回是否已执行人工等待"""
        need, reason = self.needManualHandoff(task, pendingTexts, intent, templateEmpty)
        if not need:
            return False
        friendText = '\n'.join(pendingTexts) if isinstance(pendingTexts, list) else str(pendingTexts or '')
        decision = self.waitForManualReply(task, candidateKey, candidateName, jobName, reason, friendText)
        if self.db:
            self.db.recordAction(
                candidateKey,
                'manual_reply',
                word=reason,
                success=bool(decision),
                taskRunId=self.currentTaskRunId,
                extra={'reason': reason, 'friendText': friendText[:200]},
            )
        return True

    def ensureManualResumeReview(self, task, candidateKey, candidateName, jobName, reason):
        """简历解析失败或中断恢复时，弹窗等待 HR 人工确认"""
        if self.db and self.db.hasManualResumeReviewToday(candidateKey):
            self.log(f'{candidateName} 今日已人工确认简历处理，跳过重复弹窗')
            return True
        self.waitForManualReply(task, candidateKey, candidateName, jobName, reason, '')
        if self.db:
            self.db.recordAction(
                candidateKey,
                'manual_resume_review',
                word=reason,
                success=True,
                taskRunId=self.currentTaskRunId,
                extra={'reason': reason},
            )
        return True

    def noteResumeReceived(self, candidateKey, source, partialProfile=None):
        """记录已收到简历（落库 + 动作痕迹）"""
        if not self.db:
            return
        self.db.markResumeReceived(candidateKey, source=source, partialProfile=partialProfile)
        if not self.db.hasSentToday(candidateKey, 'resume_received'):
            self.db.recordAction(
                candidateKey,
                'resume_received',
                word=source,
                success=True,
                taskRunId=self.currentTaskRunId,
            )

    def handleResumeParseFailure(self, task, candidateKey, candidateName, jobName, reason, source):
        """简历解析失败：落库、记录失败原因并切换人工确认"""
        if self.db:
            self.db.markResumeReceived(candidateKey, source=source)
            actionType = 'preview_resume_failed' if source == 'attach_preview' else 'resume_parse_failed'
            if not self.db.hasSentToday(candidateKey, actionType):
                self.db.recordAction(
                    candidateKey,
                    actionType,
                    word=reason,
                    success=False,
                    taskRunId=self.currentTaskRunId,
                )
        self.log(f'【需人工】{candidateName}: {reason}')
        self.ensureManualResumeReview(task, candidateKey, candidateName, jobName, reason)
        if self.db:
            self.db.touchCandidate(candidateKey)
        self.randomChatWait(task)
        return True

    def sendSmartReply(self, template, jobName, candidateKey=None, task=None):
        """发送智能回复话术，替换全部占位符"""
        msg = self.applyPlaceholders(template, task, jobName=jobName)
        return self.sendOrHandoff(msg, 'smart_reply', candidateKey)

    def sendGreetingIntro(self, template, jobName, candidateKey=None, task=None):
        """发送岗位介绍话术，替换全部占位符"""
        msg = self.applyPlaceholders(template, task, jobName=jobName)
        return self.sendOrHandoff(msg, 'greeting', candidateKey)

    def trySendGreetingIntro(self, task, greetingWords, candidateKey, candidateName, jobName, pendingTexts=None):
        """我方已手动打招呼且对方已回复时，发送岗位介绍"""
        pendingTexts = pendingTexts or []
        # 今日已发过或已填框则跳过
        if self.db and (self.db.hasSentToday(candidateKey, 'greeting') or self.db.hasHandoffToday(candidateKey, 'greeting')):
            return False
        if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts):
            return True
        introText = ''
        # 优先使用数据库中该岗位的 jobIntro 规则
        if self.db:
            jobRules = self.db.matchJobRules(jobName)
            if jobRules and str(jobRules.get('jobIntro') or '').strip():
                introText = str(jobRules['jobIntro'])
        # 无岗位规则则从 greeting 话术列表随机选一条
        if not introText:
            enabled = [w for w in greetingWords if w]
            if not enabled:
                self.log(f'{candidateName} 未配置岗位介绍话术（greeting / job_intro）')
                if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts, templateEmpty=True):
                    return True
                return False
            introText = self.pickTemplateWord('greeting', enabled)
        if not introText:
            if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts, templateEmpty=True):
                return True
            return False
        result = self.sendGreetingIntro(introText, jobName, candidateKey, task)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return False
        if result == 'sent':
            self.log(f'已向 {candidateName} 发送岗位介绍')
            if self.db:
                self.db.recordAction(candidateKey, 'greeting', word=introText, success=True, taskRunId=self.currentTaskRunId)
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入岗位介绍（待人工发送）')
        return True

    def trySendSmartReply(self, task, candidateKey, candidateName, jobName):
        """求职者先开口时，按意图智能回复，返回 (是否已发送, 意图)"""
        # 今日已发过、已填框或岗位介绍则跳过
        if self.db and self.db.hasSentToday(candidateKey, 'smart_reply'):
            return (False, '')
        if self.db and self.db.hasHandoffToday(candidateKey, 'smart_reply'):
            return (False, '')
        if self.db and (self.db.hasSentToday(candidateKey, 'greeting') or self.db.hasHandoffToday(candidateKey, 'greeting')):
            return (False, '')
        # 我方已有消息说明不是「对方先开口」场景
        if self.hasSelfChatMsg():
            return (False, '')
        timeline = self.readChatTimeline()
        todayOnly = (task.get('rateLimits') or {}).get('todayOnly', True)
        pendingTexts = self.getPendingFriendTexts(timeline, todayOnly=todayOnly)
        # 无待回复的对方消息则跳过
        if not pendingTexts:
            return (False, '')
        mergedText = '\n'.join(pendingTexts)
        intent = self.detectReplyIntent(mergedText)
        if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts, intent=intent):
            return (True, 'manual')
        replyText = ''
        # 「了解岗位」意图优先用 jobIntro 规则
        if intent == 'reply_learn' and self.db:
            jobRules = self.db.matchJobRules(jobName)
            if jobRules and str(jobRules.get('jobIntro') or '').strip():
                replyText = str(jobRules['jobIntro'])
        # 否则从对应意图话术列表随机选取
        if not replyText:
            words = self.pickSmartReplyWords(task, intent) if intent else []
            if not words:
                if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts, intent=intent, templateEmpty=True):
                    return (True, 'manual')
                return (False, '')
            replyText = self.pickTemplateWord(intent, words)
            if not replyText:
                if self.ensureManualReply(task, candidateKey, candidateName, jobName, pendingTexts, intent=intent, templateEmpty=True):
                    return (True, 'manual')
                return (False, '')
        result = self.sendSmartReply(replyText, jobName, candidateKey, task)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return (False, '')
        intentLabel = {'reply_resume': '主动发简历', 'reply_learn': '了解岗位', 'reply_interest': '表达兴趣'}.get(intent, intent)
        if result == 'sent':
            self.log(f'已向 {candidateName} 智能回复（{intentLabel}）')
            if self.db:
                self.db.recordAction(candidateKey, 'smart_reply', word=replyText, success=True, taskRunId=self.currentTaskRunId, extra={'intent': intent, 'friendText': mergedText[:200]})
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入智能回复（{intentLabel}，待人工发送）')
        return (True, intent)

    def msgLooksLikeResume(self, msg):
        """判断单条聊天消息是否含已接收的简历卡片"""
        # 待点击「同意」的卡片不算已接收
        if msg.ele('xpath:.//span[@class="card-btn" and text()="同意"]', timeout=0.1):
            return False
        cardEle = msg.ele('xpath:.//div[contains(@class,"message-card") or contains(@class,"chat-card")]', timeout=0.1)
        if cardEle and '简历' in str(cardEle.text or ''):
            return True
        resumeHints = ['附件简历', '对方已发送', '向您发送', '收到一份简历']
        for hint in resumeHints:
            hintEle = msg.ele(f'xpath:.//*[contains(text(),"{hint}")]', timeout=0.1)
            if hintEle:
                return True
        # PDF 附件简历（点击预览入口）
        if msg.ele('xpath:.//*[contains(text(),"点击预览附件简历")]', timeout=0.1):
            return True
        if msg.ele('xpath:.//*[contains(text(),".pdf")]', timeout=0.1):
            return True
        msgText = str(msg.text or '')
        if '.pdf' in msgText.lower() and '简历' in msgText:
            return True
        return False

    def detectResumeCard(self):
        """检测聊天区内简历卡片: none / pending_accept / accepted"""
        msgList = self.listChatMessages()
        # 优先检测待同意卡片
        for msg in msgList:
            if msg.ele('xpath:.//span[@class="card-btn" and text()="同意"]', timeout=0.1):
                return 'pending_accept'
        # 再检测已接收的简历卡片
        for msg in msgList:
            if self.msgLooksLikeResume(msg):
                return 'accepted'
        return 'none'

    def shouldReviewResume(self, status, resumeCard, attachReady=False):
        """聊天区确有简历且流程已到可审核阶段"""
        if resumeCard != 'accepted' and not attachReady:
            return False
        # 已进入面试流程或已拒绝则不再审核
        if status in ('interview_sent', 'interview_pre_sent', 'interview_awaiting_time', 'interview_reschedule_pending', 'interview_formal_pending', 'interview_cancelled', 'resume_rejected'):
            return False
        return status in ('resume_received', 'resume_requested', 'resume_passed')

    def inInterviewFlow(self, status):
        """是否处于面试预邀请/跟进流程"""
        return status in ('interview_pre_sent', 'interview_awaiting_time', 'interview_reschedule_pending', 'interview_formal_pending')

    def alreadySentInterview(self, candidateKey):
        """是否已完成正式面试邀约"""
        if not self.db:
            return False
        if self.db.getResumeStatus(candidateKey) == 'interview_sent':
            return True
        candidate = self.db.getCandidate(candidateKey)
        if candidate and candidate.get('interview_sent_at'):
            return True
        return False

    def alreadySentInterviewPre(self, candidateKey):
        """是否已发过面试预邀请"""
        if not self.db:
            return False
        status = self.db.getResumeStatus(candidateKey)
        if status in ('interview_pre_sent', 'interview_awaiting_time', 'interview_reschedule_pending', 'interview_formal_pending', 'interview_sent'):
            return True
        return self.db.hasSentToday(candidateKey, 'interview_pre') or self.db.hasHandoffToday(candidateKey, 'interview_pre')

    def interviewConfig(self, task):
        """读取任务中的面试配置"""
        return dict(task.get('interviewConfig') or {})

    def placeholderConfig(self, task):
        """读取任务中的话术占位符配置"""
        return dict(task.get('placeholderConfig') or {})

    def resolveJobText(self, task, jobName):
        """解析 {job}：优先候选人岗位，否则用配置的默认岗位名"""
        cfg = self.placeholderConfig(task)
        fallback = str(cfg.get('jobDefault') or '该岗位').strip() or '该岗位'
        return str(jobName or '').strip() or fallback

    def resolveNameText(self, task, candidateName):
        """解析 {name}：优先候选人姓名，否则用配置的默认称呼"""
        cfg = self.placeholderConfig(task)
        fallback = str(cfg.get('nameDefault') or '候选人').strip() or '候选人'
        return str(candidateName or '').strip() or fallback

    def resolveCompanyText(self, task):
        """解析 {company} 公司名称"""
        cfg = self.placeholderConfig(task)
        return str(cfg.get('company') or '我们公司').strip() or '我们公司'

    def applyPlaceholders(self, template, task, candidateName='', jobName='', dateText='', timeText='', address='', duration=''):
        """统一替换话术中的全部占位符"""
        cfg = self.placeholderConfig(task)
        icfg = self.interviewConfig(task)
        nameText = self.resolveNameText(task, candidateName)
        jobText = self.resolveJobText(task, jobName)
        companyText = self.resolveCompanyText(task)
        addr = str(address or cfg.get('address') or icfg.get('address') or '待定').strip() or '待定'
        dur = str(duration or cfg.get('duration') or icfg.get('duration') or '40-60').strip() or '40-60'
        dateVal = str(dateText or '').strip()
        if not dateVal and task:
            dayOffset = int(icfg.get('dayOffset') if icfg.get('dayOffset') is not None else cfg.get('dayOffset') or 1)
            dateVal = self.formatInterviewDate(dayOffset)
        timeVal = str(timeText or '').strip()
        return (
            str(template or '')
            .replace('{name}', nameText)
            .replace('{date}', dateVal)
            .replace('{time}', timeVal)
            .replace('{address}', addr)
            .replace('{duration}', dur)
            .replace('{job}', jobText)
            .replace('{company}', companyText)
            .replace('【XXX】', nameText)
        )

    def formatInterviewDate(self, dayOffset=1):
        """按运行日生成面试日期文案，如 6月30日（周一）"""
        target = dt.date.today() + dt.timedelta(days=int(dayOffset))
        weekNames = '一二三四五六日'
        return f'{target.month}月{target.day}日（周{weekNames[target.weekday()]}）'

    def formatSlotHour(self, hour):
        """将整点小时格式化为上午/下午几点"""
        hour = int(hour)
        if hour < 12:
            return f'上午{hour}点'
        if hour == 12:
            return '中午12点'
        if hour == 13:
            return '下午1点'
        return f'下午{hour - 12}点'

    def pickInterviewSlot(self, task):
        """从配置时段池随机选取面试整点"""
        cfg = self.interviewConfig(task)
        slots = [int(x) for x in (cfg.get('timeSlots') or [10, 11, 13, 14, 15, 16, 17])]
        if not slots:
            slots = [14]
        spread = bool(cfg.get('timeSpread', True))
        if spread and self.db:
            used = self.db.getUsedInterviewSlots()
            free = [h for h in slots if h not in used]
            if free:
                slots = free
        hour = random.choice(slots)
        if spread and self.db:
            self.db.markInterviewSlotUsed(hour)
        return hour

    def buildInterviewMsg(self, template, task, candidateName, dateText, timeText, address, duration, jobName=''):
        """替换面试话术占位符"""
        return self.applyPlaceholders(
            template,
            task,
            candidateName=candidateName,
            jobName=jobName,
            dateText=dateText,
            timeText=timeText,
            address=address,
            duration=duration,
        )

    def extractTimeFromText(self, text):
        """从对方消息中尝试提取时间描述"""
        content = str(text or '').strip()
        if not content:
            return ''
        patterns = [
            r'明天(?:上午|早上|下午|晚上)?\d{1,2}[:：点时]',
            r'后天(?:上午|早上|下午|晚上)?\d{1,2}[:：点时]',
            r'周[一二三四五六日天](?:上午|早上|下午|晚上)?\d{1,2}[:：点时]',
            r'(?:上午|早上|下午|晚上)\d{1,2}[:：点时]',
            r'\d{1,2}[:：]\d{2}',
            r'\d{1,2}点\d{0,2}分?',
        ]
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                return match.group(0)
        return ''

    def hasInterviewDecline(self, text):
        """对方是否表示时间不合适"""
        content = str(text or '')
        return any(word in content for word in self.interviewDeclineKeywords)

    def hasInterviewConfirm(self, text):
        """对方是否确认可以参加"""
        content = str(text or '').lower()
        if self.hasInterviewDecline(text):
            return False
        return any(word.lower() in content for word in self.interviewConfirmKeywords)

    def needInterviewNoReplyRemind(self, candidateKey, task):
        """预邀请超过配置小时仍无回复时是否需要追问"""
        if not self.db:
            return False
        meta = self.db.getInterviewMeta(candidateKey)
        sentAt = str(meta.get('preInviteSentAt') or '').strip()
        if not sentAt:
            return False
        if self.db.hasSentToday(candidateKey, 'interview_remind') or self.db.hasHandoffToday(candidateKey, 'interview_remind'):
            return False
        hours = float(self.interviewConfig(task).get('noReplyHours') or 1)
        try:
            sentDt = dt.datetime.fromisoformat(sentAt)
        except Exception:
            return False
        return (dt.datetime.now() - sentDt).total_seconds() >= hours * 3600

    def needInterviewAutoCancelAfterRemind(self, candidateKey, task, timeline=None, todayOnly=True):
        """预邀请且已发提醒后仍无对方回复时是否应自动停跟取消"""
        if not bool(self.interviewConfig(task).get('cancelAfterRemind', True)):
            return False
        if not self.db:
            return False
        if self.db.getResumeStatus(candidateKey) != 'interview_pre_sent':
            return False
        if not self.db.hasInterviewRemindAfterPreInvite(candidateKey):
            return False
        if timeline is not None:
            pending = self.getPendingFriendTexts(timeline, todayOnly=todayOnly)
            if '\n'.join(pending).strip():
                return False
        return True

    def shouldProcessForRemindCancel(self, candidateKey, task):
        """列表层：是否需进入流程以执行提醒后无回复自动取消"""
        if not self.db:
            return False
        if self.db.getResumeStatus(candidateKey) != 'interview_pre_sent':
            return False
        if not self.db.hasInterviewRemindAfterPreInvite(candidateKey):
            return False
        return bool(self.interviewConfig(task).get('cancelAfterRemind', True))

    def sendInterviewChat(self, text, actionType, candidateKey):
        """发送面试相关聊天话术"""
        return self.sendOrHandoff(text, actionType, candidateKey)

    def trySendInterviewPre(self, task, candidateKey, candidateName, jobName, preWords):
        """初筛通过后发送面试预邀请"""
        if self.alreadySentInterviewPre(candidateKey):
            return False
        template = self.pickTemplateWord('interview_pre', preWords)
        if not template:
            self.log(f'{candidateName} 未配置面试预邀请话术')
            return False
        cfg = self.interviewConfig(task)
        dayOffset = int(cfg.get('dayOffset') or 1)
        dateText = self.formatInterviewDate(dayOffset)
        hour = self.pickInterviewSlot(task)
        timeText = self.formatSlotHour(hour)
        address = str(cfg.get('address') or '待定')
        duration = str(cfg.get('duration') or '40-60')
        msg = self.buildInterviewMsg(template, task, candidateName, dateText, timeText, address, duration, jobName)
        result = self.sendInterviewChat(msg, 'interview_pre', candidateKey)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return False
        if result == 'sent':
            self.log(f'已向 {candidateName} 发送面试预邀请（{dateText} {timeText}）')
            if self.db:
                self.db.recordAction(candidateKey, 'interview_pre', word=msg, success=True, taskRunId=self.currentTaskRunId)
                self.db.saveInterviewPreSent(candidateKey, dateText, timeText, address, jobName)
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入面试预邀请（待人工发送）')
            if self.db:
                self.db.saveInterviewPreSent(candidateKey, dateText, timeText, address, jobName)
        return True

    def signalFormalInterviewDone(self):
        """GUI 确认人工已完成 BOSS 正式面试邀约"""
        self.formalInterviewDoneEvent.set()

    def waitForFormalInterview(self, task, candidateKey, candidateName, jobName):
        """阻塞等待人工在 BOSS 发送正式面试邀约"""
        if self.manualReplyWaitActive:
            self.log('当前正在等待人工聊天回复，无法进入正式面试等待')
            return
        meta = self.db.getInterviewMeta(candidateKey) if self.db else {}
        info = {
            'candidateKey': candidateKey,
            'candidateName': candidateName,
            'jobName': jobName or meta.get('jobName') or '',
            'agreedDate': meta.get('agreedDate') or '',
            'agreedTime': meta.get('agreedTime') or '',
            'address': meta.get('address') or '',
        }
        self.formalInterviewDoneEvent.clear()
        self.formalInterviewWaitActive = True
        self.mousePauseFlag.set()
        self.log('【需人工】请在 BOSS 聊天页点击「发送面试」，填写以下信息：')
        self.log(f"  候选人: {info['candidateName']} | 岗位: {info['jobName']}")
        self.log(f"  时间: {info['agreedDate']} {info['agreedTime']} | 地址: {info['address']}")
        self.log('完成后请点击界面「已完成正式面试发送」按钮')
        if self.db:
            self.db.setInterviewFormalPending(candidateKey)
        if self.onFormalInterviewWait:
            self.onFormalInterviewWait(info)
        lastRemind = time.time()
        done = False
        try:
            while True:
                self.waitIfPaused()
                if self.formalInterviewDoneEvent.is_set():
                    done = True
                    break
                now = time.time()
                if now - lastRemind >= self.formalInterviewRemindSec:
                    self.log('仍在等待人工发送 BOSS 正式面试邀约...')
                    lastRemind = now
                self.pauseWait(1.0)
        finally:
            self.formalInterviewWaitActive = False
            self.mousePauseFlag.clear()
            if self.onFormalInterviewWaitDone:
                self.onFormalInterviewWaitDone()
        if done and self.db:
            self.db.markInterviewSent(candidateKey)
            self.log(f'已确认 {candidateName} 正式面试邀约完成，继续下一位候选人')

    def trySendInterviewAskTime(self, task, candidateKey, candidateName, askWords):
        """追问对方方便的面试时间"""
        if self.db and (self.db.hasSentToday(candidateKey, 'interview_ask_time') or self.db.hasHandoffToday(candidateKey, 'interview_ask_time')):
            return False
        template = self.pickTemplateWord('interview_ask_time', askWords)
        if not template:
            return False
        msg = self.applyPlaceholders(template, task, candidateName=candidateName, jobName=self.currentJobName)
        result = self.sendInterviewChat(msg, 'interview_ask_time', candidateKey)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return False
        if result == 'sent':
            self.log(f'已向 {candidateName} 追问方便面试时间')
            if self.db:
                self.db.recordAction(candidateKey, 'interview_ask_time', word=msg, success=True, taskRunId=self.currentTaskRunId)
                self.db.setInterviewAwaitingTime(candidateKey)
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入追问时间话术（待人工发送）')
            if self.db:
                self.db.setInterviewAwaitingTime(candidateKey)
        return True

    def trySendInterviewRemind(self, task, candidateKey, candidateName, remindWords):
        """预邀请无回复超过时限后发追问提醒"""
        if self.db and (self.db.hasSentToday(candidateKey, 'interview_remind') or self.db.hasHandoffToday(candidateKey, 'interview_remind')):
            return False
        template = self.pickTemplateWord('interview_remind', remindWords)
        if not template:
            return False
        msg = self.applyPlaceholders(template, task, candidateName=candidateName, jobName=self.currentJobName)
        result = self.sendInterviewChat(msg, 'interview_remind', candidateKey)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return False
        if result == 'sent':
            self.log(f'已向 {candidateName} 发送面试时间跟进提醒')
            if self.db:
                self.db.recordAction(candidateKey, 'interview_remind', word=msg, success=True, taskRunId=self.currentTaskRunId)
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入面试跟进提醒（待人工发送）')
        return True

    def trySendInterviewReschedule(self, task, candidateKey, candidateName, jobName, rescheduleWords, dateText, timeText):
        """发送改期确认话术"""
        if self.db and (self.db.hasSentToday(candidateKey, 'interview_reschedule') or self.db.hasHandoffToday(candidateKey, 'interview_reschedule')):
            return False
        template = self.pickTemplateWord('interview_reschedule', rescheduleWords)
        if not template:
            return False
        cfg = self.interviewConfig(task)
        address = str(cfg.get('address') or '待定')
        duration = str(cfg.get('duration') or '40-60')
        msg = self.buildInterviewMsg(template, task, candidateName, dateText, timeText, address, duration, jobName)
        result = self.sendInterviewChat(msg, 'interview_reschedule', candidateKey)
        if result not in ('sent', 'handoff', 'handoff_done'):
            return False
        if result == 'sent':
            self.log(f'已向 {candidateName} 发送改期确认（{dateText} {timeText}）')
            if self.db:
                self.db.recordAction(candidateKey, 'interview_reschedule', word=msg, success=True, taskRunId=self.currentTaskRunId)
                self.db.updateInterviewAgreedTime(candidateKey, dateText, timeText, 'interview_reschedule_pending')
        elif result == 'handoff':
            self.log(f'已向 {candidateName} 填入改期确认（待人工发送）')
            if self.db:
                self.db.updateInterviewAgreedTime(candidateKey, dateText, timeText, 'interview_reschedule_pending')
        return True

    def trySendInterviewCancel(self, task, candidateKey, candidateName, cancelWords, reason=''):
        """发送取消面试话术并标记取消"""
        if self.db and (self.db.hasSentToday(candidateKey, 'interview_cancel') or self.db.hasHandoffToday(candidateKey, 'interview_cancel')):
            if self.db:
                self.db.markInterviewCancelled(candidateKey, reason)
            return True
        template = self.pickTemplateWord('interview_cancel', cancelWords)
        if template:
            msg = self.applyPlaceholders(template, task, candidateName=candidateName, jobName=self.currentJobName)
            result = self.sendInterviewChat(msg, 'interview_cancel', candidateKey)
            if result == 'sent' and self.db:
                self.db.recordAction(candidateKey, 'interview_cancel', word=msg, success=True, taskRunId=self.currentTaskRunId)
            elif result == 'handoff':
                self.log(f'已向 {candidateName} 填入取消面试话术（待人工发送）')
        if self.db:
            self.db.markInterviewCancelled(candidateKey, reason)
        self.log(f'{candidateName} 面试已默认取消: {reason or "未确认时间"}')
        return True

    def tryHandleInterviewFlow(self, task, candidateKey, candidateName, jobName, status, timeline):
        """处理面试预邀请后的对方回复与跟进"""
        preWords = task.get('interviewPreWords') or []
        askWords = task.get('interviewAskWords') or []
        remindWords = task.get('interviewRemindWords') or []
        rescheduleWords = task.get('interviewRescheduleWords') or []
        cancelWords = task.get('interviewCancelWords') or []
        todayOnly = (task.get('rateLimits') or {}).get('todayOnly', True)
        # 等待人工发正式邀约（任务中断后恢复）
        if status == 'interview_formal_pending':
            self.waitForFormalInterview(task, candidateKey, candidateName, jobName)
            return
        pendingTexts = self.getPendingFriendTexts(timeline, todayOnly=todayOnly)
        mergedText = '\n'.join(pendingTexts).strip()
        # 已发预邀请 + 1 次提醒后仍无回复 → 自动停跟取消
        if not mergedText and self.needInterviewAutoCancelAfterRemind(candidateKey, task, timeline, todayOnly):
            self.trySendInterviewCancel(task, candidateKey, candidateName, cancelWords, '预邀请及提醒后仍未回复')
            if self.db:
                self.db.touchCandidate(candidateKey)
            return
        # 策略 B：预邀请超时无回复则追问
        if not mergedText and self.needInterviewNoReplyRemind(candidateKey, task):
            self.trySendInterviewRemind(task, candidateKey, candidateName, remindWords)
            if self.db:
                self.db.touchCandidate(candidateKey)
            return
        if not mergedText:
            return
        proposedTime = self.extractTimeFromText(mergedText)
        declined = self.hasInterviewDecline(mergedText)
        confirmed = self.hasInterviewConfirm(mergedText)
        meta = self.db.getInterviewMeta(candidateKey) if self.db else {}
        cfg = self.interviewConfig(task)
        dayOffset = int(cfg.get('dayOffset') or 1)
        defaultDate = meta.get('agreedDate') or self.formatInterviewDate(dayOffset)
        # 对方给出新时间：发改期确认
        if proposedTime:
            dateText = defaultDate
            if '明天' in mergedText:
                dateText = self.formatInterviewDate(1)
            elif '后天' in mergedText:
                dateText = self.formatInterviewDate(2)
            timeText = proposedTime
            self.trySendInterviewReschedule(task, candidateKey, candidateName, jobName, rescheduleWords, dateText, timeText)
            if self.db:
                self.db.touchCandidate(candidateKey)
            return
        # 对方确认可参加：进入人工正式邀约等待
        if confirmed:
            self.log(f'{candidateName} 已确认面试时间，等待人工发送 BOSS 正式邀约')
            self.waitForFormalInterview(task, candidateKey, candidateName, jobName)
            return
        # 对方表示不方便且无具体时间
        if declined:
            if status == 'interview_awaiting_time':
                self.trySendInterviewCancel(task, candidateKey, candidateName, cancelWords, '已追问但仍未给出具体时间')
            else:
                self.trySendInterviewAskTime(task, candidateKey, candidateName, askWords)
            if self.db:
                self.db.touchCandidate(candidateKey)
            return

    def acceptResume(self):
        """同意接收对方发来的简历"""
        btn = self.page.ele('xpath://span[@class="card-btn" and text()="同意"]', timeout=2)
        if not btn:
            return False
        btn.click()  # 点击简历卡片上的「同意」
        self.pauseWait(self.minActionGap)
        return True

    def findAttachResumeMsg(self):
        """查找聊天区附件简历消息"""
        # 从最新消息往前找附件预览入口
        for msg in reversed(self.listChatMessages()):
            if msg.ele('xpath:.//*[contains(text(),"点击预览附件简历")]', timeout=0.1):
                return msg
            if msg.ele('xpath:.//*[contains(text(),".pdf")]', timeout=0.1):
                return msg
        return None

    def hasAttachResumeInChat(self):
        """聊天区是否出现附件简历预览入口"""
        return self.findAttachResumeMsg() is not None  # 能找到附件消息即视为有入口

    def extractAttachFileName(self, msg):
        """从附件消息提取原始文件名"""
        text = str(msg.text or '')
        nameMatch = re.search('[\\w\\u4e00-\\u9fff\\-（）()]+\\.pdf', text, re.IGNORECASE)
        if nameMatch:
            return nameMatch.group(0).strip()  # 正则匹配消息中的 .pdf 文件名
        return ''

    def safeAttachFileName(self, fileName, candidateName):
        """生成安全的附件保存文件名"""
        raw = str(fileName or '').strip() or f'{candidateName}简历.pdf'
        safe = re.sub('[<>:"/\\\\|?*]', '_', raw)  # 替换 Windows 非法文件名字符
        if not safe.lower().endswith('.pdf'):
            safe = f'{safe}.pdf'
        return safe

    def buildRequestSession(self):
        """用当前浏览器 Cookie 构建下载会话"""
        session = requests.Session()
        cookieRows = self.page.cookies() if self.page else []
        # 同步浏览器 Cookie 到 requests 会话
        for row in cookieRows:
            session.cookies.set(row.get('name', ''), row.get('value', ''), domain=row.get('domain') or '.zhipin.com')
        userAgent = ''
        with contextlib.suppress(Exception):
            userAgent = str(self.page.run_js('return navigator.userAgent') or '')
        session.headers.update({'User-Agent': userAgent or 'Mozilla/5.0', 'Referer': self.page.url or 'https://www.zhipin.com/web/chat/index'})
        return session

    def isValidPdfFile(self, filePath):
        """判断文件是否为有效 PDF"""
        if not filePath.exists() or filePath.stat().st_size < 128:
            return False
        with open(filePath, 'rb') as fp:
            return fp.read(4) == b'%PDF'  # 校验 PDF 魔数头

    def extractAttachUrlFromMsg(self, msg):
        """从聊天附件消息提取下载地址"""
        if not msg:
            return ''
        # 在消息 DOM 内查找 href/src/data-url 等属性
        url = self.page.run_js("\n            const msg = arguments[0];\n            if (!msg) return '';\n            const nodes = msg.querySelectorAll('[href],[src],[data-url],[data-src],[data-href]');\n            for (const node of nodes) {\n                const u = node.href || node.src || node.getAttribute('data-url')\n                    || node.getAttribute('data-src') || node.getAttribute('data-href') || '';\n                if (u && /pdf|attach|resume|geek|wapi|download|file/i.test(u)) {\n                    return u;\n                }\n            }\n            return '';\n            ", msg)
        return str(url or '').strip()

    def collectPdfUrlsFromPage(self):
        """收集页面里出现的 PDF/附件 URL"""
        # 从 iframe/embed/a 及 performance 资源记录中收集候选 URL
        urls = self.page.run_js("\n            const found = new Set();\n            const add = (u) => {\n                if (!u) return;\n                const text = String(u);\n                if (/pdf|attach|resume|geek|wapi|download|file/i.test(text)) {\n                    found.add(text);\n                }\n            };\n            document.querySelectorAll('iframe,embed,a,source').forEach((el) => {\n                add(el.src || el.href);\n            });\n            performance.getEntriesByType('resource').forEach((item) => add(item.name));\n            return Array.from(found);\n            ")
        if not urls:
            return []
        return [str(u).strip() for u in urls if str(u).strip()]

    def pickAttachPdfUrl(self, urls, fileName):
        """从候选 URL 中挑选最像附件 PDF 的地址"""
        if not urls:
            return ''
        baseName = fileName.lower().replace('.pdf', '')
        scored = []
        # 按特征给每个 URL 打分，取得分最高者
        for raw in urls:
            text = raw.lower()
            score = 0
            if '.pdf' in text:
                score += 5
            if 'attach' in text or 'download' in text or 'file' in text:
                score += 4
            if baseName and baseName in text:
                score += 6
            if 'resume' in text or 'geek' in text or 'wapi' in text:
                score += 2
            if text.startswith('blob:'):
                score -= 3
            scored.append((score, raw))
        scored.sort(key=lambda item: item[0], reverse=True)
        bestScore, bestUrl = scored[0]
        return bestUrl if bestScore > 0 else ''

    def downloadAttachFile(self, url, savePath):
        """按 URL 下载真实附件 PDF"""
        if not url:
            return False
        fullUrl = url if url.startswith('http') else urljoin(self.page.url or 'https://www.zhipin.com/', url)
        # blob 地址无法用 HTTP 直接下载
        if fullUrl.startswith('blob:'):
            self.log('附件为 blob 地址，无法直接 HTTP 下载')
            return False
        session = self.buildRequestSession()
        try:
            res = session.get(fullUrl, timeout=60, stream=True)
            res.raise_for_status()
            savePath.parent.mkdir(parents=True, exist_ok=True)
            with open(savePath, 'wb') as fp:
                for chunk in res.iter_content(chunk_size=8192):
                    if chunk:
                        fp.write(chunk)
        except Exception as exc:
            self.log(f'附件下载请求失败: {exc}')
            # 下载失败时清理不完整文件
            with contextlib.suppress(Exception):
                if savePath.exists():
                    savePath.unlink()
            return False
        # 校验下载结果是否为有效 PDF
        if not self.isValidPdfFile(savePath):
            self.log(f'下载内容不是有效 PDF: {savePath.name}')
            with contextlib.suppress(Exception):
                savePath.unlink()
            return False
        return True

    def clickAttachResumePreview(self):
        """点击聊天区「点击预览附件简历」"""
        msg = self.findAttachResumeMsg()
        if not msg:
            return False
        previewBtn = msg.ele('xpath:.//*[contains(text(),"点击预览附件简历")]', timeout=0.2)
        # 无专用按钮则点击整条消息
        if not previewBtn:
            previewBtn = msg
        previewBtn.click()
        self.pauseWait(self.minActionGap)
        self.ensureVerifyClear()
        return True

    def findResumePreviewModal(self):
        """查找附件简历预览弹窗元素"""
        modalXpaths = ['xpath://div[contains(@class,"resume-detail")]', 'xpath://div[contains(@class,"lib-resume")]', 'xpath://div[contains(@class,"geek-resume")]', 'xpath://div[contains(@class,"boss-dialog") and .//*[contains(text(),"附件简历")]]', 'xpath://div[contains(@class,"dialog") and .//*[contains(text(),"附件简历")]]', 'xpath://div[contains(@class,"boss-layer") and .//*[contains(text(),"附件简历")]]']
        # 依次尝试多种弹窗 xpath
        for xpath in modalXpaths:
            modal = self.page.ele(xpath, timeout=1)
            if modal:
                return modal
        return None

    def waitResumePreviewModal(self):
        """等待附件简历预览弹窗出现"""
        endAt = time.time() + 8
        while time.time() < endAt:
            self.waitIfPaused()
            if self.findResumePreviewModal():
                return True
            time.sleep(0.5)  # 每 0.5 秒轮询一次
        return False

    def resumeSaveDir(self, candidateName):
        """附件简历本地保存目录"""
        root = Path(__file__).resolve().parent.parent
        saveDir = root / 'data' / 'resumes' / candidateName
        saveDir.mkdir(parents=True, exist_ok=True)  # 按候选人姓名建子目录
        return saveDir

    def setDownloadDir(self, saveDir):
        """设置浏览器下载目录"""
        absPath = str(saveDir.resolve())
        # 通过 CDP 指定 Chrome 下载路径
        with contextlib.suppress(Exception):
            self.page.run_cdp('Page.setDownloadBehavior', behavior='allow', downloadPath=absPath)
        with contextlib.suppress(Exception):
            self.page.run_cdp('Browser.setDownloadBehavior', behavior='allowAndName', downloadPath=absPath, eventsEnabled=True)

    def waitNewDownload(self, saveDir, beforeNames, timeout=45):
        """等待下载目录出现新文件"""
        endAt = time.time() + timeout
        while time.time() < endAt:
            self.waitIfPaused()
            # 仍有 .crdownload 临时文件则继续等待
            pending = [p for p in saveDir.iterdir() if p.is_file() and p.name.endswith('.crdownload')]
            if pending:
                time.sleep(0.5)
                continue
            # 扫描目录中新出现的 PDF/DOC 文件
            for filePath in saveDir.iterdir():
                if not filePath.is_file():
                    continue
                if filePath.name in beforeNames:
                    continue
                if filePath.suffix.lower() in ('.pdf', '.doc', '.docx'):
                    return filePath
            time.sleep(0.5)
        return None

    def clickPageEle(self, ele):
        """点击页面元素，兼容 svg 等节点"""
        if not ele:
            return False
        # 依次尝试直接点击、父节点点击、actions 点击
        with contextlib.suppress(Exception):
            ele.click()
            return True
        with contextlib.suppress(Exception):
            parent = ele.parent()
            if parent:
                parent.click()
                return True
        with contextlib.suppress(Exception):
            self.page.actions.move_to(ele, duration=0.1).click()
            return True
        return False

    def tryClickResumeDownloadBtn(self, modal):
        """在预览弹窗内点击下载按钮"""
        hintXpaths = ['xpath:.//*[@title="下载"]', 'xpath:.//*[contains(@title,"下载")]', 'xpath:.//*[@aria-label="下载"]', 'xpath:.//*[contains(@aria-label,"下载")]', 'xpath:.//*[contains(@class,"download")]']
        for xpath in hintXpaths:
            btn = modal.ele(xpath, timeout=0.3)
            if not btn:
                continue
            if self.clickPageEle(btn):
                self.pauseWait(self.minActionGap)
                return True
        # 尝试工具栏第三个图标（常见下载按钮位置）
        headerXpaths = ['xpath:.//div[contains(@class,"header")]', 'xpath:.//div[contains(@class,"toolbar")]', 'xpath:.//div[contains(@class,"top-bar")]', 'xpath:.//div[contains(@class,"operate")]']
        for headerXpath in headerXpaths:
            header = modal.ele(headerXpath, timeout=0.3)
            if not header:
                continue
            toolItems = header.eles('xpath:.//div[contains(@class,"operate")]/* | .//div[contains(@class,"toolbar")]/*')
            if len(toolItems) >= 3 and self.clickPageEle(toolItems[2]):
                self.pauseWait(self.minActionGap)
                return True
            icons = header.eles('xpath:.//*[self::svg or (self::i and contains(@class,"icon")) or (self::span and contains(@class,"icon"))]')
            if len(icons) >= 3 and self.clickPageEle(icons[2]):
                self.pauseWait(self.minActionGap)
                return True
        try:
            # 兜底：用 JS 在弹窗内查找并点击下载控件
            clicked = self.page.run_js('\n                function safeClick(el) {\n                    if (!el) return false;\n                    const target = el.closest(\n                        \'a,button,[role="button"],span,i,div[class*="icon"],div[class*="operate"],div\'\n                    ) || el;\n                    try {\n                        if (typeof target.click === \'function\') {\n                            target.click();\n                            return true;\n                        }\n                    } catch (e) {}\n                    target.dispatchEvent(new MouseEvent(\'click\', {\n                        bubbles: true, cancelable: true, view: window\n                    }));\n                    return true;\n                }\n                const modal = arguments[0];\n                if (!modal) return false;\n                const byTitle = modal.querySelector(\'[title*="下载"],[aria-label*="下载"]\');\n                if (byTitle && safeClick(byTitle)) {\n                    return true;\n                }\n                const header = modal.querySelector(\'[class*="header"],[class*="toolbar"],[class*="top"]\') || modal;\n                const operateItems = header.querySelectorAll(\'[class*="operate"] > *\');\n                if (operateItems.length >= 3 && safeClick(operateItems[2])) {\n                    return true;\n                }\n                const icons = header.querySelectorAll(\n                    \'svg, i[class*="icon"], span[class*="icon"], [class*="operate"] > *\'\n                );\n                const list = Array.from(icons).filter((el) => {\n                    const rect = el.getBoundingClientRect();\n                    return rect.width > 0 && rect.height > 0;\n                });\n                if (list.length >= 3) {\n                    return safeClick(list[2]);\n                }\n                return false;\n                ', modal)
        except Exception as exc:
            self.log(f'JS 点击下载按钮失败: {exc}')
            clicked = False
        if clicked:
            self.pauseWait(self.minActionGap)
        return bool(clicked)

    def clickResumeDownload(self, saveDir):
        """在预览弹窗点击下载并等待文件落盘"""
        beforeNames = {p.name for p in saveDir.iterdir() if p.is_file()}
        self.setDownloadDir(saveDir)
        modal = self.findResumePreviewModal()
        if not modal:
            self.log('未找到简历预览弹窗，无法下载')
            return None
        # 先在主文档弹窗内尝试点击下载
        if self.tryClickResumeDownloadBtn(modal):
            self.log('已点击简历下载按钮，等待文件保存...')
            filePath = self.waitNewDownload(saveDir, beforeNames)
            if filePath:
                return filePath
            self.log('下载按钮已点击，但目标目录未出现新文件')
        # 主文档失败则尝试 iframe 内的下载按钮
        iframeEle = modal.ele('tag:iframe', timeout=1)
        if iframeEle:
            with contextlib.suppress(Exception):
                frame = iframeEle.frame
                if frame:
                    frameXpaths = ['xpath://*[@title="下载"]', 'xpath://*[contains(@title,"下载")]', 'xpath://*[contains(@aria-label,"下载")]']
                    for xpath in frameXpaths:
                        btn = frame.ele(xpath, timeout=0.5)
                        if not btn:
                            continue
                        if self.clickPageEle(btn):
                            self.pauseWait(self.minActionGap)
                            self.log('已在 iframe 内点击简历下载按钮，等待文件保存...')
                            filePath = self.waitNewDownload(saveDir, beforeNames)
                            if filePath:
                                return filePath
                    frameIcons = frame.eles('xpath://div[contains(@class,"header")]//svg')
                    if len(frameIcons) >= 3 and self.clickPageEle(frameIcons[2]):
                        self.pauseWait(self.minActionGap)
                        self.log('已在 iframe 内点击工具栏下载图标，等待文件保存...')
                        filePath = self.waitNewDownload(saveDir, beforeNames)
                        if filePath:
                            return filePath
        self.log('未找到可点击的简历下载按钮')
        return None

    def closeResumePreviewModal(self):
        """关闭附件简历预览弹窗"""
        closeXpaths = ['xpath://div[contains(@class,"boss-dialog")]//*[contains(@class,"close")]', 'xpath://div[contains(@class,"dialog") and .//*[contains(text(),"附件简历")]]//*[contains(@class,"close")]', 'xpath://*[contains(@class,"icon-close")]', 'xpath://*[@aria-label="关闭"]']
        for xpath in closeXpaths:
            btn = self.page.ele(xpath, timeout=1)
            if not btn:
                continue
            btn.click()
            self.pauseWait(self.minActionGap)
            return True
        return False  # 未找到关闭按钮

    def loadSavedAttachProfile(self, candidateKey):
        """读取已保存的预览简历解析结果"""
        if not self.db:
            return None
        row = self.db.getCandidate(candidateKey)
        if not row or not row.get('resume_json'):
            return None
        try:
            data = json.loads(str(row['resume_json']))
        except json.JSONDecodeError:
            return None
        # 必须有 rawText 才算有效缓存
        if str(data.get('rawText') or '').strip():
            return data
        return None

    def fetchAttachResume(self, candidateName):
        """打开附件预览并解析简历内容"""
        # 第一步：确认聊天区有附件简历消息
        if not self.findAttachResumeMsg():
            self.log(f'{candidateName} 聊天区未找到附件简历消息')
            return None
        # 第二步：点击预览入口
        if not self.clickAttachResumePreview():
            self.log(f'{candidateName} 无法打开附件简历预览')
            return None
        # 第三步：等待预览弹窗出现
        if not self.waitResumePreviewModal():
            self.log(f'{candidateName} 附件简历预览弹窗未出现')
            return None
        self.pauseWait(2)
        # 第四步：从预览弹窗解析简历文本
        profile = self.parser.parseFromPreviewModal(self.page)
        self.closeResumePreviewModal()
        if not str(profile.get('rawText') or '').strip():
            self.log(f'{candidateName} 预览弹窗未解析到简历内容')
            return None
        profile['source'] = 'preview_modal'
        return profile

    def sendRejectNotify(self, template, candidateName, candidateKey=None, task=None):
        """发送审核不通过话术"""
        msg = self.applyPlaceholders(template, task, candidateName=candidateName)
        return self.sendOrHandoff(msg, 'reject_notify', candidateKey)

    def startMouseThread(self):
        """启动鼠标移动线程，降低被识别为脚本的风险"""
        # 已有存活线程则不再重复启动
        if self.mouseThread and self.mouseThread.is_alive():
            return

        def moveLoop():
            """循环执行轻量鼠标移动"""
            while not self.stopFlag.is_set():
                try:
                    # 暂停或无页面时跳过本轮
                    if self.mousePauseFlag.is_set() or not self.page:
                        time.sleep(0.2)
                        continue
                    rect = self.page.rect
                    size = getattr(rect, 'size', None) or (1200, 800)
                    width, height = (int(size[0]), int(size[1]))
                    # 窗口过小则等待
                    if width < 200 or height < 200:
                        time.sleep(1)
                        continue
                    endX = random.randint(80, max(100, width - 80))
                    endY = random.randint(80, max(100, height - 80))
                    with self.actionLock:
                        if self.page and (not self.mousePauseFlag.is_set()):
                            self.page.actions.move_to((endX, endY), duration=0.3)
                    self.mousePos = [endX, endY]
                    time.sleep(random.uniform(2.0, 4.0))
                except Exception as exc:
                    if self.stopFlag.is_set():
                        break
                    time.sleep(1)
        self.mouseThread = threading.Thread(target=moveLoop, name='MouseMove', daemon=True)
        self.mouseThread.start()

    def pickTemplateWord(self, templateType, words):
        """从话术列表随机选一条"""
        enabled = [w for w in words if w]
        if not enabled:
            return ''
        return random.choice(enabled)  # 同类型话术随机轮换

    def readJobName(self, listItem):
        """读取沟通列表中的应聘岗位"""
        jobEle = listItem.ele('xpath:.//span[contains(@class, "source-job")]', timeout=0.5)
        if not jobEle or not jobEle.text:
            return ''  # 列表项无岗位标签
        return str(jobEle.text).strip()

    def markCandidateUnsuitable(self, candidateKey, candidateName, reason):
        """记录不合适决定并永久停止候选人自动流程"""
        if self.db:
            self.db.markUnsuitable(candidateKey, reason)
            if not self.db.hasSentToday(candidateKey, 'mark_unsuitable'):
                self.db.recordAction(
                    candidateKey,
                    'mark_unsuitable',
                    word=reason,
                    success=True,
                    taskRunId=self.currentTaskRunId,
                )
        self.log(f'已标记 {candidateName} 为不合适：{reason}')

    def handleCandidateReply(self, task, candidateKey, candidateName, jobName, pendingTexts):
        """等待人工回复并决定是否允许自动索要简历"""
        friendText = '\n'.join(pendingTexts)
        if self.db:
            self.db.setResumeStatus(candidateKey, 'awaiting_manual_reply')
        decision = self.waitForManualReply(
            task,
            candidateKey,
            candidateName,
            jobName,
            '候选人有待回复消息，自动流程已暂停',
            friendText,
            mode='candidate_reply',
        )
        if decision == 'unsuitable':
            self.markCandidateUnsuitable(candidateKey, candidateName, '人工判断沟通不合适')
            return 'unsuitable'
        if decision != 'continue':
            return ''
        if self.db:
            self.db.recordAction(
                candidateKey,
                'manual_reply',
                word='人工已回复并确认合适',
                success=True,
                taskRunId=self.currentTaskRunId,
                extra={'friendText': friendText[:200]},
            )
            self.db.recordAction(
                candidateKey,
                'manual_resume_approved',
                word='允许自动索要简历',
                success=True,
                taskRunId=self.currentTaskRunId,
            )
            self.db.setResumeStatus(candidateKey, 'manual_approved')
        self.log(f'人工已确认 {candidateName} 合适，允许继续索要简历')
        return 'continue'

    def processCandidate(self, task, listItem, candidateKey, candidateName, listItemKey, jobName='', wasUnread=False):
        """处理单个候选人的简历客服流程，返回是否计入单次处理上限"""
        self.currentCandidateKey = candidateKey
        self.currentCandidateName = candidateName
        self.currentJobName = jobName or ''
        self.currentTask = task
        try:
            return self.processCandidateBody(task, listItem, candidateKey, candidateName, listItemKey, jobName, wasUnread)
        finally:
            self.currentCandidateKey = ''
            self.currentCandidateName = ''
            self.currentJobName = ''
            self.currentTask = None

    def processCandidateBody(self, task, listItem, candidateKey, candidateName, listItemKey, jobName='', wasUnread=False):
        """处理单个候选人（内部实现），返回是否计入单次处理上限"""
        self.ensureVerifyClear()
        if self.manualReplyRequested and self.currentCandidateKey:
            self.ensureManualReply(task, candidateKey, candidateName, jobName or self.currentJobName, '')
        # 从任务参数读取话术与规则配置
        rules = task.get('resumeRules') or {}
        interviewPreWords = task.get('interviewPreWords') or []
        interviewAskWords = task.get('interviewAskWords') or []
        interviewRemindWords = task.get('interviewRemindWords') or []
        interviewRescheduleWords = task.get('interviewRescheduleWords') or []
        interviewCancelWords = task.get('interviewCancelWords') or []
        rejectWords = task.get('rejectWords') or []
        maxFollowDays = int(task.get('maxFollowDays') or rules.get('maxFollowDays') or 7)
        jobName = jobName or self.readJobName(listItem)
        # 超过跟进天数上限则跳过
        if self.db:
            canFollow, stopReason = self.db.canFollowCandidate(candidateKey, maxFollowDays)
            if not canFollow:
                self.log(f'跳过 {candidateName}: {stopReason}')
                return False
        # 点击左侧列表项进入该候选人聊天窗口
        nameEle = listItem.ele('xpath:.//span[contains(@class, "geek-name")]', timeout=1)
        if not nameEle:
            return False
        nameEle.click()
        self.page.wait.doc_loaded()
        self.pauseWait(2)
        self.ensureVerifyClear()
        status = self.db.getResumeStatus(candidateKey) if self.db else 'new'
        resumeCard = self.detectResumeCard()
        timeline = self.readChatTimeline()
        self.checkRiskPage()
        pendingAll = self.getPendingFriendTexts(timeline, todayOnly=False)
        replyStatuses = ('new', 'greeted', 'resume_requested', 'manual_approved', 'awaiting_manual_reply')
        # 程序打开前已读且我方未回复的历史会话，直接永久停止
        if status in replyStatuses and pendingAll and status != 'awaiting_manual_reply' and not wasUnread:
            self.markCandidateUnsuitable(candidateKey, candidateName, '历史会话我方已读未回复')
            self.skipChatWait(task)
            return True
        # 新收到的候选人回复必须由人工回复并明确作出决定
        if status in replyStatuses and (pendingAll or status == 'awaiting_manual_reply'):
            decisionTexts = pendingAll
            if not decisionTexts:
                friendItems = [item.get('text') for item in timeline if item.get('sender') == 'friend' and item.get('text')]
                decisionTexts = friendItems[-1:] if friendItems else ['候选人历史消息']
            decision = self.handleCandidateReply(task, candidateKey, candidateName, jobName, decisionTexts)
            if decision == 'unsuitable':
                self.skipChatWait(task)
                return True
            if decision != 'continue':
                self.skipChatWait(task)
                return False
            status = 'manual_approved'
            timeline = self.readChatTimeline()
        # 仅处理当天消息的候选人
        if not self.shouldProcessToday(task, candidateKey, candidateName, listItem, timeline, status, resumeCard):
            if self.db:
                self.db.touchCandidate(candidateKey)
            self.skipChatWait(task)
            return False
        # 面试预邀请 / 跟进 / 人工正式邀约等待
        if self.inInterviewFlow(status) or status == 'interview_formal_pending':
            self.tryHandleInterviewFlow(task, candidateKey, candidateName, jobName, status, timeline)
            if self.db:
                self.db.touchCandidate(candidateKey)
            self.randomChatWait(task)
            return True
        # 有待同意简历卡片时先点击同意
        if resumeCard == 'pending_accept':
            if self.acceptResume():
                self.log(f'已同意接收 {candidateName} 的简历')
                if self.db:
                    self.db.recordAction(candidateKey, 'accept_resume', success=True, taskRunId=self.currentTaskRunId)
                    self.noteResumeReceived(candidateKey, 'online_card')
                resumeCard = 'accepted'
                status = 'resume_received'
        elif resumeCard == 'accepted' and status not in ('resume_received', 'resume_requested', 'resume_passed', 'resume_rejected'):
            if self.db:
                self.noteResumeReceived(candidateKey, 'online_card')
                status = 'resume_received'
        # 人工确认合适后才允许程序发起求简历
        if status == 'manual_approved' and resumeCard == 'none':
            if self.requestResume():
                self.log(f'已向 {candidateName} 发起求简历')
                if self.db:
                    self.db.recordAction(candidateKey, 'request_resume', success=True, taskRunId=self.currentTaskRunId)
                    self.db.setResumeStatus(candidateKey, 'resume_requested')
                    status = 'resume_requested'
            resumeCard = self.detectResumeCard()
            attachReady = self.hasAttachResumeInChat()
            # 对方已主动发送简历时直接进入原有审核流程
            if attachReady or resumeCard != 'none':
                if self.db:
                    if attachReady:
                        self.noteResumeReceived(candidateKey, 'attach')
                    else:
                        self.noteResumeReceived(candidateKey, 'online_card')
                    self.db.touchCandidate(candidateKey)
            else:
                if self.db:
                    self.db.touchCandidate(candidateKey)
                self.randomChatWait(task)
                return True
        # 未收到候选人回复时只等待，不自动追问或索要简历
        if status in ('new', 'greeted') and resumeCard == 'none':
            self.log(f'等待 {candidateName} 回复，不执行自动跟进')
            if self.db:
                self.db.touchCandidate(candidateKey)
            self.skipChatWait(task)
            return False
        attachReady = self.hasAttachResumeInChat()
        if attachReady:
            resumeCard = self.detectResumeCard()
            if self.db:
                self.noteResumeReceived(candidateKey, 'attach')
        # 已求简历但对方尚未发送时只等待，不发送任何跟进话术
        if status == 'resume_requested' and resumeCard == 'none' and not attachReady:
            self.log(f'等待 {candidateName} 发送简历，不执行自动跟进')
            if self.db:
                self.db.touchCandidate(candidateKey)
            self.skipChatWait(task)
            return False
        # 未到可审核阶段则结束
        if not self.shouldReviewResume(status, resumeCard, attachReady):
            return False
        # 已完成正式面试邀约则跳过
        if self.alreadySentInterview(candidateKey):
            self.log(f'跳过 {candidateName}: 已完成正式面试邀约')
            return False
        # 今日已人工确认简历处理且未完成自动审核，不再重复弹窗
        if self.db and self.db.hasManualResumeReviewToday(candidateKey) and not self.db.hasResumeReview(candidateKey):
            self.log(f'{candidateName} 今日已人工确认简历处理，跳过自动审核')
            self.db.touchCandidate(candidateKey)
            self.randomChatWait(task)
            return True
        profile = None
        attachMode = attachReady or self.hasAttachResumeInChat()
        # 附件简历模式：优先用已解析缓存，否则打开预览弹窗解析
        if attachMode:
            if self.db:
                self.noteResumeReceived(candidateKey, 'attach')
            profile = self.loadSavedAttachProfile(candidateKey)
            if profile:
                self.log(f'使用已解析的 {candidateName} 预览简历')
            else:
                self.log(f'正在预览 {candidateName} 的附件简历')
                profile = self.fetchAttachResume(candidateName)
                if not profile:
                    return self.handleResumeParseFailure(
                        task,
                        candidateKey,
                        candidateName,
                        jobName,
                        '附件简历预览解析失败，请人工查看简历并处理',
                        'attach_preview',
                    )
                if self.db:
                    self.db.recordAction(candidateKey, 'preview_resume', success=True, taskRunId=self.currentTaskRunId)
                    self.db.markResumeReceived(candidateKey, source='attach_preview', partialProfile=profile)
        # 无附件或未解析到时从在线简历页面解析
        if not profile:
            if resumeCard == 'accepted' and self.db:
                self.noteResumeReceived(candidateKey, 'online_page')
            profile = self.parser.parseFromPage(self.page)
        if not str((profile or {}).get('rawText') or '').strip():
            return self.handleResumeParseFailure(
                task,
                candidateKey,
                candidateName,
                jobName,
                '在线简历未能自动解析，请人工查看简历并处理',
                'online_page',
            )
        # 加载岗位匹配规则并执行审核
        if self.db:
            jobRules = self.db.matchJobRules(jobName)
            if not jobRules:
                self.log(f"跳过 {candidateName}: 未匹配岗位要求「{jobName or '未知岗位'}」")
                if self.db:
                    self.db.touchCandidate(candidateKey)
                self.skipChatWait(task)
                return False
            self.matcher.loadRules(jobRules)
            self.log(f"按岗位要求审核: {jobRules.get('jobName')}")
        else:
            self.matcher.loadRules(rules)
        profile['jobName'] = jobName
        ok, reason = self.matcher.match(profile)
        self.log(f'简历审核 {candidateName}: {reason}')
        if self.db:
            self.db.saveResumeReview(candidateKey, profile, ok, reason)
            self.db.recordAction(candidateKey, 'resume_review', word=reason, success=ok, taskRunId=self.currentTaskRunId, extra=profile)
        # 审核通过 → 发送面试预邀请聊天
        if ok:
            self.trySendInterviewPre(task, candidateKey, candidateName, jobName, interviewPreWords)
        else:
            # 审核不通过 → 发送 reject 话术（每日限一次）
            rejectText = self.pickTemplateWord('reject', rejectWords)
            sentReject = self.db and self.db.hasSentToday(candidateKey, 'reject_notify')
            handoffReject = self.db and self.db.hasHandoffToday(candidateKey, 'reject_notify')
            if rejectText and (not sentReject) and (not handoffReject):
                result = self.sendRejectNotify(rejectText, candidateName, candidateKey, task)
                if result == 'sent':
                    self.log(f'已向 {candidateName} 发送未通过通知')
                    if self.db:
                        self.db.recordAction(candidateKey, 'reject_notify', word=rejectText, success=True, taskRunId=self.currentTaskRunId)
                elif result == 'handoff':
                    self.log(f'已向 {candidateName} 填入未通过通知（待人工发送）')
            if self.db:
                self.db.setResumeStatus(candidateKey, 'resume_rejected')
        if self.db:
            self.db.touchCandidate(candidateKey)
        self.randomChatWait(task)
        return True

    def runResume(self, task):
        """执行一轮推荐牛人主动联系与简历沟通任务"""
        self.applyRateLimits(task)
        self.ensureWorkWindow()
        self.ensureVerifyClear()
        maxRun = int(self.rateLimits().get('maxCandidatesPerRun') or 5)
        processedCount = 0
        testName = str(task.get('testCandidateName') or '').strip()
        # 测试指定候选人时跳过推荐页，正常任务先执行主动联系
        if not testName:
            try:
                self.runRecommend(task)
            except StopRequested:
                raise
            except Exception as exc:
                self.log(f'推荐牛人主动联系异常，本轮继续处理沟通列表：{exc}')
        # 进入沟通 Tab
        self.selectChatTab()
        keySet = set()  # 本轮已扫描过的列表项 key，避免重复
        pastTodayBoundary = False  # 列表已滚到昨天及更早，不再向下滚动
        scrollIdle = 0
        while True:
            self.waitIfPaused()
            self.ensureVerifyClear()
            roundVisited = False
            itemList = self.page.eles('xpath://div[@class="user-container"]//div[@role="group"]/div')
            if not itemList:
                break
            for listItem in itemList:
                self.waitIfPaused()
                listKey = listItem.attr('key')
                # 跳过无 key 或已扫描项
                if not listKey or listKey in keySet:
                    continue
                nameEle = listItem.ele('xpath:.//span[contains(@class, "geek-name")]', timeout=0.5)
                if not nameEle or not nameEle.text:
                    continue
                candidateName = nameEle.text
                # 测试模式只处理指定姓名
                if testName and candidateName != testName:
                    continue
                listLabel = self.readListItemTimeLabel(listItem)
                candidateKey = BossDb.buildCandidateKey(candidateName, listKey)
                # 非测试：列表时间越过当天边界（昨天及更早）则停止向下滚动
                if not testName and self.isPastTodayListBoundary(listLabel):
                    keySet.add(listKey)
                    if not self.shouldProcessForRemindCancel(candidateKey, task):
                        pastTodayBoundary = True
                        break
                    pastTodayBoundary = True
                if self.db:
                    self.db.getOrCreateCandidate(candidateName, listKey)
                wasUnread = self.isListItemUnread(listItem)
                # 列表阶段跳过确定无待办的候选人
                if not self.shouldPickListItem(task, candidateKey, candidateName, listLabel):
                    keySet.add(listKey)
                    continue
                keySet.add(listKey)
                # 过滤纯英文姓名（noForeigner 选项）
                if task.get('noForeigner') and any((c.isalpha() and c.isascii() for c in candidateName)):
                    continue
                jobName = self.readJobName(listItem)
                self.log(f'处理候选人: {candidateName}' + (f'（{jobName}）' if jobName else ''))
                roundVisited = True
                handled = self.processCandidate(task, listItem, candidateKey, candidateName, listKey, jobName, wasUnread)
                if handled:
                    processedCount += 1
                    if processedCount >= maxRun:
                        self.log(f'已达单次任务处理上限 {maxRun} 人，结束任务')
                        return
                if testName:
                    self.log(f'测试模式：已处理 {testName}，结束任务')
                    return
                # 每轮只点进一个联系人，避免列表滚动错乱
                break
            if testName and not roundVisited:
                self.log(f'测试模式：列表中未找到 {testName}')
                break
            if pastTodayBoundary:
                self.log('已到达列表当天边界（昨天及更早），停止滚动')
                break
            if roundVisited:
                scrollIdle = 0
                continue
            if self.scrollChatListDown():
                scrollIdle = 0
                continue
            scrollIdle += 1
            if scrollIdle >= 2:
                self.log('聊天列表已扫完，结束本轮扫描')
                break

    def notifyTaskDone(self, task, status, error=None):
        """通知上层任务完成"""
        if self.taskDoneCallback:
            self.taskDoneCallback(task, status, error)  # 回调 Web 层更新任务状态

    def main(self, taskQueue):
        """任务主循环"""
        browserReady = False
        try:
            while True:
                self.waitIfPaused()
                try:
                    task = taskQueue.get(timeout=0.5)
                except Empty:
                    self.pauseWait(0.5)
                    continue
                # 收到 stop 任务则退出主循环
                if task is None or task.get('taskType') == 'stop':
                    self.log('收到停止指令')
                    break
                # 首次任务前初始化浏览器、登录与安全验证
                if not browserReady:
                    if not self.page:
                        self.initBrowser()
                        self.selectChatTab()
                        self.waitForLogin()
                        self.ensureVerifyClear()
                    else:
                        self.selectChatTab()
                    self.ensureVerifyClear()
                    self.startMouseThread()
                    browserReady = True
                taskRunId = None
                if self.db:
                    taskRunId = self.db.createTaskRun(self.browserId, task.get('taskType', 'resume'), task.get('params') or {})
                self.currentTaskRunId = taskRunId
                try:
                    if task.get('taskType') == 'resume':
                        params = task.get('params') or {}
                        self.ensureVerifyClear()
                        self.runResume(params)
                        if self.db and taskRunId:
                            self.db.finishTaskRun(taskRunId, 'success')
                        self.notifyTaskDone(task, 'success')
                except StopRequested as exc:
                    if self.db and taskRunId:
                        self.db.finishTaskRun(taskRunId, 'stopped', str(exc))
                    self.notifyTaskDone(task, 'stopped', str(exc))
                    break
                except Exception as exc:
                    if self.stopFlag.is_set():
                        if self.db and taskRunId:
                            self.db.finishTaskRun(taskRunId, 'stopped', '用户请求停止')
                        self.notifyTaskDone(task, 'stopped', '用户请求停止')
                        break
                    if self.db and taskRunId:
                        self.db.finishTaskRun(taskRunId, 'failed', str(exc))
                    self.notifyTaskDone(task, 'failed', str(exc))
                    raise
                finally:
                    self.currentTaskRunId = None
        except StopRequested:
            self.log('自动化已停止')
        finally:
            # 清理：停止标志、页面引用、等待鼠标线程结束
            self.stopFlag.set()
            self.mousePauseFlag.set()
            self.page = None
            if self.mouseThread and self.mouseThread.is_alive():
                self.mouseThread.join(timeout=2)
if __name__ == '__main__':
    config = {'browserId': 'default-browser', 'userDataPath': 'D:\\boss_zhaopin_筛选简历\\boss_chrome_profile', 'connectMode': 'local'}
    boss = BossAuto(browserId=config['browserId'], userDataPath=config['userDataPath'])
    boss.connectMode = config['connectMode']
    taskQueue = Queue()
    taskQueue.put({'taskType': 'resume', 'params': {'greetingWords': ['您好，方便发一份最新简历吗？'], 'followupWords': ['您好，请问方便更新在线简历吗？'], 'interviewPreWords': ['{name}你好，恭喜通过初筛，请于{date}{time}到{address}面试，预计{duration}分钟。'], 'interviewConfig': {'dayOffset': 1, 'timeSlots': [14], 'address': '深圳', 'duration': '40-60'}, 'chatInterval': 15, 'noForeigner': True, 'resumeRules': {'ageMin': 18, 'ageMax': 45, 'educationList': ['本科', '大专'], 'workYearsMin': 0, 'mustKeywords': [], 'rejectKeywords': []}}})
    taskQueue.put({'taskType': 'stop'})
    boss.main(taskQueue)
