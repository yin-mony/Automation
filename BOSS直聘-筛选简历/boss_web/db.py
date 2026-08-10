import datetime as dt
import json
import re
import sqlite3
import threading
from pathlib import Path


class BossDb:
    """SQLite 存储：任务、候选人、简历审核、话术与浏览器配置"""

    def __init__(self, dbPath=None):
        """初始化数据库路径、连接与业务默认配置"""
        # 分发标准目录（与 exe 位置无关，便于稳定读写）
        self.appRootDir = Path('D:/boss_zhaopin_筛选简历')
        self.dbDir = self.appRootDir / 'db'
        self.chromeProfileDir = self.appRootDir / 'boss_chrome_profile'
        # 数据库文件路径，未指定则用 D:\boss_zhaopin_筛选简历\db\boss_automation.db
        self.dbPath = Path(dbPath) if dbPath else self.dbDir / 'boss_automation.db'
        # 写库锁，避免多线程并发写入冲突
        self.writeLock = threading.RLock()
        # 确保 db 目录存在
        self.dbPath.parent.mkdir(parents=True, exist_ok=True)
        # 建立 SQLite 连接，允许跨线程复用
        self.conn = sqlite3.connect(str(self.dbPath), timeout=5, check_same_thread=False)
        # 查询结果按列名映射为 Row
        self.conn.row_factory = sqlite3.Row
        # WAL 模式提升并发读写性能
        self.conn.execute('PRAGMA journal_mode=WAL')
        # 锁等待超时 5 秒
        self.conn.execute('PRAGMA busy_timeout=5000')
        # 启用外键约束
        self.conn.execute('PRAGMA foreign_keys=ON')
        # 计入每日消息上限的动作类型
        self.messageActionTypes = [
            'greeting',
            'smart_reply',
            'followup',
            'interview_pre',
            'interview_ask_time',
            'interview_remind',
            'interview_reschedule',
            'interview_cancel',
            'reject_notify',
        ]
        # 企微日报缓存键（仅存 Webhook 与 @ 手机号）
        self.wecomWebhookKey = 'wecom_webhook_url'
        self.wecomMentionMobileKey = 'wecom_mention_mobile'
        # 话术占位符取值（JSON 存 app_settings）
        self.templatePlaceholderKey = 'template_placeholders'
        # 推荐牛人每日主动联系上限（与普通聊天消息限额分开统计）
        self.recommendLimitKey = 'recommend_daily_limit'
        self.defaultRecommendLimit = 15
        # 本地回复模型配置键
        self.replyEnabledKey = 'reply_enabled'
        self.replyBaseUrlKey = 'reply_base_url'
        self.replyModelNameKey = 'reply_model_name'
        self.replyTimeoutKey = 'reply_timeout_sec'
        self.replyActiveSkillKey = 'reply_active_skill_id'
        self.defaultReplyBaseUrl = 'http://127.0.0.1:1234/v1'
        self.defaultReplyModelName = 'qwen3-8b'
        self.defaultReplyTimeout = 90
        # 首次连接时创建全部业务表
        self.initSchema()
        # 兼容旧库，补全缺失列与新表
        self.migrateSchema()

    def close(self):
        """关闭数据库连接"""
        # 释放连接资源
        self.conn.close()

    def initSchema(self):
        """创建全部业务表"""
        with self.writeLock:
            # 一次性执行建表脚本：任务、候选人、聊天动作、话术、简历规则、配置、浏览器、岗位规则
            self.conn.executescript("""
                CREATE TABLE IF NOT EXISTS task_runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    browser_id TEXT NOT NULL,
                    task_type TEXT NOT NULL,
                    params_json TEXT NOT NULL,
                    status TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    finished_at TEXT,
                    error TEXT
                );

                CREATE TABLE IF NOT EXISTS candidate_records (
                    candidate_key TEXT PRIMARY KEY,
                    candidate_name TEXT NOT NULL,
                    list_item_key TEXT NOT NULL,
                    first_contact_date TEXT NOT NULL,
                    last_contact_date TEXT,
                    resume_status TEXT NOT NULL DEFAULT 'new',
                    resume_received_at TEXT,
                    resume_pass INTEGER,
                    resume_reject_reason TEXT,
                    interview_sent_at TEXT,
                    resume_json TEXT,
                    stopped INTEGER NOT NULL DEFAULT 0,
                    stop_reason TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS chat_actions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    candidate_key TEXT NOT NULL,
                    action_type TEXT NOT NULL,
                    word TEXT,
                    success INTEGER NOT NULL,
                    action_date TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    task_run_id INTEGER,
                    extra_json TEXT,
                    FOREIGN KEY(candidate_key) REFERENCES candidate_records(candidate_key),
                    FOREIGN KEY(task_run_id) REFERENCES task_runs(id),
                    UNIQUE(candidate_key, action_type, action_date)
                );

                CREATE TABLE IF NOT EXISTS message_templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    template_type TEXT NOT NULL,
                    word TEXT NOT NULL,
                    enabled INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS resume_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    age_min INTEGER NOT NULL DEFAULT 18,
                    age_max INTEGER NOT NULL DEFAULT 45,
                    education_json TEXT NOT NULL,
                    work_years_min INTEGER NOT NULL DEFAULT 0,
                    must_keywords_json TEXT NOT NULL,
                    reject_keywords_json TEXT NOT NULL,
                    interview_time TEXT NOT NULL DEFAULT '',
                    max_follow_days INTEGER NOT NULL DEFAULT 7,
                    chat_interval INTEGER NOT NULL DEFAULT 15,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS app_settings (
                    setting_key TEXT PRIMARY KEY,
                    setting_value TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS browser_profiles (
                    browser_id TEXT PRIMARY KEY,
                    remark TEXT NOT NULL UNIQUE,
                    user_data_path TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_chat_actions_date ON chat_actions(action_date);
                CREATE INDEX IF NOT EXISTS idx_task_runs_started_at ON task_runs(started_at);

                CREATE TABLE IF NOT EXISTS job_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_name TEXT NOT NULL UNIQUE,
                    match_keys_json TEXT NOT NULL DEFAULT '[]',
                    job_intro TEXT NOT NULL DEFAULT '',
                    age_min INTEGER NOT NULL DEFAULT 18,
                    age_max INTEGER NOT NULL DEFAULT 45,
                    education_json TEXT NOT NULL DEFAULT '[]',
                    work_years_min INTEGER NOT NULL DEFAULT 0,
                    must_keywords_json TEXT NOT NULL DEFAULT '[]',
                    any_keywords_json TEXT NOT NULL DEFAULT '[]',
                    prefer_keywords_json TEXT NOT NULL DEFAULT '[]',
                    reject_keywords_json TEXT NOT NULL DEFAULT '[]',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS reply_skills (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    skill_name TEXT NOT NULL UNIQUE,
                    instruction TEXT NOT NULL,
                    examples TEXT NOT NULL DEFAULT '',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS reply_feedback (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    candidate_key TEXT NOT NULL DEFAULT '',
                    skill_id INTEGER,
                    job_name TEXT NOT NULL DEFAULT '',
                    friend_text TEXT NOT NULL DEFAULT '',
                    suggested_reply TEXT NOT NULL DEFAULT '',
                    final_reply TEXT NOT NULL DEFAULT '',
                    recommendation TEXT NOT NULL DEFAULT 'reply_only',
                    accepted INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(skill_id) REFERENCES reply_skills(id)
                );
                """)
            # 建表脚本执行完毕，提交事务
            self.conn.commit()

    def migrateSchema(self):
        """兼容旧库，补全简历相关列"""
        # 读取 candidate_records 现有列名
        columns = {row[1] for row in self.conn.execute('PRAGMA table_info(candidate_records)').fetchall()}
        # 旧库可能缺失的简历流程字段
        patches = [
            ('resume_status', "TEXT NOT NULL DEFAULT 'new'"),
            ('resume_received_at', 'TEXT'),
            ('resume_pass', 'INTEGER'),
            ('resume_reject_reason', 'TEXT'),
            ('interview_sent_at', 'TEXT'),
            ('resume_json', 'TEXT'),
            ('pre_invite_sent_at', 'TEXT'),
            ('agreed_date', 'TEXT'),
            ('agreed_time', 'TEXT'),
            ('interview_address', 'TEXT'),
            ('interview_job_name', 'TEXT'),
            ('interview_asked_at', 'TEXT'),
        ]
        with self.writeLock:
            # 逐列检查，缺失则 ALTER TABLE 追加
            for name, ddl in patches:
                if name not in columns:
                    # 为候选人表补全简历相关列
                    self.conn.execute(f'ALTER TABLE candidate_records ADD COLUMN {name} {ddl}')
            # 旧库可能没有 job_rules 表，确保存在
            self.conn.executescript("""
                CREATE TABLE IF NOT EXISTS job_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_name TEXT NOT NULL UNIQUE,
                    match_keys_json TEXT NOT NULL DEFAULT '[]',
                    job_intro TEXT NOT NULL DEFAULT '',
                    age_min INTEGER NOT NULL DEFAULT 18,
                    age_max INTEGER NOT NULL DEFAULT 45,
                    education_json TEXT NOT NULL DEFAULT '[]',
                    work_years_min INTEGER NOT NULL DEFAULT 0,
                    must_keywords_json TEXT NOT NULL DEFAULT '[]',
                    any_keywords_json TEXT NOT NULL DEFAULT '[]',
                    prefer_keywords_json TEXT NOT NULL DEFAULT '[]',
                    reject_keywords_json TEXT NOT NULL DEFAULT '[]',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS reply_skills (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    skill_name TEXT NOT NULL UNIQUE,
                    instruction TEXT NOT NULL,
                    examples TEXT NOT NULL DEFAULT '',
                    enabled INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS reply_feedback (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    candidate_key TEXT NOT NULL DEFAULT '',
                    skill_id INTEGER,
                    job_name TEXT NOT NULL DEFAULT '',
                    friend_text TEXT NOT NULL DEFAULT '',
                    suggested_reply TEXT NOT NULL DEFAULT '',
                    final_reply TEXT NOT NULL DEFAULT '',
                    recommendation TEXT NOT NULL DEFAULT 'reply_only',
                    accepted INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(skill_id) REFERENCES reply_skills(id)
                );
                """)
            # 读取 job_rules 列，检查 job_intro 是否已迁移
            jobColumns = {row[1] for row in self.conn.execute('PRAGMA table_info(job_rules)').fetchall()}
            if jobColumns and 'job_intro' not in jobColumns:
                # 旧版 job_rules 无岗位介绍列，追加默认值
                self.conn.execute("ALTER TABLE job_rules ADD COLUMN job_intro TEXT NOT NULL DEFAULT ''")
            # 迁移变更一次性提交
            self.conn.commit()

    def nowText(self):
        """返回当前 ISO 时间字符串"""
        return dt.datetime.now().isoformat(timespec='seconds')

    def todayText(self):
        """返回今天日期"""
        return dt.date.today().isoformat()

    @staticmethod
    def buildCandidateKey(candidateName, listItemKey):
        """生成候选人唯一键"""
        return f'{candidateName}:{listItemKey}'

    def rowToDict(self, row):
        """Row 转 dict"""
        return dict(row) if row else None

    def createTaskRun(self, browserId, taskType, params):
        """创建任务运行记录"""
        now = self.nowText()
        # 任务参数序列化为 JSON 入库
        paramsJson = json.dumps(params, ensure_ascii=False)
        with self.writeLock:
            # 插入 running 状态的任务记录
            cur = self.conn.execute(
                """
                INSERT INTO task_runs (browser_id, task_type, params_json, status, started_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (browserId, taskType, paramsJson, 'running', now),
            )
            self.conn.commit()
            # 返回新任务自增 ID
            return int(cur.lastrowid)

    def finishTaskRun(self, taskRunId, status, error=None):
        """标记任务结束"""
        with self.writeLock:
            # 更新任务状态、结束时间与错误信息
            self.conn.execute(
                """
                UPDATE task_runs SET status = ?, finished_at = ?, error = ? WHERE id = ?
                """,
                (status, self.nowText(), error, taskRunId),
            )
            self.conn.commit()

    def getOrCreateCandidate(self, candidateName, listItemKey):
        """获取或创建候选人"""
        candidateKey = self.buildCandidateKey(candidateName, listItemKey)
        today = self.todayText()
        now = self.nowText()
        with self.writeLock:
            # 不存在则插入新候选人，初始 resume_status 为 new
            self.conn.execute(
                """
                INSERT OR IGNORE INTO candidate_records (
                    candidate_key, candidate_name, list_item_key,
                    first_contact_date, resume_status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, 'new', ?, ?)
                """,
                (candidateKey, candidateName, listItemKey, today, now, now),
            )
            self.conn.commit()
        # 返回完整候选人记录，新建失败时给空 dict
        return self.getCandidate(candidateKey) or {}

    def getCandidate(self, candidateKey):
        """按 key 查询候选人"""
        # 按主键查询单条候选人
        row = self.conn.execute(
            'SELECT * FROM candidate_records WHERE candidate_key = ?',
            (candidateKey,),
        ).fetchone()
        return self.rowToDict(row)

    def getResumeStatus(self, candidateKey):
        """读取候选人简历流程状态"""
        row = self.getCandidate(candidateKey)
        # 无记录视为新候选人
        if not row:
            return 'new'
        return str(row.get('resume_status') or 'new')

    def setResumeStatus(self, candidateKey, status):
        """更新简历流程状态"""
        with self.writeLock:
            # 写入新状态并刷新 updated_at
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (status, self.nowText(), candidateKey),
            )
            self.conn.commit()

    def markResumeReceived(self, candidateKey, source='', partialProfile=None):
        """标记已收到简历并记录收到时间（程序中断后可恢复）"""
        now = self.nowText()
        snapshot = {}
        if partialProfile:
            snapshot = dict(partialProfile)
        if source:
            snapshot['source'] = str(source)
        if snapshot and 'parsed' not in snapshot:
            snapshot['parsed'] = bool(str(snapshot.get('rawText') or '').strip())
        resumeJson = json.dumps(snapshot, ensure_ascii=False) if snapshot else None
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'resume_received',
                    resume_received_at = COALESCE(resume_received_at, ?),
                    resume_json = CASE
                        WHEN ? IS NOT NULL AND (resume_json IS NULL OR resume_json = '') THEN ?
                        ELSE resume_json
                    END,
                    updated_at = ?
                WHERE candidate_key = ?
                """,
                (now, resumeJson, resumeJson, now, candidateKey),
            )
            self.conn.commit()

    def hasResumeReview(self, candidateKey):
        """是否已完成自动简历审核（含通过/不通过）"""
        status = self.getResumeStatus(candidateKey)
        if status in ('resume_passed', 'resume_rejected'):
            return True
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ? AND action_type = 'resume_review'
            LIMIT 1
            """,
            (candidateKey,),
        ).fetchone()
        return row is not None

    def pendingResumeReview(self, candidateKey):
        """是否已收简历但尚未完成审核"""
        if self.hasResumeReview(candidateKey):
            return False
        status = self.getResumeStatus(candidateKey)
        if status in ('resume_received', 'resume_requested'):
            return True
        row = self.getCandidate(candidateKey)
        if row and str(row.get('resume_received_at') or '').strip():
            return True
        return False

    def hasManualResumeReviewToday(self, candidateKey):
        """今日是否已人工确认简历处理"""
        return self.hasSentToday(candidateKey, 'manual_resume_review')

    def saveResumeReview(self, candidateKey, resumeJson, passed, reason):
        """保存简历审核结果"""
        now = self.nowText()
        with self.writeLock:
            # 保存简历 JSON、通过标记、拒绝原因及流程状态
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_json = ?, resume_pass = ?, resume_reject_reason = ?,
                    resume_received_at = COALESCE(resume_received_at, ?),
                    resume_status = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (
                    json.dumps(resumeJson, ensure_ascii=False),
                    1 if passed else 0,
                    reason if not passed else None,
                    now,
                    'resume_passed' if passed else 'resume_rejected',
                    now,
                    candidateKey,
                ),
            )
            self.conn.commit()

    def markInterviewSent(self, candidateKey):
        """标记已发送正式面试邀约（人工在 BOSS 完成）"""
        now = self.nowText()
        with self.writeLock:
            # 记录面试邀请时间并更新流程为 interview_sent
            self.conn.execute(
                """
                UPDATE candidate_records
                SET interview_sent_at = ?, resume_status = 'interview_sent', updated_at = ?
                WHERE candidate_key = ?
                """,
                (now, now, candidateKey),
            )
            self.conn.commit()

    def saveInterviewPreSent(self, candidateKey, agreedDate, agreedTime, address, jobName):
        """保存预邀请发出后的约定时间地点"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET pre_invite_sent_at = ?, agreed_date = ?, agreed_time = ?,
                    interview_address = ?, interview_job_name = ?,
                    resume_status = 'interview_pre_sent', updated_at = ?
                WHERE candidate_key = ?
                """,
                (now, agreedDate, agreedTime, address, jobName, now, candidateKey),
            )
            self.conn.commit()

    def updateInterviewAgreedTime(self, candidateKey, agreedDate, agreedTime, status='interview_reschedule_pending'):
        """更新改期后的约定时间"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET agreed_date = ?, agreed_time = ?, resume_status = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (agreedDate, agreedTime, status, now, candidateKey),
            )
            self.conn.commit()

    def setInterviewAwaitingTime(self, candidateKey):
        """标记已追问对方方便时间"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'interview_awaiting_time',
                    interview_asked_at = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (now, now, candidateKey),
            )
            self.conn.commit()

    def setInterviewFormalPending(self, candidateKey):
        """聊天已确认，等待人工发 BOSS 正式邀约"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'interview_formal_pending', updated_at = ?
                WHERE candidate_key = ?
                """,
                (now, candidateKey),
            )
            self.conn.commit()

    def markInterviewCancelled(self, candidateKey, reason=''):
        """标记面试已取消"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'interview_cancelled',
                    stopped = 1, stop_reason = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (reason or '候选人未确认面试时间', now, candidateKey),
            )
            self.conn.commit()

    def getInterviewMeta(self, candidateKey):
        """读取候选人面试约定信息"""
        row = self.getCandidate(candidateKey)
        if not row:
            return {}
        return {
            'agreedDate': str(row.get('agreed_date') or ''),
            'agreedTime': str(row.get('agreed_time') or ''),
            'address': str(row.get('interview_address') or ''),
            'jobName': str(row.get('interview_job_name') or ''),
            'preInviteSentAt': str(row.get('pre_invite_sent_at') or ''),
            'interviewAskedAt': str(row.get('interview_asked_at') or ''),
        }

    def hasInterviewRemindAfterPreInvite(self, candidateKey):
        """预邀请发出后是否已发过无回复提醒"""
        meta = self.getInterviewMeta(candidateKey)
        preAt = str(meta.get('preInviteSentAt') or '').strip()
        if not preAt:
            return False
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ?
              AND action_type IN ('interview_remind', 'interview_remind_handoff')
              AND created_at >= ?
            LIMIT 1
            """,
            (candidateKey, preAt),
        ).fetchone()
        return row is not None

    def interviewSlotsKey(self):
        """当日已用面试时段配置键"""
        return f'interview_slots_used_{self.todayText()}'

    def getUsedInterviewSlots(self):
        """读取今日已分配的面试时段整点列表"""
        raw = self.getSetting(self.interviewSlotsKey(), '[]')
        try:
            data = json.loads(raw)
            return [int(x) for x in data if str(x).isdigit()]
        except Exception:
            return []

    def markInterviewSlotUsed(self, hour):
        """记录今日已使用的面试时段"""
        used = self.getUsedInterviewSlots()
        hour = int(hour)
        if hour in used:
            return
        used.append(hour)
        self.setSetting(self.interviewSlotsKey(), json.dumps(used, ensure_ascii=False))

    def canFollowCandidate(self, candidateKey, maxFollowDays=7):
        """判断是否还能继续跟进"""
        candidate = self.getCandidate(candidateKey)
        # 无记录允许首次跟进
        if not candidate:
            return (True, '')
        status = str(candidate.get('resume_status') or 'new')
        # 正式面试邀约已完成则不再跟进
        if status == 'interview_sent':
            return (False, '已发送正式面试邀约')
        # 面试已取消则不再跟进
        if status == 'interview_cancelled':
            return (False, '面试已取消')
        # 已手动或自动停止跟进
        if candidate.get('stopped'):
            return (False, candidate.get('stop_reason') or '已停止跟进')
        # 简历审核未通过则终止
        if status == 'resume_rejected':
            return (False, candidate.get('resume_reject_reason') or '简历未通过')
        # 计算首次联系至今的天数
        firstContact = dt.date.fromisoformat(candidate['first_contact_date'])
        days = (dt.date.today() - firstContact).days
        # 超过最大跟进天数则自动停止
        if days >= maxFollowDays:
            self.stopCandidate(candidateKey, f'超过{maxFollowDays}天未完成流程')
            return (False, f'超过{maxFollowDays}天未完成流程')
        # 仍在跟进窗口内
        return (True, '')

    def stopCandidate(self, candidateKey, reason):
        """停止跟进候选人"""
        with self.writeLock:
            # 标记 stopped 并记录停止原因
            self.conn.execute(
                """
                UPDATE candidate_records
                SET stopped = 1, stop_reason = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (reason, self.nowText(), candidateKey),
            )
            self.conn.commit()

    def markUnsuitable(self, candidateKey, reason):
        """永久标记候选人为不合适并停止自动跟进"""
        with self.writeLock:
            # 不合适状态与停止原因同时写入，后续任务统一拦截
            self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'unsuitable',
                    stopped = 1, stop_reason = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (str(reason or '人工标记不合适'), self.nowText(), candidateKey),
            )
            self.conn.commit()

    def restoreUnsuitable(self, candidateKey):
        """人工解除不合适状态并恢复为待沟通"""
        with self.writeLock:
            # 仅允许恢复 unsuitable，避免覆盖其他已终止业务状态
            cur = self.conn.execute(
                """
                UPDATE candidate_records
                SET resume_status = 'new',
                    stopped = 0, stop_reason = NULL, updated_at = ?
                WHERE candidate_key = ? AND resume_status = 'unsuitable'
                """,
                (self.nowText(), candidateKey),
            )
            self.conn.commit()
            return int(cur.rowcount or 0) > 0

    def getUnsuitableList(self):
        """读取已标记不合适的候选人列表"""
        rows = self.conn.execute(
            """
            SELECT candidate_key, candidate_name, stop_reason, updated_at
            FROM candidate_records
            WHERE resume_status = 'unsuitable' AND stopped = 1
            ORDER BY updated_at DESC
            """
        ).fetchall()
        # 返回普通字典，供 GUI 列表直接使用
        return [dict(row) for row in rows]

    def touchCandidate(self, candidateKey):
        """更新最后联系日期"""
        with self.writeLock:
            # 刷新 last_contact_date 为今天
            self.conn.execute(
                """
                UPDATE candidate_records
                SET last_contact_date = ?, updated_at = ?
                WHERE candidate_key = ?
                """,
                (self.todayText(), self.nowText(), candidateKey),
            )
            self.conn.commit()

    def hasSentToday(self, candidateKey, actionType):
        """判断今天是否已执行某动作"""
        # 按候选人、动作类型、当天日期查重
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ? AND action_type = ? AND action_date = ?
            LIMIT 1
            """,
            (candidateKey, actionType, self.todayText()),
        ).fetchone()
        return row is not None

    def handoffActionType(self, actionType):
        """人工接管动作类型名（不计入每日消息上限）"""
        return f'{actionType}_handoff'

    def hasHandoffToday(self, candidateKey, actionType):
        """判断今天是否已对该候选人填入过待人工发送的话术"""
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ? AND action_type = ? AND action_date = ?
            LIMIT 1
            """,
            (candidateKey, self.handoffActionType(actionType), self.todayText()),
        ).fetchone()
        return row is not None

    def recordHandoff(self, candidateKey, actionType, word=None, taskRunId=None, extra=None):
        """记录话术已填入输入框、待人工发送（不占自动发送配额）"""
        extraJson = json.dumps(extra or {}, ensure_ascii=False)
        with self.writeLock:
            try:
                self.conn.execute(
                    """
                    INSERT INTO chat_actions (
                        candidate_key, action_type, word, success, action_date,
                        created_at, task_run_id, extra_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        candidateKey,
                        self.handoffActionType(actionType),
                        word,
                        1,
                        self.todayText(),
                        self.nowText(),
                        taskRunId,
                        extraJson,
                    ),
                )
                self.conn.commit()
            except sqlite3.IntegrityError:
                pass

    def countTodayMessages(self, actionType=None):
        """统计今日已成功发送的聊天消息条数"""
        if actionType:
            row = self.conn.execute(
                """
                SELECT COUNT(*) AS total FROM chat_actions
                WHERE action_type = ? AND success = 1 AND action_date = ?
                """,
                (actionType, self.todayText()),
            ).fetchone()
        else:
            placeholders = ','.join('?' for _ in self.messageActionTypes)
            row = self.conn.execute(
                f"""
                SELECT COUNT(*) AS total FROM chat_actions
                WHERE action_type IN ({placeholders}) AND success = 1 AND action_date = ?
                """,
                (*self.messageActionTypes, self.todayText()),
            ).fetchone()
        return int(row['total']) if row else 0

    def countTodayRecommend(self):
        """统计今日从推荐牛人成功发出的首次招呼次数"""
        row = self.conn.execute(
            """
            SELECT COUNT(*) AS total FROM chat_actions
            WHERE action_type = 'recommend_greeting'
              AND success = 1 AND action_date = ?
            """,
            (self.todayText(),),
        ).fetchone()
        return int(row['total']) if row else 0

    def hasRecommendContact(self, candidateKey):
        """判断该推荐牛人是否曾经成功主动联系"""
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ?
              AND action_type = 'recommend_greeting'
              AND success = 1
            LIMIT 1
            """,
            (candidateKey,),
        ).fetchone()
        return row is not None

    def canSendMessage(self, actionType, maxPerDay=50, maxPerType=10):
        """检查是否还能自动发送该类型聊天消息"""
        if actionType not in self.messageActionTypes:
            return True, ''
        total = self.countTodayMessages()
        if total >= int(maxPerDay):
            return False, f'已达每日消息总上限 {maxPerDay} 条'
        typeCount = self.countTodayMessages(actionType)
        if typeCount >= int(maxPerType):
            return False, f'已达每日 {actionType} 上限 {maxPerType} 条'
        return True, ''

    def recordRiskEvent(self, eventType):
        """记录当日风控事件并返回累计次数"""
        key = f'risk_{eventType}_{self.todayText()}'
        count = int(self.getSetting(key, '0') or '0') + 1
        self.setSetting(key, str(count))
        return count

    def countTodayRisk(self, eventType):
        """读取当日风控事件次数"""
        key = f'risk_{eventType}_{self.todayText()}'
        return int(self.getSetting(key, '0') or '0')

    def recordAction(self, candidateKey, actionType, word=None, success=True, taskRunId=None, extra=None):
        """记录一次聊天动作"""
        # 扩展信息序列化为 JSON
        extraJson = json.dumps(extra or {}, ensure_ascii=False)
        with self.writeLock:
            try:
                # 插入聊天动作，唯一约束防止同日同类型重复
                self.conn.execute(
                    """
                    INSERT INTO chat_actions (
                        candidate_key, action_type, word, success, action_date,
                        created_at, task_run_id, extra_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        candidateKey,
                        actionType,
                        word,
                        1 if success else 0,
                        self.todayText(),
                        self.nowText(),
                        taskRunId,
                        extraJson,
                    ),
                )
                self.conn.commit()
                return True
            except sqlite3.IntegrityError:
                # 违反唯一约束说明今日已记录，回滚并返回 False
                self.conn.rollback()
                return False

    def normalizeJobProfile(self, profile):
        """将 job.py 的 profile 结构转为 job_rules 表写入格式"""
        # 取出 match 子对象作为筛选条件来源
        match = dict(profile.get('match') or {})
        return {
            'jobName': str(profile.get('jobName') or ''),
            'matchKeys': list(profile.get('matchKeys') or []),
            'jobIntro': str(profile.get('intro') or ''),
            'ageMin': int(match.get('ageMin', 18)),
            'ageMax': int(match.get('ageMax', 45)),
            'educationList': list(match.get('educationList') or []),
            'workYearsMin': int(match.get('workYearsMin', 0)),
            'mustKeywords': list(match.get('mustKeywords') or []),
            'anyKeywords': list(match.get('anyKeywords') or []),
            'preferKeywords': list(match.get('preferKeywords') or []),
            'rejectKeywords': list(match.get('rejectKeywords') or []),
        }

    def countTemplates(self):
        """统计话术模板条数"""
        row = self.conn.execute('SELECT COUNT(*) AS cnt FROM message_templates').fetchone()
        return int(row['cnt'] if row else 0)

    def writeTemplatesFromBundle(self, templateBundle):
        """将 bundle 写入 message_templates（不清空，仅追加插入）"""
        now = self.nowText()
        with self.writeLock:
            for templateType, words in templateBundle.items():
                for word in words:
                    text = str(word or '').strip()
                    # 跳过空话术
                    if not text:
                        continue
                    # 按类型插入一条启用的话术模板
                    self.conn.execute(
                        """
                        INSERT INTO message_templates (template_type, word, enabled, created_at, updated_at)
                        VALUES (?, ?, 1, ?, ?)
                        """,
                        (templateType, text, now, now),
                    )
            self.conn.commit()

    def seedTemplatesIfEmpty(self, templateBundle):
        """首次运行时从 template.py 灌入话术（已有数据不覆盖）"""
        if self.countTemplates() > 0:
            return False
        self.writeTemplatesFromBundle(templateBundle)
        return True

    def reloadTemplatesFromConfig(self, templateBundle):
        """从 template.py 全量覆盖话术表（恢复默认）"""
        with self.writeLock:
            self.conn.execute('DELETE FROM message_templates')
            self.conn.commit()
        self.writeTemplatesFromBundle(templateBundle)

    def reloadTemplatesOfType(self, templateType, words):
        """恢复某一类话术为代码默认值"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute('DELETE FROM message_templates WHERE template_type = ?', (templateType,))
            for word in words or []:
                text = str(word or '').strip()
                if not text:
                    continue
                self.conn.execute(
                    """
                    INSERT INTO message_templates (template_type, word, enabled, created_at, updated_at)
                    VALUES (?, ?, 1, ?, ?)
                    """,
                    (templateType, text, now, now),
                )
            self.conn.commit()

    def countJobRules(self):
        """统计岗位规则条数"""
        row = self.conn.execute('SELECT COUNT(*) AS cnt FROM job_rules').fetchone()
        return int(row['cnt'] if row else 0)

    def insertJobProfile(self, item, enabled=True):
        """插入一条岗位规则（item 为 normalizeJobProfile 结果）"""
        now = self.nowText()
        with self.writeLock:
            self.conn.execute(
                """
                INSERT INTO job_rules (
                    job_name, match_keys_json, job_intro, age_min, age_max, education_json,
                    work_years_min, must_keywords_json, any_keywords_json,
                    prefer_keywords_json, reject_keywords_json, enabled, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item['jobName'],
                    json.dumps(item['matchKeys'], ensure_ascii=False),
                    item['jobIntro'],
                    item['ageMin'],
                    item['ageMax'],
                    json.dumps(item['educationList'], ensure_ascii=False),
                    item['workYearsMin'],
                    json.dumps(item['mustKeywords'], ensure_ascii=False),
                    json.dumps(item['anyKeywords'], ensure_ascii=False),
                    json.dumps(item['preferKeywords'], ensure_ascii=False),
                    json.dumps(item['rejectKeywords'], ensure_ascii=False),
                    1 if enabled else 0,
                    now,
                ),
            )
            self.conn.commit()

    def writeJobProfilesFromBundle(self, profiles):
        """将 job.py profiles 写入 job_rules（不清空，仅追加插入）"""
        for profile in profiles or []:
            item = self.normalizeJobProfile(profile)
            # 无岗位名则跳过
            if not item['jobName']:
                continue
            self.insertJobProfile(item, enabled=True)

    def seedJobRulesIfEmpty(self, profiles):
        """首次运行时从 job.py 灌入岗位规则（已有数据不覆盖）"""
        if self.countJobRules() > 0:
            return False
        self.writeJobProfilesFromBundle(profiles)
        return True

    def reloadJobProfilesFromConfig(self, profiles):
        """从 job.py 全量覆盖岗位规则表（恢复默认）"""
        with self.writeLock:
            self.conn.execute('DELETE FROM job_rules')
            self.conn.commit()
        self.writeJobProfilesFromBundle(profiles)

    def reloadJobProfileFromConfig(self, profile):
        """恢复单个岗位为 job.py 默认值"""
        item = self.normalizeJobProfile(profile)
        jobName = str(item.get('jobName') or '').strip()
        if not jobName:
            return
        with self.writeLock:
            self.conn.execute('DELETE FROM job_rules WHERE job_name = ?', (jobName,))
            self.conn.commit()
        self.insertJobProfile(item, enabled=True)

    def reloadResumeRulesFromConfig(self, globalSettings):
        """从 job.py globalSettings 全量覆盖全局简历规则"""
        settings = dict(globalSettings or self.defaultResumeRules())
        now = self.nowText()
        with self.writeLock:
            self.conn.execute('DELETE FROM resume_rules')
            self.conn.execute(
                """
                INSERT INTO resume_rules (
                    age_min, age_max, education_json, work_years_min,
                    must_keywords_json, reject_keywords_json, interview_time,
                    max_follow_days, chat_interval, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(settings.get('ageMin', 18)),
                    int(settings.get('ageMax', 45)),
                    json.dumps(settings.get('educationList', ['本科', '大专']), ensure_ascii=False),
                    int(settings.get('workYearsMin', 0)),
                    json.dumps(settings.get('mustKeywords', []), ensure_ascii=False),
                    json.dumps(settings.get('rejectKeywords', []), ensure_ascii=False),
                    str(settings.get('interviewTime', '')),
                    int(settings.get('maxFollowDays', 7)),
                    int(settings.get('chatInterval', 15)),
                    now,
                ),
            )
            self.conn.commit()

    def reloadJobFromConfig(self, jobBundle):
        """从 job.py 全量覆盖岗位规则与全局简历规则"""
        profiles = list(jobBundle.get('profiles') or [])
        globalSettings = dict(jobBundle.get('globalSettings') or self.defaultResumeRules())
        self.reloadJobProfilesFromConfig(profiles)
        self.reloadResumeRulesFromConfig(globalSettings)

    def getJobRulesList(self):
        """查询全部岗位规则（含 id 与 enabled）"""
        rows = self.conn.execute('SELECT * FROM job_rules ORDER BY id ASC').fetchall()
        result = []
        for row in rows:
            rules = self.jobRowToRules(row)
            rules['id'] = int(row['id'])
            rules['enabled'] = int(row['enabled'] or 0)
            result.append(rules)
        return result

    def createJobRule(self, data):
        """新增岗位规则"""
        item = self.normalizeJobForm(data)
        if not item['jobName']:
            raise ValueError('岗位名称不能为空')
        now = self.nowText()
        with self.writeLock:
            cur = self.conn.execute(
                """
                INSERT INTO job_rules (
                    job_name, match_keys_json, job_intro, age_min, age_max, education_json,
                    work_years_min, must_keywords_json, any_keywords_json,
                    prefer_keywords_json, reject_keywords_json, enabled, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item['jobName'],
                    json.dumps(item['matchKeys'], ensure_ascii=False),
                    item['jobIntro'],
                    item['ageMin'],
                    item['ageMax'],
                    json.dumps(item['educationList'], ensure_ascii=False),
                    item['workYearsMin'],
                    json.dumps(item['mustKeywords'], ensure_ascii=False),
                    json.dumps(item['anyKeywords'], ensure_ascii=False),
                    json.dumps(item['preferKeywords'], ensure_ascii=False),
                    json.dumps(item['rejectKeywords'], ensure_ascii=False),
                    1 if data.get('enabled', True) else 0,
                    now,
                ),
            )
            self.conn.commit()
            return int(cur.lastrowid)

    def updateJobRule(self, jobId, data):
        """更新岗位规则"""
        item = self.normalizeJobForm(data)
        if not item['jobName']:
            raise ValueError('岗位名称不能为空')
        with self.writeLock:
            self.conn.execute(
                """
                UPDATE job_rules SET
                    job_name = ?, match_keys_json = ?, job_intro = ?,
                    age_min = ?, age_max = ?, education_json = ?,
                    work_years_min = ?, must_keywords_json = ?,
                    any_keywords_json = ?, prefer_keywords_json = ?,
                    reject_keywords_json = ?, enabled = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    item['jobName'],
                    json.dumps(item['matchKeys'], ensure_ascii=False),
                    item['jobIntro'],
                    item['ageMin'],
                    item['ageMax'],
                    json.dumps(item['educationList'], ensure_ascii=False),
                    item['workYearsMin'],
                    json.dumps(item['mustKeywords'], ensure_ascii=False),
                    json.dumps(item['anyKeywords'], ensure_ascii=False),
                    json.dumps(item['preferKeywords'], ensure_ascii=False),
                    json.dumps(item['rejectKeywords'], ensure_ascii=False),
                    1 if data.get('enabled', True) else 0,
                    self.nowText(),
                    jobId,
                ),
            )
            self.conn.commit()

    def deleteJobRule(self, jobId):
        """删除岗位规则"""
        with self.writeLock:
            self.conn.execute('DELETE FROM job_rules WHERE id = ?', (jobId,))
            self.conn.commit()

    def normalizeJobForm(self, data):
        """将 GUI/表单数据转为岗位规则写入格式"""
        return {
            'jobName': str(data.get('jobName') or '').strip(),
            'matchKeys': list(data.get('matchKeys') or []),
            'jobIntro': str(data.get('jobIntro') or '').strip(),
            'ageMin': int(data.get('ageMin') or 18),
            'ageMax': int(data.get('ageMax') or 45),
            'educationList': list(data.get('educationList') or []),
            'workYearsMin': int(data.get('workYearsMin') or 0),
            'mustKeywords': list(data.get('mustKeywords') or []),
            'anyKeywords': list(data.get('anyKeywords') or []),
            'preferKeywords': list(data.get('preferKeywords') or []),
            'rejectKeywords': list(data.get('rejectKeywords') or []),
        }

    def reloadFromConfig(self, templateBundle, jobBundle):
        """从 template.py / job.py 全量覆盖灌库（清空后重写话术与岗位规则）"""
        self.reloadTemplatesFromConfig(templateBundle)
        self.reloadJobFromConfig(jobBundle)

    def hasIntroReply(self, candidateKey):
        """是否已发送过岗位介绍或智能回复（含人工接管填框）"""
        row = self.conn.execute(
            """
            SELECT 1 FROM chat_actions
            WHERE candidate_key = ? AND action_type IN (
                'greeting', 'smart_reply', 'greeting_handoff', 'smart_reply_handoff', 'manual_reply'
            )
            LIMIT 1
            """,
            (candidateKey,),
        ).fetchone()
        return row is not None

    def createTemplate(self, templateType, word, enabled=True):
        """新增话术模板"""
        now = self.nowText()
        with self.writeLock:
            # 插入新话术并返回自增 ID
            cur = self.conn.execute(
                """
                INSERT INTO message_templates (template_type, word, enabled, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (templateType, word, 1 if enabled else 0, now, now),
            )
            self.conn.commit()
            return int(cur.lastrowid)

    def updateTemplate(self, templateId, templateType, word, enabled):
        """更新话术模板"""
        with self.writeLock:
            # 按 ID 更新类型、内容与启用状态
            self.conn.execute(
                """
                UPDATE message_templates
                SET template_type = ?, word = ?, enabled = ?, updated_at = ?
                WHERE id = ?
                """,
                (templateType, word, 1 if enabled else 0, self.nowText(), templateId),
            )
            self.conn.commit()

    def deleteTemplate(self, templateId):
        """删除话术模板"""
        with self.writeLock:
            # 按 ID 删除单条话术
            self.conn.execute('DELETE FROM message_templates WHERE id = ?', (templateId,))
            self.conn.commit()

    def getTemplates(self, templateType=None, enabledOnly=False):
        """查询话术模板"""
        conditions = []
        params = []
        # 按模板类型过滤
        if templateType:
            conditions.append('template_type = ?')
            params.append(templateType)
        # 仅查启用项
        if enabledOnly:
            conditions.append('enabled = 1')
        # 动态拼接 WHERE 子句
        whereSql = f"WHERE {' AND '.join(conditions)}" if conditions else ''
        # 按类型与 ID 排序查询
        rows = self.conn.execute(
            f"""
            SELECT id, template_type, word, enabled, updated_at
            FROM message_templates {whereSql}
            ORDER BY template_type ASC, id ASC
            """,
            params,
        ).fetchall()
        # Row 转 dict 列表返回
        return [dict(row) for row in rows]

    def seedResumeRules(self, rules):
        """首次启动写入默认简历规则"""
        # 已有规则则不再写入
        row = self.conn.execute('SELECT COUNT(*) AS total FROM resume_rules').fetchone()
        if row and int(row['total']) > 0:
            return
        self.saveResumeRules(rules)

    def getResumeRules(self):
        """读取简历筛选规则"""
        # 取最新一条全局简历规则
        row = self.conn.execute('SELECT * FROM resume_rules ORDER BY id DESC LIMIT 1').fetchone()
        # 无记录返回内置默认
        if not row:
            return self.defaultResumeRules()
        data = dict(row)
        # JSON 字段反序列化为 Python 结构
        return {
            'ageMin': int(data['age_min']),
            'ageMax': int(data['age_max']),
            'educationList': json.loads(data['education_json']),
            'workYearsMin': int(data['work_years_min']),
            'mustKeywords': json.loads(data['must_keywords_json']),
            'rejectKeywords': json.loads(data['reject_keywords_json']),
            'interviewTime': data['interview_time'],
            'maxFollowDays': int(data['max_follow_days']),
            'chatInterval': int(data['chat_interval']),
        }

    def saveResumeRules(self, rules):
        """保存简历筛选规则"""
        now = self.nowText()
        with self.writeLock:
            # 全量替换：先删后插，保证只有一条有效规则
            self.conn.execute('DELETE FROM resume_rules')
            self.conn.execute(
                """
                INSERT INTO resume_rules (
                    age_min, age_max, education_json, work_years_min,
                    must_keywords_json, reject_keywords_json, interview_time,
                    max_follow_days, chat_interval, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(rules.get('ageMin', 18)),
                    int(rules.get('ageMax', 45)),
                    json.dumps(rules.get('educationList', ['本科', '大专']), ensure_ascii=False),
                    int(rules.get('workYearsMin', 0)),
                    json.dumps(rules.get('mustKeywords', []), ensure_ascii=False),
                    json.dumps(rules.get('rejectKeywords', []), ensure_ascii=False),
                    str(rules.get('interviewTime', '')),
                    int(rules.get('maxFollowDays', 7)),
                    int(rules.get('chatInterval', 15)),
                    now,
                ),
            )
            self.conn.commit()

    def jobRowToRules(self, row):
        """岗位要求行转匹配规则字典"""
        data = dict(row)
        eduList = json.loads(data.get('education_json') or '[]')
        # 将表字段映射为 match 模块使用的 camelCase 结构
        return {
            'jobName': str(data.get('job_name') or ''),
            'jobIntro': str(data.get('job_intro') or ''),
            'matchKeys': json.loads(data.get('match_keys_json') or '[]'),
            'ageMin': int(data.get('age_min') or 18),
            'ageMax': int(data.get('age_max') or 45),
            'educationList': list(eduList),
            'workYearsMin': int(data.get('work_years_min') or 0),
            'mustKeywords': json.loads(data.get('must_keywords_json') or '[]'),
            'anyKeywords': json.loads(data.get('any_keywords_json') or '[]'),
            'preferKeywords': json.loads(data.get('prefer_keywords_json') or '[]'),
            'rejectKeywords': json.loads(data.get('reject_keywords_json') or '[]'),
        }

    def normalizeJobText(self, text):
        """岗位名标准化，便于模糊匹配"""
        # 去空白并转小写
        return re.sub('\\s+', '', str(text or '')).lower()

    def matchJobRules(self, sourceJob):
        """按沟通列表岗位名匹配岗位要求"""
        sourceNorm = self.normalizeJobText(sourceJob)
        # 空岗位名无法匹配
        if not sourceNorm:
            return None
        # 只查启用的岗位规则，按 ID 顺序优先
        rows = self.conn.execute('SELECT * FROM job_rules WHERE enabled = 1 ORDER BY id ASC').fetchall()
        for row in rows:
            rules = self.jobRowToRules(row)
            nameNorm = self.normalizeJobText(rules['jobName'])
            # 岗位名与来源互相包含则命中
            if nameNorm and (nameNorm in sourceNorm or sourceNorm in nameNorm):
                return rules
            # 再按 matchKeys 别名逐个匹配
            for key in rules.get('matchKeys') or []:
                keyNorm = self.normalizeJobText(key)
                if keyNorm and keyNorm in sourceNorm:
                    return rules
        # 无规则命中
        return None

    def defaultResumeRules(self):
        """返回内置默认规则"""
        return {
            'ageMin': 18,
            'ageMax': 45,
            'educationList': ['本科', '大专'],
            'workYearsMin': 0,
            'mustKeywords': [],
            'rejectKeywords': [],
            'interviewTime': '明天下午14:00',
            'maxFollowDays': 7,
            'chatInterval': 15,
        }

    def seedReply(self, defaultSkill):
        """首次启动写入默认回复 Skill 与本地模型配置"""
        rows = self.getReplySkills()
        # Skill 表为空时写入代码内置的稳健回复规则
        if not rows:
            skillId = self.createReplySkill(defaultSkill)
        else:
            skillId = int(rows[0]['id'])
        self.seedSetting(self.replyEnabledKey, '1')
        self.seedSetting(self.replyBaseUrlKey, self.defaultReplyBaseUrl)
        self.seedSetting(self.replyModelNameKey, self.defaultReplyModelName)
        self.seedSetting(self.replyTimeoutKey, str(self.defaultReplyTimeout))
        # 活动 Skill 缺失或已删除时切换到第一条可用 Skill
        current = self.getSetting(self.replyActiveSkillKey, '')
        try:
            currentId = int(current)
        except (TypeError, ValueError):
            currentId = 0
        if not self.getReplySkill(currentId):
            self.setSetting(self.replyActiveSkillKey, str(skillId))
        return skillId

    def getReplySettings(self):
        """读取本地回复模型配置与当前活动 Skill"""
        try:
            timeoutSec = max(5, int(self.getSetting(self.replyTimeoutKey, self.defaultReplyTimeout)))
        except (TypeError, ValueError):
            timeoutSec = self.defaultReplyTimeout
        try:
            activeSkillId = int(self.getSetting(self.replyActiveSkillKey, '0') or 0)
        except (TypeError, ValueError):
            activeSkillId = 0
        skill = self.getReplySkill(activeSkillId)
        # 当前 Skill 不可用时回退到第一条已启用 Skill
        if not skill or not skill.get('enabled'):
            rows = self.getReplySkills(enabledOnly=True)
            skill = rows[0] if rows else None
            activeSkillId = int(skill['id']) if skill else 0
        return {
            'enabled': self.getSetting(self.replyEnabledKey, '1') == '1',
            'baseUrl': self.getSetting(self.replyBaseUrlKey, self.defaultReplyBaseUrl),
            'modelName': self.getSetting(self.replyModelNameKey, self.defaultReplyModelName),
            'timeoutSec': timeoutSec,
            'activeSkillId': activeSkillId,
            'skill': skill,
        }

    def saveReplySettings(self, data):
        """保存本地回复模型地址、模型名、超时和活动 Skill"""
        payload = dict(data or {})
        baseUrl = str(payload.get('baseUrl') or '').strip().rstrip('/')
        modelName = str(payload.get('modelName') or '').strip()
        if not baseUrl:
            raise ValueError('模型 API 地址不能为空')
        if not modelName:
            raise ValueError('模型名称不能为空')
        try:
            timeoutSec = max(5, int(payload.get('timeoutSec') or self.defaultReplyTimeout))
        except (TypeError, ValueError):
            raise ValueError('模型超时秒数必须是整数')
        try:
            activeSkillId = int(payload.get('activeSkillId') or 0)
        except (TypeError, ValueError):
            activeSkillId = 0
        if activeSkillId and not self.getReplySkill(activeSkillId):
            raise ValueError('选择的回复 Skill 不存在')
        self.setSetting(self.replyEnabledKey, '1' if payload.get('enabled') else '0')
        self.setSetting(self.replyBaseUrlKey, baseUrl)
        self.setSetting(self.replyModelNameKey, modelName)
        self.setSetting(self.replyTimeoutKey, str(timeoutSec))
        if activeSkillId:
            self.setSetting(self.replyActiveSkillKey, str(activeSkillId))
        return self.getReplySettings()

    def getReplySkills(self, enabledOnly=False):
        """读取回复 Skill 列表"""
        sql = 'SELECT id, skill_name, instruction, examples, enabled, created_at, updated_at FROM reply_skills'
        params = []
        if enabledOnly:
            sql += ' WHERE enabled = ?'
            params.append(1)
        sql += ' ORDER BY id ASC'
        rows = self.conn.execute(sql, params).fetchall()
        return [dict(row) for row in rows]

    def getReplySkill(self, skillId):
        """按编号读取单个回复 Skill"""
        if not skillId:
            return None
        row = self.conn.execute(
            'SELECT id, skill_name, instruction, examples, enabled, created_at, updated_at FROM reply_skills WHERE id = ?',
            (skillId,),
        ).fetchone()
        return self.rowToDict(row)

    def createReplySkill(self, data):
        """创建新的可切换回复 Skill"""
        payload = dict(data or {})
        skillName = str(payload.get('skillName') or '').strip()
        instruction = str(payload.get('instruction') or '').strip()
        examples = str(payload.get('examples') or '').strip()
        if not skillName:
            raise ValueError('Skill 名称不能为空')
        if not instruction:
            raise ValueError('Skill 规则不能为空')
        now = self.nowText()
        with self.writeLock:
            # 写入 Skill 规则与案例文本
            cursor = self.conn.execute(
                """
                INSERT INTO reply_skills
                (skill_name, instruction, examples, enabled, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (skillName, instruction, examples, 1 if payload.get('enabled', True) else 0, now, now),
            )
            self.conn.commit()
        return int(cursor.lastrowid)

    def updateReplySkill(self, skillId, data):
        """更新回复 Skill 的名称、规则、案例和启用状态"""
        payload = dict(data or {})
        skillName = str(payload.get('skillName') or '').strip()
        instruction = str(payload.get('instruction') or '').strip()
        examples = str(payload.get('examples') or '').strip()
        if not skillName:
            raise ValueError('Skill 名称不能为空')
        if not instruction:
            raise ValueError('Skill 规则不能为空')
        with self.writeLock:
            # 覆盖所选 Skill 内容并保留原编号
            self.conn.execute(
                """
                UPDATE reply_skills
                SET skill_name = ?, instruction = ?, examples = ?, enabled = ?, updated_at = ?
                WHERE id = ?
                """,
                (skillName, instruction, examples, 1 if payload.get('enabled', True) else 0, self.nowText(), skillId),
            )
            self.conn.commit()

    def deleteReplySkill(self, skillId):
        """删除非最后一条回复 Skill，并修复活动 Skill 指向"""
        rows = self.getReplySkills()
        if len(rows) <= 1:
            raise ValueError('至少保留一条回复 Skill')
        with self.writeLock:
            # 删除所选 Skill 后提交，反馈记录保留但解除引用会受外键保护
            used = self.conn.execute('SELECT COUNT(*) AS total FROM reply_feedback WHERE skill_id = ?', (skillId,)).fetchone()
            if used and int(used['total'] or 0) > 0:
                self.conn.execute('UPDATE reply_feedback SET skill_id = NULL WHERE skill_id = ?', (skillId,))
            self.conn.execute('DELETE FROM reply_skills WHERE id = ?', (skillId,))
            self.conn.commit()
        current = str(self.getSetting(self.replyActiveSkillKey, '') or '')
        if current == str(skillId):
            fallback = self.getReplySkills(enabledOnly=True) or self.getReplySkills()
            self.setSetting(self.replyActiveSkillKey, str(fallback[0]['id']))

    def saveReplyFeedback(self, data):
        """记录模型建议和人工最终填入内容，供后续优化 Skill"""
        payload = dict(data or {})
        with self.writeLock:
            # 保存原始建议、人工修改结果和是否采用
            self.conn.execute(
                """
                INSERT INTO reply_feedback
                (candidate_key, skill_id, job_name, friend_text, suggested_reply,
                 final_reply, recommendation, accepted, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(payload.get('candidateKey') or ''),
                    payload.get('skillId') or None,
                    str(payload.get('jobName') or ''),
                    str(payload.get('friendText') or ''),
                    str(payload.get('suggestedReply') or ''),
                    str(payload.get('finalReply') or ''),
                    str(payload.get('recommendation') or 'reply_only'),
                    1 if payload.get('accepted') else 0,
                    self.nowText(),
                ),
            )
            self.conn.commit()

    def getSetting(self, key, default=''):
        """读取配置项"""
        row = self.conn.execute(
            'SELECT setting_value FROM app_settings WHERE setting_key = ?',
            (key,),
        ).fetchone()
        return str(row['setting_value']) if row else default

    def setSetting(self, key, value):
        """写入配置项"""
        with self.writeLock:
            # upsert：存在则更新，不存在则插入
            self.conn.execute(
                """
                INSERT INTO app_settings (setting_key, setting_value, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(setting_key) DO UPDATE SET
                    setting_value = excluded.setting_value,
                    updated_at = excluded.updated_at
                """,
                (key, value, self.nowText()),
            )
            self.conn.commit()

    def seedSetting(self, key, value):
        """仅在空值时写入默认配置"""
        # 已有值则不覆盖
        if self.getSetting(key):
            return
        self.setSetting(key, value)

    def seedWecomSettings(self, webhookDefault, mobileDefault):
        """首次运行时写入企微默认配置（已有缓存不覆盖）"""
        self.seedSetting(self.wecomWebhookKey, str(webhookDefault or '').strip())
        self.seedSetting(self.wecomMentionMobileKey, str(mobileDefault or '').strip())

    def seedRecommendLimit(self):
        """首次运行时写入推荐牛人每日主动联系默认上限"""
        self.seedSetting(self.recommendLimitKey, str(self.defaultRecommendLimit))

    def getRecommendLimit(self):
        """读取推荐牛人每日主动联系上限"""
        raw = self.getSetting(self.recommendLimitKey, str(self.defaultRecommendLimit))
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return self.defaultRecommendLimit

    def saveRecommendLimit(self, dailyLimit):
        """保存推荐牛人每日主动联系上限"""
        limit = max(1, int(dailyLimit))
        # 单独配置主动联系额度，不影响普通聊天消息限额
        self.setSetting(self.recommendLimitKey, str(limit))
        return limit

    def getWecomWebhook(self, fallback=''):
        """读取缓存的企微 Webhook"""
        return str(self.getSetting(self.wecomWebhookKey, fallback) or '').strip()

    def getWecomMentionMobile(self, fallback=''):
        """读取缓存的企微 @ 手机号"""
        return str(self.getSetting(self.wecomMentionMobileKey, fallback) or '').strip()

    def saveWecomSettings(self, webhookUrl, mentionMobile):
        """保存企微 Webhook 与 @ 手机号到本地缓存"""
        self.setSetting(self.wecomWebhookKey, str(webhookUrl or '').strip())
        self.setSetting(self.wecomMentionMobileKey, str(mentionMobile or '').strip())

    def defaultPlaceholderConfig(self, jobGlobalSettings=None):
        """话术占位符默认取值（与 job.py globalSettings 对齐）"""
        gs = dict(jobGlobalSettings or {})
        slots = gs.get('interviewTimeSlots') or [10, 11, 13, 14, 15, 16, 17]
        return {
            'company': '四川伯尼森科技有限公司',
            'jobDefault': '',
            'nameDefault': '',
            'address': str(gs.get('interviewAddress') or ''),
            'duration': str(gs.get('interviewDuration') or '35-45'),
            'dayOffset': int(gs.get('interviewDayOffset') or 1),
            'timeSlots': ','.join(str(x) for x in slots),
        }

    def getPlaceholderConfig(self, jobGlobalSettings=None):
        """读取话术占位符配置，缺失字段用默认值补齐"""
        base = self.defaultPlaceholderConfig(jobGlobalSettings)
        raw = self.getSetting(self.templatePlaceholderKey, '')
        if not raw:
            return base
        try:
            saved = json.loads(raw)
        except Exception:
            return base
        if not isinstance(saved, dict):
            return base
        for key in base:
            if key in saved and saved[key] is not None and str(saved[key]).strip() != '':
                base[key] = saved[key]
        return base

    def savePlaceholderConfig(self, data):
        """保存话术占位符配置到 app_settings"""
        payload = dict(data or {})
        self.setSetting(self.templatePlaceholderKey, json.dumps(payload, ensure_ascii=False))

    def seedPlaceholderConfigIfEmpty(self, jobGlobalSettings=None):
        """首次运行时写入占位符默认配置（已有缓存不覆盖）"""
        if self.getSetting(self.templatePlaceholderKey):
            return
        self.savePlaceholderConfig(self.defaultPlaceholderConfig(jobGlobalSettings))

    def getBrowsers(self):
        """读取浏览器配置列表"""
        rows = self.conn.execute(
            """
            SELECT browser_id, remark, user_data_path, created_at, updated_at
            FROM browser_profiles ORDER BY created_at ASC
            """
        ).fetchall()
        return [dict(row) for row in rows]

    def getBrowser(self, browserId):
        """读取单个浏览器配置"""
        row = self.conn.execute(
            'SELECT * FROM browser_profiles WHERE browser_id = ?',
            (browserId,),
        ).fetchone()
        return self.rowToDict(row)

    def createBrowser(self, browserId, remark, userDataPath=''):
        """新增浏览器配置"""
        now = self.nowText()
        with self.writeLock:
            # 插入浏览器 profile 记录
            self.conn.execute(
                """
                INSERT INTO browser_profiles (browser_id, remark, user_data_path, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (browserId, remark, userDataPath, now, now),
            )
            self.conn.commit()

    def updateBrowser(self, browserId, newBrowserId, remark, userDataPath):
        """更新浏览器配置"""
        with self.writeLock:
            # 支持修改 browser_id 本身（主键变更）
            self.conn.execute(
                """
                UPDATE browser_profiles
                SET browser_id = ?, remark = ?, user_data_path = ?, updated_at = ?
                WHERE browser_id = ?
                """,
                (newBrowserId, remark, userDataPath, self.nowText(), browserId),
            )
            self.conn.commit()

    def deleteBrowser(self, browserId):
        """删除浏览器配置"""
        with self.writeLock:
            self.conn.execute('DELETE FROM browser_profiles WHERE browser_id = ?', (browserId,))
            self.conn.commit()

    def seedBrowser(self, browserId, remark='', userDataPath=''):
        """首次启动写入默认浏览器"""
        row = self.conn.execute('SELECT COUNT(*) AS total FROM browser_profiles').fetchone()
        # 已有浏览器配置则跳过
        if row and int(row['total']) > 0:
            return
        if browserId.strip():
            self.createBrowser(browserId.strip(), remark or browserId.strip(), userDataPath)

    def getDbSnapshot(self):
        """汇总数据记录供前端展示"""
        return {
            'taskRuns': self.fetchTaskRuns(),
            'todayActions': self.fetchTodayActions(),
            'resumeReviews': self.fetchResumeReviews(),
            'stoppedCandidates': self.fetchStoppedCandidates(),
        }

    def fetchTaskRuns(self, limit=50):
        """最近任务记录"""
        rows = self.conn.execute(
            """
            SELECT id, browser_id, task_type, status, started_at, finished_at, error
            FROM task_runs ORDER BY id DESC LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]

    def fetchTodayActions(self, limit=100):
        """今日动作记录"""
        rows = self.conn.execute(
            """
            SELECT a.id, c.candidate_name, a.action_type, a.word, a.success, a.created_at
            FROM chat_actions a
            LEFT JOIN candidate_records c ON c.candidate_key = a.candidate_key
            WHERE a.action_date = ? ORDER BY a.id DESC LIMIT ?
            """,
            (self.todayText(), limit),
        ).fetchall()
        return [dict(row) for row in rows]

    def fetchResumeReviews(self, limit=100):
        """简历审核记录"""
        rows = self.conn.execute(
            """
            SELECT candidate_name, resume_status, resume_pass, resume_reject_reason,
                   interview_sent_at, resume_received_at, updated_at
            FROM candidate_records
            WHERE resume_status NOT IN ('new')
            ORDER BY updated_at DESC LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]

    def countTodayActionsByType(self):
        """统计今日各动作类型次数"""
        rows = self.conn.execute(
            """
            SELECT action_type, success, COUNT(*) AS total
            FROM chat_actions
            WHERE action_date = ?
            GROUP BY action_type, success
            """,
            (self.todayText(),),
        ).fetchall()
        stats = {}
        for row in rows:
            key = str(row['action_type'] or '')
            if key not in stats:
                stats[key] = {'total': 0, 'success': 0, 'fail': 0}
            count = int(row['total'] or 0)
            stats[key]['total'] += count
            if int(row['success'] or 0):
                stats[key]['success'] += count
            else:
                stats[key]['fail'] += count
        return stats

    def countTodayFormalInterviews(self):
        """统计今日已完成正式面试邀约人数"""
        prefix = self.todayText() + '%'
        row = self.conn.execute(
            """
            SELECT COUNT(*) AS total FROM candidate_records
            WHERE resume_status = 'interview_sent' AND interview_sent_at LIKE ?
            """,
            (prefix,),
        ).fetchone()
        return int(row['total']) if row else 0

    def fetchTodayInterviewSent(self):
        """读取今日已完成正式面试邀约的候选人列表"""
        prefix = self.todayText() + '%'
        rows = self.conn.execute(
            """
            SELECT candidate_name, resume_json, interview_job_name,
                   agreed_date, agreed_time, interview_address, interview_sent_at
            FROM candidate_records
            WHERE resume_status = 'interview_sent' AND interview_sent_at LIKE ?
            ORDER BY interview_sent_at ASC
            """,
            (prefix,),
        ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            resume = {}
            try:
                resume = json.loads(str(item.get('resume_json') or '{}'))
            except Exception:
                resume = {}
            item['age'] = resume.get('age')
            item['contact'] = str(resume.get('contact') or '').strip()
            item['resumeName'] = str(resume.get('name') or '').strip()
            result.append(item)
        return result

    def fetchStoppedCandidates(self, limit=100):
        """已停止跟进列表"""
        rows = self.conn.execute(
            """
            SELECT candidate_name, resume_status, stop_reason, updated_at
            FROM candidate_records WHERE stopped = 1
            ORDER BY updated_at DESC LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]


if __name__ == '__main__':
    # 本文件独立调试配置
    config = {'dbPath': ''}
    from boss_web.job import BossJob
    from boss_web.template import BossTemplate

    # 加载话术与岗位配置 bundle
    tpl = BossTemplate()
    job = BossJob()
    # 连接数据库，空路径则用默认 data 目录
    db = BossDb(config['dbPath'] or None)
    # 补全旧库表结构
    db.migrateSchema()
    # 从 template.py / job.py 全量灌库
    db.reloadFromConfig(tpl.bundle(), job.bundle())
    # 首次启动写入默认浏览器与当前 browserId 配置
    db.seedBrowser('default-browser', '默认Chrome配置', '')
    db.seedSetting('browserId', 'default-browser')
    # 打印快照验证读写
    print(db.getDbSnapshot())
