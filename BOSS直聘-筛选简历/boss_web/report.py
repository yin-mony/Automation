import datetime as dt
import re
import time
import requests


class BossReport:
    """BOSS 招聘日报：汇总当日沟通并分批推送企业微信"""

    def __init__(self, database=None):
        self.db = database
        self.webhookUrl = ''
        self.mentionMobiles = []
        self.batchSize = 10
        self.pushGapSec = 1.5
        # 动作类型中文名，用于日报汇总
        self.actionLabels = {
            'greeting': '岗位介绍',
            'smart_reply': '智能回复',
            'request_resume': '求简历',
            'followup': '跟进话术',
            'accept_resume': '同意收简历',
            'preview_resume': '预览简历',
            'resume_review': '简历审核',
            'interview_pre': '面试预邀请',
            'interview_ask_time': '追问面试时间',
            'interview_remind': '面试无回复提醒',
            'interview_reschedule': '改期确认',
            'interview_cancel': '取消面试',
            'reject_notify': '未通过通知',
        }

    def parseMentionMobiles(self, raw):
        """解析 @ 手机号列表（逗号/空格分隔）"""
        parts = re.split(r'[,，\s]+', str(raw or ''))
        return [part.strip() for part in parts if part.strip()]

    def loadSettings(self, settings):
        """从配置加载企微推送参数"""
        cfg = dict(settings or {})
        self.webhookUrl = str(cfg.get('wecomWebhookUrl') or '').strip()
        self.mentionMobiles = self.parseMentionMobiles(cfg.get('wecomMentionMobile'))
        self.batchSize = max(1, int(cfg.get('wecomInterviewBatchSize') or 10))
        self.pushGapSec = float(cfg.get('wecomPushGapSec') or 1.5)

    def todayLabel(self):
        """今日日期展示文案"""
        today = dt.date.today()
        weekNames = '一二三四五六日'
        return f'{today.year}年{today.month}月{today.day}日（周{weekNames[today.weekday()]}）'

    def collectTodayStats(self):
        """汇总今日沟通统计数据"""
        if not self.db:
            return {'actionStats': {}, 'formalInterviewCount': 0}
        actionStats = self.db.countTodayActionsByType()
        formalCount = self.db.countTodayFormalInterviews()
        return {'actionStats': actionStats, 'formalInterviewCount': formalCount}

    def buildSummaryMarkdown(self, stats):
        """生成今日沟通汇总 Markdown"""
        dateText = self.todayLabel()
        lines = [f'## BOSS 今日沟通统计（{dateText}）', '']
        actionStats = stats.get('actionStats') or {}
        order = [
            'greeting', 'smart_reply', 'request_resume', 'followup',
            'accept_resume', 'preview_resume', 'resume_review',
            'interview_pre', 'interview_ask_time', 'interview_remind',
            'interview_reschedule', 'interview_cancel', 'reject_notify',
        ]
        used = set()
        for actionType in order:
            if actionType not in actionStats:
                continue
            used.add(actionType)
            label = self.actionLabels.get(actionType, actionType)
            data = actionStats[actionType]
            if actionType == 'resume_review':
                lines.append(f'- {label}：通过 {data.get("success", 0)} / 不通过 {data.get("fail", 0)}')
            else:
                lines.append(f'- {label}：{data.get("total", 0)}')
        for actionType, data in actionStats.items():
            if actionType in used:
                continue
            label = self.actionLabels.get(actionType, actionType)
            lines.append(f'- {label}：{data.get("total", 0)}')
        lines.append(f'- 正式面试邀约完成：{stats.get("formalInterviewCount", 0)}')
        return '\n'.join(lines)

    def formatAge(self, age):
        """格式化年龄展示"""
        if age is None or age == '':
            return '未知'
        return f'{age}岁'

    def formatContact(self, contact):
        """格式化联系方式展示"""
        text = str(contact or '').strip()
        return text or '未解析到'

    def buildInterviewBatchMarkdown(self, rows, batchIndex, batchTotal, totalCount):
        """生成单批面试邀约明细 Markdown"""
        if not rows:
            return '## 今日已邀约面试\n\n今日暂无已完成正式面试邀约的候选人。'
        header = f'## 今日已邀约面试（{batchIndex}/{batchTotal}）共 {totalCount} 人'
        lines = [header, '']
        for idx, row in enumerate(rows, start=1 + (batchIndex - 1) * self.batchSize):
            name = str(row.get('candidate_name') or row.get('resumeName') or '未知').strip()
            age = self.formatAge(row.get('age'))
            contact = self.formatContact(row.get('contact'))
            job = str(row.get('interview_job_name') or '未知岗位').strip()
            dateText = str(row.get('agreed_date') or '').strip()
            timeText = str(row.get('agreed_time') or '').strip()
            address = str(row.get('interview_address') or '').strip()
            when = ' '.join(part for part in [dateText, timeText] if part) or '待定'
            addr = address or '待定'
            lines.append(f'{idx}. **{name}** | {age} | {contact} | {job}')
            lines.append(f'   时间：{when} | 地址：{addr}')
            lines.append('')
        return '\n'.join(lines).strip()

    def postWebhook(self, body):
        """向企业微信群机器人 POST 消息"""
        url = str(self.webhookUrl or '').strip()
        if not url:
            raise RuntimeError('未配置企业微信 Webhook URL')
        res = requests.post(url, json=body, timeout=15)
        res.raise_for_status()
        data = res.json()
        if int(data.get('errcode', -1)) != 0:
            raise RuntimeError(f'企业微信推送失败: {data.get("errmsg") or data}')
        return data

    def pushMarkdown(self, content):
        """向企业微信群推送一条 Markdown 消息"""
        body = {'msgtype': 'markdown', 'markdown': {'content': str(content or '')}}
        return self.postWebhook(body)

    def pushText(self, content, mentionMobiles=None):
        """向企业微信群推送文本消息，可选 @ 手机号"""
        mobiles = mentionMobiles if mentionMobiles is not None else self.mentionMobiles
        textBody = {'content': str(content or '')}
        if mobiles:
            textBody['mentioned_mobile_list'] = mobiles
        body = {'msgtype': 'text', 'text': textBody}
        return self.postWebhook(body)

    def pushMentionNotice(self):
        """日报推送完成后 @ 指定企微手机号"""
        if not self.mentionMobiles:
            return None
        dateText = self.todayLabel()
        content = f'BOSS 今日日报（{dateText}）已推送，请查阅'
        return self.pushText(content, self.mentionMobiles)

    def pushToday(self):
        """推送今日日报：先汇总，再分批推送面试明细，最后 @ 通知"""
        if not self.db:
            raise RuntimeError('数据库未初始化')
        stats = self.collectTodayStats()
        summary = self.buildSummaryMarkdown(stats)
        sent = 0
        self.pushMarkdown(summary)
        sent += 1
        time.sleep(self.pushGapSec)
        rows = self.db.fetchTodayInterviewSent()
        total = len(rows)
        if total == 0:
            self.pushMarkdown(self.buildInterviewBatchMarkdown([], 1, 1, 0))
            sent += 1
        else:
            batchCount = (total + self.batchSize - 1) // self.batchSize
            for batchIndex in range(batchCount):
                start = batchIndex * self.batchSize
                batch = rows[start:start + self.batchSize]
                text = self.buildInterviewBatchMarkdown(batch, batchIndex + 1, batchCount, total)
                self.pushMarkdown(text)
                sent += 1
                if batchIndex < batchCount - 1:
                    time.sleep(self.pushGapSec)
        if self.mentionMobiles:
            time.sleep(self.pushGapSec)
            self.pushMentionNotice()
            sent += 1
        batchCount = 1 if total == 0 else (total + self.batchSize - 1) // self.batchSize
        return {'sent': sent, 'interviewTotal': total, 'batchCount': batchCount}


if __name__ == '__main__':
    # 独立调试：构造样例数据预览 Markdown（不实际推送）
    config = {
        'wecomWebhookUrl': '',
        'wecomMentionMobile': '13800138000',
        'wecomInterviewBatchSize': 10,
        'wecomPushGapSec': 1.5,
    }
    report = BossReport()
    report.loadSettings(config)
    sampleStats = {
        'actionStats': {
            'greeting': {'total': 2, 'success': 2, 'fail': 0},
            'resume_review': {'total': 3, 'success': 2, 'fail': 1},
        },
        'formalInterviewCount': 2,
    }
    print(report.buildSummaryMarkdown(sampleStats))
    sampleRows = [
        {
            'candidate_name': '张三',
            'age': 28,
            'contact': '13812345678',
            'interview_job_name': '后端开发工程师',
            'agreed_date': '7月1日（周二）',
            'agreed_time': '下午2点',
            'interview_address': '深圳市南山区',
        },
    ]
    print('---')
    print(report.buildInterviewBatchMarkdown(sampleRows, 1, 1, 1))
