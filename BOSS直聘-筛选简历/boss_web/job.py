import os


class BossJob:
    """岗位默认配置：职责介绍 intro + 简历筛选 match，启动时灌入 job_rules 表"""

    def __init__(self):
        # 全局默认筛选参数，写入 resume_rules 表
        self.globalSettings = {
            'ageMin': 18,
            'ageMax': 45,
            'educationList': ['本科', '大专'],
            'workYearsMin': 0,
            'mustKeywords': [],
            'rejectKeywords': [],
            'interviewTime': '明天下午14:00',
            # 面试日期相对运行日的偏移：0=当天 1=明天
            'interviewDayOffset': 1,
            # 预邀请可选时段（24 小时制整点）
            'interviewTimeSlots': [10, 11, 13, 14, 15, 16, 17],
            # 当天尽量不重复同一时段
            'interviewTimeSpread': True,
            'interviewAddress': '四川省成都市成华区亚太并购大厦15楼',
            'interviewDuration': '40-60',
            # 预邀请无回复超过此小时数发追问
            'interviewNoReplyHours': 1,
            # 预邀请且已发 1 次提醒后仍无回复则自动停跟取消
            'interviewCancelAfterRemind': True,
            'maxFollowDays': 7,
            'chatInterval': 25,
            # 仅处理当天有消息的候选人
            'todayOnly': True,
            # 单次任务最多处理人数
            'maxCandidatesPerRun': 5,
            # 每日聊天消息总上限（五类合计）
            'maxMessagesPerDay': 50,
            # 每类话术每日上限
            'maxPerActionType': 10,
            # 有效处理完一人后的随机等待秒数
            'chatIntervalMin': 25,
            'chatIntervalMax': 50,
            # 跳过或无有效操作时的短等待秒数
            'skipIntervalMin': 3,
            'skipIntervalMax': 6,
            # 页面操作随机间隔秒数
            'minActionGapMin': 2,
            'minActionGapMax': 5,
            # 当日安全验证超过此次数停跑
            'maxVerifyPerDay': 2,
            # 允许程序运行的工作时段（仅限制任务启动时刻，不筛选列表消息时间）
            'workWindows': [['9:00', '11:30'], ['13:00', '18:00']],
            # 页面风控关键词，命中即停跑
            'riskKeywords': ['发送过于频繁', '操作频繁', '账号异常', '请稍后再试', '访问过于频繁'],
            # 话术配额用尽时填入输入框由人工发送，不停任务
            'handoffWhenLimit': True,
            # 企业微信群机器人 Webhook 完整 URL（留空则不推送日报）
            'wecomWebhookUrl': os.getenv('BOSS_WECOM_WEBHOOK_URL', ''),
            # 日报推送完成后 @ 的企微绑定手机号（逗号分隔多个）
            'wecomMentionMobile': '',
            # 企微日报面试明细每批条数
            'wecomInterviewBatchSize': 10,
            # 企微连续发送间隔秒数
            'wecomPushGapSec': 1.5,
            # 人工切入聊天回复：命中关键词则等待 HR 自由回复
            'manualHandoffKeywords': ['薪资', '工资', '待遇', '外包', '远程', '兼职', '实习', '社保', '公积金'],
            'manualWhenUnknownIntent': True,
            'manualWhenNoTemplate': True,
        }
        # 四个内置岗位：intro + match 合并配置
        self.profiles = [
            {
                'jobName': 'AI海外短剧制作',
                'matchKeys': ['AI海外短剧', '海外短剧制作', 'AI短剧'],
                'intro': '您好！感谢回复。{job}岗位我们目前在招，主要负责海外短剧剪辑与多平台发布，需要熟悉 TikTok、YouTube 等渠道，有短剧或短视频剪辑经验优先。您要是感兴趣，方便先发一份最新简历吗？',
                'match': {
                    'ageMin': 18,
                    'ageMax': 45,
                    'educationList': [],
                    'workYearsMin': 0,
                    'mustKeywords': ['剪辑'],
                    'anyKeywords': [
                        ['短剧', 'AI短剧', '海外短剧', '短视频'],
                        ['TikTok', 'tiktok', 'TK', 'YouTube', 'Shorts', 'Facebook', 'Reels'],
                        ['半年', '6个月', '6月', '一年', '1年', '2年', '经验'],
                    ],
                    'preferKeywords': ['剪映', 'Runway', 'Pika', 'AI配音', 'AI字幕', '英文字幕', '海外', '英语', '英文'],
                    'rejectKeywords': ['无经验', '小白', '应届生', '勿扰'],
                },
            },
            {
                'jobName': '海外短剧全域运营',
                'matchKeys': ['海外短剧', '全域运营', '短剧运营'],
                'intro': '您好！感谢回复。{job}岗位我们目前在招，主要负责海外短剧账号运营、内容分发与数据复盘，需要熟悉 TikTok/YouTube/Facebook 等平台。欢迎发份简历，我们详细沟通～',
                'match': {
                    'ageMin': 18,
                    'ageMax': 45,
                    'educationList': [],
                    'workYearsMin': 0,
                    'mustKeywords': ['短剧'],
                    'anyKeywords': [
                        ['运营', '剪辑', '切片', '高光'],
                        ['TikTok', 'tiktok', 'TK', 'YouTube', 'Facebook', 'Instagram', 'Reels'],
                        ['海外', '英语', '英文'],
                    ],
                    'preferKeywords': ['养号', '矩阵', '爆款', '完播', '限流', '漫剧', '小说推文', '影视切片', '数据复盘', '剪映'],
                    'rejectKeywords': ['无经验', '小白', '纯小白', '勿扰', '未接触过'],
                },
            },
            {
                'jobName': '跨境电商产品拍摄师',
                'matchKeys': ['跨境电商', '产品拍摄', '亚马逊拍摄'],
                'intro': '您好！感谢回复。{job}岗位我们目前在招，主要负责跨境电商产品图与视频拍摄、亚马逊 A+ 页面素材制作，需要熟练使用相机/灯光及 PR 等后期工具。方便先发一份最新简历吗？',
                'match': {
                    'ageMin': 18,
                    'ageMax': 45,
                    'educationList': [],
                    'workYearsMin': 1,
                    'mustKeywords': [],
                    'anyKeywords': [
                        ['拍摄', '摄影', '剪辑', '视频'],
                        ['亚马逊', 'Amazon', '跨境电商', '电商', 'A+'],
                        ['Premiere', 'PR', 'Final Cut', 'FCP', 'After Effects', 'AE', '单反'],
                        ['1年', '2年', '3年', '三年', '两年', '经验'],
                    ],
                    'preferKeywords': ['主图', '白底', '场景图', 'Shopify', 'eBay', '3D', '动画'],
                    'rejectKeywords': ['无经验', '小白', '未接触过拍摄'],
                },
            },
            {
                'jobName': '后端开发工程师',
                'matchKeys': ['后端开发', '后端工程师', 'Java开发', 'Python开发', '服务端开发'],
                'intro': '您好！感谢回复。{job}岗位我们目前在招，主要负责服务端接口开发、数据库设计与系统维护，技术栈以 Python/Java 为主，有 Spring/Django 经验优先。您要是感兴趣，方便先发一份最新简历吗？',
                'match': {
                    'ageMin': 18,
                    'ageMax': 40,
                    'educationList': ['本科', '大专'],
                    'workYearsMin': 1,
                    'mustKeywords': [],
                    'anyKeywords': [
                        ['Python', 'python', 'Java', 'java', 'Golang', 'golang', 'Go语言', 'Node.js', 'NodeJS', 'PHP', 'C#', '.NET'],
                        ['后端', '服务端', 'Server', 'API', '接口开发', '微服务'],
                        ['Spring', 'SpringBoot', 'Spring Boot', 'Django', 'Flask', 'FastAPI', 'MyBatis', 'Hibernate'],
                        ['MySQL', 'Redis', 'MongoDB', 'PostgreSQL', 'Oracle', '数据库'],
                        ['1年', '2年', '3年', '三年', '两年', '经验', '开发经验'],
                    ],
                    'preferKeywords': ['分布式', '高并发', 'Kafka', 'RabbitMQ', 'Docker', 'Kubernetes', 'k8s', 'Linux', 'Git', 'RESTful', '消息队列', 'Elasticsearch', 'Nginx'],
                    'rejectKeywords': ['无经验', '小白', '纯前端', '仅前端', '前端开发', '运维', '测试工程师'],
                },
            },
        ]

    def bundle(self):
        """返回岗位 bundle，供 BossDb.reloadFromConfig 写入数据库"""
        return {'profiles': self.profiles, 'globalSettings': self.globalSettings}

    def profileNames(self):
        """返回所有内置岗位名称"""
        return [str(p.get('jobName') or '') for p in self.profiles]

if __name__ == '__main__':
    # 独立调试：打印 globalSettings 与各岗位 intro/match 字段数量
    job = BossJob()
    bundle = job.bundle()
    print('globalSettings', bundle['globalSettings'])
    for profile in bundle['profiles']:
        print(profile['jobName'], len(profile.get('intro') or ''), 'match keys', len(profile.get('match') or {}))
