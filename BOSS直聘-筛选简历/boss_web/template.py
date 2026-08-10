class BossTemplate:
    """话术默认配置：唯一代码入口，启动时灌入 message_templates 表"""

    def __init__(self):
        # 七类话术，key 与 message_templates.template_type 一致
        self.words = {
            # 对方回复后的岗位介绍（也可被 job.py 的 intro 覆盖）
            "greeting": [
                "您好！感谢回复。{job}这个岗位我们目前在招，主要负责相关核心业务，团队氛围好、成长空间也不错。您要是感兴趣，方便先发一份最新简历吗？"
            ],
            # 求简历后仍未收到时的跟进
            "followup": [
                "您好，请问方便更新一下在线简历吗？",
            ],
            # 审核通过后发面试预邀请（聊天），确认后再由人工发 BOSS 正式邀约
            "interview_pre": [
                "{name}你好，恭喜通过简历初筛！\n我们想邀请你参加面试，时间和地点如下：\n{date} {time}\n{address}\n预计{duration}分钟。\n这个时间方便吗？确认后我发正式邀请函。",
            ],
            # 对方表示时间不合适时追问
            "interview_ask_time": [
                "收到，请问您什么时候方便参加面试呢？回复具体时间即可，谢谢～",
            ],
            # 预邀请 1 小时无回复时追问
            "interview_remind": [
                "您好，想跟进一下面试时间是否方便？若可以请回复确认，若不方便也请告知合适的时间，谢谢～",
            ],
            # 对方提出新时间后确认改期
            "interview_reschedule": [
                "好的，已为您调整至{date} {time}，请回复确认是否方便，确认后我发正式邀请函。",
            ],
            # 默认取消面试
            "interview_cancel": [
                "好的，已为您取消本次面试安排。祝您求职顺利，后续有合适机会再联系您～",
            ],
            # 审核不通过通知
            "reject": [
                "{name}你好，感谢你的简历投递！\n很抱歉，经过评估，虽然你的背景很优秀，但目前这个岗位的要求和你的经验/技能方向匹配度不是特别高。\n祝你求职顺利，找到心仪的工作！",
            ],
            # 智能回复：对方表达兴趣
            "reply_interest": [
                "你好啊，可以聊一聊～感谢关注，方便先发一份最新简历吗？",
                "您好，很高兴收到您的消息！这个岗位还在招聘，方便发份简历吗？",
            ],
            # 智能回复：对方主动发简历
            "reply_resume": [
                "好的，欢迎发简历，我这边看完后尽快给您反馈～",
                "可以的，您直接发附件简历或在线简历都可以，谢谢！",
            ],
            # 智能回复：对方想了解岗位
            "reply_learn": [
                "您好！{job}这个岗位我们还在招人，感兴趣的话欢迎先发份简历，我们详聊～",
                "您好，岗位详情在招聘信息里都有介绍，您要是觉得合适，方便先发份简历吗？",
            ],
        }

    def bundle(self):
        """返回完整话术 bundle，供 BossDb.reloadFromConfig 写入数据库"""
        return self.words

    def types(self):
        """返回所有 template_type 名称列表"""
        return list(self.words.keys())

    def wordsOf(self, templateType):
        """读取某一类话术文本列表"""
        return list(self.words.get(templateType) or [])


if __name__ == "__main__":
    # 独立调试：打印各类话术条数
    config = {}
    tpl = BossTemplate()
    for templateType, wordList in tpl.bundle().items():
        print(templateType, len(wordList))
