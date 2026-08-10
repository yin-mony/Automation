class ResumeMatch:
    """简历条件匹配引擎：按 job_rules 中的 match 字段审核简历"""

    def __init__(self):
        # 年龄下限（含）
        self.ageMin = 18
        # 年龄上限（含）
        self.ageMax = 45
        # 允许的学历列表，空列表表示不限
        self.educationList = ['本科', '大专']
        # 最低工作年限（年），0 表示不限
        self.workYearsMin = 0
        # 必须全部出现在 rawText 中的关键词
        self.mustKeywords = []
        # 每组至少命中一个词，组与组之间为「且」关系
        self.anyKeywords = []
        # 优先项关键词，命中则在通过原因中标注
        self.preferKeywords = []
        # 命中任一即拒绝的关键词
        self.rejectKeywords = []

    def loadRules(self, rules):
        """从岗位规则字典加载筛选条件"""
        self.ageMin = int(rules.get('ageMin', self.ageMin))
        self.ageMax = int(rules.get('ageMax', self.ageMax))
        self.educationList = list(rules.get('educationList') or [])
        self.workYearsMin = int(rules.get('workYearsMin', self.workYearsMin))
        self.mustKeywords = list(rules.get('mustKeywords') or [])
        self.anyKeywords = list(rules.get('anyKeywords') or [])
        self.preferKeywords = list(rules.get('preferKeywords') or [])
        self.rejectKeywords = list(rules.get('rejectKeywords') or [])

    def match(self, profile):
        """判断简历是否满足条件，返回 (是否通过, 原因)"""
        age = profile.get('age')
        education = str(profile.get('education') or '')
        rawText = str(profile.get('rawText') or '')
        workYearsText = str(profile.get('workYears') or '')
        # 年龄已知且不在 [ageMin, ageMax] 范围内则拒绝
        if age is not None and (not self.ageMin <= int(age) <= self.ageMax):
            return (False, f'年龄不符: {age}')
        # 学历不在允许列表则拒绝（列表为空则跳过此检查）
        if self.educationList and education and (education not in self.educationList):
            return (False, f'学历不符: {education}')
        # 工作年限不足则拒绝（未解析出年限时不做此检查）
        if self.workYearsMin > 0 and workYearsText:
            try:
                if int(workYearsText) < self.workYearsMin:
                    return (False, f'工作年限不足: {workYearsText}年')
            except ValueError:
                # 年限文本非数字则忽略此项
                pass
        # 命中任一排除词则拒绝
        for keyword in self.rejectKeywords:
            if keyword and keyword in rawText:
                return (False, f'命中排除词: {keyword}')
        # 必须关键词缺一不可
        for keyword in self.mustKeywords:
            if keyword and keyword not in rawText:
                return (False, f'缺少必要关键词: {keyword}')
        # 任一组 anyKeywords 至少命中一个词
        for group in self.anyKeywords:
            words = [w for w in group if w]
            if not words:
                continue
            if not any((word in rawText for word in words)):
                return (False, f"未满足条件: {'/'.join(words)}")
        # 统计命中的优先项，用于通过原因标注
        preferHits = [word for word in self.preferKeywords if word and word in rawText]
        if preferHits:
            return (True, f"通过（优先项: {', '.join(preferHits)}）")
        # 全部检查通过且无优先项命中
        return (True, '通过')

if __name__ == '__main__':
    # 独立调试：加载样例规则并对样例简历执行 match
    config = {'rules': {'ageMin': 20, 'ageMax': 35, 'educationList': ['本科'], 'workYearsMin': 1, 'mustKeywords': ['剪辑'], 'anyKeywords': [['短剧', 'AI短剧'], ['TikTok', 'YouTube']], 'preferKeywords': ['剪映'], 'rejectKeywords': ['无经验']}, 'profile': {'age': 25, 'education': '本科', 'workYears': '3', 'rawText': '3年短剧剪辑经验，熟悉TikTok发布'}}
    matcher = ResumeMatch()
    matcher.loadRules(config['rules'])
    ok, reason = matcher.match(config['profile'])
    print({'ok': ok, 'reason': reason})
