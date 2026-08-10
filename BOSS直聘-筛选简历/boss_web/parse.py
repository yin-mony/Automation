import re

class ResumeParse:
    """从 BOSS 聊天页解析候选人简历信息"""

    def __init__(self):
        # 聊天页顶部基础信息区域（年龄、学历等）
        self.baseInfoXpath = 'xpath://div[@class="base-info-single-detial"]/div[string-length(@class)=0]'
        # 简历详情正文区域
        self.detailXpath = 'xpath://div[contains(@class,"resume") or contains(@class,"geek-detail")]'

    def parseFromPage(self, page):
        """从当前打开的聊天页解析简历"""
        # 初始化空 profile 结构
        profile = {'name': '', 'age': None, 'education': '', 'workYears': '', 'contact': '', 'recentJobs': [], 'skills': [], 'rawText': ''}
        # 读取顶部基础信息 div 列表
        baseEles = page.eles(self.baseInfoXpath, timeout=2)
        if baseEles:
            # 第一个 div 通常含年龄
            ageText = str(baseEles[0].text or '')
            ageMatch = re.search('(\\d+)', ageText.replace('岁', ''))
            if ageMatch:
                profile['age'] = int(ageMatch.group(1))
            # 最后一个 div 通常为学历
            profile['education'] = str(baseEles[-1].text or '').strip()
        # 读取简历详情正文
        detailEle = page.ele(self.detailXpath, timeout=2)
        if detailEle:
            profile['rawText'] = str(detailEle.text or '')
        # 汇总聊天消息中的文本补充 rawText
        msgText = self.collectMessageText(page)
        if msgText:
            profile['rawText'] = (profile['rawText'] + '\n' + msgText).strip()
        # 从合并文本提取工作年限
        profile['workYears'] = self.extractWorkYears(profile['rawText'])
        profile['contact'] = self.extractContact(profile['rawText'])
        return profile

    def parseFromPreviewModal(self, page):
        """从附件简历预览弹窗解析简历"""
        profile = {'name': '', 'age': None, 'education': '', 'workYears': '', 'contact': '', 'recentJobs': [], 'skills': [], 'rawText': '', 'filePath': ''}
        # 多种弹窗 xpath，按顺序尝试匹配
        modalXpaths = ['xpath://div[contains(@class,"resume-detail")]', 'xpath://div[contains(@class,"lib-resume")]', 'xpath://div[contains(@class,"geek-resume")]', 'xpath://div[contains(@class,"boss-dialog") and .//*[contains(text(),"附件简历")]]', 'xpath://div[contains(@class,"dialog") and .//*[contains(text(),"附件简历")]]']
        modalEle = None
        for xpath in modalXpaths:
            modalEle = page.ele(xpath, timeout=2)
            if modalEle:
                break
        # 未找到弹窗则返回空 profile
        if not modalEle:
            return profile
        # 取弹窗全文作为 rawText
        rawText = str(modalEle.text or '').strip()
        profile['rawText'] = rawText
        if not rawText:
            return profile
        # 从开头匹配姓名（2～10 个非数字字符）
        nameMatch = re.search('^([^\\n\\d]{2,10})', rawText)
        if nameMatch:
            profile['name'] = nameMatch.group(1).strip()
        # 匹配「XX岁」提取年龄
        ageMatch = re.search('(\\d{1,2})\\s*岁', rawText)
        if ageMatch:
            profile['age'] = int(ageMatch.group(1))
        # 按优先级匹配学历关键词
        for edu in ('博士', '硕士', '本科', '大专', '中专', '高中'):
            if edu in rawText:
                profile['education'] = edu
                break
        # 从全文提取工作年限
        profile['workYears'] = self.extractWorkYears(rawText)
        profile['contact'] = self.extractContact(rawText)
        return profile

    def collectMessageText(self, page):
        """汇总聊天消息文本"""
        chunks = []
        # 查找所有消息条目
        msgEles = page.eles('xpath://div[@class="message-item"]', timeout=2)
        for msg in msgEles or []:
            # 每条消息取 text-content  span
            textEle = msg.ele('xpath:.//span[@class="text-content"]', timeout=0.2)
            if textEle and textEle.text:
                chunks.append(str(textEle.text))
        # 用换行拼接全部消息
        return '\n'.join(chunks)

    def extractContact(self, text):
        """从简历文本中提取手机或微信联系方式"""
        content = str(text or '')
        if not content:
            return ''
        # 优先匹配带标签的手机号
        phoneLabel = re.search(r'(?:手机|电话|联系方式|联系电话|手机号)[:：\s]*([0-9+\-\s]{8,20})', content)
        if phoneLabel:
            digits = re.sub(r'\D', '', phoneLabel.group(1))
            if len(digits) >= 11 and digits[-11:][0] == '1':
                return digits[-11:]
        # 匹配独立 11 位大陆手机号
        phoneMatch = re.search(r'(?<!\d)(1[3-9]\d{9})(?!\d)', content)
        if phoneMatch:
            return phoneMatch.group(1)
        # 匹配微信号
        wxLabel = re.search(r'(?:微信|微信号|wx|WX)[:：\s]*([a-zA-Z][a-zA-Z0-9_-]{5,19})', content, re.IGNORECASE)
        if wxLabel:
            return f'微信:{wxLabel.group(1)}'
        return ''

    def extractWorkYears(self, text):
        """从文本中提取工作年限描述"""
        if not text:
            return ''
        # 多种「X年经验」表述的正则
        patterns = ['(\\d+)\\s*年(?:工作)?经验', '工作(\\d+)年', '经验(\\d+)年']
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                # 返回数字部分作为年限
                return match.group(1)
        return ''

if __name__ == '__main__':
    # 独立调试：用样本文本测试解析字段
    config = {'sampleText': '张三 手机：13812345678 3年工作经验，本科，熟悉Python'}
    parser = ResumeParse()
    print({'workYears': parser.extractWorkYears(config['sampleText']), 'contact': parser.extractContact(config['sampleText'])})
