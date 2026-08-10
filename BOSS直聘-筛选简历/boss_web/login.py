import threading
import time
try:
    from DrissionPage import ChromiumOptions, ChromiumPage
except ImportError:
    ChromiumOptions = None
    ChromiumPage = None

class BossLogin:
    """BOSS 招聘端登录：APP 扫码（默认）或手机号验证码（保留）"""

    def __init__(self):
        # Chrome 用户数据目录，复用已登录会话
        self.userDataPath = ''
        # 登录方式：scan 扫码 / sms 手机验证码
        self.loginMode = 'scan'
        # 手机号，sms 模式时由程序填入
        self.phone = ''
        # BOSS 招聘端登录页 URL
        self.loginUrl = 'https://login.zhipin.com/?ka=header-boss'
        # 登录成功后 URL 中应包含的片段
        self.loggedInUrlPart = 'zhipin.com/web'
        # 轮询登录状态的间隔秒数
        self.pollSec = 2.0
        # 等待登录时周期性提醒的间隔秒数
        self.loginRemindSec = 30.0
        # 二维码失效页面可能出现的文案
        self.qrExpiredTexts = ['二维码已失效', '已失效', '过期', '点击刷新', '请重新刷新']
        # 招聘方登录界面特征文案
        self.recruiterMarkers = ['招聘效果好', '跳槽牛人', '牛人简历', '人才匹配', '招人才']
        # 求职者登录界面特征文案（用于排除误判）
        self.seekerMarkers = ['各大行业职位任你选', '任性选']
        # DrissionPage 页面对象
        self.page = None
        # 停止等待登录的事件标志
        self.stopFlag = threading.Event()

    def log(self, message, logCallback=None):
        """输出日志"""
        text = str(message)
        # 打印到控制台
        print(text)
        # 若有回调则同步推送给上层
        if logCallback:
            logCallback(text)

    def requestStop(self):
        """请求停止登录等待"""
        # 置位停止标志，waitScanLoop / waitSmsLoginLoop 会中断
        self.stopFlag.set()

    def ensureNotStopped(self):
        """未停止才继续"""
        # 用户已请求停止则抛出中断异常
        if self.stopFlag.is_set():
            raise InterruptedError('用户已停止登录')

    def createPage(self):
        """创建本地 Chrome 页面"""
        # 未安装 DrissionPage 时无法启动浏览器
        if ChromiumOptions is None:
            raise RuntimeError('缺少 DrissionPage，请执行: pip install DrissionPage')
        # 配置 Chrome 启动参数
        co = ChromiumOptions()
        # 降低自动化特征被检测的概率
        co.set_argument('--disable-blink-features=AutomationControlled')
        # 指定用户数据目录以复用 Cookie
        if self.userDataPath:
            co.set_paths(user_data_path=self.userDataPath)
        # 返回页面对象供后续操作
        return ChromiumPage(co)

    def clickByTexts(self, texts, timeout=3):
        """按文本列表依次尝试点击元素"""
        for text in texts:
            # 方式一：DrissionPage 精确 text= 选择器
            ele = self.page.ele(f'text={text}', timeout=timeout)
            if ele:
                ele.click()
                return True
            # 方式二：text: 模糊匹配选择器
            ele = self.page.ele(f'text:{text}', timeout=0.5)
            if ele:
                ele.click()
                return True
            # 方式三：XPath 按 normalize-space 包含文本查找
            ele = self.page.ele(f'xpath://*[contains(normalize-space(.),"{text}")]', timeout=0.5)
            if ele:
                ele.click()
                return True
        # 所有文本均未匹配到可点击元素
        return False

    def pageText(self):
        """读取页面可见文本"""
        try:
            # 通过 JS 取 body 内可见文本，用于身份判断
            return self.page.run_js("return document.body ? document.body.innerText : ''") or ''
        except Exception:
            # 页面未就绪或 JS 执行失败时返回空串
            return ''

    def isRecruiterMode(self):
        """当前是否为招聘方登录界面"""
        text = self.pageText()
        # 页面文案含招聘方特征词则判定为招聘方
        if any((marker in text for marker in self.recruiterMarkers)):
            return True
        # 备选：读取身份 Tab 当前激活项文案
        activeTab = self.page.run_js("\n            const tab = document.querySelector('.identity-tab li.active');\n            return tab ? tab.textContent.trim() : '';\n            ")
        # 「我要招聘」Tab 激活即为招聘方
        return activeTab == '我要招聘'

    def ensureIdentityTabs(self, logCallback=None):
        """确保身份切换 Tab 可见（扫码视图会隐藏 Tab，需先切回表单视图）"""
        # Tab 已可见则无需切换
        if self.page.ele('css:ul.identity-tab', timeout=1):
            return
        self.log('当前在扫码视图，先切回表单以切换招聘方身份...', logCallback)
        # 优先点击「切换到手机登录」按钮
        switchBtn = self.page.ele('css:.btn-sign-switch.phone-switch', timeout=2)
        if switchBtn:
            switchBtn.click()
            # 等待表单视图渲染完成
            time.sleep(1.2)
            return
        # 兜底：按文案点击验证码登录入口
        self.clickByTexts(['验证码登录', '验证码登录/注册', '短信登录'], timeout=2)

    def openLoginPage(self, logCallback=None):
        """打开 BOSS 招聘端登录页"""
        self.log('正在启动浏览器...', logCallback)
        # 创建 Chrome 页面对象
        self.page = self.createPage()
        # 导航到 BOSS 登录 URL
        self.page.get(self.loginUrl)
        # 等待文档加载完成
        self.page.wait.doc_loaded()
        # 额外等待动态内容渲染
        time.sleep(1.5)
        self.log('已打开 BOSS 登录页', logCallback)

    def switchRecruiter(self, logCallback=None):
        """切换到「我要招聘」身份 Tab"""
        self.log('正在切换到招聘方身份...', logCallback)
        # 扫码页会隐藏 Tab，先确保 Tab 可见
        self.ensureIdentityTabs(logCallback)
        # 优先用 XPath 定位「我要招聘」Tab
        recruitTab = self.page.ele('xpath://ul[contains(@class,"identity-tab")]//li[contains(normalize-space(.),"我要招聘")]', timeout=3)
        if recruitTab:
            recruitTab.click()
            # 等待左侧文案切换
            time.sleep(1.2)
        else:
            # 兜底：按文本点击
            self.clickByTexts(['我要招聘'], timeout=2)
        # 校验是否已处于招聘方界面
        if self.isRecruiterMode():
            self.log('已切换到招聘方（左侧应显示招人/招聘相关文案）', logCallback)
            return
        # 未能确认时提示用户人工检查
        self.log('未能确认招聘方界面，请检查浏览器窗口', logCallback)

    def clickScanLogin(self, logCallback=None):
        """进入 APP 扫码登录"""
        # 已在扫码页则直接返回
        if self.page.ele('xpath://*[contains(text(),"BOSS直聘APP扫码登录")]', timeout=1):
            self.log('已在扫码登录界面', logCallback)
            return
        self.log('正在切换到扫码登录...', logCallback)
        # 优先点击二维码切换按钮
        scanBtn = self.page.ele('css:.btn-sign-switch.ewm-switch', timeout=2)
        if scanBtn:
            scanBtn.click()
            time.sleep(1.2)
            self.log('已进入扫码登录', logCallback)
            return
        # 兜底：按文案点击扫码入口
        if self.clickByTexts(['APP扫码登录', '扫码登录'], timeout=2):
            time.sleep(1)
            self.log('已进入扫码登录', logCallback)
            return
        # 两种方式均未找到入口
        self.log('未找到扫码登录入口，请检查页面', logCallback)

    def isLoggedIn(self):
        """检测是否已登录成功"""
        url = self.page.url or ''
        # 已跳转到 web 工作台且不在登录页、非个人中心中间页
        if self.loggedInUrlPart in url and 'login' not in url.lower() and ('/web/user' not in url):
            return True
        # 页面出现用户名元素
        if self.page.ele('xpath://span[@class="user-name"]', timeout=0.5):
            return True
        # 出现「沟通」导航且不在个人中心页
        if self.page.ele('text:沟通', timeout=0.5) and '/web/user' not in url:
            return True
        # 以上条件均不满足，视为未登录
        return False

    def isQrExpired(self):
        """检测二维码是否已失效"""
        html = self.page.html or ''
        text = self.pageText()
        # 匹配预设失效关键词
        for keyword in self.qrExpiredTexts:
            if keyword in html or keyword in text:
                return True
        # 匹配失效/过期相关 CSS 类名
        if self.page.ele('xpath://*[contains(@class,"invalid") or contains(@class,"expire")]', timeout=0.3):
            return True
        return False

    def refreshQr(self, logCallback=None):
        """点击刷新二维码"""
        refreshTexts = ['点击刷新', '刷新二维码', '请重新刷新', '刷新', '重新获取']
        # 按文案点击刷新按钮
        clicked = self.clickByTexts(refreshTexts, timeout=1)
        if clicked:
            self.log('二维码已刷新，请重新扫码', logCallback)
            time.sleep(1)
            return True
        # 兜底：点击二维码图片区域尝试刷新
        qrBox = self.page.ele('xpath://img[contains(@class,"qrcode") or contains(@alt,"二维码")]', timeout=0.5)
        if qrBox:
            qrBox.click()
            self.log('已尝试点击二维码区域刷新', logCallback)
            time.sleep(1)
            return True
        # 未找到可刷新入口
        return False

    def waitScanLoop(self, logCallback=None):
        """持续等待扫码，失效则自动刷新"""
        self.log('请使用 BOSS 直聘 APP（招聘方身份）扫描浏览器中的二维码', logCallback)
        self.log('登录成功或点击停止后将结束等待', logCallback)
        lastRemind = time.time()
        while True:
            # 用户停止则中断循环
            self.ensureNotStopped()
            # 检测到登录成功则结束
            if self.isLoggedIn():
                self.log('扫码登录成功！', logCallback)
                return
            # 二维码失效则自动刷新
            if self.isQrExpired():
                self.log('检测到二维码已失效，正在自动刷新...', logCallback)
                self.refreshQr(logCallback)
            # 周期性提醒用户继续扫码
            now = time.time()
            if now - lastRemind >= 30:
                self.log('仍在等待扫码，请在 APP 中确认登录...', logCallback)
                lastRemind = now
            # 间隔 pollSec 秒后再检测
            time.sleep(self.pollSec)

    def ensurePhoneLoginView(self, logCallback=None):
        """确保停留在手机号验证码表单页（不进入扫码页）"""
        # 当前在扫码页则需切回表单
        onScanPage = self.page.ele('xpath://*[contains(text(),"BOSS直聘APP扫码登录")]', timeout=0.5)
        if onScanPage:
            self.log('当前在扫码页，切回手机号验证码登录...', logCallback)
            self.ensureIdentityTabs(logCallback)
            time.sleep(1.2)
            return
        # 查找手机号输入框
        phoneEle = self.page.ele('css:input[type="tel"]', timeout=0.5)
        if not phoneEle:
            phoneEle = self.page.ele('xpath://input[contains(@placeholder,"手机")]', timeout=0.5)
        # 未找到输入框则先确保身份 Tab 可见
        if not phoneEle:
            self.ensureIdentityTabs(logCallback)
            time.sleep(1.0)

    def findPhoneInput(self):
        """查找手机号输入框"""
        # 优先 type=tel 输入框
        phoneEle = self.page.ele('css:input[type="tel"]', timeout=2)
        if phoneEle:
            return phoneEle
        # 兜底：placeholder 含「手机」的输入框
        return self.page.ele('xpath://input[contains(@placeholder,"手机")]', timeout=2)

    def maskPhone(self, phone):
        """脱敏显示手机号"""
        # 只保留数字
        digits = ''.join((ch for ch in phone if ch.isdigit()))
        # 长度足够则中间四位打星
        if len(digits) >= 7:
            return f'{digits[:3]}****{digits[-4:]}'
        return phone

    def fillPhone(self, logCallback=None):
        """填写手机号"""
        phone = str(self.phone or '').strip()
        # 未配置手机号则拒绝继续
        if not phone:
            raise ValueError('请填写手机号')
        # 确保在手机号表单页
        self.ensurePhoneLoginView(logCallback)
        phoneEle = self.findPhoneInput()
        # 找不到输入框则报错
        if not phoneEle:
            raise RuntimeError('未找到手机号输入框，请检查登录页')
        phoneEle.clear()
        phoneEle.input(phone)
        # 日志中脱敏显示
        self.log(f'已填写手机号: {self.maskPhone(phone)}', logCallback)
        time.sleep(0.8)

    def clickSendCode(self, logCallback=None):
        """点击发送验证码"""
        sendTexts = ['发送验证码', '获取验证码', '重新发送']
        # 按文案点击发送按钮
        if self.clickByTexts(sendTexts, timeout=3):
            self.log('已点击发送验证码，请查看手机短信', logCallback)
            time.sleep(1)
            return True
        # 未找到按钮时提示用户手动操作
        self.log('未找到发送验证码按钮，请在浏览器中手动点击', logCallback)
        return False

    def waitSmsLoginLoop(self, logCallback=None):
        """等待用户在浏览器中输入验证码并完成登录"""
        self.log('请在浏览器中输入短信验证码并点击登录', logCallback)
        self.log('登录成功或点击停止后将结束等待', logCallback)
        lastRemind = time.time()
        while True:
            # 用户停止则中断
            self.ensureNotStopped()
            # 检测到登录成功则结束
            if self.isLoggedIn():
                self.log('手机号验证码登录成功！', logCallback)
                return
            # 周期性提醒用户输入验证码
            now = time.time()
            if now - lastRemind >= self.loginRemindSec:
                self.log('仍在等待登录，请在浏览器输入验证码并点击登录...', logCallback)
                lastRemind = now
            # 间隔 pollSec 秒后再检测
            time.sleep(self.pollSec)

    def runSmsLogin(self, logCallback=None):
        """手机号验证码登录：程序填号发码，用户在浏览器完成验证"""
        # 确保在手机号表单页
        self.ensurePhoneLoginView(logCallback)
        # 自动填写手机号
        self.fillPhone(logCallback)
        # 点击发送验证码
        self.clickSendCode(logCallback)
        # 等待用户在浏览器完成验证并登录
        self.waitSmsLoginLoop(logCallback)

    def run(self, logCallback=None):
        """完整登录流程"""
        # 每次 run 前清除停止标志
        self.stopFlag.clear()
        try:
            # 打开登录页
            self.openLoginPage(logCallback)
            self.ensureNotStopped()
            # 切换到招聘方身份
            self.switchRecruiter(logCallback)
            self.ensureNotStopped()
            # 浏览器已有有效会话则直接结束
            if self.isLoggedIn():
                self.log('当前浏览器已处于登录状态', logCallback)
                return
            # 按 loginMode 走扫码或短信流程
            if self.loginMode == 'scan':
                self.clickScanLogin(logCallback)
                self.ensureNotStopped()
                self.waitScanLoop(logCallback)
            else:
                self.runSmsLogin(logCallback)
        except InterruptedError as exc:
            # 用户主动停止，仅记录日志不向上抛
            self.log(str(exc), logCallback)
        except Exception as exc:
            # 其他异常记录后重新抛出
            self.log(f'登录流程异常: {exc}', logCallback)
            raise

if __name__ == '__main__':
    # 本文件独立调试：python boss_web/login.py 或 python -m boss_web.login
    # 修改下方 config 后可直接测试扫码/短信登录，无需启动完整 Web 服务
    config = {'userDataPath': 'D:\\boss_zhaopin_筛选简历\\boss_chrome_profile', 'loginMode': 'scan', 'phone': ''}
    login = BossLogin()
    login.userDataPath = config['userDataPath']
    login.loginMode = config['loginMode']
    login.phone = config['phone']
    login.run()
