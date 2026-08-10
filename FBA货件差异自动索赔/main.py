"""FBA 货件差异自动索赔配置与流程调用入口。"""

import os
from datetime import date, timedelta

from DrissionPage import ChromiumPage

from auto import Auto
from export import PopExport
from saihu import Saihu


class Main:
    """统一维护 GUI 默认配置，并调用赛狐流程与易得客流程"""

    def __init__(self):
        # 项目根目录，用于定位资源文件、配置文件和默认输出目录
        self.baseDir = PopExport.getBaseDir()
        # GUI 本地缓存配置文件，删除后会按本类默认值重新生成
        self.configFile = self.baseDir / "run_config.json"
        # 默认 POP 输出目录，赛狐流程生成的 PDF 默认保存到这里
        self.defaultExportDir = str(self.baseDir / "file")
        # 默认 POP 模板文件，GUI 未选择模板时使用项目内置模板
        self.defaultTemplatePath = str(PopExport.getResourceDir() / "服务商模板.docx")
        # 默认运行环境，False 表示线下环境，True 表示线上环境
        self.defaultIsOnline = False
        # 默认邮件发送开关，False 表示不发送邮件
        self.defaultSendEmail = False
        # 默认邮件接收邮箱，用于 GUI 首次启动时回填
        self.defaultEmail = "yinkaiyuan@bonison.net"
        # 默认邮件发件邮箱，用于 SMTP 登录并发送通知邮件
        self.defaultSenderEmail = "1974419863@qq.com"
        # 默认 SMTP 授权码，用于发件邮箱登录 SMTP 服务
        self.defaultSmtpAuthCode = os.getenv("SMTP_AUTH_CODE", "")
        # 默认 SMTP 服务地址，QQ 邮箱使用 smtp.qq.com
        self.defaultSmtpServer = "smtp.qq.com"
        # 默认 SMTP SSL 端口，QQ 邮箱 SSL 发送端口为 465
        self.defaultSmtpPort = "465"
        # 默认企业微信发送开关，False 表示不发送企业微信通知
        self.defaultSendWechat = False
        # 默认企业微信群机器人 Webhook，用于推送 CASE 结果汇总
        self.defaultWechatWebhook = os.getenv("FBA_WECOM_WEBHOOK_URL", "")
        # 默认企业微信 @ 手机号，GUI 可按实际接收人修改
        self.defaultWechatMobile = "18280194086"
        # 默认赛狐账号，GUI 首次启动时回填
        self.defaultSaihuUsername = os.getenv("SAIHU_USERNAME", "")
        # 默认赛狐密码，GUI 首次启动时回填
        self.defaultSaihuPassword = os.getenv("SAIHU_PASSWORD", "")
        # 默认赛狐筛选站点
        self.defaultSiteName = "美国"
        # 默认赛狐店铺主体名，运行时会按站点拼接后缀
        self.defaultShopBaseName = "Lydia deal"
        # 默认易得客账号，GUI 首次启动时回填
        self.defaultYidekeUsername = os.getenv("YIDEKE_USERNAME", "")
        # 默认易得客密码，GUI 首次启动时回填
        self.defaultYidekePassword = os.getenv("YIDEKE_PASSWORD", "")
        # 默认易得客店铺站点，用于进店访问
        self.defaultAutoSiteName = "美国"
        # 默认 Amazon 后台站点，用于 Seller Central 站点切换
        self.defaultAmazonSiteName = "美国"
        # 默认店铺 IP，GUI 首次启动时回填
        self.defaultShopIp = os.getenv("FBA_SHOP_IP", "")
        # 默认店铺调试端口，用于接管易得客浏览器
        self.defaultShopPort = "8888"
        # 默认 Amazon 登录邮箱，GUI 首次启动时回填
        self.defaultAmazonEmail = os.getenv("FBA_AMAZON_EMAIL", "")
        # 默认 Amazon 登录密码，GUI 首次启动时回填
        self.defaultAmazonPassword = os.getenv("FBA_AMAZON_PASSWORD", "")
        # 当前月份第一天，用于计算默认筛选时间
        firstDayThisMonth = date.today().replace(day=1)
        # 上月最后一天，用于默认结束时间
        lastDayLastMonth = firstDayThisMonth - timedelta(days=1)
        # 上月第一天，用于默认开始时间
        firstDayLastMonth = lastDayLastMonth.replace(day=1)
        # 默认赛狐筛选开始时间
        self.defaultStartDate = firstDayLastMonth.strftime("%Y-%m-%d")
        # 默认赛狐筛选结束时间
        self.defaultEndDate = lastDayLastMonth.strftime("%Y-%m-%d")
        # 赛狐支持筛选的站点名称
        self.siteNames = [
            "美国", "加拿大", "墨西哥", "巴西",
            "英国", "法国", "德国", "意大利", "西班牙", "荷兰", "瑞典", "波兰", "比利时", "爱尔兰",
            "日本", "新加坡", "澳大利亚", "印度", "阿联酋", "沙特阿拉伯", "土耳其", "埃及", "南非",
        ]
        # 赛狐支持筛选的店铺主体名，运行时根据站点补全后缀
        self.shopNames = [
            "Hoople", "TONOS", "KORCCI", "BPG", "dpd",
            "Lydia deal", "TOPOKO", "Bofoho", "TOPULORS", "KK",
            "EZ-COZY", "SUANDSU", "SERIX", "EVERTIX", "7star",
        ]
        # 赛狐店铺站点后缀映射，用于把店铺主体名拼成页面完整店铺名
        self.siteShopSuffixMap = {
            "美国": "US",
            "加拿大": "CA",
            "墨西哥": "MX",
            "巴西": "BR",
            "英国": "UK",
            "法国": "FR",
            "德国": "DE",
            "意大利": "IT",
            "西班牙": "ES",
            "荷兰": "NL",
            "瑞典": "SE",
            "波兰": "PL",
            "比利时": "BE",
            "爱尔兰": "IE",
            "日本": "JP",
            "新加坡": "SG",
            "澳大利亚": "AU",
            "印度": "IN",
            "阿联酋": "AE",
            "沙特阿拉伯": "SA",
            "土耳其": "TR",
            "埃及": "EG",
            "南非": "ZA",
        }
        # 易得客与 Amazon 后台站点映射，键用于 GUI 展示，值用于 Amazon 站点切换
        self.autoSiteMap = {
            "美国": "United States",
            "加拿大": "Canada",
            "墨西哥": "Mexico",
            "巴西": "Brazil",
            "英国": "United Kingdom",
            "法国": "France",
            "德国": "Germany",
            "意大利": "Italy",
            "西班牙": "Spain",
            "荷兰": "Netherlands",
            "瑞典": "Sweden",
            "波兰": "Poland",
            "比利时": "Belgium",
            "爱尔兰": "Ireland",
            "日本": "Japan",
            "新加坡": "Singapore",
            "澳大利亚": "Australia",
            "印度": "India",
            "阿联酋": "United Arab Emirates",
            "沙特阿拉伯": "Saudi Arabia",
            "土耳其": "Turkey",
            "埃及": "Egypt",
            "南非": "South Africa",
        }
        # 易得客流程 GUI 站点下拉项
        self.autoSiteNames = list(self.autoSiteMap.keys())

    def runSaihu(self, config):
        """执行赛狐流程"""
        # 赛狐业务逻辑统一由 Saihu 类负责
        Saihu(config).run()

    def runAuto(self, config):
        """执行易得客流程"""
        # 易得客业务逻辑统一由 Auto 类负责
        Auto(config).run()


if __name__ == "__main__":
    # 本文件独立调试入口，默认调起赛狐流程
    service = Main()
    # 调试时创建浏览器页面实例
    page = ChromiumPage()
    # 调试配置仅用于本文件直接运行
    config = {
        "page": page,
        "username": service.defaultSaihuUsername,
        "password": service.defaultSaihuPassword,
        "exportDir": service.defaultExportDir,
        "baseDir": str(service.baseDir),
        "isOnline": service.defaultIsOnline,
        "siteName": service.defaultSiteName,
        "shopName": f"{service.defaultShopBaseName}-{service.siteShopSuffixMap[service.defaultSiteName]}",
        "shopBaseName": service.defaultShopBaseName,
        "startDate": service.defaultStartDate,
        "endDate": service.defaultEndDate,
        "templatePath": service.defaultTemplatePath,
        "sendEmail": service.defaultSendEmail,
        "email": service.defaultEmail,
        "sender_email": service.defaultSenderEmail,
        "smtp_auth_code": service.defaultSmtpAuthCode,
        "smtp_server": service.defaultSmtpServer,
        "smtp_port": service.defaultSmtpPort,
        "sendWechat": service.defaultSendWechat,
        "wechatWebhook": service.defaultWechatWebhook,
        "wechatMobile": service.defaultWechatMobile,
    }
    service.runSaihu(config)
