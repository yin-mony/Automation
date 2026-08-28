"""读取人事制单工具的本地配置。"""

from pathlib import Path

import config


DEFAULT_OFFER_CC_RECIPIENTS = [
    {"name": "何倩怡", "email": "heqianyi@bonison.net"},
    {"name": "宁致远", "email": "ningzhiyuan@bonison.net"},
]


class Settings:
    """集中管理界面、文件、邮箱草稿和公司默认信息。"""

    def __init__(self):
        self.baseDir = Path(__file__).resolve().parent
        self.host = getattr(config, "PORTAL_HOST", "127.0.0.1")
        self.port = int(getattr(config, "PORTAL_PORT", 8700))
        self.dataDir = self.baseDir / "data"
        self.jobDir = self.dataDir / "jobs"
        self.outputDir = self.baseDir / "output"
        self.maxResumeMb = int(getattr(config, "MAX_RESUME_MB", 15))
        self.smtpHost = getattr(config, "SMTP_HOST", "smtp.exmail.qq.com")
        self.smtpPort = int(getattr(config, "SMTP_PORT", 465))
        self.smtpUseSsl = bool(getattr(config, "SMTP_USE_SSL", True))
        self.smtpUsername = getattr(config, "SMTP_USERNAME", "")
        self.smtpPassword = getattr(config, "SMTP_PASSWORD", "")
        self.mailFrom = getattr(config, "MAIL_FROM", self.smtpUsername)
        self.imapHost = getattr(config, "IMAP_HOST", "imap.exmail.qq.com")
        self.imapPort = int(getattr(config, "IMAP_PORT", 993))
        self.imapUseSsl = bool(getattr(config, "IMAP_USE_SSL", True))
        self.draftFolder = getattr(config, "DRAFT_FOLDER", "Drafts")
        self.companyName = getattr(config, "COMPANY_NAME", "四川伯尼森科技有限公司")
        self.hrName = getattr(config, "HR_NAME", "人事")
        self.hrEmail = getattr(config, "HR_EMAIL", self.mailFrom)
        self.hrPhone = getattr(config, "HR_PHONE", "")
        self.reportLocation = getattr(config, "REPORT_LOCATION", "")
        self.reportTime = getattr(config, "REPORT_TIME", "上午 09:00-09:30")
        self.offerCcRecipients = self.normalizeCcRecipients()
        self.offerCcEmails = [item["email"] for item in self.offerCcRecipients]
        self.offerCcNames = [item["name"] for item in self.offerCcRecipients]
        self.offerCcDisplay = "、".join(self.offerCcNames)
        self.departments = list(getattr(config, "DEPARTMENTS", [
            "总经办", "人事行政部", "财务部", "技术部", "产品部", "运营部", "市场部", "销售部",
        ]))
        self.supervisorDepartments = dict(getattr(config, "SUPERVISOR_DEPARTMENTS", {}))
        configuredFont = getattr(config, "FONT_PATH", "")
        fontCandidates = [
            configuredFont,
            r"C:\Windows\Fonts\simsun.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        ]
        self.fontPath = next((Path(path) for path in fontCandidates if path and Path(path).exists()), None)
        if self.fontPath is None:
            raise RuntimeError("未找到中文字体，请安装 Noto CJK 字体或在 config.py 设置 FONT_PATH")
        self.jobDir.mkdir(parents=True, exist_ok=True)
        self.outputDir.mkdir(parents=True, exist_ok=True)

    def validateDraft(self):
        """校验写入企业邮箱草稿箱所需配置。"""
        if not all([self.imapHost, self.smtpUsername, self.smtpPassword, self.mailFrom]):
            raise ValueError("邮箱草稿配置不完整")

    def validateSend(self):
        """校验真实发送邮件所需配置。"""
        if not all([self.smtpHost, self.smtpPort, self.smtpUsername, self.smtpPassword, self.mailFrom]):
            raise ValueError("邮箱发送配置不完整")

    def normalizeList(self, value):
        """把配置中的字符串或列表统一为去重后的列表。"""
        if isinstance(value, str):
            items = [item.strip() for item in value.replace("；", ",").replace(";", ",").replace("，", ",").replace("|", ",").replace("\n", ",").split(",")]
        else:
            items = [str(item or "").strip() for item in value]
        return list(dict.fromkeys(item for item in items if item))

    def normalizeCcRecipients(self):
        """读取固定抄送人配置，前端展示姓名，邮件发送使用邮箱。"""
        configured = getattr(config, "OFFER_CC_RECIPIENTS", None)
        if configured:
            recipients = []
            for item in configured:
                if isinstance(item, dict):
                    name = str(item.get("name") or "").strip()
                    email = str(item.get("email") or "").strip()
                else:
                    name, email = "", str(item or "").strip()
                if email:
                    recipients.append({"name": name or self.nameForCcEmail(email), "email": email})
            return self.uniqueRecipients(recipients)
        legacyEmails = self.normalizeList(getattr(config, "OFFER_CC_EMAILS", []))
        if legacyEmails:
            return self.uniqueRecipients([
                {"name": self.nameForCcEmail(email), "email": email}
                for email in legacyEmails
            ])
        return list(DEFAULT_OFFER_CC_RECIPIENTS)

    def nameForCcEmail(self, email):
        """把默认抄送邮箱映射为姓名。"""
        mapping = {item["email"].lower(): item["name"] for item in DEFAULT_OFFER_CC_RECIPIENTS}
        return mapping.get(str(email or "").lower(), str(email or ""))

    def uniqueRecipients(self, recipients):
        """按邮箱去重固定抄送人。"""
        result = []
        seen = set()
        for item in recipients:
            email = str(item.get("email") or "").strip()
            if not email or email.lower() in seen:
                continue
            seen.add(email.lower())
            name = str(item.get("name") or "").strip() or email
            result.append({"name": name, "email": email})
        return result
