"""审核通过后的推广邮件（草稿或真实发送）。"""

import hashlib
import html
import imaplib
import mimetypes
import smtplib
import sys
import time
from email.header import Header
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

from openpyxl import Workbook


class Mail:
    """只处理审核通过的联系方式，每家公司只发一封。"""

    def __init__(self, data, config, outputDir, dataDir, log=None):
        """初始化邮件固定配置和文件目录。"""
        self.data = data
        self.config = config
        self.outputDir = Path(outputDir)
        self.dataDir = Path(dataDir)
        self.log = log or (lambda message: None)
        self.promotionServer = str(config.get("promotionSmtpServer") or "smtp.qiye.aliyun.com")
        self.promotionPort = int(config.get("promotionSmtpPort") or 465)
        self.promotionImapServer = str(config.get("promotionImapServer") or "imap.qiye.aliyun.com")
        self.promotionImapPort = int(config.get("promotionImapPort") or 993)
        self.draftMailboxName = str(config.get("promotionDraftMailbox") or "Drafts")
        self.logoFileName = str(config.get("promotionLogoFile") or "time2renew-logo.png")
        self.recordName = "邮件发送记录.xlsx"
        self.skippedRecords = []
        self.defaultPromotionSubject = "Quick question about your agents' CE renewals"
        self.defaultPromotionBody = (
            "Hi,\n"
            "Your agents' renewal season is coming up. We offer a full 24-hour "
            "TDI-approved CE package at $39.99 – probably the lowest price they'll find. "
            "TDI Provider #233836.\n\n"
            "For every agent in your office who uses our package, we can provide a 10% "
            "referral fee. If you're open to a quick conversation, please reply to this email.\n\n"
            "Best,\nTime2renew Support Team\nWebsite: https://time2renew.com"
        )
        self.defaultPersonSubject = "Quick reminder – your insurance license renewal is coming up"
        self.defaultPersonBody = (
            "Hi ,\n\n"
            "Noticed your Texas insurance license is up for renewal. We're running a sale on our 24-hour CE package — $39.99 (regularly $99.99).\n\n"
            "It covers everything you need:\n"
            "3-hour Ethics (mandatory)\n"
            "21-hour electives (Emerging Risks, Fraud, Life & Health, Property & Casualty, Texas Law)\n\n"
            "Texas-licensed (Provider #233836). Fully online. Take it anytime from your phone or computer.\n"
            "No pressure — just wanted to put this on your radar before you pay full price elsewhere.\n\n"
            "Check it out here:https://shop.time2renew.com/products/texas-insurance-ce-courses\n\n"
            "Best,\n\n"
            "QIANYI Ho"
        )

    def approvedRecords(self, action="draft"):
        """读取审核通过且每个对象（公司/个人）只取第一个邮箱的唯一记录。"""
        companySubject = str(self.config.get("promotionSubject") or self.defaultPromotionSubject)
        companyBody = str(self.config.get("promotionBody") or self.defaultPromotionBody)
        personSubject = str(self.config.get("personPromotionSubject") or self.defaultPersonSubject)
        personBody = str(self.config.get("personPromotionBody") or self.defaultPersonBody)
        records = []
        seenObject = set()
        seenEmail = set()
        self.skippedRecords = []
        for result in self.data.contactResults():
            if result.get("reviewStatus") != "approved":
                continue
            mode = str(result.get("mode") or "company")
            objectKey = result.get("objectKey", "")
            if (mode, objectKey) in seenObject:
                continue
            seenObject.add((mode, objectKey))
            emails = result.get("emails") or []
            if not emails:
                continue
            receiver = str(emails[0] or "").strip().lower()
            subject = personSubject if mode == "person" else companySubject
            body = personBody if mode == "person" else companyBody
            if mode == "person":
                first = str(result.get("objectName") or "").split()
                first = first[0].capitalize() if first else ""
                if first:
                    body = body.replace("Hi ,", f"Hi {first},")
            baseRecord = {
                "来源类型": "个人" if mode == "person" else "公司",
                "对象键": result.get("objectKey", ""),
                "对象名称": result.get("objectName", ""),
                "许可证号": result.get("licenseNumber", ""),
                "邮箱": receiver,
                "邮件主题": subject,
                "邮件正文": body,
                "来源链接": "; ".join(
                    result.get("verifiedUrls") or result.get("detailUrls") or result.get("sourceUrls") or []
                ),
                "审核状态": "已通过",
                "发送状态": "待处理",
                "发送结果": "",
                "失败原因": "",
            }
            if receiver in seenEmail:
                duplicate = dict(baseRecord)
                duplicate["发送状态"] = "已跳过"
                duplicate["发送结果"] = "同一批次邮箱重复"
                self.skippedRecords.append(duplicate)
                continue
            seenEmail.add(receiver)
            if self.data.hasMailAction(receiver, action):
                duplicate = dict(baseRecord)
                duplicate["发送状态"] = "已跳过"
                duplicate["发送结果"] = "SQLite 邮件台账已存在成功记录"
                self.skippedRecords.append(duplicate)
                continue
            records.append(baseRecord)
        return records

    def writeRecords(self, records):
        """把推广预览或发送结果写入原中文记录文件。"""
        headers = [
            "来源类型", "对象键", "对象名称", "许可证号", "邮箱", "邮件主题", "邮件正文",
            "来源链接", "审核状态", "发送状态", "发送结果", "失败原因",
        ]
        path = self.outputDir / self.recordName
        path.parent.mkdir(parents=True, exist_ok=True)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "推广发送记录"
        sheet.append(headers)
        for record in records:
            sheet.append([record.get(header, "") for header in headers])
        sheet.auto_filter.ref = sheet.dimensions
        workbook.save(path)
        return path

    def logoPath(self):
        """优先读取外部 file 目录，打包后兼容 PyInstaller 内置资源。"""
        localPath = self.dataDir / self.logoFileName
        if localPath.is_file():
            return localPath
        bundleRoot = str(getattr(sys, "_MEIPASS", "") or "")
        if bundleRoot:
            bundlePath = Path(bundleRoot) / "file" / self.logoFileName
            if bundlePath.is_file():
                return bundlePath
        return localPath

    def htmlBody(self, body, hasLogo):
        """把纯文本正文安全转换为兼容邮箱客户端的 HTML。"""
        safeBody = html.escape(str(body or "")).replace("\n", "<br>\n")
        logoBlock = ""
        if hasLogo:
            logoBlock = (
                '<div style="margin-top:20px;">'
                '<img src="cid:time2renew-logo" alt="Time2Renew" width="220" '
                'style="display:block;width:220px;max-width:100%;height:auto;border:0;">'
                "</div>"
            )
        return (
            '<!doctype html><html><body style="margin:0;padding:0;">'
            '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;'
            f'line-height:1.6;color:#1f2937;">{safeBody}{logoBlock}</div>'
            "</body></html>"
        )

    def buildMessage(self, record, sender):
        """构造带纯文本备用正文和 CID Logo 的标准 MIME 邮件。"""
        message = MIMEMultipart("related")
        fromName = "QIANYI Ho" if record.get("来源类型") == "个人" else "Time2renew Support Team"
        message["From"] = formataddr((str(Header(fromName, "utf-8")), sender))
        message["To"] = record["邮箱"]
        message["Subject"] = Header(record["邮件主题"], "utf-8").encode()
        logoPath = self.logoPath()
        hasLogo = logoPath.is_file()
        alternatives = MIMEMultipart("alternative")
        alternatives.attach(MIMEText(record["邮件正文"], "plain", "utf-8"))
        alternatives.attach(MIMEText(self.htmlBody(record["邮件正文"], hasLogo), "html", "utf-8"))
        message.attach(alternatives)
        if hasLogo:
            mimeType = mimetypes.guess_type(str(logoPath))[0] or "image/png"
            subtype = mimeType.split("/", 1)[-1]
            image = MIMEImage(logoPath.read_bytes(), _subtype=subtype)
            image.add_header("Content-ID", "<time2renew-logo>")
            image.add_header("Content-Disposition", "inline", filename=logoPath.name)
            image.add_header("Content-Location", logoPath.name)
            message.attach(image)
        return message

    def findDraftMailbox(self, server):
        """优先识别 IMAP 特殊草稿箱，无法识别时使用配置名称。"""
        status, rows = server.list()
        if status == "OK":
            for raw in rows or []:
                text = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else str(raw)
                if "\\Drafts" not in text:
                    continue
                name = text.rsplit(" ", 1)[-1].strip().strip('"')
                if name:
                    return name
        return self.draftMailboxName

    def subjectHash(self, subject):
        """生成不包含正文的邮件主题稳定指纹。"""
        return hashlib.sha256(str(subject or "").encode("utf-8")).hexdigest()

    def processApproved(self, action="draft"):
        """默认写入阿里邮箱草稿箱，明确选择 send 时才真实发送。"""
        if action not in {"draft", "send"}:
            raise ValueError("邮件模式只能是 draft 或 send")
        records = self.approvedRecords(action)
        summary = {
            "mode": action,
            "total": len(records),
            "drafted": 0,
            "sent": 0,
            "skipped": len(self.skippedRecords),
            "failed": 0,
            "recordFile": "",
        }
        sender = str(self.config.get("promotionSenderEmail") or "").strip()
        password = str(self.config.get("promotionSmtpAuthCode") or "")
        if not sender or not password:
            raise RuntimeError("阿里邮箱账号或第三方客户端安全密码未配置")
        if not records:
            summary["recordFile"] = str(self.writeRecords(self.skippedRecords))
            return summary

        waitSeconds = max(0.0, float(self.config.get("promotionWaitSeconds", 8)))
        if action == "draft":
            server = imaplib.IMAP4_SSL(self.promotionImapServer, self.promotionImapPort, timeout=30)
            try:
                server.login(sender, password)
                mailbox = self.findDraftMailbox(server)
                for record in records:
                    try:
                        message = self.buildMessage(record, sender)
                        status, response = server.append(
                            mailbox,
                            "(\\Draft)",
                            imaplib.Time2Internaldate(time.time()),
                            message.as_bytes(),
                        )
                        if status != "OK":
                            raise RuntimeError(str(response)[:240])
                        record["发送状态"] = "草稿已创建"
                        record["发送结果"] = "已写入阿里邮箱草稿箱"
                        self.data.recordMailAction(
                            record["邮箱"], "draft", record["对象键"], record["对象名称"],
                            self.subjectHash(record["邮件主题"]), mailbox,
                        )
                        summary["drafted"] += 1
                        self.log(f"草稿已创建：{record['对象名称']} -> {record['邮箱']}")
                    except Exception as error:
                        record["发送状态"] = "草稿创建失败"
                        record["发送结果"] = "邮箱服务器未保存草稿"
                        record["失败原因"] = str(error)[:300]
                        summary["failed"] += 1
                        self.log(f"草稿创建失败：{record['对象名称']} -> {record['邮箱']}，{str(error)[:200]}")
            finally:
                try:
                    server.logout()
                except Exception:
                    pass
        else:
            server = smtplib.SMTP_SSL(self.promotionServer, self.promotionPort, timeout=30)
            try:
                server.login(sender, password)
                for index, record in enumerate(records):
                    try:
                        message = self.buildMessage(record, sender)
                        server.sendmail(sender, [record["邮箱"]], message.as_bytes())
                        record["发送状态"] = "发送成功"
                        record["发送结果"] = "邮件已真实发送"
                        self.data.recordMailAction(
                            record["邮箱"], "send", record["对象键"], record["对象名称"],
                            self.subjectHash(record["邮件主题"]), "SMTP 发送成功",
                        )
                        summary["sent"] += 1
                        self.log(f"已发送：{record['对象名称']} -> {record['邮箱']}")
                    except Exception as error:
                        record["发送状态"] = "发送失败"
                        record["发送结果"] = "邮件未发送"
                        record["失败原因"] = str(error)[:300]
                        summary["failed"] += 1
                        self.log(f"发送失败：{record['对象名称']} -> {record['邮箱']}，{str(error)[:200]}")
                    if waitSeconds and index < len(records) - 1:
                        time.sleep(waitSeconds)
            finally:
                try:
                    server.quit()
                except Exception:
                    pass
        summary["recordFile"] = str(self.writeRecords(records + self.skippedRecords))
        return summary
