"""审核通过后的推广邮件和运行结果通知。"""

import hashlib
import html
import imaplib
import mimetypes
import smtplib
import sys
import time
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

from openpyxl import Workbook

from data import Data


class Mail:
    """只发送审核通过的联系方式，并保留原中文发送记录。"""

    def __init__(
        self,
        data,
        config,
        outputDir,
        dataDir,
        log=None,
    ):
        """初始化邮件固定配置和文件目录。"""
        self.data = data
        self.config = config
        self.outputDir = Path(outputDir)
        self.dataDir = Path(dataDir)
        self.log = log or (lambda message: None)
        self.projectName = "TREC 公司+个人合作推广"
        self.promotionServer = str(
            config.get("promotionSmtpServer") or "smtp.qiye.aliyun.com"
        )
        self.promotionPort = int(config.get("promotionSmtpPort") or 465)
        self.promotionImapServer = str(
            config.get("promotionImapServer") or "imap.qiye.aliyun.com"
        )
        self.promotionImapPort = int(config.get("promotionImapPort") or 993)
        self.draftMailboxName = str(config.get("promotionDraftMailbox") or "Drafts")
        self.logoFileName = str(config.get("promotionLogoFile") or "time2renew-logo.png")
        self.reportServer = self.promotionServer
        self.reportPort = self.promotionPort
        self.recordName = "邮件发送记录.xlsx"
        self.skippedRecords = []
        self.defaultPromotionSubject = "Partner with us on agent CE renewals"
        self.defaultPromotionBody = (
            "Hi,\n"
            "Your agents' renewal season is coming up. We offer a full 18-hour "
            "TREC-approved CE package at $49.99 -- probably the lowest price they'll "
            "find. TREC Provider #11011-CEP.\n\n"
            "For every agent in your office who uses our package, we can provide a "
            "20% referral fee. If you are open to a quick conversation, please reply "
            "to this email.\n\n"
            "Best,\nTime2renew Support Team\nWebsite: https://time2renew.com"
        )

    def splitReceivers(self, value):
        """拆分逗号、分号和中文标点分隔的邮箱。"""
        text = str(value or "").replace("；", ";").replace("，", ",")
        output = []
        for item in text.replace(";", ",").split(","):
            email = item.strip().lower()
            if email and email not in output:
                output.append(email)
        return output

    def approvedRecords(self, action="draft"):
        """读取审核通过且未执行同类邮件动作的唯一邮箱。"""
        subject = str(self.config.get("promotionSubject") or self.defaultPromotionSubject)
        body = str(self.config.get("promotionBody") or self.defaultPromotionBody)
        records = []
        seen = set()
        self.skippedRecords = []
        for result in self.data.contactResults():
            if result.get("reviewStatus") != "approved":
                continue
            for email in result.get("emails") or []:
                receiver = str(email or "").strip().lower()
                if not receiver:
                    continue
                baseRecord = {
                    "来源类型": "公司" if result.get("mode") == "company" else "个人",
                    "对象键": result.get("objectKey", ""),
                    "对象名称": result.get("objectName", ""),
                    "许可证号": result.get("licenseNumber", ""),
                    "邮箱": receiver,
                    "邮件主题": subject,
                    "邮件正文": body,
                    "来源链接": "; ".join(
                        result.get("detailUrls") or result.get("sourceUrls") or []
                    ),
                    "审核状态": "已通过",
                    "发送状态": "待处理",
                    "发送结果": "",
                    "失败原因": "",
                }
                if receiver in seen:
                    duplicate = dict(baseRecord)
                    duplicate["发送状态"] = "已跳过"
                    duplicate["发送结果"] = "同一批次邮箱重复"
                    self.skippedRecords.append(duplicate)
                    continue
                seen.add(receiver)
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
        message["From"] = formataddr((str(Header("Time2renew Support Team", "utf-8")), sender))
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
        else:
            self.log(f"邮件正文 Logo 未找到：{logoPath}")
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

        waitSeconds = max(0.0, float(self.config.get("promotionWaitSeconds", 10)))
        if action == "draft":
            server = imaplib.IMAP4_SSL(
                self.promotionImapServer,
                self.promotionImapPort,
                timeout=30,
            )
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
                            record["邮箱"],
                            "draft",
                            record["对象键"],
                            record["对象名称"],
                            self.subjectHash(record["邮件主题"]),
                            mailbox,
                        )
                        summary["drafted"] += 1
                    except Exception as error:
                        record["发送状态"] = "草稿创建失败"
                        record["发送结果"] = "邮箱服务器未保存草稿"
                        record["失败原因"] = str(error)[:300]
                        summary["failed"] += 1
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
                            record["邮箱"],
                            "send",
                            record["对象键"],
                            record["对象名称"],
                            self.subjectHash(record["邮件主题"]),
                            "SMTP 发送成功",
                        )
                        summary["sent"] += 1
                    except Exception as error:
                        record["发送状态"] = "发送失败"
                        record["发送结果"] = "邮件未发送"
                        record["失败原因"] = str(error)[:300]
                        summary["failed"] += 1
                    if waitSeconds and index < len(records) - 1:
                        time.sleep(waitSeconds)
            finally:
                try:
                    server.quit()
                except Exception:
                    pass
        summary["recordFile"] = str(self.writeRecords(records + self.skippedRecords))
        return summary

    def sendApproved(self, execute=False):
        """兼容旧入口；默认创建服务器草稿，execute=True 才真实发送。"""
        return self.processApproved("send" if execute else "draft")

    def addAttachment(self, message, path):
        """把一个存在的 Excel 文件加入通知邮件。"""
        if not path.exists():
            return False
        mimeType, _ = mimetypes.guess_type(str(path))
        mainType, subType = (mimeType or "application/octet-stream").split("/", 1)
        with path.open("rb") as fileObject:
            attachment = MIMEBase(mainType, subType)
            attachment.set_payload(fileObject.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition", "attachment", filename=("utf-8", "", path.name)
        )
        message.attach(attachment)
        return True

    def sendReport(self, summary):
        """按开关发送四个固定数据表附件，不发送推广邮件。"""
        if not bool(self.config.get("sendEmail", False)):
            return True
        sender = str(self.config.get("reportSenderEmail") or "").strip()
        password = str(self.config.get("reportSmtpAuthCode") or "")
        receivers = self.splitReceivers(self.config.get("reportReceivers"))
        if not sender or not password or not receivers:
            raise RuntimeError("结果通知邮箱配置不完整")
        paths = [
            self.dataDir / "初始总量数据未清洗.xlsx",
            self.dataDir / "已获取到的初始总数据.xlsx",
            self.outputDir / str(self.config.get("companyResultFileName")),
            self.outputDir / str(self.config.get("personResultFileName")),
        ]
        message = MIMEMultipart()
        message["From"] = formataddr((str(Header(self.projectName, "utf-8")), sender))
        message["To"] = ", ".join(receivers)
        message["Subject"] = Header(
            str(self.config.get("reportSubject") or "自动化_TREC公司+个人合作推广数据"),
            "utf-8",
        ).encode()
        body = (
            "TREC 公司+个人合作推广数据已完成。\n\n"
            f"公司完成：{summary.get('completed', {}).get('company', 0)}\n"
            f"个人完成：{summary.get('completed', {}).get('person', 0)}\n"
            f"待审核：{summary.get('reviewPending', 0)}"
        )
        message.attach(MIMEText(body, "plain", "utf-8"))
        added = sum(1 for path in paths if self.addAttachment(message, path))
        if not added:
            return False
        server = smtplib.SMTP_SSL(self.reportServer, self.reportPort, timeout=30)
        try:
            server.login(sender, password)
            server.sendmail(sender, receivers, message.as_bytes())
        finally:
            server.quit()
        self.log(f"结果通知邮件已发送，附件数量：{added}")
        return True
