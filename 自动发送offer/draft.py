"""生成本地 EML 并写入企业邮箱草稿箱。"""

import imaplib
import re
import smtplib
from datetime import datetime
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid
from html import escape
from pathlib import Path


EMAIL_PATTERN = re.compile(r"[^@\s,;，；|]+@[^@\s,;，；|]+\.[^@\s,;，；|]+")


def parseAddressList(value):
    """解析逗号、分号、换行或竖线分隔的邮箱列表。"""
    if isinstance(value, (list, tuple, set)):
        rawItems = []
        for item in value:
            rawItems.extend(re.split(r"[,;，；|\n]+", str(item or "")))
    else:
        rawItems = re.split(r"[,;，；|\n]+", str(value or ""))
    addresses = [item.strip() for item in rawItems if item.strip()]
    invalid = [item for item in addresses if not EMAIL_PATTERN.fullmatch(item)]
    if invalid:
        raise ValueError(f"邮箱格式不正确：{'; '.join(invalid)}")
    return list(dict.fromkeys(addresses))


class MailDraft:
    """创建带 Offer PDF 附件的可编辑邮件草稿。"""

    def __init__(self, settings):
        self.settings = settings

    def buildMessage(self, values, pdfPath):
        """构建候选人 Offer 邮件。"""
        message = EmailMessage()
        message["From"] = self.settings.mailFrom
        message["To"] = values["email"]
        ccEmails = parseAddressList(values.get("ccEmails", ""))
        if ccEmails:
            message["Cc"] = ", ".join(ccEmails)
        message["Subject"] = f"{self.settings.companyName}录用通知书 - {values['name']}"
        message["Date"] = format_datetime(datetime.now().astimezone())
        message["Message-ID"] = make_msgid(domain=self.settings.mailFrom.split("@")[-1])
        message["X-Unsent"] = "1"
        textBody = (
            f"{values['name']}，您好：\n\n"
            f"很高兴通知您，经过面试与评估，我们诚邀您加入{self.settings.companyName}，"
            f"担任{values['position']}，入职部门为{values['department']}。\n\n"
            "正式录用通知书见附件，请仔细核对岗位、薪酬、入职日期及报到安排。\n\n"
            f"入职日期：{values['entryDate']}\n"
            f"报到时间：{values['reportTime']}\n"
            f"报到地点：{values['reportLocation']}\n\n"
            "如确认接受录用，请准备以下入职资料：\n"
            "1. 原单位出具并加盖公章的离职证明原件（应届毕业生无需提供）；\n"
            "2. 身份证原件及复印件、最高学历学位证书、技术职称或资历证书原件；\n"
            f"3. {values['salaryBank']}银行卡复印件，并注明卡号及开户行；\n"
            "4. 三个月内且包含胸透项目的体检报告。\n\n"
            f"请您于{values['responseDeadline']}前仔细审阅本邮件及附件，并通过电子邮件“全部回复”的方式给予书面确认。"
            "如确认接受本次录用，请回复：“本人已完整阅读并理解本录用通知书所载内容，接受贵公司的录用安排，"
            "并确认将按通知约定的时间办理报到及入职手续。”如您对录用事项存在疑问、需要协商，或无法接受本次录用，"
            "请在上述期限内回复本邮件并说明具体情况。\n\n"
            f"联系人：{values['hrName']}\n联系电话：{values['hrPhone']}\n\n"
            f"{self.settings.companyName}"
        )
        htmlBody = (
            f"<p>{escape(values['name'])}，您好：</p>"
            f"<p>很高兴通知您，经过面试与评估，我们诚邀您加入{escape(self.settings.companyName)}，"
            f"担任<strong>{escape(values['position'])}</strong>，入职部门为<strong>{escape(values['department'])}</strong>。</p>"
            "<p>正式录用通知书见附件，请仔细核对岗位、薪酬、入职日期及报到安排。</p>"
            f"<p><strong>入职日期：</strong>{escape(values['entryDate'])}<br>"
            f"<strong>报到时间：</strong>{escape(values['reportTime'])}<br>"
            f"<strong>报到地点：</strong>{escape(values['reportLocation'])}</p>"
            "<p><strong>如确认接受录用，请准备以下入职资料：</strong></p>"
            "<ol>"
            "<li>原单位出具并加盖公章的离职证明原件（应届毕业生无需提供）；</li>"
            "<li>身份证原件及复印件、最高学历学位证书、技术职称或资历证书原件；</li>"
            f"<li>{escape(values['salaryBank'])}银行卡复印件，并注明卡号及开户行；</li>"
            "<li>三个月内且包含胸透项目的体检报告。</li>"
            "</ol>"
            f"<p>请您于<strong>{escape(values['responseDeadline'])}</strong>前仔细审阅本邮件及附件，并通过电子邮件“全部回复”的方式给予书面确认。"
            "如确认接受本次录用，请回复：<strong>“本人已完整阅读并理解本录用通知书所载内容，接受贵公司的录用安排，"
            "并确认将按通知约定的时间办理报到及入职手续。”</strong>如您对录用事项存在疑问、需要协商，或无法接受本次录用，"
            "请在上述期限内回复本邮件并说明具体情况。</p>"
            f"<p>联系人：{escape(values['hrName'])}<br>联系电话：{escape(values['hrPhone'])}</p>"
            f"<p>{escape(self.settings.companyName)}</p>"
        )
        message.set_content(textBody)
        message.add_alternative(htmlBody, subtype="html")
        pdf = Path(pdfPath)
        message.add_attachment(pdf.read_bytes(), maintype="application", subtype="pdf", filename=pdf.name)
        return message

    def saveLocal(self, message, outputPath):
        """保存可由桌面邮件客户端打开的 EML 草稿。"""
        path = Path(outputPath)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(message.as_bytes())
        return path

    def appendServer(self, message):
        """通过 IMAP 将邮件追加到企业邮箱草稿箱。"""
        self.settings.validateDraft()
        if self.settings.imapUseSsl:
            client = imaplib.IMAP4_SSL(self.settings.imapHost, self.settings.imapPort)
        else:
            client = imaplib.IMAP4(self.settings.imapHost, self.settings.imapPort)
            client.starttls()
        try:
            client.login(self.settings.smtpUsername, self.settings.smtpPassword)
            status, _ = client.append(
                self.settings.draftFolder,
                r"(\Draft \Seen)",
                imaplib.Time2Internaldate(datetime.now().timestamp()),
                message.as_bytes(),
            )
            if status != "OK":
                raise RuntimeError("企业邮箱拒绝写入草稿箱")
        finally:
            try:
                client.logout()
            except Exception:
                pass

    def sendMessage(self, message):
        """通过 SMTP 真实发送已经确认的 Offer 邮件。"""
        self.settings.validateSend()
        if "X-Unsent" in message:
            del message["X-Unsent"]
        if self.settings.smtpUseSsl:
            client = smtplib.SMTP_SSL(self.settings.smtpHost, self.settings.smtpPort, timeout=20)
        else:
            client = smtplib.SMTP(self.settings.smtpHost, self.settings.smtpPort, timeout=20)
            client.starttls()
        try:
            client.login(self.settings.smtpUsername, self.settings.smtpPassword)
            client.send_message(message)
        finally:
            try:
                client.quit()
            except Exception:
                pass

    def create(self, values, pdfPath, emlPath, saveServer=True, sendNow=False):
        """创建并留存草稿，按确认结果决定是否真实发送。"""
        message = self.buildMessage(values, pdfPath)
        ccEmails = parseAddressList(values.get("ccEmails", ""))
        ccNames = str(values.get("ccNames") or "").strip()
        localPath = self.saveLocal(message, emlPath)
        serverSaved = False
        serverError = ""
        sent = False
        sendError = ""
        if saveServer:
            try:
                self.appendServer(message)
                serverSaved = True
            except Exception as exc:
                serverError = str(exc)
        if sendNow:
            if not serverSaved:
                sendError = f"草稿箱保存失败，邮件未发送：{serverError or '未启用企业邮箱草稿留存'}"
            else:
                try:
                    self.sendMessage(message)
                    sent = True
                except Exception as exc:
                    sendError = str(exc)
        return {
            "emlPath": str(localPath),
            "serverSaved": serverSaved,
            "serverError": serverError,
            "sent": sent,
            "sendError": sendError,
            "ccEmails": ccEmails,
            "ccNames": ccNames,
        }
