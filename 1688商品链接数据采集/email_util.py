"""本项目的邮件发送工具（独立模块，不与其他子项目共享）。"""

import mimetypes
import os
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

PROJECT_NAME = "1688商品链接数据采集"
DEFAULT_SUBJECT = f"自动化_{PROJECT_NAME}"
DEFAULT_SENDER_EMAIL = "1974419863@qq.com"
DEFAULT_SMTP_AUTH_CODE = os.getenv("SMTP_AUTH_CODE", "")


def send_email_with_files(sender_email, auth_code, receiver_email, subject, content, file_paths):
    """发送带多个附件的邮件（修复附件变成 bin 的问题）。"""
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    smtp_server = "smtp.qq.com"
    smtp_port = 465

    msg = MIMEMultipart()
    msg["From"] = formataddr(["发件人", sender_email])
    msg["To"] = formataddr(["收件人", receiver_email])
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(content, "plain", "utf-8"))

    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"警告：文件不存在，跳过: {file_path}")
            continue

        try:
            filename = os.path.basename(file_path)
            mime_type, _ = mimetypes.guess_type(file_path)
            if mime_type is None:
                mime_type = "application/octet-stream"

            main_type, sub_type = mime_type.split("/", 1)

            with open(file_path, "rb") as f:
                attachment = MIMEBase(main_type, sub_type)
                attachment.set_payload(f.read())
                encoders.encode_base64(attachment)

                try:
                    attachment.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=("utf-8", "", filename),
                    )
                except Exception:
                    ascii_filename = filename.encode("ascii", "ignore").decode("ascii") or "attachment"
                    attachment.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=ascii_filename,
                    )

                msg.attach(attachment)
                print(f"已添加附件: {filename} (MIME: {mime_type})")

        except Exception as e:
            print(f"添加附件失败 {file_path}: {e}")

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, auth_code)
        server.sendmail(sender_email, [receiver_email], msg.as_bytes())
        server.quit()
        print("邮件发送成功！")
        return True
    except Exception as e:
        print(f"发送失败: {e}")
        return False


def deliver_outputs(config, file_paths, subject=None, content=None):
    """sendEmail=True 时发送邮件附件。"""
    if not config.get("sendEmail"):
        return True

    paths = [file_paths] if isinstance(file_paths, str) else list(file_paths or [])
    paths = [str(p) for p in paths if p and Path(p).exists()]

    receiver = (config.get("email") or "").strip()
    if not receiver:
        raise ValueError("发送邮件时必须填写 email")

    sender = (
        config.get("sender_email")
        or os.getenv("SMTP_SENDER")
        or DEFAULT_SENDER_EMAIL
    ).strip()
    auth_code = (
        config.get("smtp_auth_code")
        or os.getenv("SMTP_AUTH_CODE")
        or DEFAULT_SMTP_AUTH_CODE
    ).strip()

    if not paths:
        print("未找到可发送的汇总文件，跳过邮件")
        return False

    return send_email_with_files(
        sender_email=sender,
        auth_code=auth_code,
        receiver_email=receiver,
        subject=subject or DEFAULT_SUBJECT,
        content=content or f"自动化_{PROJECT_NAME}导出的规格汇总文件",
        file_paths=paths,
    )
