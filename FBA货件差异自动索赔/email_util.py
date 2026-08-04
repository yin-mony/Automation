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

PROJECT_NAME = "FBA货件差异自动索赔"
DEFAULT_SUBJECT = f"自动化_{PROJECT_NAME}"
DEFAULT_SENDER_EMAIL = "1974419863@qq.com"
DEFAULT_SMTP_AUTH_CODE = "ucvopobstjhobbef"
DEFAULT_SMTP_SERVER = "smtp.qq.com"
DEFAULT_SMTP_PORT = 465


def formatMailAddr(name, email):
    """按 UTF-8 编码邮件显示名，避免部分邮箱客户端显示乱码。"""
    return formataddr((str(Header(name, "utf-8")), email))


def send_email_with_files(sender_email, auth_code, receiver_email, subject, content, file_paths, smtp_server=None, smtp_port=None):
    """发送带多个附件的邮件。"""
    # 统一附件入参格式，调用方可传单个路径或路径列表
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    # SMTP 服务地址和端口由配置传入，缺省时使用项目默认值
    smtp_server = str(smtp_server or DEFAULT_SMTP_SERVER).strip()
    smtp_port = int(smtp_port or DEFAULT_SMTP_PORT)

    # 邮件标题、显示名与正文全部在 Python UTF-8 字符串内构造，避免 PowerShell 中转导致中文变问号
    msg = MIMEMultipart()
    msg["From"] = formatMailAddr("发件人", sender_email)
    msg["To"] = formatMailAddr("收件人", receiver_email)
    msg["Subject"] = Header(str(subject), "utf-8").encode()
    msg.attach(MIMEText(str(content), "plain", "utf-8"))

    for file_path in file_paths:
        # 附件不存在时跳过当前文件，避免单个坏路径阻断整封邮件
        if not os.path.exists(file_path):
            print(f"警告：文件不存在，跳过: {file_path}")
            continue

        try:
            # 根据文件名推断 MIME 类型，无法识别时按二进制附件处理
            filename = os.path.basename(file_path)
            mime_type, _ = mimetypes.guess_type(file_path)
            if mime_type is None:
                mime_type = "application/octet-stream"

            main_type, sub_type = mime_type.split("/", 1)

            with open(file_path, "rb") as f:
                # 读取附件二进制并使用 base64 编码，兼容 PDF 等非文本文件
                attachment = MIMEBase(main_type, sub_type)
                attachment.set_payload(f.read())
                encoders.encode_base64(attachment)

                try:
                    # 优先按 UTF-8 写入附件文件名，保留中文与空格
                    attachment.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=("utf-8", "", filename),
                    )
                except Exception:
                    # 极端客户端不支持 UTF-8 文件名时，回退为 ASCII 文件名
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
        # 登录 SMTP 并发送完整 MIME 邮件
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
    """sendEmail=True 时发送 POP 文档附件。"""
    # GUI 未开启邮件通知时直接视为发送成功
    if not config.get("sendEmail"):
        return True

    # 只保留真实存在的 POP 文件，避免邮件携带空附件
    paths = [file_paths] if isinstance(file_paths, str) else list(file_paths or [])
    paths = [str(p) for p in paths if p and Path(p).exists()]

    # 邮件接收人来自 GUI 公共配置
    receiver = (config.get("email") or "").strip()
    if not receiver:
        raise ValueError("发送邮件时必须填写 email")

    # SMTP 账号优先读取运行配置，缺省时使用项目内默认值
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
    smtpServer = (
        config.get("smtp_server")
        or os.getenv("SMTP_SERVER")
        or DEFAULT_SMTP_SERVER
    ).strip()
    smtpPort = (
        config.get("smtp_port")
        or os.getenv("SMTP_PORT")
        or DEFAULT_SMTP_PORT
    )

    if not paths:
        print("未找到可发送的 POP 文件，跳过邮件")
        return False

    # 赛狐流程邮件：只发送本轮生成的 POP 文件
    return send_email_with_files(
        sender_email=sender,
        auth_code=auth_code,
        receiver_email=receiver,
        subject=subject or DEFAULT_SUBJECT,
        content=content or f"自动化_{PROJECT_NAME}生成的 POP 文档，共 {len(paths)} 个文件。",
        file_paths=paths,
        smtp_server=smtpServer,
        smtp_port=smtpPort,
    )


def deliverCase(config, resultList, failList=None, skipList=None, caseResultPath=None):
    """sendEmail=True 时发送易得客 CASE 结果汇总与对应 POP 附件。"""
    # GUI 未开启邮件通知时直接视为发送成功
    if not config.get("sendEmail"):
        return True

    # CASE 结果邮件使用公共接收邮箱
    receiver = (config.get("email") or "").strip()
    if not receiver:
        raise ValueError("发送 CASE 邮件时必须填写 email")

    # SMTP 账号优先读取运行配置，缺省时使用项目内默认值
    sender = (
        config.get("sender_email")
        or os.getenv("SMTP_SENDER")
        or DEFAULT_SENDER_EMAIL
    ).strip()
    authCode = (
        config.get("smtp_auth_code")
        or os.getenv("SMTP_AUTH_CODE")
        or DEFAULT_SMTP_AUTH_CODE
    ).strip()
    smtpServer = (
        config.get("smtp_server")
        or os.getenv("SMTP_SERVER")
        or DEFAULT_SMTP_SERVER
    ).strip()
    smtpPort = (
        config.get("smtp_port")
        or os.getenv("SMTP_PORT")
        or DEFAULT_SMTP_PORT
    )

    resultList = list(resultList or [])
    failList = list(failList or [])
    skipList = list(skipList or [])
    # 相对 POP 文件名以 CASE 结果文件所在目录为基准解析
    baseDir = Path(caseResultPath).parent if caseResultPath else None
    filePaths = []
    seenPaths = set()

    # CASE 汇总正文包含成功、失败、跳过三类结果，方便业务直接查阅
    lines = [
        "FBA 货件差异自动索赔 CASE 提交结果",
        "",
        f"本次成功提交/记录 {len(resultList)} 个货件。",
    ]

    if resultList:
        lines.append("")
        lines.append("CASE 结果：")
        for index, item in enumerate(resultList, start=1):
            # 成功结果保留货件编号、CASE 问题编号、POP 文件名与状态
            shipmentId = str(item.get("shipmentId") or "").strip()
            caseId = str(item.get("caseId") or "").strip()
            popFile = str(item.get("popFile") or "").strip()
            popPathText = str(item.get("popPath") or popFile).strip()
            status = str(item.get("status") or "提交成功").strip()
            lines.append(f"{index}. 差异货件编号：{shipmentId}，CASE 问题编号：{caseId}，状态：{status}")
            if popFile:
                lines.append(f"   POP 文件：{popFile}")
            if popPathText:
                # 同一 POP 附件只添加一次，避免重复附件
                popPath = Path(popPathText)
                if not popPath.is_absolute() and baseDir:
                    popPath = baseDir / popPathText
                if popPath.exists() and str(popPath) not in seenPaths:
                    filePaths.append(str(popPath))
                    seenPaths.add(str(popPath))

    if failList:
        lines.append("")
        lines.append("失败明细：")
        for index, item in enumerate(failList, start=1):
            # 失败明细保留原因，便于后续人工补跑
            shipmentId = str(item.get("shipmentId") or "").strip()
            reason = str(item.get("reason") or "").strip()
            lines.append(f"{index}. 货件编号：{shipmentId}，原因：{reason}")

    if skipList:
        lines.append("")
        lines.append("跳过明细：")
        for index, item in enumerate(skipList, start=1):
            # 跳过明细通常用于无差异货件或已有 CASE 的情况
            shipmentId = str(item.get("shipmentId") or "").strip()
            reason = str(item.get("reason") or "").strip()
            lines.append(f"{index}. 货件编号：{shipmentId}，原因：{reason}")

    if caseResultPath:
        lines.append("")
        lines.append(f"本地结果文件：{caseResultPath}")

    lines.append("")
    lines.append("附件为本次成功提交/记录货件对应的 POP PDF 文件。")

    # 易得客流程邮件：发送 CASE 汇总正文与对应 POP PDF 附件
    return send_email_with_files(
        sender_email=sender,
        auth_code=authCode,
        receiver_email=receiver,
        subject="FBA货件差异自动索赔-CASE提交结果",
        content="\n".join(lines),
        file_paths=filePaths,
        smtp_server=smtpServer,
        smtp_port=smtpPort,
    )
