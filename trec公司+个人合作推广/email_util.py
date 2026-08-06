"""TREC 公司+个人合作推广邮件发送工具。"""

import mimetypes
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path


class EmailUtil:
    """只负责发送 TREC 项目最终四个数据表附件。"""

    # 项目名称：用于邮件标题、发件显示名和日志输出。
    projectName = "TREC 公司+个人合作推广"

    # 默认邮件标题：调用方没有覆盖 emailSubject 时使用。
    defaultSubject = "自动化_TREC公司+个人合作推广数据"

    # SMTP 服务器：当前沿用 QQ 邮箱 SMTP SSL。
    smtpServer = "smtp.qq.com"

    # SMTP SSL 端口：QQ 邮箱固定使用 465。
    smtpPort = 465

    @staticmethod
    def cleanText(value):
        """把空值统一转换成安全文本。"""
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def formatMailAddr(name, email):
        """按 UTF-8 编码邮件显示名，避免中文乱码。"""
        return formataddr((str(Header(name, "utf-8")), email))

    @staticmethod
    def splitReceivers(receiverText):
        """把逗号或分号分隔的收件邮箱拆成列表。"""
        text = EmailUtil.cleanText(receiverText)
        text = text.replace("；", ";").replace("，", ",")
        receivers = []
        for part in text.replace(";", ",").split(","):
            email = part.strip()
            if email:
                receivers.append(email)
        return receivers

    @classmethod
    def collectTrecFiles(cls, config, outputDir, dataDir):
        """按固定白名单收集四个 TREC 数据表路径。"""
        # dataPath：内置底表目录，固定对应项目 file 目录。
        dataPath = Path(dataDir or "file")

        # outputPath：最终搜索结果表所在目录，不能和 file 内置数据目录混用。
        outputPath = Path(outputDir or config.get("outputDir") or "output")

        # rawFileName：TREC 网站采集得到的未清洗全量底表。
        rawFileName = config.get("rawFileName", "初始总量数据未清洗.xlsx")

        # cleanFileName：融合 Open Data 后得到的已清洗初始表。
        cleanFileName = config.get("cleanFileName", "已获取到的初始总数据.xlsx")

        # companyFileName：公司模式最终搜索匹配结果表。
        companyFileName = config.get("companyResultFileName", "已完成搜索匹配的公司联系信息数据.xlsx")

        # personFileName：个人模式最终搜索匹配结果表。
        personFileName = config.get("personResultFileName", "已完成搜索匹配的个人联系信息数据.xlsx")

        # fileList：邮件附件白名单，顺序固定，数量最多四个。
        fileList = [
            dataPath / rawFileName,
            dataPath / cleanFileName,
            outputPath / companyFileName,
            outputPath / personFileName,
        ]
        return fileList

    @classmethod
    def buildContent(cls, config, filePaths, missingPaths, summary=None):
        """生成 TREC 数据邮件正文。"""
        # summary：main.py 传入的本次运行数量统计。
        summary = dict(summary or {})

        # environment：本机/线上运行状态，方便收到邮件后判断来源。
        environment = summary.get("environment") or ("线上" if config.get("isOnline") else "本机")

        # lines：正文只描述本次数据表，不再出现其他子项目的 POP 或 CASE 文案。
        lines = [
            "TREC 公司+个人合作推广数据已完成。",
            "",
            f"运行环境：{environment}",
            f"初始数据行数：{summary.get('rows', 0)}",
            f"公司模式结果数：{summary.get('companyResults', 0)}",
            f"个人模式结果数：{summary.get('personResults', 0)}",
            "",
            "本邮件附件只包含以下四类数据表：",
        ]

        # existingNameList：实际存在并准备发送的附件文件名。
        existingNameList = [Path(path).name for path in filePaths]
        for index, fileName in enumerate(existingNameList, start=1):
            lines.append(f"{index}. {fileName}")

        # missingPaths：缺失文件会写入正文提醒，但不会添加其他替代附件。
        if missingPaths:
            lines.append("")
            lines.append("以下数据表当前不存在，未作为附件发送：")
            for path in missingPaths:
                lines.append(f"- {Path(path).name}")

        return "\n".join(lines)

    @classmethod
    def addAttachment(cls, message, filePath):
        """向邮件对象添加一个 Excel 附件。"""
        # path：附件必须是真实存在的文件。
        path = Path(filePath)
        if not path.exists():
            print("附件不存在，跳过:", path)
            return False

        # mimeType：根据文件名推断 MIME 类型，Excel 识别失败时按二进制处理。
        mimeType, _ = mimetypes.guess_type(str(path))
        if not mimeType:
            mimeType = "application/octet-stream"
        mainType, subType = mimeType.split("/", 1)

        # attachment：读取二进制内容并做 base64 编码。
        with path.open("rb") as fileObject:
            attachment = MIMEBase(mainType, subType)
            attachment.set_payload(fileObject.read())
        encoders.encode_base64(attachment)

        # filename：按 UTF-8 写入中文附件名。
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=("utf-8", "", path.name),
        )
        message.attach(attachment)
        print("已添加邮件附件:", path.name)
        return True

    @classmethod
    def sendEmailWithFiles(cls, senderEmail, authCode, receiverEmail, subject, content, filePaths):
        """发送带多个 Excel 附件的 TREC 数据邮件。"""
        # receivers：支持一个或多个收件邮箱，多个邮箱用英文逗号或分号分隔。
        receivers = cls.splitReceivers(receiverEmail)
        if not receivers:
            raise ValueError("发送邮件时必须填写接收邮箱")

        # message：构建标准 MIME 邮件，正文使用 UTF-8 文本。
        message = MIMEMultipart()
        message["From"] = cls.formatMailAddr(cls.projectName, senderEmail)
        message["To"] = ", ".join(receivers)
        message["Subject"] = Header(str(subject), "utf-8").encode()
        message.attach(MIMEText(str(content), "plain", "utf-8"))

        # addedCount：实际成功加入邮件的附件数量。
        addedCount = 0
        for filePath in filePaths:
            if cls.addAttachment(message, filePath):
                addedCount += 1

        if addedCount == 0:
            print("没有可发送的 TREC 数据表附件，跳过邮件。")
            return False

        # server：登录 SMTP 并发送邮件。
        server = smtplib.SMTP_SSL(cls.smtpServer, cls.smtpPort)
        try:
            server.login(senderEmail, authCode)
            server.sendmail(senderEmail, receivers, message.as_bytes())
        finally:
            server.quit()

        print("TREC 数据邮件发送成功，附件数量:", addedCount)
        return True

    @classmethod
    def deliverOutputs(cls, config, outputDir=None, dataDir=None, summary=None):
        """sendEmail=True 时发送 TREC 最终四个数据表附件。"""
        # 未开启邮件时直接返回 True，主流程不需要额外判断。
        if not config.get("sendEmail"):
            return True

        # receiverEmail：收件邮箱来自 GUI 或 main.py 公共配置。
        receiverEmail = cls.cleanText(config.get("email"))
        if not receiverEmail:
            raise ValueError("已开启邮件发送，但没有填写接收邮箱 email")

        # senderEmail：发件邮箱只读取 main.py 固定配置，不交给 GUI 或环境变量修改。
        senderEmail = cls.cleanText(config.get("sender_email"))
        if not senderEmail:
            raise ValueError("已开启邮件发送，但 main.py 固定发件邮箱 sender_email 为空")

        # authCode：SMTP 授权码只读取 main.py 固定配置，不交给 GUI 或环境变量修改。
        authCode = cls.cleanText(config.get("smtp_auth_code"))
        if not authCode:
            raise ValueError("已开启邮件发送，但 main.py 固定 SMTP 授权码 smtp_auth_code 为空")

        # expectedPaths：只按四个固定 TREC 数据表白名单发送。
        expectedPaths = cls.collectTrecFiles(config=config, outputDir=outputDir, dataDir=dataDir)

        # validPaths：真实存在的白名单附件。
        validPaths = []

        # missingPaths：白名单里当前不存在的附件，只进入正文提醒。
        missingPaths = []
        for path in expectedPaths:
            path = Path(path)
            if path.exists():
                validPaths.append(path)
            else:
                missingPaths.append(path)

        # content：正文说明运行情况、附件范围和缺失文件。
        content = cls.buildContent(
            config=config,
            filePaths=validPaths,
            missingPaths=missingPaths,
            summary=summary,
        )

        # subject：邮件标题优先读取公共配置。
        subject = cls.cleanText(config.get("emailSubject")) or cls.defaultSubject

        return cls.sendEmailWithFiles(
            senderEmail=senderEmail,
            authCode=authCode,
            receiverEmail=receiverEmail,
            subject=subject,
            content=content,
            filePaths=validPaths,
        )
