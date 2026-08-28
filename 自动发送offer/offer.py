"""根据人事复核结果生成正式 Offer PDF。"""

from html import escape
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


class OfferPdf:
    """按录用通知模板内容输出中文 PDF 附件。"""

    def __init__(self, settings):
        self.settings = settings
        self.fontName = "OfferChinese"
        pdfmetrics.registerFont(TTFont(self.fontName, str(settings.fontPath)))

    def text(self, value):
        """转义 PDF 富文本中的动态内容。"""
        return escape(str(value or ""))

    def noBreak(self, value):
        """保护日期、金额等短语，避免在短语内部换行。"""
        return f"<nobr>{self.text(value)}</nobr>"

    def styles(self):
        """创建正式录用通知的版式样式。"""
        base = getSampleStyleSheet()
        return {
            "title": ParagraphStyle("OfferTitle", parent=base["Title"], fontName=self.fontName, fontSize=24, leading=32, alignment=TA_CENTER, textColor=colors.HexColor("#17221d"), spaceAfter=3 * mm),
            "subtitle": ParagraphStyle("OfferSubtitle", parent=base["Normal"], fontName=self.fontName, fontSize=10, leading=14, alignment=TA_CENTER, textColor=colors.HexColor("#4d5b54"), spaceAfter=9 * mm),
            "heading": ParagraphStyle("OfferHeading", parent=base["Heading2"], fontName=self.fontName, fontSize=13, leading=18, textColor=colors.HexColor("#176b45"), spaceBefore=3 * mm, spaceAfter=2 * mm),
            "body": ParagraphStyle("OfferBody", parent=base["BodyText"], fontName=self.fontName, fontSize=10, leading=16, alignment=TA_LEFT, textColor=colors.HexColor("#1f2924"), spaceAfter=2 * mm),
            "bodyIndented": ParagraphStyle("OfferBodyIndented", parent=base["BodyText"], fontName=self.fontName, fontSize=10, leading=16, alignment=TA_JUSTIFY, firstLineIndent=20, wordWrap="CJK", textColor=colors.HexColor("#1f2924"), spaceAfter=2 * mm),
            "small": ParagraphStyle("OfferSmall", parent=base["BodyText"], fontName=self.fontName, fontSize=9, leading=15, textColor=colors.HexColor("#526059")),
        }

    def valueTable(self, values, styles):
        """生成录用信息表。"""
        rows = [
            ("姓名", values["name"]), ("职位", values["position"]),
            ("部门", values["department"]), ("直属上级", values["reportPosition"]),
            ("薪酬定级", values["salaryGrade"]), ("试用期", f"{values['probationMonths']}个月"),
            ("试用期薪酬", values["probationSalary"]), ("转正薪酬", values["regularSalary"]),
            ("入职日期", values["entryDate"]), ("试岗最后日期", values["trialEndDate"]),
            ("报到时间", values["reportTime"]),
        ]
        data = [[Paragraph(self.text(label), styles["small"]), Paragraph(self.text(value), styles["body"])] for label, value in rows]
        table = Table(data, colWidths=[35 * mm, 120 * mm], repeatRows=0)
        table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), self.fontName),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eef4f1")),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#355148")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd7d1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return table

    def generate(self, values, outputPath):
        """生成并返回 Offer PDF 路径。"""
        output = Path(outputPath)
        output.parent.mkdir(parents=True, exist_ok=True)
        styles = self.styles()
        document = SimpleDocTemplate(
            str(output), pagesize=A4, leftMargin=23 * mm, rightMargin=23 * mm,
            topMargin=20 * mm, bottomMargin=18 * mm,
            title=f"{values['name']}录用通知书", author=self.settings.companyName,
        )
        story = [
            Paragraph("录 用 通 知", styles["title"]),
            Paragraph("OFFER LETTER", styles["subtitle"]),
            Paragraph(f"尊敬的 {self.text(values['name'])} 先生/女士，您好！", styles["body"]),
            Paragraph(f"经严格甄选，我们非常高兴地邀请您加入公司，希望 {self.text(self.settings.companyName)} 成为您创造价值、实现自我的新平台。公司尊重追求美好精神和物质生活的愿望，崇尚大家用自己的辛勤劳动和专业技能，与企业一道实现理想和愿景，共达双赢！", styles["bodyIndented"]),
            Paragraph("录用情况", styles["heading"]),
            self.valueTable(values, styles),
            Spacer(1, 4 * mm),
            Paragraph(
                f"我们邀请您于 {self.noBreak(values['entryDate'])} 入职。双方约定先试岗5个工作日，"
                f"{self.noBreak(values['trialEndDate'])} 为最后一个试岗日。试岗期间，如经公司评估未达到岗位要求，"
                f"公司将终止试岗，并按{self.noBreak('100元/日')}结算试岗期间的劳动所得；如员工在试岗期间主动提出退出，"
                f"则按{self.noBreak('50元/日')}结算试岗期间的劳动所得。入职购买社保（当月15号后入职的次月购买社保）；"
                f"试用期 {self.text(values['probationMonths'])} 个月结束后，绩效考核符合公司岗位需求的，以转正薪资发放并购买公积金。"
                "绩效考核考察价值观、可量化的关键工作指标及公司经营指标。根据个人能力和表现，符合或超出职位要求者，可以提前转正。",
                styles["bodyIndented"],
            ),
            Paragraph("试岗及绩效考核说明", styles["heading"]),
            Paragraph("1. 试岗期（5个工作日）不参与绩效考核，只计算固定薪资；", styles["body"]),
            Paragraph("2. 试岗通过后第一个工作日开始绩效考核周期；", styles["body"]),
            Paragraph("3. 当月绩效考核周期天数少于当月工作日的50%，绩效分数为0；", styles["body"]),
            Paragraph("4. 当月绩效考核周期天数达到当月工作日的50%但不足100%，考核原则上按合格分数计算，工作表现良好可酌情加分。", styles["body"]),
            Paragraph("入职相关手续", styles["heading"]),
            Paragraph(f"报到地点：{self.text(values['reportLocation'])}", styles["body"]),
            Paragraph(f"入职日期及报到时间：{self.text(values['entryDate'])} {self.text(values['reportTime'])}", styles["body"]),
            Paragraph("入职携带资料", styles["heading"]),
            Paragraph("1. 原单位出具并加盖公章的离职证明原件，应届毕业生无需提供；", styles["body"]),
            Paragraph("2. 身份证原件及复印件、最高学历学位证书、技术职称或资历证书原件；", styles["body"]),
            Paragraph(f"3. 公司使用 {self.text(values['salaryBank'])} 作为工资卡，请在入职时提供银行卡复印件，并注明卡号及开户行；", styles["body"]),
            Paragraph("4. 三个月内体检报告，体检须包含胸透项目；体检不符合录用条件的，不予录用。", styles["body"]),
            Paragraph("入职须知", styles["heading"]),
            Paragraph("以上所有资料均为入职必备资料，将作为您的人事资料存档。", styles["bodyIndented"]),
            Paragraph("您若与其他公司签有竞业限制协议，请及时告知。若未如实提供信息，由此产生的责任由本人承担。", styles["bodyIndented"]),
            Paragraph("若电话号码、现住址或身份证地址有变化，请及时告知公司，否则由本人承担相应责任。", styles["bodyIndented"]),
            Paragraph("新员工请按通知内容完善入职手续；入职需提交的资料未齐全者，请在入职七日内补齐。", styles["bodyIndented"]),
            Paragraph("薪酬保密制度是公司管理的重要原则，请勿打探或向他人透露薪酬情况。违反薪酬保密制度的，公司可据此解除本聘用书或劳动合同关系。", styles["bodyIndented"]),
            Paragraph("在双方未正式建立劳动关系之前，如出现包括但不限于以下情形，我司有权取消录用决定：与原单位签订的竞业限制协议尚未解除；存在违反职业道德和职业操守的行为；涉嫌犯罪正被调查处理或曾被追究刑事责任；存在严重违纪行为；背景调查结果不符；公司组织架构调整、岗位变动或用工需求发生变更等。", styles["bodyIndented"]),
            Paragraph("录用确认", styles["heading"]),
            Paragraph(f"请认真阅读本录用通知。若确认无误，请于 {self.noBreak(values['responseDeadline'])} 前通过电子邮件点击“全部回复”，并回复：“本人了解并同意此安排，将准时报到”。如不接受聘用，烦请注明具体原因。欢迎您加入我们！", styles["bodyIndented"]),
            Paragraph(f"如准备上述资料存在困难，请联系 {self.text(values['hrName'])}，电话：{self.noBreak(values['hrPhone'])}，我们将全力协助您。", styles["bodyIndented"]),
            Spacer(1, 18 * mm),
            Paragraph(f"聘用单位：{self.text(self.settings.companyName)}", styles["body"]),
            Paragraph(f"发出日期：{self.text(values['issueDate'])}", styles["body"]),
        ]
        document.build(story, onFirstPage=self.drawPage, onLaterPages=self.drawPage)
        return output

    def drawPage(self, canvas, document):
        """绘制页脚页码和公司名称。"""
        canvas.saveState()
        canvas.setFont(self.fontName, 8)
        canvas.setFillColor(colors.HexColor("#77827d"))
        canvas.drawString(23 * mm, 10 * mm, self.settings.companyName)
        canvas.drawRightString(A4[0] - 23 * mm, 10 * mm, f"第 {document.page} 页")
        canvas.restoreState()
