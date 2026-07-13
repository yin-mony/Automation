"""POP 文档导出：按模板填充 detailInfo 生成 PDF，以及多 SKU 模板检查"""
import calendar
import re
import sys
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from shutil import copyfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class PopExport:
    """FBA 货件 POP 文档导出与多 SKU 模板检查"""

    @classmethod
    def getBaseDir(cls):
        """获取脚本或 exe 所在目录"""
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    @classmethod
    def getResourceDir(cls):
        """获取模板等资源所在目录"""
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS)
        return Path(__file__).resolve().parent

    def __init__(self, baseDir=None):
        # 项目根目录（输出、配置、OCR 临时文件）
        self.baseDir = baseDir or self.getBaseDir()
        # 模板目录始终从资源路径读取（打包后在 _MEIPASS，与 baseDir 分离）
        resourceDir = self.getResourceDir()
        # 单 SKU 模板
        self.templatePath = resourceDir / "db53060fa183_发票模板.docx"
        # 多 SKU 模板
        self.multiTemplatePath = resourceDir / "服务商模板.docx"
        # POP 输出目录
        self.exportDir = self.baseDir / "output"
        # 汇总表占位标签
        self.summaryLabels = [
            "Packing List Date:",
            "Shipment ID:",
            "Shipment Name:",
            "SKU Total:",
            "Unit Total:",
        ]
        # 汇总表占位示例值
        self.summaryValues = [
            "12/12/2025",
            "FBA1961YHSPH",
            "FBA ASDN (01/12/2026 01:27)-SBD1",
            "1",
            "132",
        ]
        # 签名区姓名表宽度与左侧 X 坐标
        self.nameTableWidth = 5485
        self.nameTableLeftX = 299
        # 多 SKU 商品表预置数据行数（不含表头）
        self.multiItemPresetRows = 6
        # 多 SKU 商品表模板行高，用于超出预置行时同步下移签名表
        self.multiItemRowHeight = 457
        # 授权签名姓名，默认沿用服务商模板
        self.signatureName = "Xiaoyu Wang"
        # 授权签名图片路径，为空时保留模板自带签名图片
        self.signatureImagePath = ""

    def setParaText(self, paragraph, text):
        """设置段落文本，尽量保留首 run 样式"""
        if paragraph.runs:
            paragraph.runs[0].text = text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.text = text

    def setCellText(self, cell, text):
        """设置单元格文本（单段落单元格）"""
        if cell.paragraphs:
            self.setParaText(cell.paragraphs[0], text)
        else:
            cell.text = text

    def setCellParaText(self, cell, paraIndex, text):
        """设置单元格内指定段落文本，保留其余段落与样式结构"""
        if paraIndex < len(cell.paragraphs):
            self.setParaText(cell.paragraphs[paraIndex], text)

    def formatDate(self, createTime):
        """YYYY-MM-DD 减一个月后格式化为 MM/DD/YYYY，无效则返回 None"""
        if not createTime:
            return None
        try:
            dt = datetime.strptime(str(createTime).strip()[:10], "%Y-%m-%d")
        except ValueError:
            return None
        month = dt.month - 1
        year = dt.year
        if month == 0:
            month = 12
            year -= 1
        day = min(dt.day, calendar.monthrange(year, month)[1])
        shifted = datetime(year, month, day)
        return shifted.strftime("%m/%d/%Y")

    def cleanName(self, name):
        """清理 Windows 文件名非法字符"""
        text = str(name or "").strip()
        # 将 Windows 文件名非法字符替换为空格
        text = re.sub(r'[\\/:*?"<>|]+', " ", text)
        # 合并多余空格，避免生成难读文件名
        text = re.sub(r"\s+", " ", text).strip()
        return text or "Unknown"

    def fillSummary(self, summaryTable, shipmentId, name, skuTotal, unitTotal, packingListDate=None):
        """汇总表：多 SKU 模板 5 行按行填值；单 SKU 模板 1 行按段落填值"""
        if len(summaryTable.rows) >= 5:
            if packingListDate:
                self.setCellText(summaryTable.rows[0].cells[1], packingListDate)
            self.setCellText(summaryTable.rows[1].cells[1], shipmentId)
            self.setCellText(summaryTable.rows[2].cells[1], str(name or ""))
            self.setCellText(summaryTable.rows[3].cells[1], str(skuTotal))
            self.setCellText(summaryTable.rows[4].cells[1], str(unitTotal))
            return
        rightCell = summaryTable.rows[0].cells[1]
        if len(rightCell.paragraphs) < 5:
            return
        if packingListDate:
            self.setCellParaText(rightCell, 0, packingListDate)
        self.setCellParaText(rightCell, 1, shipmentId)
        self.setCellParaText(rightCell, 2, str(name or ""))
        self.setCellParaText(rightCell, 3, str(skuTotal))
        self.setCellParaText(rightCell, 4, str(unitTotal))

    def fillCommon(self, doc, detailInfo, shipmentId, skuTotal, unitTotal):
        """填充单条/多条模板共用的 Ship From、Ship To、汇总表字段"""
        source = detailInfo.get("source")
        # 当前模板第 10 段为 Ship From 地址
        if len(doc.paragraphs) > 10 and source:
            self.setParaText(doc.paragraphs[10], str(source))
        # 当前模板第 11 段为 Ship From 地址后续空行
        if len(doc.paragraphs) > 11:
            self.setParaText(doc.paragraphs[11], "")

        fulfillmentCenterId = str(detailInfo.get("fulfillmentCenterId") or "").strip()
        # 当前模板第 17 段为 Ship To 仓库编码段落
        if len(doc.paragraphs) > 17 and fulfillmentCenterId:
            para = doc.paragraphs[17]
            centerCode = fulfillmentCenterId[:4]
            newText = re.sub(r"\([^)]*\)", f"({centerCode})", para.text)
            self.setParaText(para, newText)

        packingListDate = self.formatDate(detailInfo.get("createTime"))
        self.fillSummary(
            doc.tables[0],
            shipmentId,
            detailInfo.get("name"),
            skuTotal,
            unitTotal,
            packingListDate,
        )

    def getPictureSize(self, cell):
        """读取模板签名图片原始尺寸，避免替换图片后破坏版式"""
        extents = cell._tc.xpath(".//wp:extent")
        if not extents:
            return 3423920, 807720
        try:
            return int(extents[0].get("cx")), int(extents[0].get("cy"))
        except (TypeError, ValueError):
            return 3423920, 807720

    def clearCellKeepPr(self, cell):
        """清空单元格正文内容，但保留单元格属性和段落样式"""
        oldPara = cell.paragraphs[0]._p if cell.paragraphs else None
        oldParaPr = deepcopy(oldPara.find(qn("w:pPr"))) if oldPara is not None else None
        tc = cell._tc
        # 只移除正文节点，保留 tcPr 中的宽度、边框、内边距等模板样式
        for child in list(tc):
            if child.tag != qn("w:tcPr"):
                tc.remove(child)
        para = OxmlElement("w:p")
        if oldParaPr is not None:
            para.append(oldParaPr)
        tc.append(para)
        return cell.paragraphs[0]

    def fillSignature(self, doc, detailInfo):
        """填充授权签名姓名与签名图片"""
        if len(doc.tables) < 4:
            return

        signatureName = str(
            detailInfo.get("signatureName") or self.signatureName or "Xiaoyu Wang"
        ).strip() or "Xiaoyu Wang"
        signatureImagePath = str(
            detailInfo.get("signatureImagePath") or self.signatureImagePath or ""
        ).strip()

        nameCell = doc.tables[3].rows[0].cells[0]
        if len(nameCell.paragraphs) >= 2:
            self.setParaText(nameCell.paragraphs[1], signatureName)
        elif nameCell.paragraphs:
            nameCell.add_paragraph(signatureName)

        if not signatureImagePath:
            return
        imagePath = Path(signatureImagePath)
        if not imagePath.is_file():
            raise FileNotFoundError(f"签名图片不存在: {imagePath}")

        imageCell = doc.tables[2].rows[1].cells[0]
        width, height = self.getPictureSize(imageCell)
        para = self.clearCellKeepPr(imageCell)
        run = para.add_run()
        # 按模板原签名图片尺寸写入，保持签名框版式不变
        run.add_picture(str(imagePath), width=width, height=height)

    def fillSingle(self, itemTable, items):
        """单 MSKU 使用服务商模板，只保留表头和一行数据"""
        if not items:
            return
        self.keepRows(itemTable, 1)
        row = itemTable.rows[1]
        it = items[0]
        self.setCellText(row.cells[0], str(it.get("fnSku") or ""))
        self.setCellText(row.cells[1], str(it.get("msku") or ""))
        self.setCellText(row.cells[2], str(it.get("quantity") or ""))

    def keepRows(self, itemTable, dataCount):
        """商品表只保留表头和指定数量的数据行"""
        keepCount = dataCount + 1
        # 从表尾删除多余预留行，避免单 SKU 文件留下空白商品行
        while len(itemTable.rows) > keepCount:
            itemTable._tbl.remove(itemTable.rows[-1]._tr)

    def cloneRow(self, itemTable):
        """复制商品表最后一行（空白数据行）作为样式模板并追加"""
        templateRow = itemTable.rows[-1]
        newTr = deepcopy(templateRow._tr)
        itemTable._tbl.append(newTr)
        return itemTable.rows[-1]

    def ensureRows(self, itemTable, needCount):
        """商品表数据行不足时按模板行 deepcopy 增行，保留行样式"""
        dataRowCount = len(itemTable.rows) - 1
        if needCount <= dataRowCount:
            return
        extra = needCount - dataRowCount
        for _ in range(extra):
            newRow = self.cloneRow(itemTable)
            for cell in newRow.cells:
                self.setCellText(cell, "")

    def removeFloat(self, tbl):
        """移除表格页面绝对定位，使其回到正文流式排版"""
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            return
        for tag in ("tblpPr", "tblOverlap"):
            old = tblPr.find(qn(f"w:{tag}"))
            if old is not None:
                tblPr.remove(old)

    def getRowHeight(self, itemTable):
        """读取多 SKU 商品行高度，模板未设置时使用默认高度"""
        if len(itemTable.rows) < 2:
            return self.multiItemRowHeight
        trPr = itemTable.rows[1]._tr.find(qn("w:trPr"))
        trHeight = trPr.find(qn("w:trHeight")) if trPr is not None else None
        if trHeight is None:
            return self.multiItemRowHeight
        try:
            return int(trHeight.get(qn("w:val")) or self.multiItemRowHeight)
        except (TypeError, ValueError):
            return self.multiItemRowHeight

    def shiftFloatY(self, tbl, offsetY):
        """保持浮动表样式并按指定距离纵向移动"""
        if offsetY == 0:
            return
        tblPr = tbl.find(qn("w:tblPr"))
        tblp = tblPr.find(qn("w:tblpPr")) if tblPr is not None else None
        if tblp is None:
            return
        try:
            oldY = int(tblp.get(qn("w:tblpY")) or 0)
        except (TypeError, ValueError):
            return
        # 只调整 Y 坐标，其余浮动定位、边距、重叠规则保持模板原样
        tblp.set(qn("w:tblpY"), str(oldY + offsetY))

    def shiftSignature(self, doc, itemCount):
        """SKU 超过预置行时保持签名表浮动样式并随商品表下移"""
        if itemCount <= self.multiItemPresetRows:
            return
        if len(doc.tables) < 3:
            return
        extraRows = itemCount - self.multiItemPresetRows
        rowHeight = self.getRowHeight(doc.tables[1])
        # 新增几行商品，就按模板商品行高同步下移右侧签名表
        self.shiftFloatY(doc.tables[2]._tbl, extraRows * rowHeight)

    def shiftSingleSignature(self, doc, itemCount):
        """单 SKU 裁剪商品预留行后，同步上移右侧签名表"""
        if itemCount <= 0 or itemCount >= self.multiItemPresetRows:
            return
        if len(doc.tables) < 3:
            return
        deletedRows = self.multiItemPresetRows - itemCount
        rowHeight = self.getRowHeight(doc.tables[1])
        # 删除几行商品预留行，就按模板商品行高同步上移右侧签名表
        self.shiftFloatY(doc.tables[2]._tbl, -deletedRows * rowHeight)

    def fillMulti(self, itemTable, items):
        """多 MSKU 模板：预置行内填入；超出时复制空白行增行且保留行样式"""
        if not items:
            return
        self.ensureRows(itemTable, len(items))
        dataRows = itemTable.rows[1:]
        for idx, it in enumerate(items):
            row = dataRows[idx]
            self.setCellText(row.cells[0], str(it.get("fnSku") or ""))
            self.setCellText(row.cells[1], str(it.get("msku") or ""))
            self.setCellText(row.cells[2], str(it.get("quantity") or ""))

    def toPdf(self, docxPath):
        """将 docx 转为 pdf（Windows 需已安装 Microsoft Word）。"""
        docxPath = Path(docxPath)
        pdfPath = docxPath.with_suffix(".pdf")
        try:
            from docx2pdf import convert
            convert(str(docxPath), str(pdfPath))
        except Exception as e:
            # Word 可能在已保存 PDF 后，于退出 COM 时抛远程过程调用失败
            if pdfPath.is_file() and pdfPath.stat().st_size > 0:
                print(f"DOCX 转 PDF 已生成，但 Word 退出异常，继续使用: {pdfPath}", flush=True)
                return pdfPath
            raise RuntimeError(f"DOCX 转 PDF 失败（需安装 Microsoft Word）: {e}") from e
        if not pdfPath.is_file() or pdfPath.stat().st_size <= 0:
            raise RuntimeError(f"PDF 未生成: {pdfPath}")
        return pdfPath

    def build(self, detailInfo):
        """按模板填充 detailInfo，生成 POP PDF 文档，调试时可保留 DOCX"""
        items = detailInfo.get("items") or []
        # SKU 总数
        skuTotal = len(items)
        # 装箱总数量
        unitTotal = 0
        for it in items:
            try:
                unitTotal += int(it.get("quantity") or 0)
            except (TypeError, ValueError):
                pass

        shopName = str(detailInfo.get("shopName") or "Unknown").strip()
        shipmentId = str(detailInfo.get("shipmentId") or "").strip()
        # 货件编号为空时无法生成可追踪的 POP 文件
        if not shipmentId:
            raise ValueError("货件编号不能为空，无法生成 POP")

        baseName = self.cleanName(f"{shopName}_{shipmentId}_POP")
        self.exportDir.mkdir(parents=True, exist_ok=True)
        docxPath = self.exportDir / f"{baseName}.docx"
        pdfPath = self.exportDir / f"{baseName}.pdf"

        useMulti = len(items) > 1
        # 调试查看版式时可只保留 docx，不进入 PDF 转换
        keepDocx = bool(detailInfo.get("keepDocx"))
        templatePath = self.multiTemplatePath
        # 模板不存在时提前报错，避免 copyfile 抛出不直观异常
        if not templatePath.is_file():
            raise FileNotFoundError(f"POP 模板不存在: {templatePath}")

        # 复制模板到临时 docx，填充后转 PDF
        copyfile(templatePath, docxPath)
        doc = Document(str(docxPath))
        # POP 模板至少需要汇总表与商品表
        if len(doc.tables) < 2:
            raise ValueError(f"POP 模板表格数量不足，至少需要 2 个表格: {templatePath}")

        # 当前模板第 0 个表为汇总表，第 1 个表为商品表
        self.fillCommon(doc, detailInfo, shipmentId, skuTotal, unitTotal)
        self.fillSignature(doc, detailInfo)

        itemTable = doc.tables[1]
        if useMulti:
            self.fillMulti(itemTable, items)
            self.shiftSignature(doc, len(items))
        else:
            self.fillSingle(itemTable, items)
            self.shiftSingleSignature(doc, len(items))

        doc.save(str(docxPath))
        if keepDocx:
            print(f"POP DOCX 已生成: {docxPath}", flush=True)
            return str(docxPath)
        pdfPath = self.toPdf(docxPath)
        try:
            docxPath.unlink()
        except OSError:
            pass
        print(f"POP 已生成: {pdfPath}", flush=True)
        return str(pdfPath)

    def addBorder(self, parent, tag, val="dotted", sz="4", color="000000"):
        """向表格边框父节点追加单边样式"""
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), val)
        el.set(qn("w:color"), color)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        parent.append(el)

    def setDashBorder(self, tblPr):
        """汇总表设置为虚线边框"""
        old = tblPr.find(qn("w:tblBorders"))
        if old is not None:
            tblPr.remove(old)
        borders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            self.addBorder(borders, side)
        tblPr.append(borders)

    def setSpacing(self, pElem):
        """段前/段后归零，单倍行距，避免固定大行高"""
        pPr = pElem.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            pElem.insert(0, pPr)
        oldSp = pPr.find(qn("w:spacing"))
        if oldSp is not None:
            pPr.remove(oldSp)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        pPr.append(spacing)

    def clearHeight(self, tr):
        """移除表格行固定高度"""
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            return
        hr = trPr.find(qn("w:trHeight"))
        if hr is not None:
            trPr.remove(hr)

    def setCellSingle(self, tc, text, tcTemplate):
        """按模板单元格样式写入单行文本"""
        # 取模板单元格第一个段落作为样式来源
        templateParas = tcTemplate.findall(qn("w:p"))
        srcPara = templateParas[0] if templateParas else None
        # 清空目标单元格内容，保留后续重新写入的样式结构
        for child in list(tc):
            tc.remove(child)
        # 复制模板单元格属性，保持边框、宽度等样式
        tcPr = tcTemplate.find(qn("w:tcPr"))
        if tcPr is not None:
            tc.append(deepcopy(tcPr))
        if srcPara is not None:
            # 复制模板段落并替换其中的第一个文本节点
            newPara = deepcopy(srcPara)
            texts = newPara.findall(".//" + qn("w:t"))
            if texts:
                texts[0].text = text
                for node in texts[1:]:
                    node.text = ""
            # 统一段落行距，避免模板固定行高影响展示
            self.setSpacing(newPara)
            tc.append(newPara)

    def rebuildSummary(self, doc):
        """多 SKU 模板汇总表拆为 5 行 2 列"""
        oldTable = doc.tables[0]
        oldRow = oldTable.rows[0]
        # 备份左右单元格样式，后续新行复用
        leftTpl = deepcopy(oldRow.cells[0]._tc)
        rightTpl = deepcopy(oldRow.cells[1]._tc)
        tbl = oldTable._tbl
        tblPr = tbl.find(qn("w:tblPr"))

        # 清空原汇总表数据行
        for tr in list(tbl.findall(qn("w:tr"))):
            tbl.remove(tr)

        if tblPr is not None:
            self.setDashBorder(tblPr)

        # 按固定 5 个汇总字段重建行
        templateTr = deepcopy(oldRow._tr)
        for label, value in zip(self.summaryLabels, self.summaryValues):
            tr = deepcopy(templateTr)
            self.clearHeight(tr)
            leftTc, rightTc = tr.findall(qn("w:tc"))
            self.setCellSingle(leftTc, label, leftTpl)
            self.setCellSingle(rightTc, value, rightTpl)
            tbl.append(tr)

    def relaxRows(self, doc):
        """商品表去除固定行高并统一段落行距"""
        itemTbl = doc.tables[1]._tbl
        for tr in itemTbl.findall(qn("w:tr")):
            self.clearHeight(tr)
            for tc in tr.findall(qn("w:tc")):
                for p in tc.findall(qn("w:p")):
                    self.setSpacing(p)

    def setCellWidth(self, tc, width):
        """设置单元格宽度"""
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is None:
            tcW = OxmlElement("w:tcW")
            tcPr.append(tcW)
        tcW.set(qn("w:w"), str(width))
        tcW.set(qn("w:type"), "dxa")

    def setFloat(self, tbl, tblpX, tblpY):
        """表格相对页面绝对定位"""
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl.insert(0, tblPr)

        # 先清理旧定位，避免重复定位节点
        for tag in ("tblpPr", "tblOverlap"):
            old = tblPr.find(qn(f"w:{tag}"))
            if old is not None:
                tblPr.remove(old)

        # 写入新的页面绝对定位参数
        tblp = OxmlElement("w:tblpPr")
        tblp.set(qn("w:leftFromText"), "180")
        tblp.set(qn("w:rightFromText"), "180")
        tblp.set(qn("w:vertAnchor"), "page")
        tblp.set(qn("w:horzAnchor"), "page")
        tblp.set(qn("w:tblpX"), str(tblpX))
        tblp.set(qn("w:tblpY"), str(tblpY))
        tblPr.insert(1, tblp)

        overlap = OxmlElement("w:tblOverlap")
        overlap.set(qn("w:val"), "never")
        tblPr.insert(2, overlap)

        # 移除普通缩进，避免与绝对定位叠加
        tblInd = tblPr.find(qn("w:tblInd"))
        if tblInd is not None:
            tblPr.remove(tblInd)

    def alignSignature(self, doc):
        """姓名表与签名表同 Y 绝对定位，并拆成两行与右侧等高对齐"""
        sigTbl = doc.tables[2]._tbl
        nameTbl = doc.tables[3]._tbl
        sigTblPr = sigTbl.find(qn("w:tblPr"))
        sigTblp = sigTblPr.find(qn("w:tblpPr")) if sigTblPr is not None else None
        if sigTblp is None:
            return

        # 姓名表跟随签名表的 Y 坐标对齐
        tblpY = sigTblp.get(qn("w:tblpY"), "10620")
        nameRows = nameTbl.findall(qn("w:tr"))
        if len(nameRows) == 1:
            nameTc = nameRows[0].find(qn("w:tc"))
            paras = nameTc.findall(qn("w:p")) if nameTc is not None else []
            sigRows = sigTbl.findall(qn("w:tr"))

            # 复用签名表两行结构，拆出姓名标题行和姓名内容行
            labelTr = deepcopy(sigRows[0])
            bodyTr = deepcopy(sigRows[1])

            # 清空原姓名表单行结构
            for tr in list(nameTbl.findall(qn("w:tr"))):
                nameTbl.remove(tr)

            # 写入姓名标题行
            labelTc = labelTr.find(qn("w:tc"))
            for p in list(labelTc.findall(qn("w:p"))):
                labelTc.remove(p)
            if paras:
                labelTc.append(deepcopy(paras[0]))
            self.setCellWidth(labelTc, self.nameTableWidth)

            # 写入姓名内容行
            bodyTc = bodyTr.find(qn("w:tc"))
            for p in list(bodyTc.findall(qn("w:p"))):
                bodyTc.remove(p)
            if len(paras) > 1:
                bodyTc.append(deepcopy(paras[1]))
            else:
                bodyTc.append(OxmlElement("w:p"))
            self.setCellWidth(bodyTc, self.nameTableWidth)

            nameTbl.append(labelTr)
            nameTbl.append(bodyTr)

            # 同步表格网格宽度，避免 Word 自动压缩姓名表
            tblGrid = nameTbl.find(qn("w:tblGrid"))
            if tblGrid is not None:
                for gridCol in tblGrid.findall(qn("w:gridCol")):
                    gridCol.set(qn("w:w"), str(self.nameTableWidth))

        self.setFloat(nameTbl, self.nameTableLeftX, tblpY)

    def checkTemplate(self):
        """只读检查多 SKU 模板结构，避免误覆盖服务商模板.docx"""
        if not self.multiTemplatePath.is_file():
            raise FileNotFoundError(f"多 SKU 模板不存在: {self.multiTemplatePath}")

        # 只读取模板结构，不复制、不重建、不保存，避免破坏已确认样式
        doc = Document(str(self.multiTemplatePath))
        if len(doc.tables) < 4:
            raise ValueError(f"多 SKU 模板表格数量不足，当前为 {len(doc.tables)} 个")

        summaryTable = doc.tables[0]
        itemTable = doc.tables[1]
        signatureTable = doc.tables[2]
        nameTable = doc.tables[3]

        signaturePr = signatureTable._tbl.find(qn("w:tblPr"))
        signatureFloat = signaturePr.find(qn("w:tblpPr")) if signaturePr is not None else None
        signatureAttrs = None
        if signatureFloat is not None:
            signatureAttrs = {key.split("}")[1]: value for key, value in signatureFloat.attrib.items()}

        summaryOk = (
            len(summaryTable.rows) == 1
            and len(summaryTable.columns) == 2
            and len(summaryTable.rows[0].cells[1].paragraphs) >= 5
        )
        itemOk = len(itemTable.rows) >= self.multiItemPresetRows + 1 and len(itemTable.columns) == 3
        signatureOk = signatureFloat is not None
        nameOk = len(nameTable.rows) == 1 and len(nameTable.columns) == 1

        print(f"多 SKU 模板检查: {self.multiTemplatePath}", flush=True)
        print(f"表格数量: {len(doc.tables)}", flush=True)
        print(
            f"汇总表: {len(summaryTable.rows)} 行 {len(summaryTable.columns)} 列，"
            f"右侧段落 {len(summaryTable.rows[0].cells[1].paragraphs)} 个",
            flush=True,
        )
        print(f"商品表: {len(itemTable.rows)} 行 {len(itemTable.columns)} 列", flush=True)
        print(f"签名表浮动定位: {signatureAttrs}", flush=True)
        print(f"姓名表: {len(nameTable.rows)} 行 {len(nameTable.columns)} 列", flush=True)

        if not (summaryOk and itemOk and signatureOk and nameOk):
            raise ValueError("多 SKU 模板结构不符合当前确认样式，请检查服务商模板.docx")
        return True


if __name__ == "__main__":
    config = {
        "mode": "build",
        "detailInfo": {
            "shipmentId": "FBA19TEST01",
            "shopName": "Lydia deal-CA",
            "name": "FBA STA test export",
            "fulfillmentCenterId": "YEG2",
            "createTime": "2026-06-17",
            "source": "JUNMALL CN Hangzhou Zhejiang 311100",
            "items": [
                {"fnSku": "X004YE2RTB", "msku": "2pc Toilet Flapper Red-CA", "quantity": 250},
                {"fnSku": "X004YEDA37", "msku": "Toilet Flush Lever White-CA", "quantity": 150},
            ],
        },
    }

    svc = PopExport()
    if config["mode"] == "prepareTemplate":
        svc.checkTemplate()
    else:
        savePath = svc.build(config["detailInfo"])
        print(f"调试输出: {savePath}")
