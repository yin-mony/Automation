# -*- coding: utf-8 -*-
"""运行入口：支持 CLI 与 Qt 双模式。"""

import argparse
import sys
from pathlib import Path

from main import ExcelFile


class RunGui:
    """匹配回填工具的命令行与图形界面入口。"""

    def __init__(self):
        self.pipeline = ExcelFile()

    def runCli(self):
        """启动命令行交互模式。"""
        self.pipeline.runInteractive()

    def runGui(self):
        """启动 Qt 图形界面模式。"""
        try:
            from PySide6.QtWidgets import (
                QApplication,
                QFileDialog,
                QHBoxLayout,
                QLabel,
                QLineEdit,
                QMessageBox,
                QPushButton,
                QTextEdit,
                QVBoxLayout,
                QWidget,
            )
        except ImportError as exc:
            raise RuntimeError("未安装 PySide6，无法启动 Qt 模式。请先安装: pip install PySide6") from exc

        pipeline = self.pipeline

        class MainWindow(QWidget):
            """匹配回填工具主窗口。"""

            def __init__(self):
                super().__init__()
                self.setWindowTitle("匹配与ASIN回填工具")
                self.resize(820, 520)
                self.buildUi()

            def buildUi(self):
                """创建文件选择、执行按钮和日志输出区域。"""
                layout = QVBoxLayout(self)

                self.totalEdit = QLineEdit("")
                self.subEdit = QTextEdit("")
                self.totalEdit.setPlaceholderText("请选择主表文件（Excel）")
                self.subEdit.setPlaceholderText(
                    "可填写多个副表路径：每行一个；也可用分号分隔；或下方「选择文件」一次多选追加"
                )
                self.subEdit.setMinimumHeight(80)
                self.logOutput = QTextEdit()
                self.logOutput.setReadOnly(True)

                layout.addLayout(self.createFileRow("主表路径", self.totalEdit))
                layout.addLayout(self.createSubMultiRow())

                runButton = QPushButton("开始执行")
                runButton.clicked.connect(self.runPipeline)
                layout.addWidget(runButton)
                layout.addWidget(QLabel("执行日志："))
                layout.addWidget(self.logOutput)

            def createFileRow(self, labelText, lineEdit):
                """创建单个文件选择行。"""
                row = QHBoxLayout()
                row.addWidget(QLabel(labelText))
                row.addWidget(lineEdit)

                browseButton = QPushButton("选择文件")
                browseButton.clicked.connect(lambda: self.pickFile(lineEdit))
                row.addWidget(browseButton)
                return row

            def createSubMultiRow(self):
                """创建支持多选副表的文件区域。"""
                col = QVBoxLayout()
                head = QHBoxLayout()
                head.addWidget(QLabel("副表路径（可多选，依次匹配）"))
                browseMulti = QPushButton("选择文件")
                browseMulti.clicked.connect(self.pickSubFilesMulti)
                head.addWidget(browseMulti)
                head.addStretch()
                col.addLayout(head)
                col.addWidget(self.subEdit)
                return col

            def pickFile(self, lineEdit):
                """选择单个 Excel 文件并写入输入框。"""
                filePath, _ = QFileDialog.getOpenFileName(
                    self,
                    "选择Excel文件",
                    str(Path.home()),
                    "Excel Files (*.xlsx *.xls)",
                )
                if filePath:
                    lineEdit.setText(filePath)

            def pickSubFilesMulti(self):
                """多选副表 Excel 文件并追加到输入框。"""
                paths, _ = QFileDialog.getOpenFileNames(
                    self,
                    "选择副表Excel（可多选）",
                    str(Path.home()),
                    "Excel Files (*.xlsx *.xls)",
                )
                if not paths:
                    return

                existing = self.subEdit.toPlainText().replace(";", "\n")
                lines = [line.strip() for line in existing.splitlines() if line.strip()]
                seen = set(lines)
                for filePath in paths:
                    if filePath not in seen:
                        seen.add(filePath)
                        lines.append(filePath)
                self.subEdit.setPlainText("\n".join(lines))

            def runPipeline(self):
                """读取界面路径并执行匹配回填。"""
                totalPath = self.totalEdit.text().strip()
                rawSub = self.subEdit.toPlainText().replace(";", "\n")
                subPaths = [line.strip() for line in rawSub.splitlines() if line.strip()]
                if not totalPath or not subPaths:
                    QMessageBox.warning(self, "提示", "请先选择主表并填写至少一个副表路径")
                    return

                totalPathObj = Path(totalPath)
                if not totalPathObj.exists():
                    QMessageBox.warning(self, "路径错误", f"主表文件不存在:\n{totalPathObj}")
                    return

                subPathObjs = []
                for subPath in subPaths:
                    pathObj = Path(subPath)
                    if not pathObj.exists():
                        QMessageBox.warning(self, "路径错误", f"副表文件不存在:\n{pathObj}")
                        return
                    subPathObjs.append(pathObj)

                try:
                    result = pipeline.run(
                        totalPath=totalPathObj,
                        subPath=subPathObjs,
                        printSummary=False,
                        saveResult=True,
                        outputPath=totalPathObj,
                    )
                    self.renderResultLog(result, subPathObjs)
                    QMessageBox.information(self, "完成", "已成功完成匹配且进行回填")
                except Exception as exc:  # noqa: BLE001
                    QMessageBox.critical(self, "执行失败", str(exc))

            def renderResultLog(self, result, subPathObjs):
                """渲染执行后的匹配统计和主表预览。"""
                targetColUsed = result["target_col_used"]
                self.logOutput.clear()
                self.logOutput.append(f"副表数量: {len(subPathObjs)}（按列表顺序依次匹配回填）")
                self.logOutput.append("")

                for index, block in enumerate(result.get("per_sub_match_results", []), start=1):
                    subResult = block["sub_result"]
                    self.logOutput.append(f"--- 副表 {index}: {Path(block['sub_path']).name} ---")
                    self.logOutput.append(f"  行数: {len(subResult)}")
                    self.logOutput.append(f"  匹配成功: {int(subResult['is_match'].sum())}")
                    self.logOutput.append(f"  匹配失败: {int((~subResult['is_match']).sum())}")
                    self.logOutput.append("")

                subResultAll = result["sub_result"]
                self.logOutput.append(
                    f"副表合计 - 行数: {len(subResultAll)}，匹配成功: {int(subResultAll['is_match'].sum())}，"
                    f"失败: {int((~subResultAll['is_match']).sum())}"
                )
                self.logOutput.append("")
                self.logOutput.append("主表回填预览（前10行）：")
                self.logOutput.append(
                    result["total_df_filled"][[pipeline.totalCol, targetColUsed]]
                    .head(10)
                    .to_string(index=False)
                )

        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        app.exec()

    def main(self):
        """按启动参数分发 CLI 或 GUI 模式。"""
        parser = argparse.ArgumentParser(description="主表描述与副表订单ID匹配，并回填ASIN到主表编码列")
        parser.add_argument(
            "--mode",
            choices=("cli", "gui"),
            default="gui",
            help="运行模式：cli(命令行交互) 或 gui(Qt界面)",
        )
        args = parser.parse_args()

        if args.mode == "gui":
            self.runGui()
        else:
            self.runCli()


if __name__ == "__main__":
    RunGui().main()
