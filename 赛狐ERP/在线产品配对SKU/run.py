import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

try:
    from PyQt5.QtWidgets import (
        QApplication,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMessageBox,
        QTextEdit,
        QPushButton,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:
    raise ImportError("请先安装 PyQt5：pip inspipipiptall PyQt5") from exc

from SaihuERPLogin import SaihuERPLogin
from main import EXCEL_FILE_PATH, EXCEL_SHEET_NAME, SaiHuERP_WebPage, read_Excel


class SaihuRunnerWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("纯新品 → 通过新品开发流程生成SKU并在线配对（仅限于已完成提交流程的商品）")
        self.resize(640, 220)
        self._build_ui()

    def _build_ui(self):
        form = QFormLayout()

        self.excel_input = QLineEdit(str(EXCEL_FILE_PATH))
        browse_btn = QPushButton("选择文件")
        browse_btn.clicked.connect(self._choose_excel)
        excel_row = QHBoxLayout()
        excel_row.addWidget(self.excel_input)
        excel_row.addWidget(browse_btn)
        excel_hint = QLabel("请上传，已导出的工作计划表")

        self.sheet_input = QLineEdit(EXCEL_SHEET_NAME)
        self.username_input = QLineEdit(SaihuERPLogin.DEFAULT_USERNAME)
        self.password_input = QLineEdit(SaihuERPLogin.DEFAULT_PASSWORD)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        self.status_box.setPlaceholderText("状态栏：将显示待处理数量、对应数据和执行结果。")

        form.addRow("表格文件：", excel_row)
        form.addRow("", excel_hint)
        form.addRow("工作表名称：", self.sheet_input)
        form.addRow("赛狐登录账号：", self.username_input)
        form.addRow("赛狐登录密码：", self.password_input)
        form.addRow("状态栏：", self.status_box)

        self.run_btn = QPushButton("开始执行")
        self.run_btn.clicked.connect(self._run_task)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addWidget(self.run_btn)
        self.setLayout(layout)

    def _choose_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 文件",
            str(EXCEL_FILE_PATH.parent),
            "Excel Files (*.xlsx *.xls)",
        )
        if file_path:
            self.excel_input.setText(file_path)

    def _run_task(self):
        excel_path = self.excel_input.text().strip()
        sheet_name = self.sheet_input.text().strip() or EXCEL_SHEET_NAME
        username = self.username_input.text().strip() or SaihuERPLogin.DEFAULT_USERNAME
        password = self.password_input.text() or SaihuERPLogin.DEFAULT_PASSWORD

        if not excel_path:
            QMessageBox.warning(self, "参数错误", "请先选择 Excel 文件。")
            return

        self.run_btn.setEnabled(False)
        try:
            grouped_result = read_Excel(excel_file_path=excel_path, sheet_name=sheet_name)
            records = grouped_result.get("records", [])
            lines = [f"待操作数量：{len(records)}"]
            for idx, item in enumerate(records, start=1):
                lines.append(
                    f"{idx}. 编号={item['request_no']} | sku={item['new_prefix']} | ASIN={item['new_asin']} | 人员={item['new_name']}"
                )
            lines.append("状态：执行中...")
            self.status_box.setPlainText("\n".join(lines))
            QApplication.processEvents()

            SaiHuERP_WebPage(
                excel_file_path=excel_path,
                username=username,
                password=password,
                sheet_name=sheet_name,
            )
            done_text = self.status_box.toPlainText() + "\n状态：已完成"
            self.status_box.setPlainText(done_text)
            QMessageBox.information(self, "执行完成", "流程执行完成，请查看浏览器结果。")
        except Exception as exc:
            fail_text = self.status_box.toPlainText()
            if fail_text:
                fail_text += "\n"
            fail_text += f"状态：执行失败 - {exc}"
            self.status_box.setPlainText(fail_text)
            QMessageBox.critical(self, "执行失败", str(exc))
        finally:
            self.run_btn.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    window = SaihuRunnerWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
