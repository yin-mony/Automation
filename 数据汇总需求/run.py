import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QMessageBox, QTextEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from main import Excel_file

class WorkerThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths

    def run(self):
        try:
            results = Excel_file(self.file_paths).process_multiple_files()
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("数据汇总处理工具")
        self.setGeometry(100, 100, 500, 400)

        self.file_list = []

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        title_label = QLabel("数据汇总处理工具")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; text-align: center;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        layout.addWidget(QLabel("已选择的文件:"))
        hint = QLabel("在弹出的对话框中可按住 Ctrl 或 Shift 一次选中多个 Excel 文件。")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(hint)

        self.file_list_label = QLabel("未选择文件")
        self.file_list_label.setStyleSheet("color: #666;")
        layout.addWidget(self.file_list_label)

        self.select_button = QPushButton("选择文件")
        self.select_button.clicked.connect(self.select_files)
        layout.addWidget(self.select_button)

        self.process_button = QPushButton("开始处理")
        self.process_button.clicked.connect(self.start_processing)
        layout.addWidget(self.process_button)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(QLabel("处理结果:"))
        layout.addWidget(self.result_text)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", "", 
                                                "Excel文件 (*.xlsx *.xls)")
        self.file_list = files
        if self.file_list:
            self.file_list_label.setText(f"已选择 {len(self.file_list)} 个文件")
        else:
            self.file_list_label.setText("未选择文件")

    def start_processing(self):
        if not self.file_list:
            QMessageBox.warning(self, "警告", "请先选择要处理的文件")
            return

        self.select_button.setEnabled(False)
        self.process_button.setEnabled(False)
        self.result_text.clear()

        self.worker = WorkerThread(self.file_list)
        self.worker.finished.connect(self.on_process_finished)
        self.worker.error.connect(self.on_process_error)
        self.worker.start()

    def on_process_finished(self, results):
        self.select_button.setEnabled(True)
        self.process_button.setEnabled(True)

        result_text = "处理完成！\n\n"
        success_count = 0
        fail_count = 0

        for file_path, result in results.items():
            if result.get('success'):
                result_text += f"✓ {file_path}\n  记录数: {result['count']}\n\n"
                success_count += 1
            else:
                result_text += f"✗ {file_path}\n  处理失败: {result.get('error', '未知错误')}\n\n"
                fail_count += 1

        result_text += f"总计: {success_count} 成功, {fail_count} 失败"
        self.result_text.setText(result_text)

        QMessageBox.information(self, "完成", f"处理完成！\n成功: {success_count} 个文件\n失败: {fail_count} 个文件")

    def on_process_error(self, error_msg):
        self.select_button.setEnabled(True)
        self.process_button.setEnabled(True)
        QMessageBox.critical(self, "错误", f"处理过程中发生错误:\n{error_msg}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())