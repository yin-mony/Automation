import sys
import traceback
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (
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

from main import run as run_main

DEFAULT_DIR = Path(r"C:\Users\admin\Desktop\采购单与商品单汇总对比")


class _QtLogStream:
    def __init__(self, emit_func, prefix=""):
        self.emit_func = emit_func
        self.prefix = prefix
        self.buffer = ""

    def write(self, text):
        if not text:
            return
        self.buffer += str(text)
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            line = line.rstrip()
            if line:
                self.emit_func(f"{self.prefix}{line}")

    def flush(self):
        line = self.buffer.rstrip()
        if line:
            self.emit_func(f"{self.prefix}{line}")
        self.buffer = ""


class Worker(QThread):
    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(bool, str)

    def __init__(self, download_dir):
        super().__init__()
        self.download_dir = download_dir

    def run(self):
        stdout_stream = _QtLogStream(self.log_signal.emit)
        stderr_stream = _QtLogStream(self.log_signal.emit, "[stderr] ")
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_stream
        sys.stderr = stderr_stream
        try:
            run_main(download_dir=self.download_dir)
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(True, "执行完成。")
        except Exception as exc:
            traceback.print_exc()
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(False, f"执行失败: {exc}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class RunWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.setWindowTitle("采购单与商品单汇总对比")
        self.resize(760, 520)
        self._build_ui()

    def _build_ui(self):
        self.path_input = QLineEdit(str(DEFAULT_DIR))
        self.path_input.setPlaceholderText("请选择下载和输出目录")

        browse_btn = QPushButton("选择目录")
        browse_btn.clicked.connect(self._choose_dir)

        path_layout = QHBoxLayout()
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(browse_btn)

        self.run_btn = QPushButton("开始执行")
        self.run_btn.clicked.connect(self._start_run)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志将在这里显示...")

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel("下载/输出目录："))
        main_layout.addLayout(path_layout)
        main_layout.addWidget(self.run_btn)
        main_layout.addWidget(QLabel("日志："))
        main_layout.addWidget(self.log_box)
        self.setLayout(main_layout)

    def _choose_dir(self):
        selected = QFileDialog.getExistingDirectory(
            self,
            "选择目录",
            self.path_input.text().strip() or str(DEFAULT_DIR),
        )
        if selected:
            self.path_input.setText(selected)

    def _start_run(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "提示", "任务正在执行，请勿重复点击。")
            return

        download_dir = self.path_input.text().strip()
        if not download_dir:
            QMessageBox.warning(self, "参数错误", "请先选择目录。")
            return

        Path(download_dir).mkdir(parents=True, exist_ok=True)
        self.log_box.clear()
        self.log_box.append(f"目录: {download_dir}")
        self.log_box.append("-" * 60)
        self.run_btn.setEnabled(False)

        self.worker = Worker(download_dir)
        self.worker.log_signal.connect(self.log_box.append)
        self.worker.done_signal.connect(self._on_done)
        self.worker.start()

    def _on_done(self, success, message):
        self.run_btn.setEnabled(True)
        self.log_box.append("-" * 60)
        self.log_box.append(message)
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "失败", message)


def main():
    app = QApplication(sys.argv)
    window = RunWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
