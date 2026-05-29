import os
import sys
import traceback
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
import runpy

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


if getattr(sys, "frozen", False):
    BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
else:
    BASE_DIR = Path(__file__).resolve().parent

CURRENT_DIR = BASE_DIR
TEST_SCRIPT = BASE_DIR / "test.py"


class RunWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("请款汇总")
        self.resize(680, 360)
        self._build_ui()

    def _build_ui(self):
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("请选择要处理的目标文件夹")

        browse_btn = QPushButton("选择文件夹")
        browse_btn.clicked.connect(self._choose_folder)

        run_btn = QPushButton("开始执行")
        run_btn.clicked.connect(self._run_test)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志将显示在这里...")

        row_layout = QHBoxLayout()
        row_layout.addWidget(self.folder_input)
        row_layout.addWidget(browse_btn)

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel("目标文件夹："))
        main_layout.addLayout(row_layout)
        main_layout.addWidget(run_btn)
        main_layout.addWidget(QLabel("输出日志："))
        main_layout.addWidget(self.log_box)
        self.setLayout(main_layout)

    def _choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择目标文件夹")
        if folder:
            self.folder_input.setText(folder)

    def _run_test(self):
        folder = self.folder_input.text().strip()
        if not folder:
            QMessageBox.warning(self, "参数错误", "请先选择目标文件夹。")
            return

        folder_path = Path(folder)
        if not folder_path.exists() or not folder_path.is_dir():
            QMessageBox.warning(self, "路径无效", "所选路径不存在或不是文件夹。")
            return

        if not TEST_SCRIPT.exists():
            QMessageBox.critical(self, "文件缺失", f"未找到脚本: {TEST_SCRIPT}")
            return

        self.log_box.clear()
        self.log_box.append(f"目标文件夹: {folder_path}")
        self.log_box.append("开始执行 test.py ...")
        self.log_box.append("-" * 60)

        env = os.environ.copy()
        env["TARGET_FOLDER"] = str(folder_path)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        old_target = os.environ.get("TARGET_FOLDER")
        old_io_encoding = os.environ.get("PYTHONIOENCODING")
        old_utf8 = os.environ.get("PYTHONUTF8")
        os.environ.update(env)

        stdout_buffer = StringIO()
        stderr_buffer = StringIO()

        return_code = 0
        try:
            with redirect_stdout(stdout_buffer), redirect_stderr(stderr_buffer):
                try:
                    runpy.run_path(str(TEST_SCRIPT), run_name="__main__")
                except SystemExit as exc:
                    code = exc.code
                    if code not in (0, None):
                        return_code = int(code) if isinstance(code, int) else 1
        except Exception:
            return_code = 1
            traceback.print_exc(file=stderr_buffer)
        finally:
            if old_target is None:
                os.environ.pop("TARGET_FOLDER", None)
            else:
                os.environ["TARGET_FOLDER"] = old_target

            if old_io_encoding is None:
                os.environ.pop("PYTHONIOENCODING", None)
            else:
                os.environ["PYTHONIOENCODING"] = old_io_encoding

            if old_utf8 is None:
                os.environ.pop("PYTHONUTF8", None)
            else:
                os.environ["PYTHONUTF8"] = old_utf8

        stdout_text = stdout_buffer.getvalue().strip()
        stderr_text = stderr_buffer.getvalue().strip()

        if stdout_text:
            self.log_box.append(stdout_text)
        if stderr_text:
            self.log_box.append("[stderr]")
            self.log_box.append(stderr_text)

        self.log_box.append("-" * 60)
        if return_code == 0:
            self.log_box.append("执行完成。")
            QMessageBox.information(self, "完成", "执行完成。")
        else:
            self.log_box.append(f"执行失败，退出码: {return_code}")
            QMessageBox.warning(self, "失败", f"执行失败，退出码: {return_code}")


def main():
    app = QApplication(sys.argv)
    window = RunWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
