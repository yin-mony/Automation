import json
import sys
import traceback
from pathlib import Path

# 避免 pywinauto / comtypes 在 Qt 环境下出现 COM 线程模型冲突（RPC_E_CHANGED_MODE）
if not hasattr(sys, "coinit_flags"):
    sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED (STA)

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from test import main as run_test
from SaihuERPLogin import SaihuERPLogin

CURRENT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = CURRENT_DIR / "run_config.json"

MODE_7_DAYS = "recent_7_days"
MODE_30_DAYS = "recent_30_days"
MODE_THIS_MONTH = "this_month"
MODE_LAST_MONTH = "last_month"
DEFAULT_USERNAME = SaihuERPLogin.DEFAULT_USERNAME
DEFAULT_PASSWORD = SaihuERPLogin.DEFAULT_PASSWORD


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


class DownloadWorker(QThread):
    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(bool, str)

    def __init__(self, params):
        super().__init__()
        self.params = params

    def run(self):
        stdout_stream = _QtLogStream(self.log_signal.emit)
        stderr_stream = _QtLogStream(self.log_signal.emit, "[stderr] ")
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_stream
        sys.stderr = stderr_stream
        try:
            run_test(
                mode=self.params["mode"],
                start_date=self.params.get("start_date", ""),
                end_date=self.params.get("end_date", ""),
                username=self.params.get("username"),
                password=self.params.get("password"),
                export_dir=self.params.get("export_dir", ""),
            )
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(True, "下载流程执行完成。")
        except Exception as exc:
            traceback.print_exc()
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(False, f"下载失败: {exc}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class RunWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.config = self._load_config()
        self.setWindowTitle("赛狐请款单下载")
        self.resize(640, 460)
        self._build_ui()
        self._load_defaults_to_form()

    def _build_ui(self):
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("请输入赛狐账号")

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("请输入赛狐密码")

        mode_group = QGroupBox("下载模式")
        mode_layout = QVBoxLayout()
        self.mode_button_group = QButtonGroup(self)
        self.radio_7_days = QRadioButton("最近7天")
        self.radio_30_days = QRadioButton("最近30天")
        self.radio_this_month = QRadioButton("本月")
        self.radio_last_month = QRadioButton("上个月")
        self.mode_button_group.addButton(self.radio_7_days)
        self.mode_button_group.addButton(self.radio_30_days)
        self.mode_button_group.addButton(self.radio_this_month)
        self.mode_button_group.addButton(self.radio_last_month)
        mode_layout.addWidget(self.radio_7_days)
        mode_layout.addWidget(self.radio_30_days)
        mode_layout.addWidget(self.radio_this_month)
        mode_layout.addWidget(self.radio_last_month)
        mode_group.setLayout(mode_layout)

        self.tip_label = QLabel(
            "说明：可选择下载模式；默认“最近7天”。下载目录使用浏览器当前默认设置。"
        )

        self.download_btn = QPushButton("下载")
        self.download_btn.clicked.connect(self._start_download)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志将在这里显示...")

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel("赛狐账号："))
        main_layout.addWidget(self.username_input)
        main_layout.addWidget(QLabel("赛狐密码："))
        main_layout.addWidget(self.password_input)
        main_layout.addWidget(mode_group)
        main_layout.addWidget(self.tip_label)
        main_layout.addWidget(self.download_btn)
        main_layout.addWidget(QLabel("日志："))
        main_layout.addWidget(self.log_box)
        self.setLayout(main_layout)

    def _load_defaults_to_form(self):
        self.username_input.setText(self.config.get("last_username", DEFAULT_USERNAME))
        self.password_input.setText(self.config.get("last_password", DEFAULT_PASSWORD))
        # 模式不记忆，每次打开界面都默认“最近7天”
        self.radio_7_days.setChecked(True)

    def _current_mode(self):
        if self.radio_30_days.isChecked():
            return MODE_30_DAYS
        if self.radio_this_month.isChecked():
            return MODE_THIS_MONTH
        if self.radio_last_month.isChecked():
            return MODE_LAST_MONTH
        return MODE_7_DAYS

    def _start_download(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "提示", "任务正在执行，请勿重复点击。")
            return

        username = self.username_input.text().strip()
        password = self.password_input.text().strip()
        if not username:
            QMessageBox.warning(self, "参数错误", "请输入赛狐账号。")
            return
        if not password:
            QMessageBox.warning(self, "参数错误", "请输入赛狐密码。")
            return

        params = {
            "username": username,
            "password": password,
            # 与 test.py 现状对齐：不从界面传出导出目录，交给 test.py 内部默认规则处理
            "export_dir": "",
            "mode": self._current_mode(),
            "start_date": "",
            "end_date": "",
        }
        self._save_config(params)

        mode_name = {
            MODE_7_DAYS: "最近7天",
            MODE_30_DAYS: "最近30天",
            MODE_THIS_MONTH: "本月",
            MODE_LAST_MONTH: "上个月",
        }.get(params["mode"], "最近7天")

        self.log_box.clear()
        self.log_box.append("下载参数确认：")
        self.log_box.append(f"- 账号: {params['username']}")
        self.log_box.append(f"- 模式: {mode_name}")
        self.log_box.append("- 下载目录: 浏览器当前默认下载目录")
        self.log_box.append("-" * 60)

        self.download_btn.setEnabled(False)
        self.worker = DownloadWorker(params)
        self.worker.log_signal.connect(self.log_box.append)
        self.worker.done_signal.connect(self._on_done)
        self.worker.start()

    def _on_done(self, success, message):
        self.download_btn.setEnabled(True)
        self.log_box.append("-" * 60)
        self.log_box.append(message)
        if success:
            self.log_box.append("已完成下载，文件保存在浏览器当前默认下载目录。")
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "失败", message)

    def _load_config(self):
        if not CONFIG_FILE.exists():
            return {}
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_config(self, params):
        data = {
            "last_username": params["username"],
            "last_password": params["password"],
        }
        try:
            CONFIG_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass


def main():
    app = QApplication(sys.argv)
    window = RunWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
