import json
import os
import sys
import traceback
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from OnlyMain import SaiHuMain
from SaihuERPLogin import SaihuERPLogin

CURRENT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = CURRENT_DIR / "onlyrun_config.json"
CONDA312_PYTHON = Path(r"C:\Users\admin\miniconda3\envs\saihu312\python.exe")
CONDA312_QT_PLUGIN_DIR = Path(r"C:\Users\admin\miniconda3\envs\saihu312\Lib\site-packages\PyQt5\Qt5\plugins")
CONDA312_QT_PLATFORM_DIR = CONDA312_QT_PLUGIN_DIR / "platforms"
DEFAULT_USERNAME = SaihuERPLogin.DEFAULT_USERNAME
DEFAULT_PASSWORD = SaihuERPLogin.DEFAULT_PASSWORD

MODE_DEW = "dew"
MODE_LOW = "low"


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


class RunnerThread(QThread):
    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(bool, str)

    def __init__(self, mode, excel_path, username, password):
        super().__init__()
        self.mode = mode
        self.excel_path = excel_path
        self.username = username
        self.password = password

    def run(self):
        stdout_stream = _QtLogStream(self.log_signal.emit)
        stderr_stream = _QtLogStream(self.log_signal.emit, "[stderr] ")
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = stdout_stream
        sys.stderr = stderr_stream
        try:
            main_client = SaiHuMain(username=self.username, password=self.password)
            main_client.run(mode=self.mode, excel_file_path=self.excel_path)
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(True, "流程执行完成")
        except Exception as exc:
            traceback.print_exc()
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(False, str(exc))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class OnlyRunnerWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.config = self._load_config()
        self.paths_by_mode = {
            MODE_DEW: self.config.get("last_excel_dew", ""),
            MODE_LOW: self.config.get("last_excel_low", ""),
        }
        self.setWindowTitle("赛狐ERP - 统一运行入口")
        self.resize(760, 500)
        self._build_ui()
        self._refresh_path_by_mode()

    def _build_ui(self):
        remembered_username = self.config.get("last_username", DEFAULT_USERNAME)
        remembered_password = self.config.get("last_password", DEFAULT_PASSWORD)

        self.username_input = QLineEdit(remembered_username)
        self.password_input = QLineEdit(remembered_password)
        self.password_input.setEchoMode(QLineEdit.Password)

        mode_group = QGroupBox("运行模式")
        mode_layout = QHBoxLayout()
        self.mode_dew_radio = QRadioButton("纯新品列表创建商品并在线配对")
        self.mode_low_radio = QRadioButton("低价商品列表创建商品并在线配对")
        self.mode_dew_radio.setChecked(True)
        self.mode_dew_radio.toggled.connect(self._refresh_path_by_mode)
        self.mode_low_radio.toggled.connect(self._refresh_path_by_mode)
        mode_layout.addWidget(self.mode_dew_radio)
        mode_layout.addWidget(self.mode_low_radio)
        mode_group.setLayout(mode_layout)

        self.file_input = QLineEdit("")
        self.file_input.setPlaceholderText("请选择当前模式的 Excel 文件")
        self.choose_btn = QPushButton("选择Excel")
        self.choose_btn.clicked.connect(self._choose_file)

        file_row = QHBoxLayout()
        file_row.addWidget(self.file_input)
        file_row.addWidget(self.choose_btn)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("运行日志将在这里实时显示...")

        self.run_btn = QPushButton("开始执行")
        self.run_btn.clicked.connect(self._run_main)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("赛狐账号："))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel("赛狐密码："))
        layout.addWidget(self.password_input)
        layout.addWidget(mode_group)
        layout.addWidget(QLabel("Excel 文件："))
        layout.addLayout(file_row)
        layout.addWidget(self.log_box)
        layout.addWidget(self.run_btn)
        self.setLayout(layout)

    def _current_mode(self):
        return MODE_DEW if self.mode_dew_radio.isChecked() else MODE_LOW

    def _refresh_path_by_mode(self):
        mode = self._current_mode()
        self.file_input.setText(self.paths_by_mode.get(mode, ""))

    def _choose_file(self):
        mode = self._current_mode()
        title = "选择纯新品模式 Excel 文件" if mode == MODE_DEW else "选择低价模式 Excel 文件"
        initial_dir = str(CURRENT_DIR)
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            initial_dir,
            "Excel Files (*.xlsx *.xls)",
        )
        if file_path:
            self.paths_by_mode[mode] = file_path
            self.file_input.setText(file_path)

    def _run_main(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "提示", "任务正在执行，请勿重复启动。")
            return

        mode = self._current_mode()
        excel_path = self.file_input.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "参数错误", "请先选择 Excel 文件。")
            return
        if not Path(excel_path).exists():
            QMessageBox.warning(self, "参数错误", f"Excel 文件不存在：{excel_path}")
            return

        username = self.username_input.text().strip() or DEFAULT_USERNAME
        password = self.password_input.text() or DEFAULT_PASSWORD
        self.paths_by_mode[mode] = excel_path
        self._save_config(username, password)

        self.log_box.clear()
        mode_text = "纯新品列表创建商品并在线配对" if mode == MODE_DEW else "低价商品列表创建商品并在线配对"
        self.log_box.append(f"执行模式: {mode_text}")
        self.log_box.append(f"执行参数: --excel {excel_path} --username {username} --password ******")
        self.log_box.append("-" * 60)

        self.run_btn.setEnabled(False)
        self.worker = RunnerThread(mode, excel_path, username, password)
        self.worker.log_signal.connect(self.log_box.append)
        self.worker.done_signal.connect(self._on_done)
        self.worker.start()

    def _on_done(self, success, message):
        self.run_btn.setEnabled(True)
        self.log_box.append("-" * 60)
        self.log_box.append(message)
        if success:
            QMessageBox.information(self, "执行完成", "流程执行完成。")
        else:
            QMessageBox.warning(self, "执行失败", message)

    def _load_config(self):
        if not CONFIG_FILE.exists():
            return {}
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_config(self, username, password):
        data = {
            "last_username": username,
            "last_password": password,
            "last_excel_dew": self.paths_by_mode.get(MODE_DEW, ""),
            "last_excel_low": self.paths_by_mode.get(MODE_LOW, ""),
        }
        try:
            CONFIG_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass


def main():
    _ensure_conda312_runtime()
    _ensure_qt_plugin_env()
    app = QApplication(sys.argv)
    window = OnlyRunnerWindow()
    window.show()
    sys.exit(app.exec_())


def _ensure_conda312_runtime():
    """
    固定当前项目使用 conda 的 3.12 环境运行：
    - 若当前解释器不是 saihu312 的 python.exe，则自动切换并重启本脚本；
    - 若 conda 环境不存在，则保持当前解释器并输出提示。
    """
    # 打包后的 exe 运行时不可再切换解释器，否则会导致启动失败。
    if getattr(sys, "frozen", False):
        return

    current_exe = Path(sys.executable).resolve()
    target_exe = CONDA312_PYTHON.resolve()

    if not target_exe.exists():
        print(f"[warn] 未找到 conda 3.12 解释器: {target_exe}", flush=True)
        return

    if current_exe == target_exe:
        return

    print(f"检测到当前解释器: {current_exe}", flush=True)
    print(f"自动切换到 conda 3.12 环境: {target_exe}", flush=True)
    _set_qt_env_for_conda312()
    os.execv(str(target_exe), [str(target_exe), str(Path(__file__).resolve()), *sys.argv[1:]])


def _set_qt_env_for_conda312():
    if CONDA312_QT_PLATFORM_DIR.exists():
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(CONDA312_QT_PLATFORM_DIR)
    if CONDA312_QT_PLUGIN_DIR.exists():
        os.environ["QT_PLUGIN_PATH"] = str(CONDA312_QT_PLUGIN_DIR)


def _ensure_qt_plugin_env():
    """
    修复部分环境变量被置空导致的：
    Could not find the Qt platform plugin "windows" in ""
    """
    platform_env = os.environ.get("QT_QPA_PLATFORM_PLUGIN_PATH", "").strip()
    if platform_env:
        return
    _set_qt_env_for_conda312()


if __name__ == "__main__":
    main()
