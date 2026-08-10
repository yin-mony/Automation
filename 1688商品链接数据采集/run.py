"""1688 商品链接采集 — PyQt5 GUI 入口。

仅 import main.Ali1688，不引用 test.py。
路径由界面输入 / run_config.json 提供，后台线程执行采集并实时显示日志。
"""

import json
import sys
import traceback
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from main import Ali1688


def get_app_base_dir():
    """脚本目录；PyInstaller 单文件打包时取 exe 所在目录。"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CURRENT_DIR = get_app_base_dir()
CONFIG_FILE = CURRENT_DIR / 'run_config.json'  # 持久化上次界面填写的路径
PREVIEW_MAX_ROWS = 50                          # 预览表格每个 sheet 最多显示行数


class _QtLogStream:
    """将 print 输出按行切分后通过回调发送到 Qt 信号（供日志区实时显示）。"""

    def __init__(self, emit_func, prefix=''):
        self.emit_func = emit_func
        self.prefix = prefix
        self.buffer = ''

    def write(self, text):
        if not text:
            return
        self.buffer += str(text)
        while '\n' in self.buffer:
            line, self.buffer = self.buffer.split('\n', 1)
            line = line.rstrip()
            if line:
                self.emit_func(f'{self.prefix}{line}')

    def flush(self):
        """刷新缓冲区中未换行的尾部内容。"""
        line = self.buffer.rstrip()
        if line:
            self.emit_func(f'{self.prefix}{line}')
        self.buffer = ''


class Worker(QThread):
    """后台执行 Ali1688 任务，避免阻塞 GUI 主线程。"""

    log_signal = pyqtSignal(str)       # 单行日志
    done_signal = pyqtSignal(bool, str)  # (是否成功, 结束消息)

    def __init__(self, task, config):
        super().__init__()
        self.task = task    # 'run' | 'data' | 'excel'（GUI 目前仅用 'run'）
        self.config = config

    def run(self):
        """重定向 stdout/stderr 到日志区，执行完毕后恢复。"""
        stdout_stream = _QtLogStream(self.log_signal.emit)
        stderr_stream = _QtLogStream(self.log_signal.emit, '[stderr] ')
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_stream
        sys.stderr = stderr_stream
        page = None
        try:
            if self.task == 'excel':
                # 仅导出：不需浏览器
                ali = Ali1688(page=None, config=self.config)
                ali.excel_df()
            else:
                page = ChromiumPage()
                ali = Ali1688(page=page, config=self.config)
                if self.task == 'data':
                    ali.data()
                elif self.task == 'run':
                    ali.run()
                else:
                    raise ValueError(f'未知任务: {self.task}')
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(True, '任务执行完成。')
        except Exception as exc:
            traceback.print_exc()
            stdout_stream.flush()
            stderr_stream.flush()
            self.done_signal.emit(False, f'执行失败: {exc}')
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class PreviewDialog(QDialog):
    """在窗口内预览 规格汇总.xlsx（简表，不含合并单元格等完整样式）。"""

    def __init__(self, excel_path, parent=None):
        super().__init__(parent)
        self.excel_path = Path(excel_path)
        self.sheets = {}
        self.setWindowTitle('规格汇总预览')
        self.resize(720, 480)
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        """构建说明、工作表下拉框、表格与关闭按钮。"""
        hint = QLabel(
            '预览为数据简表（每表最多显示前 50 行）。'
            '合并单元格、列宽等完整样式请打开导出的 Excel 文件查看。'
        )
        hint.setWordWrap(True)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._show_sheet)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)

        close_btn = QPushButton('关闭')
        close_btn.clicked.connect(self.accept)

        top = QHBoxLayout()
        top.addWidget(QLabel('工作表：'))
        top.addWidget(self.sheet_combo, 1)

        layout = QVBoxLayout(self)
        layout.addWidget(hint)
        layout.addLayout(top)
        layout.addWidget(self.table, 1)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

    def _load_data(self):
        """一次性读入全部 sheet，填充下拉框并展示第一个。"""
        self.sheets = pd.read_excel(self.excel_path, sheet_name=None)
        names = list(self.sheets.keys())
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(names)
        self.sheet_combo.blockSignals(False)
        if names:
            self._show_sheet(names[0])

    def _show_sheet(self, name):
        """将当前 sheet 的前 PREVIEW_MAX_ROWS 行渲染到 QTableWidget。"""
        if not name or name not in self.sheets:
            self.table.clear()
            return
        df = self.sheets[name].head(PREVIEW_MAX_ROWS)
        rows, cols = df.shape
        self.table.clear()
        self.table.setRowCount(rows)
        self.table.setColumnCount(cols)
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for r in range(rows):
            for c in range(cols):
                val = df.iat[r, c]
                text = '' if pd.isna(val) else str(val)
                self.table.setItem(r, c, QTableWidgetItem(text))
        self.table.resizeColumnsToContents()


class RunWindow(QWidget):
    """主窗口：路径配置、一键采集、预览汇总、运行日志。"""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.saved_config = self._load_config()
        self.setWindowTitle('1688 商品链接数据采集')
        self.resize(520, 560)
        self.setMinimumSize(460, 500)
        self.setMaximumWidth(640)
        self._build_ui()
        self._load_defaults()

    def _build_ui(self):
        """组装路径输入、操作按钮与日志区。"""
        self.file_path_input = QLineEdit()
        self.file_path_input.setPlaceholderText('选择含「描述」「链接」列的 Excel')
        browse_excel_btn = QPushButton('浏览')
        browse_excel_btn.setFixedWidth(64)
        browse_excel_btn.clicked.connect(self._choose_excel)

        self.output_path_input = QLineEdit()
        self.output_path_input.setPlaceholderText('JSON 与规格汇总输出目录')
        browse_out_btn = QPushButton('浏览')
        browse_out_btn.setFixedWidth(64)
        browse_out_btn.clicked.connect(self._choose_output)

        excel_row = QHBoxLayout()
        excel_row.setSpacing(8)
        excel_row.addWidget(self.file_path_input, 1)
        excel_row.addWidget(browse_excel_btn)

        out_row = QHBoxLayout()
        out_row.setSpacing(8)
        out_row.addWidget(self.output_path_input, 1)
        out_row.addWidget(browse_out_btn)

        # 运行环境：线下 / 线上（与下载美国站子项目一致，目前仅作环境标记）
        env_row = QHBoxLayout()
        env_row.addWidget(QLabel('运行环境'))
        self.env_offline = QRadioButton('线下')
        self.env_online = QRadioButton('线上')
        self.env_offline.setChecked(True)
        env_group = QButtonGroup(self)
        env_group.addButton(self.env_offline, 0)
        env_group.addButton(self.env_online, 1)
        env_row.addWidget(self.env_offline)
        env_row.addWidget(self.env_online)
        env_row.addStretch(1)

        # 邮件通知
        mail_row = QHBoxLayout()
        mail_row.addWidget(QLabel('邮件通知'))
        self.mail_no = QRadioButton('不发送')
        self.mail_yes = QRadioButton('发送')
        self.mail_no.setChecked(True)
        mail_group = QButtonGroup(self)
        mail_group.addButton(self.mail_no, 0)
        mail_group.addButton(self.mail_yes, 1)
        self.mail_no.toggled.connect(self._toggle_email_entry)
        self.mail_yes.toggled.connect(self._toggle_email_entry)
        mail_row.addWidget(self.mail_no)
        mail_row.addWidget(self.mail_yes)
        mail_row.addStretch(1)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText('选择发送邮件时填写接收邮箱')

        self.btn_run = QPushButton('一键采集并导出')
        self.btn_preview = QPushButton('预览汇总')
        for btn in (self.btn_run, self.btn_preview):
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_run.clicked.connect(self._start_run)
        self.btn_preview.clicked.connect(self._preview)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        btn_row.addWidget(self.btn_run)
        btn_row.addWidget(self.btn_preview)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText('运行日志将显示在这里...')
        self.log_box.setMinimumHeight(180)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        layout.addWidget(QLabel('Excel 路径'))
        layout.addLayout(excel_row)
        layout.addWidget(QLabel('输出目录'))
        layout.addLayout(out_row)
        layout.addLayout(env_row)
        layout.addLayout(mail_row)
        layout.addWidget(QLabel('接收邮箱'))
        layout.addWidget(self.email_input)
        layout.addLayout(btn_row)
        layout.addWidget(QLabel('运行日志'))
        layout.addWidget(self.log_box, 1)
        self._toggle_email_entry()

    def _toggle_email_entry(self):
        """选择发送邮件时才启用邮箱输入框。"""
        enabled = self.mail_yes.isChecked()
        self.email_input.setEnabled(enabled)

    def _load_config(self):
        """从 run_config.json 恢复上次界面配置。"""
        empty = {
            'file_path': '',
            'path': '',
            'isOnline': False,
            'sendEmail': False,
            'email': '',
        }
        if not CONFIG_FILE.is_file():
            return empty
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
            return {
                'file_path': str(data.get('file_path', '') or '').strip(),
                'path': str(data.get('path', '') or '').strip(),
                'isOnline': bool(data.get('isOnline', False)),
                'sendEmail': bool(data.get('sendEmail', False)),
                'email': str(data.get('email', '') or '').strip(),
            }
        except (json.JSONDecodeError, OSError):
            return empty

    def _save_config(self):
        """将当前界面路径写入 run_config.json。"""
        cfg = self._current_config()
        try:
            CONFIG_FILE.write_text(
                json.dumps(cfg, ensure_ascii=False, indent=2),
                encoding='utf-8',
            )
        except OSError:
            pass

    def _load_defaults(self):
        """启动时把持久化配置填回输入框。"""
        self.file_path_input.setText(self.saved_config.get('file_path', ''))
        self.output_path_input.setText(self.saved_config.get('path', ''))
        if self.saved_config.get('isOnline'):
            self.env_online.setChecked(True)
        else:
            self.env_offline.setChecked(True)
        if self.saved_config.get('sendEmail'):
            self.mail_yes.setChecked(True)
        else:
            self.mail_no.setChecked(True)
        self.email_input.setText(self.saved_config.get('email', ''))
        self._toggle_email_entry()

    def _current_config(self):
        """读取界面当前配置，组装为 main.Ali1688 所需的 config 字典。"""
        return {
            'file_path': self.file_path_input.text().strip(),
            'path': self.output_path_input.text().strip(),
            'isOnline': self.env_online.isChecked(),
            'sendEmail': self.mail_yes.isChecked(),
            'email': self.email_input.text().strip(),
        }

    def _choose_excel(self):
        """文件对话框选择输入 Excel 并保存配置。"""
        path, _ = QFileDialog.getOpenFileName(
            self,
            '选择 Excel 文件',
            self.file_path_input.text().strip() or str(Path.home()),
            'Excel 文件 (*.xlsx *.xls)',
        )
        if path:
            self.file_path_input.setText(path)
            self._save_config()

    def _choose_output(self):
        """文件夹对话框选择 JSON/汇总输出目录并保存配置。"""
        path = QFileDialog.getExistingDirectory(
            self,
            '选择输出目录',
            self.output_path_input.text().strip() or str(Path.home()),
        )
        if path:
            self.output_path_input.setText(path)
            self._save_config()

    def _validate_config(self, need_excel=True):
        """校验路径合法性；通过时顺带持久化。返回 config 或 None。"""
        cfg = self._current_config()
        if not cfg['path']:
            QMessageBox.warning(self, '参数错误', '请填写输出目录 (path)。')
            return None
        out = Path(cfg['path'])
        if not out.is_dir():
            QMessageBox.warning(self, '路径无效', '输出目录不存在。')
            return None
        if cfg.get('sendEmail') and not cfg.get('email'):
            QMessageBox.warning(self, '参数错误', '选择发送邮件时必须填写接收邮箱。')
            return None
        if need_excel:
            if not cfg['file_path']:
                QMessageBox.warning(self, '参数错误', '请填写 Excel 路径 (file_path)。')
                return None
            if not Path(cfg['file_path']).is_file():
                QMessageBox.warning(self, '路径无效', 'Excel 文件不存在。')
                return None
        self._save_config()
        return cfg

    def _set_buttons_enabled(self, enabled):
        """任务运行期间禁用操作按钮，防止重复提交。"""
        self.btn_run.setEnabled(enabled)
        self.btn_preview.setEnabled(enabled)

    def _start_run(self):
        """在后台线程执行 Ali1688.run()（采集 + 导出）。"""
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, '运行中', '请等待当前任务结束。')
            return

        cfg = self._validate_config(need_excel=True)
        if not cfg:
            return

        self.log_box.clear()
        self.log_box.append('=== 一键采集并导出 ===')
        self.log_box.append(f"Excel: {cfg.get('file_path', '-')}")
        self.log_box.append(f"输出: {cfg['path']}")
        self.log_box.append(f"运行环境: {'线上' if cfg.get('isOnline') else '线下'}")
        self.log_box.append(f"邮件通知: {'发送' if cfg.get('sendEmail') else '不发送'}")
        if cfg.get('sendEmail'):
            self.log_box.append(f"接收邮箱: {cfg.get('email', '')}")
        self.log_box.append('-' * 60)

        self._set_buttons_enabled(False)
        self.worker = Worker('run', cfg)
        self.worker.log_signal.connect(self._append_log)
        self.worker.done_signal.connect(self._on_done)
        self.worker.start()

    def _append_log(self, line):
        """追加日志并滚动到底部。"""
        self.log_box.append(line)
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum(),
        )

    def _on_done(self, ok, message):
        """Worker 结束：恢复按钮、写结束日志并弹窗提示。"""
        self._set_buttons_enabled(True)
        self.log_box.append('-' * 60)
        self.log_box.append(message)
        if ok:
            QMessageBox.information(self, '完成', message)
        else:
            QMessageBox.critical(self, '失败', message)

    def _preview(self):
        """打开 PreviewDialog 预览输出目录下的 规格汇总.xlsx。"""
        cfg = self._validate_config(need_excel=False)
        if not cfg:
            return
        excel_path = Path(cfg['path']) / '规格汇总.xlsx'
        if not excel_path.is_file():
            QMessageBox.warning(
                self,
                '文件不存在',
                '未找到规格汇总.xlsx，请先执行「一键采集并导出」。',
            )
            return
        try:
            dlg = PreviewDialog(excel_path, self)
            dlg.exec_()
        except Exception as exc:
            QMessageBox.critical(self, '预览失败', str(exc))


def main():
    """启动 PyQt 应用并显示主窗口。"""
    app = QApplication(sys.argv)
    win = RunWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
