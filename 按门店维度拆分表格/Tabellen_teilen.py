# -*- coding: utf-8 -*-
"""
程序入口：按店铺列拆分 xlsx 的 Qt 图形界面；可配合 PyInstaller 打包为 exe。
拆分算法在 excel_store_split.py 中实现。
依赖: pip install pandas openpyxl PySide6
"""

import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import excel_store_split as core


# =============================================================================
# 后台线程：在子线程里跑 pandas/openpyxl，避免卡住 Qt 主界面
# =============================================================================


class HeaderReadWorker(QObject):
    """在子线程读取 xlsx 表头，避免大文件或慢盘时卡住界面。"""

    finished = Signal(list)
    failed = Signal(str)

    def __init__(self, xlsx_path):
        super().__init__()
        self._path = Path(xlsx_path)

    def run(self):
        try:
            cols = core.read_headers(self._path)
            self.finished.emit(cols)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


class SplitWorker(QObject):
    """放到 QThread 里执行拆分；不要在主线程直接跑耗时逻辑。"""

    finished = Signal(int)  # 成功：生成文件数
    failed = Signal(str)  # 失败：错误信息

    # -------------------------------------------------------------------------
    # 传参：xlsx_path — 待拆分的 Excel 路径（pathlib.Path）
    #       store_col — 用作分组的列名（与表头一致的对象，通常来自下拉框 itemData）
    #       out_dir — 输出目录（pathlib.Path）
    # 返回：无（仅保存到 self._xlsx、self._col、self._out）
    # -------------------------------------------------------------------------
    def __init__(
        self,
        xlsx_path,
        store_col,
        out_dir,
        name_type=None,
        name_time=None,
    ):
        super().__init__()
        self._xlsx = xlsx_path
        self._col = store_col
        self._out = out_dir
        self._name_type = name_type
        self._name_time = name_time

    # -------------------------------------------------------------------------
    # 传参：无（使用 __init__ 中保存的路径与列）
    # 返回：无；成功则 emit finished(文件个数)，失败则 emit failed(错误字符串)
    # 作用：由 QThread.started 触发，内部调用 core.split_by_store
    # -------------------------------------------------------------------------
    def run(self):
        try:
            n = core.split_by_store(
                self._xlsx,
                self._col,
                self._out,
                name_type=self._name_type,
                name_time=self._name_time,
            )
            self.finished.emit(n)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


# =============================================================================
# 主窗口：选文件、选列、选目录、开始拆分、日志
# =============================================================================


class MainWindow(QMainWindow):
    """按门店拆分表格主界面。"""

    # -------------------------------------------------------------------------
    # 传参：无
    # 返回：无；初始化窗口、控件与 self._xlsx / self._thread / self._worker 状态
    # -------------------------------------------------------------------------
    def __init__(self):
        super().__init__()
        self.setWindowTitle("按门店拆分表格")
        self.setMinimumSize(520, 520)

        self._xlsx = None
        self._thread = None
        self._worker = None
        self._hdr_thread = None
        self._hdr_worker = None

        self._build_ui()
        self._set_naming_widgets_active(False)

    # -------------------------------------------------------------------------
    # 传参：无
    # 返回：无；创建中央控件、表单、下拉框、按钮与日志区
    # -------------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        file_box = QGroupBox("文件")
        form = QFormLayout(file_box)

        row_in = QHBoxLayout()
        self.ed_in = QLineEdit()
        self.ed_in.setReadOnly(True)
        self.ed_in.setPlaceholderText("请选择 .xlsx 文件")
        self.btn_in = QPushButton("浏览…")
        self.btn_in.clicked.connect(self._pick_file)
        row_in.addWidget(self.ed_in, 1)
        row_in.addWidget(self.btn_in)
        w_in = QWidget()
        w_in.setLayout(row_in)
        form.addRow("Excel 文件:", w_in)

        row_out = QHBoxLayout()
        self.ed_out = QLineEdit()
        self.ed_out.setReadOnly(True)
        self.ed_out.setPlaceholderText("请选择输出文件夹")
        btn_out = QPushButton("浏览…")
        btn_out.clicked.connect(self._pick_dir)
        row_out.addWidget(self.ed_out, 1)
        row_out.addWidget(btn_out)
        w_out = QWidget()
        w_out.setLayout(row_out)
        form.addRow("输出目录:", w_out)

        root.addWidget(file_box)

        col_box = QGroupBox("店铺列")
        col_lay = QVBoxLayout(col_box)
        col_lay.addWidget(QLabel("选择用于拆分的列（选择文件后会自动尝试识别）："))
        self.combo = QComboBox()
        self.combo.setEnabled(False)
        col_lay.addWidget(self.combo)
        root.addWidget(col_box)

        name_box = QGroupBox("拆分文件命名")
        name_root = QVBoxLayout(name_box)
        self.lbl_name_hint = QLabel(
            "默认按「店铺名称」单独作为文件名。"
            "若在下方勾选「类型」或「时间」任意一项并填写完整，则输出为「类型」-店铺名称，或「时间」-店铺名称；"
            "若「类型」与「时间」两项均勾选并填写完整，则输出为「类型」-「时间」-店铺名称；"
            "否则仍为仅店铺名称。"
        )
        self.lbl_name_hint.setWordWrap(True)
        name_root.addWidget(self.lbl_name_hint)

        form_name = QFormLayout()
        row_type = QHBoxLayout()
        row_type.setContentsMargins(0, 0, 0, 0)
        self.chk_type = QCheckBox("类型")
        self.chk_type.setToolTip("勾选后请在右侧填写用于文件名的类型文字")
        self.chk_type.toggled.connect(self._on_type_time_toggled)
        self.ed_type = QLineEdit()
        self.ed_type.setPlaceholderText("勾选「类型」后填写")
        self.ed_type.textChanged.connect(self._sync_run_btn)
        row_type.addWidget(self.chk_type)
        row_type.addWidget(self.ed_type, 1)
        w_type = QWidget()
        w_type.setLayout(row_type)
        form_name.addRow(w_type)

        row_time = QHBoxLayout()
        row_time.setContentsMargins(0, 0, 0, 0)
        self.chk_time = QCheckBox("时间")
        self.chk_time.setToolTip("勾选后请在右侧填写用于文件名的时间文字")
        self.chk_time.toggled.connect(self._on_type_time_toggled)
        self.ed_time = QLineEdit()
        self.ed_time.setPlaceholderText("勾选「时间」后填写")
        self.ed_time.textChanged.connect(self._sync_run_btn)
        row_time.addWidget(self.chk_time)
        row_time.addWidget(self.ed_time, 1)
        w_time = QWidget()
        w_time.setLayout(row_time)
        form_name.addRow(w_time)

        name_root.addLayout(form_name)
        root.addWidget(name_box)

        self.combo.currentIndexChanged.connect(self._sync_run_btn)

        self.btn_run = QPushButton("开始拆分")
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self._run_split)
        root.addWidget(self.btn_run)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("运行日志…")
        root.addWidget(self.log, 1)

    # -------------------------------------------------------------------------
    # 传参：text — 追加到日志区的字符串
    # 返回：无
    # -------------------------------------------------------------------------
    def _log(self, text):
        self.log.append(text)

    # -------------------------------------------------------------------------
    # 传参：无（从文件对话框取路径）
    # 返回：无；成功则设置 self._xlsx、填充下拉列名并可能自动选中店铺列
    # -------------------------------------------------------------------------
    def _pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 表格",
            "",
            "Excel 工作簿 (*.xlsx);;所有文件 (*.*)",
        )
        if not path:
            return

        if self._hdr_thread is not None and self._hdr_thread.isRunning():
            return

        self._xlsx = Path(path)
        self.ed_in.setText(str(self._xlsx))
        self._clear_cols(keep_path=True)
        self._start_header_read()

    def _start_header_read(self):
        self._log("正在后台读取表头…")
        self.btn_in.setEnabled(False)

        self._hdr_thread = QThread()
        self._hdr_worker = HeaderReadWorker(self._xlsx)
        self._hdr_worker.moveToThread(self._hdr_thread)

        self._hdr_thread.started.connect(self._hdr_worker.run)
        self._hdr_worker.finished.connect(self._on_header_read_finished)
        self._hdr_worker.failed.connect(self._on_header_read_failed)
        self._hdr_worker.finished.connect(self._hdr_thread.quit)
        self._hdr_worker.failed.connect(self._hdr_thread.quit)
        self._hdr_thread.finished.connect(self._hdr_thread.deleteLater)
        self._hdr_thread.finished.connect(self._on_header_thread_finished)
        self._hdr_thread.start()

    def _on_header_read_finished(self, cols):
        self.combo.clear()
        self.combo.setEnabled(True)
        for c in cols:
            self.combo.addItem(str(c), c)

        hit = core.guess_store_col(cols)
        if hit is not None:
            self._select_col(hit)
            self._log(f"已加载列名，自动选中店铺列：{hit!s}")
        else:
            self.combo.setCurrentIndex(0)
            self._log("已加载列名，请手动选择「店铺」对应的列。")

        self._set_naming_widgets_active(True)
        self._sync_run_btn()

    def _on_header_read_failed(self, msg):
        QMessageBox.critical(self, "读取失败", f"无法读取表头：\n{msg}")
        self._clear_cols(keep_path=False)
        self.ed_in.clear()
        self._xlsx = None
        self._sync_run_btn()

    def _on_header_thread_finished(self):
        self.btn_in.setEnabled(True)
        if self._hdr_worker is not None:
            self._hdr_worker.deleteLater()
            self._hdr_worker = None
        self._hdr_thread = None

    # -------------------------------------------------------------------------
    # 传参：col — 要在下拉框中选中的列名（与 itemData 比较时忽略首尾空格）
    # 返回：无；找不到匹配项则保持当前选中不变
    # -------------------------------------------------------------------------
    def _select_col(self, col):
        for i in range(self.combo.count()):
            d = self.combo.itemData(i)
            if d is not None and str(d).strip() == str(col).strip():
                self.combo.setCurrentIndex(i)
                return

    def _on_type_time_toggled(self, _checked=False):
        self._refresh_type_time_edits()
        self._sync_run_btn()

    def _refresh_type_time_edits(self):
        """勾选后才允许编辑对应填写框；取消勾选时清空。"""
        active = self.combo.isEnabled()
        self.ed_type.setEnabled(active and self.chk_type.isChecked())
        self.ed_time.setEnabled(active and self.chk_time.isChecked())
        if not self.chk_type.isChecked():
            self.ed_type.clear()
        if not self.chk_time.isChecked():
            self.ed_time.clear()

    def _set_naming_widgets_active(self, active):
        """已加载 Excel 表头后启用「类型」「时间」区域。"""
        self.chk_type.setEnabled(active)
        self.chk_time.setEnabled(active)
        if not active:
            self.chk_type.setChecked(False)
            self.chk_time.setChecked(False)
        self._refresh_type_time_edits()

    # -------------------------------------------------------------------------
    # 传参：keep_path — 为 True 时保留已选路径（重新读表头前清空下拉用）
    # 返回：无；清空店铺列下拉并禁用，刷新「开始拆分」按钮状态
    # -------------------------------------------------------------------------
    def _clear_cols(self, keep_path=False):
        self.combo.clear()
        self.combo.setEnabled(False)
        self.chk_type.setChecked(False)
        self.chk_time.setChecked(False)
        self.ed_type.clear()
        self.ed_time.clear()
        self._set_naming_widgets_active(False)
        if not keep_path:
            self._xlsx = None
        self._sync_run_btn()

    # -------------------------------------------------------------------------
    # 传参：无（从目录对话框取路径）
    # 返回：无；将所选目录写入输出路径输入框并刷新按钮状态
    # -------------------------------------------------------------------------
    def _pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if d:
            self.ed_out.setText(d)
        self._sync_run_btn()

    # -------------------------------------------------------------------------
    # 传参：无
    # 返回：无；根据文件、输出目录、列是否就绪设置「开始拆分」是否可点
    # -------------------------------------------------------------------------
    def _sync_run_btn(self):
        ok = (
            self._xlsx is not None
            and self._xlsx.is_file()
            and bool(self.ed_out.text().strip())
            and self.combo.currentIndex() >= 0
        )
        if ok and self.chk_type.isChecked() and not self.ed_type.text().strip():
            ok = False
        if ok and self.chk_time.isChecked() and not self.ed_time.text().strip():
            ok = False
        self.btn_run.setEnabled(ok)

    # -------------------------------------------------------------------------
    # 传参：无（从界面读取路径、列、输出目录）
    # 返回：无；启动 QThread + SplitWorker；结果由 _on_ok / _on_err 处理
    # -------------------------------------------------------------------------
    def _run_split(self):
        if not self._xlsx or not self._xlsx.is_file():
            QMessageBox.warning(self, "提示", "请先选择有效的 Excel 文件。")
            return
        out = self.ed_out.text().strip()
        if not out:
            QMessageBox.warning(self, "提示", "请选择输出目录。")
            return

        col = self.combo.currentData()
        if col is None and self.combo.count():
            col = self.combo.itemData(0)

        if self.chk_type.isChecked() and not self.ed_type.text().strip():
            QMessageBox.warning(self, "提示", "已勾选「类型」，请填写类型内容。")
            self._sync_run_btn()
            return
        if self.chk_time.isChecked() and not self.ed_time.text().strip():
            QMessageBox.warning(self, "提示", "已勾选「时间」，请填写时间内容。")
            self._sync_run_btn()
            return

        name_type = self.ed_type.text().strip() if self.chk_type.isChecked() else None
        name_time = self.ed_time.text().strip() if self.chk_time.isChecked() else None

        if name_type and name_time:
            self._log("命名方式：类型-时间-店铺名称。")
        elif name_type:
            self._log("命名方式：类型-店铺名称。")
        elif name_time:
            self._log("命名方式：时间-店铺名称。")
        else:
            self._log("命名方式：仅店铺名称。")

        self.btn_run.setEnabled(False)
        self._log("正在拆分，请稍候…")

        self._thread = QThread()
        self._worker = SplitWorker(
            self._xlsx,
            col,
            Path(out),
            name_type=name_type,
            name_time=name_time,
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_ok)
        self._worker.failed.connect(self._on_err)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_done)
        self._thread.start()

    # -------------------------------------------------------------------------
    # 传参：n — 拆分生成的 xlsx 文件个数（来自 SplitWorker.finished 信号）
    # 返回：无；写日志并弹出成功提示框
    # -------------------------------------------------------------------------
    def _on_ok(self, n):
        out = self.ed_out.text().strip()
        self._log(f"完成：共生成 {n} 个 xlsx 文件。\n输出目录：{out}")
        QMessageBox.information(
            self,
            "完成",
            f"已按店铺列拆分为 {n} 个文件。\n\n保存位置：\n{out}",
        )

    # -------------------------------------------------------------------------
    # 传参：msg — 错误信息字符串（来自 SplitWorker.failed 信号）
    # 返回：无；写日志并弹出错误提示框
    # -------------------------------------------------------------------------
    def _on_err(self, msg):
        self._log(f"错误：{msg}")
        QMessageBox.critical(self, "错误", msg)

    # -------------------------------------------------------------------------
    # 传参：无（在线程 finished 时调用）
    # 返回：无；释放 worker 引用、清空 thread 引用并恢复「开始拆分」可用状态
    # -------------------------------------------------------------------------
    def _on_done(self):
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None
        self._thread = None
        self._sync_run_btn()


# -----------------------------------------------------------------------------
# 传参：无（通过 sys.argv 判断模式）
# 返回：无；--cli 时拆分后 print 并 return，否则启动 Qt 事件循环（进程退出码由 app.exec 决定）
# 说明：命令行用法 Tabellen_teilen.py --cli <xlsx路径> <列名> [输出目录]
#       未给输出目录时默认为「输入文件同目录下的 拆分结果」
# -----------------------------------------------------------------------------
def main():
    if len(sys.argv) >= 4 and sys.argv[1] == "--cli":
        inp = Path(sys.argv[2])
        col = sys.argv[3]
        out = Path(sys.argv[4]) if len(sys.argv) > 4 else inp.parent / "拆分结果"
        n = core.split_by_store(inp, col, out)
        print(f"已拆分 {n} 个文件 -> {out}")
        return

    app = QApplication(sys.argv)
    app.setApplicationName("按门店拆分表格")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
