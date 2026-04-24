# -*- coding: utf-8 -*-
"""
程序入口：按店铺列拆分 xlsx 的 Qt 图形界面；可配合 PyInstaller 打包为 exe。
拆分算法在 excel_store_split.py 中实现。
依赖: pip install pandas openpyxl PySide6
"""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
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


class SplitWorker(QObject):
    """放到 QThread 里执行拆分；不要在主线程直接跑耗时逻辑。"""

    finished = Signal(int)  # 成功：生成文件数
    failed = Signal(str)  # 失败：错误信息

    # -------------------------------------------------------------------------
    # 作用：保存本次任务的路径、店铺列、输出目录。
    # -------------------------------------------------------------------------
    def __init__(self, xlsx_path: Path, store_col: object, out_dir: Path) -> None:
        super().__init__()
        self._xlsx = xlsx_path
        self._col = store_col
        self._out = out_dir

    # -------------------------------------------------------------------------
    # 作用：由 QThread.started 触发；调 core.split_by_store，再通过信号回传结果。
    # -------------------------------------------------------------------------
    def run(self) -> None:
        try:
            n = core.split_by_store(self._xlsx, self._col, self._out)
            self.finished.emit(n)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


# =============================================================================
# 主窗口：选文件、选列、选目录、开始拆分、日志
# =============================================================================


class MainWindow(QMainWindow):
    """按门店拆分表格主界面。"""

    # -------------------------------------------------------------------------
    # 作用：初始化状态、搭建界面。
    # -------------------------------------------------------------------------
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("按门店拆分表格")
        self.setMinimumSize(520, 420)

        self._xlsx: Path | None = None
        self._thread: QThread | None = None
        self._worker: SplitWorker | None = None

        self._build_ui()

    # -------------------------------------------------------------------------
    # 作用：组装布局与信号绑定。
    # -------------------------------------------------------------------------
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        file_box = QGroupBox("文件")
        form = QFormLayout(file_box)

        row_in = QHBoxLayout()
        self.ed_in = QLineEdit()
        self.ed_in.setReadOnly(True)
        self.ed_in.setPlaceholderText("请选择 .xlsx 文件")
        btn_in = QPushButton("浏览…")
        btn_in.clicked.connect(self._pick_file)
        row_in.addWidget(self.ed_in, 1)
        row_in.addWidget(btn_in)
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

        self.btn_run = QPushButton("开始拆分")
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self._run_split)
        root.addWidget(self.btn_run)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("运行日志…")
        root.addWidget(self.log, 1)

    # -------------------------------------------------------------------------
    # 作用：追加一行日志。
    # -------------------------------------------------------------------------
    def _log(self, text: str) -> None:
        self.log.append(text)

    # -------------------------------------------------------------------------
    # 作用：选 xlsx，读表头填下拉框，尽量自动选中店铺列。
    # -------------------------------------------------------------------------
    def _pick_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 表格",
            "",
            "Excel 工作簿 (*.xlsx);;所有文件 (*.*)",
        )
        if not path:
            return

        self._xlsx = Path(path)
        self.ed_in.setText(str(self._xlsx))

        try:
            cols = core.read_headers(self._xlsx)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "读取失败", f"无法读取表头：\n{e}")
            self._clear_cols()
            return

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

        self._sync_run_btn()

    # -------------------------------------------------------------------------
    # 作用：在下拉框里选中 userData 与目标列（去空格）一致的一项。
    # -------------------------------------------------------------------------
    def _select_col(self, col: object) -> None:
        for i in range(self.combo.count()):
            d = self.combo.itemData(i)
            if d is not None and str(d).strip() == str(col).strip():
                self.combo.setCurrentIndex(i)
                return

    # -------------------------------------------------------------------------
    # 作用：读表头失败时清空下拉框并禁用。
    # -------------------------------------------------------------------------
    def _clear_cols(self) -> None:
        self.combo.clear()
        self.combo.setEnabled(False)
        self._sync_run_btn()

    # -------------------------------------------------------------------------
    # 作用：选输出目录。
    # -------------------------------------------------------------------------
    def _pick_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if d:
            self.ed_out.setText(d)
        self._sync_run_btn()

    # -------------------------------------------------------------------------
    # 作用：输入文件、输出目录、列都就绪时启用「开始拆分」。
    # -------------------------------------------------------------------------
    def _sync_run_btn(self) -> None:
        ok = (
            self._xlsx is not None
            and self._xlsx.is_file()
            and bool(self.ed_out.text().strip())
            and self.combo.currentIndex() >= 0
        )
        self.btn_run.setEnabled(ok)

    # -------------------------------------------------------------------------
    # 作用：后台启动拆分线程。
    # -------------------------------------------------------------------------
    def _run_split(self) -> None:
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

        self.btn_run.setEnabled(False)
        self._log("正在拆分，请稍候…")

        self._thread = QThread()
        self._worker = SplitWorker(self._xlsx, col, Path(out))
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
    # 作用：拆分成功：日志 + 弹窗。
    # -------------------------------------------------------------------------
    def _on_ok(self, n: int) -> None:
        out = self.ed_out.text().strip()
        self._log(f"完成：共生成 {n} 个 xlsx 文件。\n输出目录：{out}")
        QMessageBox.information(
            self,
            "完成",
            f"已按店铺列拆分为 {n} 个文件。\n\n保存位置：\n{out}",
        )

    # -------------------------------------------------------------------------
    # 作用：拆分失败：日志 + 弹窗。
    # -------------------------------------------------------------------------
    def _on_err(self, msg: str) -> None:
        self._log(f"错误：{msg}")
        QMessageBox.critical(self, "错误", msg)

    # -------------------------------------------------------------------------
    # 作用：线程结束，释放 worker，恢复按钮。
    # -------------------------------------------------------------------------
    def _on_done(self) -> None:
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None
        self._thread = None
        self._sync_run_btn()


# -----------------------------------------------------------------------------
# 作用：--cli 无界面拆分，否则启动 Qt。
# -----------------------------------------------------------------------------
def main() -> None:
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
