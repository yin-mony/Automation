# -*- coding: utf-8 -*-
"""
按店铺列拆分 xlsx（Qt 界面）。可配合 PyInstaller 打包为 exe。
依赖: pip install pandas openpyxl PySide6
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

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

# 预设的「店铺」相关表头名称列表；auto_detect_store_column 按此顺序在 Excel 列名里查找第一个匹配项
STORE_COLUMN_CANDIDATES = (
    "店铺",
    "门店",
    "店名",
    "店铺名称",
    "门店名称",
    "网店",
    "店铺名",
    "Store",
    "store",
    "Shop",
    "shop",
    "门店编码",
    "店铺编码",
)


def _sanitize_filename(name: str, max_len: int = 120) -> str:
    """
    将单元格中的店铺名称转成可安全用作 Windows 文件名的字符串。
    去掉首尾空格、替换非法字符、压缩空白；空值或 nan 则固定为「未填写店铺」。
    若过长则截断到 max_len，避免路径过长。
    """
    s = str(name).strip()
    if not s or s.lower() == "nan":
        return "未填写店铺"
    for ch in r'\/:*?"<>|':
        s = s.replace(ch, "_")
    s = re.sub(r"\s+", " ", s).strip()
    return s[:max_len] if len(s) > max_len else s


def read_xlsx_columns(xlsx_path: Path) -> list:
    """
    仅读取 xlsx 的表头行（不加载数据行），用于快速填充界面上的「店铺列」下拉框。
    返回 pandas 解析得到的列名列表（元素类型与完整读表时一致）。
    """
    df = pd.read_excel(xlsx_path, engine="openpyxl", nrows=0)
    return list(df.columns)


def auto_detect_store_column(columns: list) -> object | None:
    """
    根据 STORE_COLUMN_CANDIDATES 在传入的列名列表中自动识别「店铺列」。
    先按去空格后的精确匹配，再按不区分大小写匹配；若无任一命中则返回 None。
    """
    stripped_map = {str(c).strip(): c for c in columns}
    normalized = {str(c).strip().lower(): c for c in columns}
    for cand in STORE_COLUMN_CANDIDATES:
        key = cand.strip()
        if key in stripped_map:
            return stripped_map[key]
        low = key.lower()
        if low in normalized:
            return normalized[low]
    return None


def split_by_store(xlsx_path: Path, store_col: object, out_dir: Path) -> int:
    """
    读取完整 Excel，按指定列（店铺列）分组，每个分组写入单独的一个 xlsx 文件。
    若同一店铺名经 _sanitize_filename 后文件名冲突，则自动追加 _2、_3 等后缀。
    返回成功写出的文件个数。
    """
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    if df.empty:
        raise ValueError("表格为空，无法拆分。")

    columns = list(df.columns)
    actual = next(
        (c for c in columns if str(c).strip() == str(store_col).strip()),
        None,
    )
    if actual is None:
        raise ValueError("所选「店铺列」在当前表中不存在，请重新加载文件后选择列名。")

    out_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    for store_val, part in df.groupby(actual, dropna=False):
        label = _sanitize_filename(store_val if pd.notna(store_val) else "未填写店铺")
        out_path = out_dir / f"{label}.xlsx"
        if out_path.exists():
            base = label
            n = 2
            while True:
                cand = out_dir / f"{base}_{n}.xlsx"
                if not cand.exists():
                    out_path = cand
                    break
                n += 1
        part.to_excel(out_path, index=False, engine="openpyxl")
        count += 1
    return count


class _SplitWorker(QObject):
    """
    在子线程中执行拆分逻辑：避免大文件读写时阻塞 Qt 主界面。
    通过 finished_ok / failed 信号把结果或错误信息发回主线程。
    """

    finished_ok = Signal(int)  # 拆分成功时携带生成的文件数量
    failed = Signal(str)  # 拆分失败时携带错误说明字符串

    def __init__(self, xlsx_path: Path, store_col: object, out_dir: Path) -> None:
        """保存待处理的输入路径、店铺列标识、输出目录，供 run() 使用。"""
        super().__init__()
        self._xlsx_path = xlsx_path
        self._store_col = store_col
        self._out_dir = out_dir

    def run(self) -> None:
        """由 QThread.started 触发：调用 split_by_store，成功则 emit finished_ok，异常则 emit failed。"""
        try:
            n = split_by_store(self._xlsx_path, self._store_col, self._out_dir)
            self.finished_ok.emit(n)
        except Exception as e:
            self.failed.emit(str(e))


class MainWindow(QMainWindow):
    """主窗口：选择输入 xlsx、输出目录、店铺列，并触发后台拆分与日志展示。"""

    def __init__(self) -> None:
        """搭建界面布局，绑定按钮信号，初始化路径/线程/工作对象引用。"""
        super().__init__()
        self.setWindowTitle("按门店拆分表格")
        self.setMinimumSize(520, 420)

        self._input_path: Path | None = None
        self._thread: QThread | None = None
        self._worker: _SplitWorker | None = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        file_box = QGroupBox("文件")
        file_form = QFormLayout(file_box)
        row_in = QHBoxLayout()
        self.edit_input = QLineEdit()
        self.edit_input.setReadOnly(True)
        self.edit_input.setPlaceholderText("请选择 .xlsx 文件")
        btn_in = QPushButton("浏览…")
        btn_in.clicked.connect(self._pick_input)
        row_in.addWidget(self.edit_input, 1)
        row_in.addWidget(btn_in)
        w_in = QWidget()
        w_in.setLayout(row_in)
        file_form.addRow("Excel 文件:", w_in)

        row_out = QHBoxLayout()
        self.edit_output = QLineEdit()
        self.edit_output.setReadOnly(True)
        self.edit_output.setPlaceholderText("请选择输出文件夹")
        btn_out = QPushButton("浏览…")
        btn_out.clicked.connect(self._pick_output)
        row_out.addWidget(self.edit_output, 1)
        row_out.addWidget(btn_out)
        w_out = QWidget()
        w_out.setLayout(row_out)
        file_form.addRow("输出目录:", w_out)
        layout.addWidget(file_box)

        col_box = QGroupBox("店铺列")
        col_layout = QVBoxLayout(col_box)
        col_layout.addWidget(QLabel("选择用于拆分的列（选择文件后会自动尝试识别）："))
        self.combo_column = QComboBox()
        self.combo_column.setEnabled(False)
        col_layout.addWidget(self.combo_column)
        layout.addWidget(col_box)

        self.btn_run = QPushButton("开始拆分")
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self._start_split)
        layout.addWidget(self.btn_run)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("运行日志…")
        layout.addWidget(self.log, 1)

    def _append_log(self, text: str) -> None:
        """在窗口底部只读文本框中追加一行日志。"""
        self.log.append(text)

    def _pick_input(self) -> None:
        """
        弹出文件对话框选择 xlsx；读取表头填充「店铺列」下拉框；
        尝试 auto_detect_store_column 自动选中店铺列，并刷新「开始拆分」是否可点。
        """
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 表格",
            "",
            "Excel 工作簿 (*.xlsx);;所有文件 (*.*)",
        )
        if not path:
            return
        self._input_path = Path(path)
        self.edit_input.setText(str(self._input_path))
        try:
            cols = read_xlsx_columns(self._input_path)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法读取表头：\n{e}")
            self._clear_columns()
            return

        self.combo_column.clear()
        self.combo_column.setEnabled(True)
        for c in cols:
            self.combo_column.addItem(str(c), c)

        detected = auto_detect_store_column(cols)
        if detected is not None:
            for i in range(self.combo_column.count()):
                data = self.combo_column.itemData(i)
                if data is not None and str(data).strip() == str(detected).strip():
                    self.combo_column.setCurrentIndex(i)
                    break
            self._append_log(f"已加载列名，自动选中店铺列：{detected!s}")
        else:
            self.combo_column.setCurrentIndex(0)
            self._append_log("已加载列名，请手动选择「店铺」对应的列。")

        self._update_run_enabled()

    def _clear_columns(self) -> None:
        """清空店铺列下拉框并禁用，用于读表头失败等场景；同时更新「开始拆分」状态。"""
        self.combo_column.clear()
        self.combo_column.setEnabled(False)
        self._update_run_enabled()

    def _pick_output(self) -> None:
        """弹出目录对话框选择拆分结果保存路径，写入输出路径输入框并更新「开始拆分」状态。"""
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if d:
            self.edit_output.setText(d)
        self._update_run_enabled()

    def _update_run_enabled(self) -> None:
        """
        根据是否已选有效输入文件、输出目录、以及下拉框是否有当前列，
        决定「开始拆分」按钮是否可用。
        """
        ok = (
            self._input_path is not None
            and self._input_path.is_file()
            and self.edit_output.text().strip()
            and self.combo_column.currentIndex() >= 0
        )
        self.btn_run.setEnabled(bool(ok))

    def _start_split(self) -> None:
        """
        校验输入与输出；禁用按钮防止重复点击；创建 QThread 与 _SplitWorker，
        在子线程执行拆分，通过信号连接完成/失败回调与线程退出清理。
        """
        if not self._input_path or not self._input_path.is_file():
            QMessageBox.warning(self, "提示", "请先选择有效的 Excel 文件。")
            return
        out = self.edit_output.text().strip()
        if not out:
            QMessageBox.warning(self, "提示", "请选择输出目录。")
            return
        store_col = self.combo_column.currentData()
        if store_col is None and self.combo_column.count():
            store_col = self.combo_column.itemData(0)

        self.btn_run.setEnabled(False)
        self._append_log("正在拆分，请稍候…")

        self._thread = QThread()
        self._worker = _SplitWorker(self._input_path, store_col, Path(out))
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished_ok.connect(self._on_split_ok)
        self._worker.failed.connect(self._on_split_err)
        self._worker.finished_ok.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_thread_finished)
        self._thread.start()

    def _on_split_ok(self, n: int) -> None:
        """拆分成功：写日志、弹出完成提示框（携带输出目录与文件个数）。"""
        out = self.edit_output.text().strip()
        self._append_log(f"完成：共生成 {n} 个 xlsx 文件。\n输出目录：{out}")
        QMessageBox.information(
            self,
            "完成",
            f"已按店铺列拆分为 {n} 个文件。\n\n保存位置：\n{out}",
        )

    def _on_split_err(self, msg: str) -> None:
        """拆分失败：写日志并弹出错误对话框。"""
        self._append_log(f"错误：{msg}")
        QMessageBox.critical(self, "错误", msg)

    def _on_thread_finished(self) -> None:
        """
        线程结束后释放 Worker（deleteLater）、清空线程引用，
        并重新计算「开始拆分」是否可点（通常恢复为可点）。
        """
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None
        self._thread = None
        self._update_run_enabled()


def main() -> None:
    """
    程序入口：若命令行带 --cli 则走无界面拆分（便于脚本/调试）；
    否则启动 Qt 应用并显示主窗口。
    """
    if len(sys.argv) >= 4 and sys.argv[1] == "--cli":
        inp = Path(sys.argv[2])
        col = sys.argv[3]
        out = Path(sys.argv[4]) if len(sys.argv) > 4 else inp.parent / "拆分结果"
        n = split_by_store(inp, col, out)
        print(f"已拆分 {n} 个文件 -> {out}")
        return

    app = QApplication(sys.argv)
    app.setApplicationName("按门店拆分表格")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
