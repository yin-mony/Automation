# -*- coding: utf-8 -*-
"""运行入口：支持 CLI 与 Qt 双模式。"""

import argparse
import sys
from pathlib import Path

from Filter_add import DEFAULT_TOTAL_COL, run_interactive, run_pipeline


def run_cli_mode():
    # 参数: 无
    # 返回: 无返回值; 调用命令行交互流程
    run_interactive()


def run_gui_mode():
    # 参数: 无
    # 返回: 无返回值; 启动 Qt 图形界面并执行流程
    try:
        from PySide6.QtWidgets import (
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
    except ImportError as exc:
        raise RuntimeError("未安装 PySide6，无法启动 Qt 模式。请先安装: pip install PySide6") from exc

    class MainWindow(QWidget):
        def __init__(self):
            # 参数: 无
            # 返回: 无返回值; 初始化主窗口与控件
            super().__init__()
            self.setWindowTitle("匹配与ASIN回填工具")
            self.resize(820, 520)
            self._build_ui()

        def _build_ui(self):
            # 参数: 无
            # 返回: 无返回值; 组装窗口布局与控件
            layout = QVBoxLayout(self)

            self.total_edit = QLineEdit("")
            self.sub_edit = QLineEdit("")
            self.total_edit.setPlaceholderText("请选择主表文件（Excel）")
            self.sub_edit.setPlaceholderText("请选择副表文件（Excel）")
            self.log_output = QTextEdit()
            self.log_output.setReadOnly(True)

            layout.addLayout(self._create_file_row("主表路径", self.total_edit))
            layout.addLayout(self._create_file_row("副表路径", self.sub_edit))

            run_button = QPushButton("开始执行")
            run_button.clicked.connect(self._run_pipeline)
            layout.addWidget(run_button)
            layout.addWidget(QLabel("执行日志："))
            layout.addWidget(self.log_output)

        def _create_file_row(self, label_text, line_edit):
            # 参数: label_text=行标题, line_edit=对应输入框控件
            # 返回: 一行文件选择布局对象
            row = QHBoxLayout()
            row.addWidget(QLabel(label_text))
            row.addWidget(line_edit)

            browse_button = QPushButton("选择文件")
            browse_button.clicked.connect(lambda: self._pick_file(line_edit))
            row.addWidget(browse_button)
            return row

        def _pick_file(self, line_edit):
            # 参数: line_edit=需要回写路径的输入框控件
            # 返回: 无返回值; 选择文件后更新输入框
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择Excel文件",
                str(Path.home()),
                "Excel Files (*.xlsx *.xls)",
            )
            if file_path:
                line_edit.setText(file_path)

        def _run_pipeline(self):
            # 参数: 无
            # 返回: 无返回值; 读取输入路径并执行主流程
            total_path = self.total_edit.text().strip()
            sub_path = self.sub_edit.text().strip()
            if not total_path or not sub_path:
                QMessageBox.warning(self, "提示", "请先选择主表和副表路径")
                return

            total_path_obj = Path(total_path)
            sub_path_obj = Path(sub_path)
            if not total_path_obj.exists():
                QMessageBox.warning(self, "路径错误", f"主表文件不存在:\n{total_path_obj}")
                return
            if not sub_path_obj.exists():
                QMessageBox.warning(self, "路径错误", f"副表文件不存在:\n{sub_path_obj}")
                return

            try:
                result = run_pipeline(
                    total_path=total_path_obj,
                    sub_path=sub_path_obj,
                    print_summary=False,
                    save_result=True,
                    output_path=total_path_obj,
                )
                total = len(result["sub_result"])
                success = int(result["sub_result"]["is_match"].sum())
                failed = int((~result["sub_result"]["is_match"]).sum())
                target_col_used = result["target_col_used"]
                self.log_output.clear()
                self.log_output.append(f"总匹配行数: {total}")
                self.log_output.append(f"匹配成功: {success}")
                self.log_output.append(f"匹配失败: {failed}")
                self.log_output.append("")
                self.log_output.append("主表回填预览（前10行）：")
                self.log_output.append(
                    result["total_df_filled"][[DEFAULT_TOTAL_COL, target_col_used]].head(10).to_string(index=False)
                )
                QMessageBox.information(self, "完成", "已成功完成匹配且进行回填")
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "执行失败", str(exc))

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()


def main():
    # 参数: 无
    # 返回: 无返回值; 根据 mode 分发 CLI 或 GUI 入口
    parser = argparse.ArgumentParser(description="主表描述与副表订单ID匹配，并回填ASIN到主表编码列")
    parser.add_argument(
        "--mode",
        choices=("cli", "gui"),
        default="gui",
        help="运行模式：cli(命令行交互) 或 gui(Qt界面)",
    )
    args = parser.parse_args()

    if args.mode == "gui":
        run_gui_mode()
    else:
        run_cli_mode()


if __name__ == "__main__":
    main()