# 请款汇总

将请款单 PDF 批量转换为 Excel，并提取关键字段生成汇总表。

## 当前实现状态

- 已实现可运行的 Qt 图形界面入口：`run.py`。
- 支持从界面选择目标文件夹，并将路径传入处理脚本。
- 处理脚本 `test.py` 保留原有数据处理逻辑，负责 PDF 转换与汇总输出。
- 已支持打包为可执行文件（`PyInstaller`）。

## 运行方式

- 开发环境运行：`python run.py`
- 打包运行：执行 `dist` 目录内生成的 `请款汇总.exe` 或 `请款汇总-单文件.exe`

## 依赖说明

- Python 3.14（本地环境）
- PyQt5
- pandas
- pdfplumber
- openpyxl
- numpy
