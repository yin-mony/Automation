# 请款汇总

将请款单 PDF 批量转换为 Excel，并提取关键字段生成汇总表。

## 代码结构

```
请款汇总/
├── run.py                  # PyQt5 GUI：选文件夹并执行 test.py
├── test.py                 # PDF→Excel + 字段汇总主逻辑
├── ceshi.py                # 已合并至 test.py 的占位说明
├── 请款汇总.spec
├── 请款汇总-单文件.spec
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `run.py` | `RunWindow`：选择目标文件夹，设置环境变量 `TARGET_FOLDER`，`runpy.run_path(test.py)` 并捕获日志 |
| `test.py` | 遍历子文件夹 PDF → `pdf_to_excel`；汇总 `采购单号`、`品名/SKU`、`采购单价` 等字段 |
| `ceshi.py` | 废弃提示，非正式入口 |

### 环境变量

| 变量 | 含义 |
| --- | --- |
| `TARGET_FOLDER` | 请款单批量下载根目录（含各子文件夹 PDF） |

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
