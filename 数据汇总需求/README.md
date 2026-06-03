# 数据汇总需求

按订单（`myp_order_id`）汇总 Excel 中的 `msku` 与 `charged_amount`；剔除「同一订单仅一种 `msku`」的整单数据。提供图形界面选择多个 `.xlsx` / `.xls` 批量处理。

## 代码结构

```
数据汇总需求/
├── main.py                 # openpyxl 汇总逻辑
├── run.py                  # PyQt5 GUI 入口
├── requirements.txt
├── run_gui.bat
├── build_exe.bat
├── 数据汇总处理工具.spec
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | `process_excel_file()`：按订单分组，剔除单 msku 订单；写 Sheet1 筛选结果 + sheet2 汇总 |
| `run.py` | `MainWindow` + `WorkerThread` 多文件批量调用 `process_multiple_files` |
| `build_exe.bat` | 打包为 `dist\数据汇总处理工具.exe` |

### 输入表要求

| 列名 | 用途 |
| --- | --- |
| `myp_order_id` | 订单分组键 |
| `msku` | 判断订单是否含多种 SKU |
| `charged_amount` | sheet2 汇总金额 |

## 依赖

- Python 3.10+（建议；需与本机已安装的 PyQt5 / PyInstaller wheel 匹配）
- 见 `requirements.txt`

## 本地运行 GUI

双击 `run_gui.bat`，或在目录下执行：

```bat
python -m pip install -r requirements.txt
python run.py
```

## 打包为单文件 exe

1. 在本目录打开命令行，或双击 `build_exe.bat`。
2. 成功后生成：`dist\数据汇总处理工具.exe`（无控制台窗口）。

说明：`build/`、`dist/` 已在仓库根 `.gitignore` 中忽略，构建产物不会提交到 Git。
