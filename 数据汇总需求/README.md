# 数据汇总需求

按订单（`myp_order_id`）汇总 Excel 中的 `msku` 与 `charged_amount`；剔除「同一订单仅一种 `msku`」的整单数据。提供图形界面选择多个 `.xlsx` / `.xls` 批量处理。

## 依赖

- Python 3.10+（建议；需与本机已安装的 PyQt5 / PyInstaller  wheel 匹配）
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

## 输入表要求

首行表头需包含列：`myp_order_id`、`msku`、`charged_amount`。处理结果写回原文件（新增/覆盖 Sheet1 与 sheet2）。
