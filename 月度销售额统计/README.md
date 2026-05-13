# 月度销售额统计

易得客（eDecker）浏览器自动化：按达人账号列表导出视频表现报表，并与本地汇总表路径配置合并；含 Tkinter 图形界面（`run.py`）与 Excel 合并逻辑（`analysis.py`）。

## 依赖

- Windows（依赖 `pywinauto`、易得客客户端路径）
- Python 3.10+ 推荐
- 安装：`python -m pip install -r requirements.txt`

## 运行

- **GUI（推荐）**：`python run.py`
- **仅跑自动化脚本（需自行改 `main.py` 末尾 config）**：`python main.py`

首次使用请将 `monthly_sales_gui_config.example.txt` 复制为 `monthly_sales_gui_config.txt`（与 `run.py` 同目录），填写汇总表路径、导出目录、易得客账号与店铺 IP 等；**不要将含真实密码的配置文件提交到 Git**。

## 打包 exe

在项目目录执行 `build_exe.bat`（另需 `requirements-build.txt` 中的 PyInstaller）。产物为 `dist\月度销售额统计\` 目录下的可执行文件，分发时需带上该目录。

## 仓库说明

`flie/`、`logs/`、`*.xlsx`、`monthly_sales_gui_config.txt` 等已在 `.gitignore` 中排除，避免将业务数据与密钥推送到远程。
