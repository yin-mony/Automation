# 亚马逊评论分析

易得客浏览器自动化抓取亚马逊商品各星级评论并导出 Excel；再调用 DeepSeek API 生成好评卖点、差评痛点与改进建议报告。提供两个 Tkinter GUI 入口，可分别打包为 exe。

## 模块说明

| 文件 | 用途 |
| --- | --- |
| `main.py` | `Comment` 类：登录易得客、启动店铺浏览器、抓取评论、导出 `亚马逊评论.xlsx` |
| `analysis.py` | `CommentAnalyzer` 类：读取评论 Excel，调用 DeepSeek 生成分析报告 |
| `YidekeLogin.py` | 易得客登录与浏览器启动 |
| `down_run.py` | **评论下载 GUI** 入口 |
| `excel_run.py` | **AI 分析 GUI** 入口 |

## 依赖

- Windows（依赖 `pywinauto`、易得客客户端）
- Python 3.10+ 推荐
- 安装：`python -m pip install -r requirements.txt`

## 运行

- **评论下载 GUI**：`python down_run.py`
- **AI 分析 GUI**：`python excel_run.py`
- **命令行（需自行改各文件末尾 config）**：`python main.py` / `python analysis.py`

界面会读写同目录下的 `comment_download_gui_config.txt`、`comment_analyzer_gui_config.txt`；**不要将含真实密码或 API Key 的配置文件提交到 Git**。

## 打包 exe

在项目目录执行：

- `build_down_exe.bat` → `dist\亚马逊评论下载\亚马逊评论下载.exe`
- `build_excel_exe.bat` → `dist\亚马逊评论分析\亚马逊评论分析.exe`
- `build_all_exe.bat` → 以上两个依次打包

打包前脚本会自动执行 `pip install -r requirements.txt`。分发时需带上整个 `dist\…` 文件夹。

## 仓库说明

`logs/`、`dist/`、`build/`、`*.xlsx`、`comment_*_gui_config.txt` 等已在 `.gitignore` 中排除，避免将业务数据与密钥推送到远程。
