# 亚马逊评论工具

统一 GUI 工具，包含两个独立功能页：

- **评论下载**：通过易得客启动店铺浏览器，按需进入 Amazon Seller 后台，再在 Amazon 搜索 ASIN，按 1～5 星抓取评论并导出 Excel。
- **AI 分析**：读取评论 Excel，调用 OpenAI API 生成好评卖点、差评痛点与改进建议报告。

## 规范说明

本子项目按 Automation 通用 `docs/NAMING.md` 规范整理：

- 运行代码按「一文件一主类」拆分。
- 文件名保持简短，方法和变量使用 `camelCase`。
- 每个 Python 文件保留 `if __name__ == "__main__"` 便于独立调试。
- 评论下载主流程保留在 `main.py`，旧版主流程代码在文件底部以注释形式归档，便于后续对照参考。

## 模块说明

| 文件 | 主类 | 用途 |
| --- | --- | --- |
| `run.py` | `RunGui` | 统一窗口入口，内部子类 `DownloadPage`、`AnalysisPage` 与 `LogStream` 分别实现下载页、分析页和 GUI 日志路由 |
| `main.py` | `Auto` | 易得客进店、Amazon 后台登录、Amazon 评论抓取、Excel 导出主流程，文件底部保留旧流程注释归档 |
| `analysis.py` | `CommentAnalyzer` | 评论分析核心流程：读取 Excel、调用 OpenAI、生成报告 |
| `YidekeLogin.py` | `YidekeLogin` | 易得客客户端启动、登录、调试端口清理，内部子类 `AmazonSeller` 负责 Seller 登录辅助 |

## 依赖

- Windows（依赖 `pywinauto`、易得客客户端）
- Python 3.10+ 推荐
- 安装：`python -m pip install -r requirements.txt`

## 运行

```bat
python run.py
```

## 打包 exe

在项目目录执行：

```bat
build_all_exe.bat
```

打包脚本会直接以 `run.py` 作为主窗口入口调用 PyInstaller。

输出：

```text
dist\亚马逊评论工具\亚马逊评论工具.exe
dist\亚马逊评论工具.zip
```

## 输出

- 评论下载：保存为所选目录下的 `亚马逊评论.xlsx`
- AI 分析：保存到评论 Excel 同目录下的 `分析报告` 文件夹
- 日志：保存到程序目录下的 `logs` 文件夹

## 本地配置

界面会读写同目录下的：

- `comment_download_gui_config.txt`
- `comment_analyzer_gui_config.txt`

这些文件可能包含账号、密码或 API Key，已由 `.gitignore` 排除，不要提交到远程仓库。
