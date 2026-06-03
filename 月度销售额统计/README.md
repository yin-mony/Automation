# 月度销售额统计

易得客（eDecker）+ TikTok 店铺后台自动化：按达人账号导出视频 GMV 报表，并将下载结果合并进汇总 Excel。提供**两个独立 GUI 入口**（下载 / 汇总）及 PyInstaller 打包脚本。

## 代码结构

```
月度销售额统计/
├── main.py          # 核心自动化：登录易得客 → 访问店铺 → 导出各达人 xlsx
├── YidekeLogin.py   # 易得客登录（DrissionPage + pywinauto）
├── run.py           # GUI：数据下载（调用 main.Automation.Start）
├── 运行文件.py       # GUI：Excel 汇总（调用 analysis.ExcelUtil.MergeData）
├── analysis.py      # 读取导出 xlsx，按规则合并到汇总表
├── requirements.txt
├── requirements-build.txt
├── build_exe.bat    # 一键打包两个 exe
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | `Automation` 类：`Start()` 完成登录、点访问、按 profile 启动店铺浏览器、等待进入卖家后台、逐达人 Filter 并 Export |
| `YidekeLogin.py` | 启动易得客 9222、点击登录、pywinauto 填账号密码 |
| `run.py` | Tkinter 界面：配置账号 / 店铺 IP / 调试端口 / 达人列表 / 下载目录，后台线程执行 `Automation.Start()` |
| `运行文件.py` | Tkinter 界面：选择汇总表与数据文件夹，执行 `ExcelUtil.MergeData` |
| `analysis.py` | `ExcelUtil`：读取达人报表 xlsx，筛选 `Video items sold > 0`，写回汇总表 |

### 配置说明（`main.py` / `run.py`）

| 字段 | 含义 |
| --- | --- |
| `username` / `password` | 易得客账号 |
| `ip` | 店铺 VPS IP 列表（工作台显示的 IP，**不是** `127.0.0.1`） |
| `port` | 各店铺 profile 浏览器调试端口，与 `ip` 一一对应（如 `8945`） |
| `experts` | 达人账号列表 |
| `file_path` | xlsx 下载目录 |

`127.0.0.1:9222` 为易得客管理端；`127.0.0.1:{port}` 为店铺 profile 浏览器（DrissionPage 接管地址）。

## 依赖

- Windows（`pywinauto`、易得客客户端）
- Python 3.10+ 推荐
- 安装：`python -m pip install -r requirements.txt`

## 运行

- **数据下载 GUI**：`python run.py`
- **Excel 汇总 GUI**：`python 运行文件.py`
- **命令行（需改 `main.py` 末尾 config）**：`python main.py`

GUI 内可填账号与路径；**勿将含真实密码的配置文件提交到 Git**。

## 打包 exe

在项目目录执行 `build_exe.bat`（需 `requirements-build.txt`）。产物位于 `dist/`：

| exe | 入口 | 功能 |
| --- | --- | --- |
| `月度销售额下载.exe` | `run.py` | 易得客登录 + TikTok 报表下载 |
| `月度销售额汇总.exe` | `运行文件.py` | 汇总表与下载数据合并 |

打包前请关闭正在运行的同名 exe，避免文件占用导致失败。

## 仓库说明

`flie/`、`logs/`、`*.xlsx`、`build/`、`dist/`、`.venv/` 等已在 `.gitignore` 中排除。
