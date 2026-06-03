# 审核请款单

用于在赛狐 ERP 中批量下载请款单及附件，支持通过图形界面选择下载模式（最近 7 天、最近 30 天、本月、上月）并自动分页执行下载动作。

## 代码结构

```
审核请款单/
├── run.py                  # PyQt5 GUI 入口
├── test.py                 # 自动化主流程（登录、筛选、下载）
├── SaihuERPLogin.py        # 赛狐 ERP 登录（验证码 OCR）
├── 1111.py                 # 下载目录 zip 探测草稿
├── .gitignore
├── 审核请款单.spec
├── 请款单下载.spec
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `run.py` | GUI：账号/密码/下载模式/日志；后台线程调用 `test.main`；可读写 `run_config.json` |
| `test.py` | `main(mode, …)`：DrissionPage 登录 → 财务/请款单 → 按模式筛选 → 下载请款单及附件 |
| `SaihuERPLogin.py` | `SaihuERPLogin.login()`：会话复用、验证码识别、手动兜底 |
| `1111.py` | 在 Downloads 查找最新 `请款单*.zip` 的实验脚本 |

### 配置说明（`run.py` / `run_config.json`）

| 字段 | 含义 |
| --- | --- |
| `mode` | `recent_7_days` / `recent_30_days` / `this_month` / `last_month` |
| `username` / `password` | 赛狐账号（**勿将含真实密码的配置提交到 Git**） |

## 当前实现状态

- 已实现 `run.py` 图形界面入口（账号、密码、下载模式、日志输出）。
- 已实现 `test.py` 自动化主流程（登录、进入请款单页面、按模式筛选、逐页下载）。
- 已支持 PyInstaller 打包，当前可生成 `请款单下载.exe`。

## 运行方式

在本目录命令行执行：

```bat
python run.py
```

运行前建议：

- 本机已安装 Python 及依赖（`PyQt5`、`DrissionPage`、`pywinauto` 等）。
- Edge 可正常访问赛狐 ERP。

## 说明与依赖

- 下载文件保存位置使用浏览器当前默认下载目录。
- 如需重新打包可使用：

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "请款单下载" run.py
```
