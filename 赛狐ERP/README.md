# 赛狐ERP

赛狐 ERP 相关自动化脚本项目，包含浏览器拉起、赛狐网页登录、在线产品配对 SKU 以及 GUI/EXE 运行入口。

## 用途

- 拉起本地 Edge 浏览器并连接自动化会话。
- 自动检测赛狐登录状态，必要时执行账号密码与验证码登录。
- 从工作计划表读取 `XP` 需求，按条数循环执行在线产品配对流程。
- 提供 Qt 图形界面入口，支持选择表格路径并执行整批任务。

## 当前实现状态

- 已实现 `在线产品配对SKU` 主流程（读取 Excel -> 登录赛狐 -> 商品创建/在线产品配对）。
- 已实现 `run.py` Qt 界面（账号、密码、工作表、文件路径、状态栏）。
- 已支持 PyInstaller 打包 `在线产品配对SKU.exe`。

## 运行与依赖

### 环境依赖

- Python 3.x
- 主要依赖：`PyQt5`、`pandas`、`openpyxl`、`DrissionPage`、`psutil`、`ddddocr`

安装示例：

```bash
pip install PyQt5 pandas openpyxl DrissionPage psutil ddddocr
```

### 运行方式

- GUI 方式（推荐）：

```bash
python "在线产品配对SKU/run.py"
```

- 脚本方式：

```bash
python "在线产品配对SKU/main.py"
```

### 打包方式

在 `赛狐ERP/在线产品配对SKU` 目录下执行：

```bash
pyinstaller --noconfirm --clean "在线产品配对SKU.spec"
```
