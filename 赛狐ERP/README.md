# 赛狐ERP

赛狐 ERP 自动化项目，统一入口支持两种模式，均可自动创建 SKU 并完成在线配对。

## 两种自动化模式

### 1) 低价模式（低价表创建 SKU 并配对）

- 数据来源：低价业务表（按 `LowMain.py` 的读取规则）。
- 流程目标：自动创建普通商品 SKU，并在在线产品中完成配对。
- 适用场景：低价表批量建品与配对。

### 2) 纯新品模式（工作计划表识别新品并创建 SKU）

- 数据来源：工作计划表中“新品sku配对自动提醒”工作表（按 `DewMain.py` 的筛选规则，`情况=未配对` 且包含 `XP` 编号）。
- 流程目标：自动生成普通商品并完成在线产品配对。
- 适用前提：仅适用于**已走完提交流程**的新品需求单。

## 核心文件说明

- `OnlyRun.py`
  - Qt 图形化统一入口。
  - 支持账号/密码输入、模式切换、按模式选择 Excel、实时日志输出。
- `OnlyMain.py`
  - 全局调度入口 `SaiHuMain`。
  - 负责浏览器会话管理、登录态复用、统一登录、按模式调用流程。
- `SaihuERPLogin.py`
  - 赛狐登录流程（账号密码、验证码自动优先+手动兜底、稳定进入业务页）。
- `LowMain.py`
  - 低价模式的主流程自动化逻辑（创建 SKU + 在线配对）。
- `DewMain.py`
  - 纯新品模式的主流程自动化逻辑（根据工作计划表识别新品并创建/配对）。

## 运行方式

### GUI 方式（推荐）

```bash
python "OnlyRun.py"
```

### 脚本方式

```bash
python "OnlyMain.py"
```

## 依赖说明

- Python（建议 3.12）
- 主要依赖：`PyQt5`、`DrissionPage`、`pandas`、`openpyxl`、`pywin32`、`ddddocr`、`onnxruntime`

安装示例：

```bash
pip install PyQt5 DrissionPage pandas openpyxl pywin32 ddddocr onnxruntime
```
