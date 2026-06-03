# 赛狐ERP

赛狐 ERP 自动化项目，统一入口支持两种模式，均可自动创建 SKU 并完成在线配对。

## 代码结构

```
赛狐ERP/
├── run.py                  # Tkinter 统一 GUI 入口（推荐）
├── NewSet.py               # 模式一：新品工作计划表 → 建品配对
├── LowPrice.py             # 模式二：低价表 → 建 SKU 配对
├── main.py                 # 脚本式 new_set_pairing / low_price_pairing（与类版并行）
├── SaihuERPLogin.py        # SaiHuERPLogin 登录
├── login.py                # 另一套登录类（遗留）
├── test.py                 # SaihuERPLogin 副本/测试
├── QiYeVxLogin.py          # 企业微信登录辅助
├── YidekeLogin.py          # 易得客相关（遗留/辅助）
├── onlyrun_config.json     # GUI 上次账号与 Excel 路径
├── 赛狐ERP统一入口.spec
├── 赛狐ERP运行入口.spec
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `run.py` | `RunnerApp`（Tk）：模式切换、Excel 路径、账号；`RunnerThread` 调度 `NewSetPage.main()` 或 `LowPricePage.main()` |
| `NewSet.py` | `NewSetPage`：读「新品sku配对自动提醒」，`情况=未配对` 且含 `XP` → 赛狐「生成普通商品」+ 在线配对 |
| `LowPrice.py` | `LowPricePage`：读低价表最新日期行 → 商品列表建 SKU + 配对 |
| `main.py` | 函数式 `new_set_pairing` / `low_price_pairing`（内联 DrissionPage，与类版逻辑重复） |
| `SaihuERPLogin.py` | 赛狐登录：账号密码、验证码 OCR、进入业务页 |

### 配置说明（`onlyrun_config.json`）

| 字段 | 含义 |
| --- | --- |
| `last_username` / `last_password` | 上次赛狐账号（**勿提交真实密码**） |
| `last_excel_dew` | 模式一 Excel（工作计划表） |
| `last_excel_low` | 模式二 Excel（低价商城表） |

## 两种自动化模式

### 1) 低价模式（低价表创建 SKU 并配对）

- 数据来源：低价业务表（`LowPrice.py` 读取规则）。
- 流程目标：自动创建普通商品 SKU，并在在线产品中完成配对。
- 适用场景：低价表批量建品与配对。

### 2) 纯新品模式（工作计划表识别新品并创建 SKU）

- 数据来源：工作计划表「新品sku配对自动提醒」（`NewSet.py`：`情况=未配对` 且包含 `XP` 编号）。
- 流程目标：自动生成普通商品并完成在线产品配对。
- 适用前提：仅适用于**已走完提交流程**的新品需求单。

## 运行方式

### GUI 方式（推荐）

```bash
python run.py
```

### 脚本方式

```bash
python main.py
```

## 依赖说明

- Python（建议 3.12）
- 主要依赖：`PyQt5` / Tkinter、`DrissionPage`、`pandas`、`openpyxl`、`pywin32`、`ddddocr`、`onnxruntime`

安装示例：

```bash
pip install PyQt5 DrissionPage pandas openpyxl pywin32 ddddocr onnxruntime
```
