# 赛狐ERP

赛狐 ERP 自动化项目，统一入口支持两种模式，均可自动创建 SKU 并完成在线配对。

## 当前架构

```
赛狐ERP/
├── run.py                  # Tkinter 统一 GUI 入口：读取配置、选择模式、调度 main.SaihuERP
├── main.py                 # 脚本统一入口：SaihuERP(config) 按模式分发
├── NewSet.py               # 模式一：新品工作计划表 → 建品配对
├── LowPrice.py             # 模式二：低价表 → 建 SKU 配对
├── SaihuERPLogin.py        # 唯一正式赛狐登录类：验证码 OCR、协议勾选、公告关闭
├── login.py                # 遗留兼容入口：委托 SaihuERPLogin.py
├── test.py                 # 赛狐登录调试类
├── QiYeVxLogin.py          # 企业微信登录辅助
├── YidekeLogin.py          # 易得客登录辅助
├── onlyrun_config.json     # GUI 上次账号、密码与 Excel 路径
├── 赛狐ERP统一入口.spec
├── 赛狐ERP运行入口.spec
└── README.md
```

| 文件 | 主类 | 职责 |
| --- | --- | --- |
| `run.py` | `SaihuERPRun` | GUI 面板、配置读写、线程调度、日志输出 |
| `main.py` | `SaihuERP` | 统一脚本入口，根据 `mode` 调度模式一或模式二 |
| `NewSet.py` | `NewSetPage` | 读取「新品sku配对自动提醒」，筛选 `情况=未配对` 且含 `XP` 的行，生成普通商品并在线配对 |
| `LowPrice.py` | `LowPricePage` | 读取低价表最新日期行，在商品列表创建 SKU，并在线配对 |
| `SaihuERPLogin.py` | `SaiHuERPLogin` | 参考 FBA 子项目登录逻辑，处理赛狐登录、验证码和公告 |
| `login.py` | `SaihuERPLogin` | 兼容旧入口，不再维护独立登录逻辑 |

## 配置说明

当前正式流程均通过 `config` 传递配置参数。常用字段如下：

| 字段 | 含义 |
| --- | --- |
| `page` | DrissionPage 的 `ChromiumPage` 实例 |
| `mode` | `mode_one` / `mode1` 为纯新品；`mode_two` / `mode2` 为低价模式 |
| `username` / `password` | 赛狐账号密码 |
| `excel_path` | 当前模式使用的 Excel 文件路径 |
| `base_dir` | 临时验证码图片和运行基础目录 |

`onlyrun_config.json` 会被 GUI 读写：

| 字段 | 含义 |
| --- | --- |
| `last_username` / `last_password` | 上次赛狐账号密码 |
| `last_excel_dew` | 模式一 Excel（工作计划表） |
| `last_excel_low` | 模式二 Excel（低价商城表） |

## 两种自动化模式

### 1) 纯新品模式

- 数据来源：工作计划表「新品sku配对自动提醒」。
- 筛选规则：`情况=未配对`，且「赛狐新品开发编号」包含 `XP+数字`。
- 流程目标：赛狐「新品开发」生成普通商品，再进入「在线产品」完成 ASIN 配对。

### 2) 低价模式

- 数据来源：低价业务表「工作表1」。
- 筛选规则：自动识别「时间」列最新日期，并处理该日期所有行。
- 流程目标：赛狐「商品列表」创建普通商品 SKU，再进入「在线产品」完成 ASIN 配对。

## 运行方式

### GUI 方式（推荐）

```bash
python run.py
```

### 脚本方式

```bash
$env:SAIHU_USERNAME = "赛狐账号"
$env:SAIHU_PASSWORD = "赛狐密码"
python main.py --mode mode_one --path C:\Users\admin\Desktop\工作计划表.xlsx
python main.py --mode mode_two --path C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx
```

## 依赖说明

- Python（建议 3.12）
- 主要依赖：`DrissionPage`、`pandas`、`openpyxl`、`pywin32`、`ddddocr`、`onnxruntime`

安装示例：

```bash
pip install DrissionPage pandas openpyxl pywin32 ddddocr onnxruntime
```
