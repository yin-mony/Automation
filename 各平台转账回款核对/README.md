# 各平台转账回款核对

## 项目简介

本项目用于自动登录赛狐 ERP，分别下载 TikTok、eBay、Walmart、亚马逊四个平台的转账明细，并将下载文件统一整理、合并，再与收款平台流水记录进行到账核对。

项目将文件下载与表格处理拆分为两个独立操作，避免平台下载尚未全部完成时提前执行合并。

## 项目结构

```text
各平台转账回款核对/
├─ main.py             正式的平台自动化流程
├─ test.py             测试与定位调试文件
├─ Excel.py            Walmart 汇总、四平台合并及到账核对
├─ SaihuERPLogin.py    赛狐 ERP 登录与验证码识别
├─ run.py              PyQt5 图形界面入口
├─ README.md           开发说明
└─ 使用说明.md          用户操作说明
```

## 运行环境

- Windows
- Python 3.12 或 Python 3.14
- Chrome 或 Chromium 内核浏览器
- PyQt5
- DrissionPage
- pandas
- openpyxl
- ddddocr
- onnxruntime

`run.py` 会在 PyQt5 之前预加载 `onnxruntime`，用于避免 OCR 依赖与 PyQt5 的 DLL 加载顺序冲突。

## 模块职责

### main.py

`Payment` 是业务主流程类。

- `__init__()`：读取登录信息、输出目录、流水文件、日期、平台和执行状态。
- `multi()`：处理多平台模式下的 TikTok、eBay、Walmart 页面流程。
- `amazon()`：处理亚马逊平台的结算汇总页面流程。
- `download()`：点击导出与立即下载，检测新文件并按规则重命名。
- `excel()`：调用 `Excel.walmart()` 和 `Excel.merge()`。
- `main()`：切换多平台或亚马逊平台，并调用对应流程。
- `run()`：根据 `download_only` 判断执行下载还是表格处理。

### Excel.py

`Excel` 负责所有表格数据处理。

- `__init__()`：接收 config、开始日期和结束日期。
- `walmart()`：按店铺和付款周期结束时间汇总 Walmart 金额，输出独立汇总文件。
- `merge()`：读取四个平台文件，统一字段，生成基础合并明细，再与流水记录核对并生成独立对账结果。

### SaihuERPLogin.py

负责打开赛狐 ERP、判断登录状态、输入账号密码、识别图形验证码并关闭登录后的公告弹窗。业务流程复用登录模块创建的同一个浏览器页面。

### run.py

PyQt5 GUI 入口。下载与表格处理使用后台线程执行，避免界面在自动化运行期间无响应。

## 配置说明

```python
config = {
    "username": "赛狐登录账号",
    "password": "赛狐登录密码",
    "file_path": r"输出目录",
    "receipt_file": r"收款流水文件路径",
    "start_date": "2026-07-01",
    "end_date": "2026-07-31",
    "download_only": True,
    "platformName": "多平台",
    "mode": "TikTok",
}
```

| 字段 | 说明 |
| --- | --- |
| `username` | 赛狐 ERP 登录账号 |
| `password` | 赛狐 ERP 登录密码 |
| `file_path` | 下载文件和处理结果的输出目录 |
| `receipt_file` | 收款平台流水记录文件 |
| `start_date` | 开始日期，格式为 `YYYY-MM-DD` |
| `end_date` | 结束日期，格式为 `YYYY-MM-DD` |
| `download_only` | `True` 只下载文件；`False` 只处理表格 |
| `platformName` | `多平台` 或 `亚马逊` |
| `mode` | 多平台下使用：`TikTok`、`eBay` 或 `Walmart` |

## 执行流程

### 下载流程

当 `download_only=True` 时：

1. 创建浏览器页面并设置下载目录。
2. 调用 `SaiHuERPLogin` 登录赛狐 ERP。
3. 切换至 config 指定的平台。
4. 执行指定页面的筛选和导出。
5. 等待文件落盘并按项目命名规则重命名。

一次下载操作只执行当前选中的一个平台流程。四个平台需要分别运行下载。

### 表格处理流程

当 `download_only=False` 时：

1. 不创建浏览器，不执行登录。
2. 汇总最新的 Walmart 源文件。
3. 读取四个平台最新的转账明细。
4. 统一为店铺、追踪编号、转账时间、到账金额、银行尾号、平台、币种、到账状态。
5. 生成基础合并明细，初始到账状态为“暂无匹配”。
6. 使用收款流水进行一对一核对。
7. 将核对状态写入新的对账结果文件，不覆盖基础合并明细。

## 到账匹配规则

一条平台回款记录与一条收款流水同时满足以下条件时，判定为“已到账”：

- 币种一致。
- 金额保留两位小数后完全一致。
- 收款日期不早于平台转账日期。
- 收款日期在平台转账日期当日及之后 5 天内。
- 一条收款流水最多匹配一条平台回款记录。

未找到匹配流水的记录保留“暂无匹配”。

## 输出文件

输出目录清空后完整执行四个平台下载和表格处理，最终生成 7 个文件：

1. `TikTok转账明细-YYYY年MM月DD日至YYYY年MM月DD日.xlsx`
2. `eBay转账明细-YYYY年MM月DD日至YYYY年MM月DD日.xlsx`
3. `Walmart转账明细-YYYY年MM月.xlsx`
4. `亚马逊转账明细-YYYY年MM月DD日至YYYY年MM月DD日.xlsx`
5. `Walmart转账明细-YYYY年MM月-汇总完成.xlsx`
6. `各平台回款明细表-YYYY年MM月.xlsx`
7. `各平台回款明细表（对账结果）-YYYY年MM月.xlsx`

Walmart 使用付款周期筛选，周期可能跨月；文件名按配置开始月份命名。

## 开发约定

- 正式代码修改在 `main.py` 完成；页面定位验证可先在 `test.py` 进行。
- 不在代码中增加类型注解。
- 页面定位优先使用 DrissionPage 和 XPath。
- `SaihuERPLogin.py` 是共用登录模块，业务页面逻辑不应写入该文件。
- 下载文件与表格处理保持独立，不在单个平台下载完成后调用 Excel 处理。
- 不覆盖平台源文件、Walmart 源文件或基础合并明细。
- 新增平台时，应在 `Excel.merge()` 中明确字段映射、金额字段和币种规则。
- 修改 Excel 逻辑后，应核对总行数、平台分布、到账状态数量和输出格式。

## 常见问题

### onnxruntime DLL 加载失败

必须通过 `run.py` 启动 GUI，并确保 `onnxruntime` 在 PyQt5 前加载。不要随意调整 `run.py` 顶部的导入顺序。

### 文件下载完成后程序仍在等待

`download()` 根据输出目录中新建或更新时间发生变化的 `.xlsx` 文件判断下载完成。检查输出目录权限，并确认浏览器没有弹出另存为窗口。

### 表格处理提示缺少平台文件

确认四个平台源文件位于同一个输出目录，文件名符合项目命名规则，并且没有以 `~$` 开头的临时文件。

### PermissionError

目标文件正被 WPS 或 Excel 占用。关闭对应文件后重新执行表格处理。
