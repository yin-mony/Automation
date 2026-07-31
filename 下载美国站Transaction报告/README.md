# 下载 Transaction 报告

易得客浏览器自动化：登录 Amazon Seller Central，按配置依次切换一个或多个后台站点，并请求各站点上一个自然月的 Transaction 报告。默认站点为美国。

## 当前状态

- `YidekeLogin.py`：易得客登录、浏览器启动及 Amazon 登录页/账户选择页/二步验证检查
- `main.py`：多店铺启动、Amazon 中文界面与后台站点检查、Reports Repository 表单填写与请求报告
- `run.py`：Tkinter GUI（店铺站点、Amazon 登录信息、后台站点、运行环境和邮件通知）
- `email_util.py`：本项目专用邮件发送（`sendEmail=True` 时发送桌面报告附件）
- `build_exe.bat`：安装依赖并打包 exe
- `下载美国站Transaction报告.spec`：PyInstaller 配置
- `test.py`：调试脚本

## 依赖

- Python 3
- DrissionPage、psutil、pywinauto、python-dateutil

## 运行

### GUI

```bash
python run.py
```

### 命令行

默认初始值统一定义在 `TestPage.__init__()` 中；调用方传入 `config` 时，只覆盖其中提供的配置项。

| 字段 | 说明 |
| --- | --- |
| `username` / `password` | 易得客账号；也可使用环境变量 `YIDEKE_USERNAME` / `YIDEKE_PASSWORD` |
| `autoSiteName` | 易得客店铺站点，访问店铺时使用，默认 `美国` |
| `ip` / `port` | 店铺 IP 与调试端口 |
| `amazonSiteNames` | Amazon 后台站点列表，按列表顺序逐站请求报告，默认 `["美国"]` |
| `amazonSiteName` | 兼容旧配置的单个 Amazon 后台站点 |
| `amazonEmail` / `amazonPassword` | Amazon 登录信息；密码为空时优先等待浏览器保存的密码 |
| `isOnline` | 仅用于 GUI 模式选择和日志标识；线上、线下模式都会先登录易得客并访问店铺 |
| `sendEmail` | `True` 完成后发送邮件，`False` 不发送 |
| `email` | 选择发送邮件时必填 |
| `sender_email` / `smtp_auth_code` | SMTP 发件邮箱与授权码；也可使用环境变量 `SMTP_SENDER` / `SMTP_AUTH_CODE` |
| `file_path` | 报告目录，默认桌面 |

易得客账号密码、发件邮箱与 SMTP 授权码均在 `TestPage.__init__()` 中完成初始化。为避免敏感信息进入版本库，建议使用环境变量：

```powershell
$env:YIDEKE_USERNAME = "易得客账号"
$env:YIDEKE_PASSWORD = "易得客密码"
$env:SMTP_SENDER = "发件邮箱"
$env:SMTP_AUTH_CODE = "SMTP 授权码"
python run.py
```

报告文件默认在**桌面**，文件名需包含 `下载美国站Transaction报告`。
```bash
python main.py
```

## 打包

```bash
build_exe.bat
```

或：

```bash
pyinstaller --clean 下载美国站Transaction报告.spec
```

## 说明

- 接管 Seller Central 后先检查登录页、密码页、账户选择页和二步验证码，再进入报告流程
- Amazon 后台已是中文简体时不会重复切换；检测为英文时才切换到中文简体
- 站点、账户选择、付款菜单、报告库、报告筛选、月份和请求报告按钮均兼容中文与英文界面文案
- 美国站校验并处理商城、账户类型、报告类型 3 个筛选；其他站点校验并处理账户类型、报告类型 2 个筛选
- 每个筛选会先记录当前值和可点击状态：目标值已正确时跳过；不可点击且值不正确时停止；账户类型没有 `全部 / All` 选项时保留站点当前有效值
- 商城存在时选择 `所有商城 / All Stores`，报告类型选择 `交易 / Transaction`，日期范围选择 `月 / Month`
- GUI 的 Amazon 后台站点使用复选框多选；流程按照界面站点顺序逐站处理，单选时保持原有流程
- 后台站点默认 **美国 / United States**；若当前非目标站点，会先打开 `查看所有/See all`，再选择目标站点并点击 `选择账户/Select account`
- 报告类型为 `SELLER_TRANSACTION_DATE_RANGE`，按系统时间自动请求上一个自然月，并同步选择对应年份
- 每个所选站点均以点击 **请求报告 / Request Report** 为完成，不持续等待报告生成或下载
