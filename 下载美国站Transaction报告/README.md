# 下载美国站 Transaction 报告

易得客浏览器自动化：登录 TikTok Shop 卖家后台，固定美国站点并请求当月 Transaction 报告。

## 当前状态

- `YidekeLogin.py`：易得客登录与浏览器启动（`Specification`）
- `main.py`：多店铺启动、美国站点检查/切换、Reports Repository 表单填写与 Request Report
- `run.py`：Tkinter GUI 入口
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

在 `main.py` 底部 `config` 中配置易得客账号、店铺 IP 与调试端口：

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

- 站点固定为 **United States**；若当前非美国站，会先打开 See all 再切换
- 报告类型为 `SELLER_TRANSACTION_DATE_RANGE`，按月请求当前月份
- 当前流程到 **Request Report** 为止，报告生成后的下载步骤待补充
