# 下载全站点汇总报告

易得客浏览器自动化：登录 TikTok Shop 卖家后台，按店铺下载全站点汇总报告。

## 当前状态

- `YidekeLogin.py`：易得客登录与浏览器启动（`Specification`）
- `main.py`：易得客登录、多店铺浏览器启动与报告下载主流程
- `run.py`：Tkinter GUI 入口
- `下载全站点汇总报告.spec`：PyInstaller 打包配置
- `test.py`：调试脚本

## 依赖

- Python 3
- DrissionPage、psutil、pywinauto

## 运行

### GUI

```bash
python run.py
```

### 命令行

在 `main.py` 底部 `config` 中配置账号与店铺信息后：

```bash
python main.py
```

## 打包

```bash
pyinstaller --clean 下载全站点汇总报告.spec
```

产物在 `dist/`（已被 `.gitignore` 忽略）。
