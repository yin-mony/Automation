# 产品主要流量词监控

易得客浏览器自动化：按 ASIN 在西柚找词中下载「关键词反查结果」，解析自然排名变化，并通过企业微信群机器人推送告警。

## 代码结构

```
产品主要流量词监控/
├── main.py                      # Comment 主流程：易得客、西柚下载与汇总推送
├── YidekeLogin.py               # 易得客登录封装（Specification）
├── run.py                       # Tkinter GUI 入口（推荐）
├── test.py                      # 解析已下载 Excel 并测试企业微信推送
├── 产品主要流量词监控.spec       # PyInstaller 打包配置
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | 启动易得客、按 IP/端口接管浏览器、西柚流量词下载、汇总后 `@` 指定手机号推送 |
| `YidekeLogin.py` | 易得客账号登录 |
| `run.py` | GUI 收集配置，后台线程执行 `Comment.run()` 并显示日志 |
| `test.py` | 离线读取 `file_path` 下 xlsx/csv，筛选自然排名异常并发送 webhook |

## 环境要求

- Windows（依赖易得客客户端、`pywinauto`、OpenCV 等）
- Python 3.10+
- 浏览器需安装西柚找词/西柚洞察插件
- 企业微信群机器人 Webhook

## 运行

### GUI（推荐）

```bash
python run.py
```

填写易得客账号、店铺 IP/端口、ASIN、保存目录与 Webhook 后点击「开始运行」。

### 命令行

修改 `main.py` 末尾 `config` 后：

```bash
python main.py
```

## 打包

```bash
pyinstaller --clean 产品主要流量词监控.spec
```

产物在 `dist/产品主要流量词监控.exe`（`dist/` 已被 `.gitignore` 忽略）。

## 注意事项

- 运行前会在本地启动/清理 `edecker` 进程，请勿与其他自动化任务冲突
- **请勿将真实密码、Webhook Key 提交到 Git**
