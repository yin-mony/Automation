# 产品主要流量词监控

易得客浏览器自动化：按 ASIN 在卖家精灵等工具中下载「关键词反查结果」，解析自然排名变化，并通过企业微信群机器人推送告警。

## 代码结构

```
产品主要流量词监控/
├── main.py           # Comment 主流程：登录易得客、多店铺浏览器、下载与汇总
├── YidekeLogin.py    # 易得客登录封装（Specification）
├── test.py           # 解析已下载 Excel 并测试企业微信推送
├── package-lock.json # Node 依赖锁定（若使用 @wecom/aibot-node-sdk 等）
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | 启动易得客、按 IP/端口接管浏览器、下载关键词报表、汇总后 `@` 指定手机号推送 |
| `YidekeLogin.py` | 易得客账号登录 |
| `test.py` | 离线读取 `file_path` 下 xlsx/csv，筛选自然排名异常并发送 webhook |

## 环境要求

- Windows（依赖易得客客户端、`pywinauto`、OpenCV 等）
- Python 3.10+
- 企业微信群机器人 Webhook（在 `main.py` / `test.py` 中配置）

## 运行

修改 `main.py` 末尾 `config`：

| 字段 | 说明 |
| --- | --- |
| `username` / `password` | 易得客账号 |
| `ip` / `port` | 店铺浏览器 IP 与端口列表（一一对应） |
| `asin` | 待监控 ASIN 列表 |
| `file_path` | 关键词报表下载目录 |
| `number` | 企业微信 @ 的手机号 |

```bash
python main.py
```

仅测试 Excel 解析与推送：

```bash
python test.py
```

## 注意事项

- 运行前会在本地启动/清理 `edecker` 与 Chrome 进程，请勿与其他自动化任务冲突。
- `node_modules/`、浏览器临时配置目录（UUID 命名文件夹）已在仓库 `.gitignore` 中排除。
- **请勿将真实密码、Webhook Key 提交到 Git**；本地修改 `config` 后自行保管。
