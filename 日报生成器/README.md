# 日报生成器

Tkinter GUI 填写当日工作项，生成工作日报文本并通过企业微信群机器人推送。

## 代码结构

| 文件 | 职责 |
| --- | --- |
| `main.py` | `Comment`：组装日报内容、发送企业微信 webhook |
| `run.py` | GUI 入口：录入事项、预览并触发发送 |
| `日报生成器.spec` | PyInstaller 打包配置 |

## 依赖

- Python 3.10+
- `requests`

## 运行

```bash
python run.py
```

在界面填写企业微信 Webhook、@ 手机号及当日工作事项后发送。

## 打包

```bash
pyinstaller --clean 日报生成器.spec
```

产物在 `dist/`（已被 `.gitignore` 忽略）。

## 说明

- 生成的 `工作日报-*.txt` 为本地产物，勿提交 Git
- Webhook 与手机号请在界面填写，避免将真实密钥提交到仓库
