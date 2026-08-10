# 日报生成器



Tkinter GUI 填写**今日工作事项**与**明日待办**（均为标题 + 负责人 + 详情），通过企业微信群机器人以 **Markdown 分两条消息**推送。



## 代码结构



| 文件 | 职责 |

| --- | --- |

| `main.py` | `Comment`：Markdown 排版、今日/明日分条发送、@ 提醒 |

| `run.py` | GUI 入口：录入事项、预览、配置缓存 |

| `run_config.json` | 本地缓存 webhook、手机号及已添加事项（勿提交 Git） |

| `日报生成器.spec` | PyInstaller 打包配置 |



## 依赖



- Python 3.10+

- `requests`



## 运行



```bash

python run.py

```



界面默认填充企业微信手机号与 webhook（与 `main.py` 默认值一致），关闭窗口时自动保存到 `run_config.json`。



推送规则：



1. 有今日事项 → 发送一条「工作日报 {日期}」Markdown

2. 有明日待办 → 再发送一条「明日待办 {次日日期}」Markdown

3. 最后发送一条文本消息 @ 配置的手机号



## 打包



```bash

pyinstaller --clean 日报生成器.spec

```



产物在 `dist/`（已被 `.gitignore` 忽略）。



## 说明

- 生成的 `工作日报-*.txt` 为本地产物，勿提交 Git

- `run_config.json` 含 webhook，已加入 `.gitignore`
