# BOSS 直聘 · 筛选简历

BOSS 招聘端扫码登录后，按岗位浏览推荐牛人并主动联系近期活跃候选人；候选人回复后由人工回复和确认，收到简历后继续自动审核并进入面试流程。支持话术模板、岗位规则、不合适名单与企业微信日报推送。代码规范见 [`docs/NAMING.md`](./docs/NAMING.md)。

## 当前流程

1. 登录逻辑保持招聘端 APP 扫码方式不变。
2. 推荐牛人严格按刚刚活跃/在线、今日活跃、三天内活跃的顺序扫描，超过三天或活跃时间未知时跳过。
3. 推荐牛人首次主动联系按自然日独立计数，默认最多 15 人，可在 GUI 修改；普通聊天回复、求简历和面试消息不占该额度。
4. 候选人未回复时不追问、不索要简历；有新回复时暂停自动流程，等待人工回复并选择继续索要简历或标记不合适。
5. 候选人待回复时可调用本机 Qwen 生成结构化建议，HR 可修改后一键填入 BOSS 聊天框；模型不点击发送，也不代替人工决定是否索要简历。
6. 回复 Skill 可在 GUI 新建、编辑和切换，模型建议与人工最终填入内容记录在 SQLite，供后续持续优化。
7. 程序启动前已经由我方读过但没有回复的会话标记为不合适，后续自动流程永久跳过，可在 GUI 的不合适名单中人工解除。
8. 收到简历后的解析、岗位规则审核、面试预邀请与正式面试流程保持不变。

## 代码结构

| 文件 / 目录 | 职责 |
| --- | --- |
| `main.py` | GUI 入口，启动 `RunGui` |
| `run.py` | Tkinter 主界面：登录、推荐额度、人工决策、话术/岗位管理、企微日报 |
| `boss_web/auto.py` | `BossAuto`：DrissionPage 浏览器自动化核心 |
| `boss_web/login.py` | `BossLogin`：招聘端 APP 扫码登录 |
| `boss_web/db.py` | `BossDb`：SQLite 持久化（候选人、话术、岗位规则等） |
| `boss_web/job.py` | `BossJob`：岗位与全局筛选默认配置 |
| `boss_web/template.py` | `BossTemplate`：话术默认配置 |
| `boss_web/parse.py` | `ResumeParse`：简历解析 |
| `boss_web/match.py` | `ResumeMatch`：简历匹配审核 |
| `boss_web/report.py` | `BossReport`：企业微信日报推送 |
| `boss_web/reply.py` | `BossReply`：本地 OpenAI 兼容模型调用与结构化回复建议 |
| `boss.spec` | PyInstaller 打包配置 |
| `docs/NAMING.md` | 项目编码与命名规范 |
| `.cursor/rules/boss-coding-standards.mdc` | Cursor 规则摘要 |

## 数据目录

本项目的 Chrome 用户数据、SQLite 数据库等统一存放在：

`D:\boss_zhaopin_筛选简历`

- Chrome 配置：`D:\boss_zhaopin_筛选简历\boss_chrome_profile`
- 数据库：`D:\boss_zhaopin_筛选简历\db\boss_automation.db`

与参考项目 `F:\boss直聘` 数据隔离，互不影响。

## 依赖

- Python 3.10+
- `DrissionPage`、`requests`（见 `requirements.txt`）
- 标准库：`tkinter`（Windows 自带）

## 运行

```bash
cd BOSS直聘-筛选简历
pip install -r requirements.txt
python main.py
```

## 本地回复模型

回复助手通过 LM Studio 调用本机 Qwen，默认 OpenAI 兼容地址：

`http://127.0.0.1:1234/v1`

推荐配置：

- LM Studio 0.4 或更高版本
- `Qwen/Qwen3-8B-GGUF`
- `Q4_K_M` 量化版本
- API Model Identifier：`qwen3-8b`
- 上下文长度：`8192`
- Developer → Local Server 端口：`1234`
- 关闭局域网服务、CORS 与 per-request MCPs

启动 LM Studio、加载模型并开启 Local Server 后，进入程序的“本地回复模型与 Skill”，点击“测试本地服务”。候选人回复触发人工等待时，可生成建议并填入聊天框，最后仍由人工在 BOSS 页面检查并点击发送。

程序拒绝把候选人消息发送到非本机模型地址。模型不可用时不影响原人工回复流程。

各模块可单独调试，例如：

```bash
python boss_web/login.py
```

## 打包为单文件 exe

```bash
build_exe.bat
```

或手动：

```bash
pip install -r requirements.txt
pip install -r requirements-build.txt
pyinstaller --noconfirm --clean boss.spec
```

产物：`dist/BOSS直聘筛选简历.exe`（单文件，无控制台窗口）。目标机器需已安装 Chrome/Edge；数据仍读写 `D:\boss_zhaopin_筛选简历`。

## 实现状态

- 已从参考项目 `F:\boss直聘` 迁移完整实现
- 数据目录已改为 `D:\boss_zhaopin_筛选简历`
