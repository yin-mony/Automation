# TREC 公司与个人合作推广

本项目从 Texas Open Data 同步 TREC 许可证数据，按公司或个人模式使用 SerpApi 搜索公开网页，再通过 DrissionPage 提取普通二级页面的公开联系方式。Facebook 只记录主页链接，不登录、不访问页面。所有联系方式先进入人工审核，通过后才能生成推广记录。

## 业务链路

1. 直连同步 Texas Open Data，统一写入 `output/data.db`。
2. 公司模式按挂靠公司去重汇总；个人模式筛选 Active、无挂靠且在到期窗口内的数据。
3. 优先读取搜索缓存，未命中时才调用一次 SerpApi Google 第一页搜索。
4. 普通二级页面默认由 DrissionPage 直接访问；Facebook 主页只保存链接。
5. 公开邮箱和电话写入 SQLite，并导出原中文名称的 Excel 文件。
6. 人工审核联系方式；支持单选、多选和全选当前筛选结果，只有“已通过”记录可以创建邮箱草稿或真实发送。
7. GUI 内置“操作说明”页，按页面列出用途、操作步骤和注意事项。
8. “运行设置”可修改并保存每日新搜索上限，工作台会同步显示当天已用次数和当前上限。

## 文件结构

| 文件 | 主类 | 职责 |
| --- | --- | --- |
| `run.py` | `RunGui` | PySide6 桌面入口和后台线程 |
| `main.py` | `Main` | 完整业务编排与暂停、停止控制 |
| `data.py` | `Data` | TREC、SQLite、审核和 Excel |
| `serp.py` | `Serp` | SerpApi 搜索、缓存和免费额度 |
| `proxy.py` | `Proxy` | SOCKS5 到本地 HTTP 的代理桥 |
| `browser.py` | `Browser` | DrissionPage 普通网页提取和 Facebook 链接识别 |
| `mail.py` | `Mail` | 阿里邮箱服务器草稿、真实发送和结果通知 |

每个 Python 文件只保留一个顶层主类，非 GUI 模块不嵌套辅助类。方法使用 `camelCase`，不使用类型注解，关键流程使用中文 docstring 和中文注释。

## 安装与启动

```powershell
py -3.14 -m pip install -r requirements.txt
py -3.14 run.py
```

SerpApi Key、阿里邮箱账号和授权码集中写在 `Main.__init__` 的默认参数中，`config.local.json` 仍可覆盖这些默认值。凭据已经进入源码，因此提交或分享项目之前必须先清除或更换；`config.local.json` 和 `ipfiy.py` 继续保持 Git 忽略。

## 免费额度

- 月额度：250 次。
- 日常搜索预算：180 次。
- 人工复查预留：20 次。
- 硬预留：20 次。
- 额外安全余量：30 次。
- 每日新搜索上限：6 次；公司与个人合并运行时各 3 次。
- 缓存命中、普通网页提取和 Facebook 链接记录不额外消耗 SerpApi 搜索额度。
- GUI 启动后自动调用不计费的 SerpApi Account API，以官方 `used` 和 `remaining` 为准。

SerpApi 请求不传固定 `num`，使用 Google 第一页默认结果。已收到但无法解析的响应只记账一次且不自动重试。

## 网络路径

| 服务 | 网络方式 |
| --- | --- |
| Texas Open Data | 直连 |
| SerpApi 搜索与账户额度 | 直连 |
| 阿里邮箱草稿 IMAP 和真实发送 SMTP | SSL 直连 |
| 普通二级页面 | DrissionPage 直接访问 |
| Facebook 主页链接 | 不访问页面，只保存已发现链接 |

默认配置 `proxyRequired=false`，因此完整流程不会读取、测试或启动 `ipfiy.py` 代理，代理节点异常也不会阻断任务。代理桥代码继续保留；以后手动改为 `true` 时，才会恢复 SOCKS5 多目标检测和失败关闭策略。

## 数据与输出

- `output/data.db`：唯一内部状态库。
- `file/初始总量数据未清洗.xlsx`：兼容旧流程的官方原始底表。
- `file/已获取到的初始总数据.xlsx`：兼容旧流程的清洗底表。
- `output/已完成搜索匹配的公司联系信息数据.xlsx`：公司结果。
- `output/已完成搜索匹配的个人联系信息数据.xlsx`：个人结果。
- `output/邮件发送记录.xlsx`：服务器草稿、真实发送、重复跳过和错误审计记录。

首次运行若只存在旧 `output/trec_automation.db`，程序会复制为 `data.db`，不会删除旧文件。

## Facebook 范围

- 只记录搜索结果或普通公开网页发现的 Facebook 主页链接。
- 程序不会打开 Facebook、不会登录账号，也不会监听 Facebook 网络请求。
- Facebook 链接保存在 SQLite、GUI“结果数据”页和公司/个人 Excel 结果中。

## 邮件模式

- 默认“生成邮箱草稿”：通过 IMAP SSL 把邮件写入阿里邮箱服务器草稿箱。
- “真实发送”：通过 SMTP SSL 立即发送，GUI 执行前必须再次确认。
- 草稿和真实发送都使用 HTML 正文，并以内嵌 CID 图片显示 `file/time2renew-logo.png`。
- 本地 Excel 只作为动作审计记录，不代替服务器草稿。

## 重复控制

- TREC 数据按许可证号写入，公司按挂靠许可证汇总，个人按详情 ID 或许可证号识别。
- 同一次队列按 `objectKey` 去重，已完成对象由 SQLite 主键跳过。
- 搜索按对象和搜索词生成 SHA-256 缓存键，缓存命中不再次调用 SerpApi。
- 邮箱统一转为小写，同批只保留一次；成功草稿和成功发送写入 `mail_actions`，后续运行自动跳过。

## 打包

已使用 `trec.spec` 生成稳定目录版，入口为 `dist/TREC推广工具/TREC推广工具.exe`。

- 必须保留同目录的 `_internal`，不能只把 EXE 单独移走。
- `output/data.db` 保存现有数据、审核、额度和去重记录。
- `ipfiy.py` 与 `config.local.json` 含本机代理或账号配置，发布目录仅限可信环境使用。
