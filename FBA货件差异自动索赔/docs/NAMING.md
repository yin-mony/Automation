# FBA 货件差异自动索赔 — 代码结构与命名规范

> 本文档为项目**强制规范**。所有新建、修改代码必须遵守。  
> 核心原则：**一文件一类、可独立调试、常量进 init、配置进 main、命名简短、逐步中文注释**。

**项目简介**：登录赛狐 ERP（Sellfox），在 FBA 货件相关页面识别收发差异并执行索赔相关自动化；提供 GUI 配置账号与导出路径，核心流程可 `python xxx.py` 单文件调试。

---

## 1. 文件结构规范（强制）

### 1.1 一文件一类

- 每个 `.py` 文件**只允许有一个主类**，文件名与类名对应。
- **禁止**在一个文件里写多个业务类。
- **禁止**单独建 `constants.py`、`config.py`、`defaults.py` 这类纯常量文件。

| 文件名 | 类名 | 说明 |
|--------|------|------|
| `login.py` | `SaihuLogin` | 赛狐 ERP 登录（验证码 OCR、公告关闭） |
| `claim.py` | `FbaClaim` | FBA 货件差异查询与索赔页面自动化 |
| `run.py` | `RunGui` | GUI 入口（账号、路径、日志、一键运行） |

### 1.2 标准文件模板

每个 Python 文件必须按以下结构组织：

```python
class XxxService:
    """类的职责说明（中文）"""

    def __init__(self):
        # 所有初始值、常量、默认配置，全部写在这里
        self.host = "0.0.0.0"
        self.httpPort = 8765
        self.wsPort = 8766
        self.defaultWords = ["话术1", "话术2"]
        self.filterDefaults = {"活跃度": "不限", "性别": "不限"}

    def run(self):
        pass


if __name__ == "__main__":
    # 本文件独立调试时的配置，只写在这里
    config = {
        "host": "127.0.0.1",
        "httpPort": 8765,
        "wsPort": 8766,
        "browserId": "test-browser-id",
    }

    service = XxxService()
    service.run()
```

### 1.3 `if __name__ == "__main__"` 要求

- **每个文件必须有**，用于单独运行、单独调试。
- 该文件专属的调试配置（host、port、browserId、测试数据等）**只能写在 main 块里**。
- main 块里允许实例化本类并调用方法，方便 `python xxx.py` 直接测试。

### 1.4 常量与初始值存放规则

| ✅ 允许 | ❌ 禁止 |
|---------|---------|
| 写在类的 `__init__` 里，用 `self.xxx` | 模块级 `DEFAULT_XXX = ...` |
| 实例属性：`self.defaultWords` | 单独建 `Defaults` 类 |
| 类内固定值：`self.maxRetry = 3` | 单独建 `Config` 类 |
| 调试配置写在 `if __main__` 的 `config` 字典 | 用 `.env` / 独立 config 文件（除非后续统一约定） |

**示例：**

```python
class FbaClaim:
    def __init__(self):
        # 赛狐入口与 FBA 货件列表页
        self.portalUrl = "https://www.sellfox.com/amzup-web-main/web/dashboard.html"
        self.shipmentListUrl = ""
        # 默认筛选条件（与赛狐页面字段一致时用中文键）
        self.filterDefaults = {
            "店铺": "全部",
            "时间范围": "最近30天",
        }
        # 导出目录
        self.exportDir = ""
        # 运行状态
        self.stopFlag = False
        self.page = None
```

### 1.5 文件命名（强制）

- **优先单词文件名**（无下划线）：如 `auto.py`、`db.py`、`service.py`。
- 仅在单词无法表达职责时，才用 `snake_case` 两词组合（最多 2 词）。
- 文件名必须准确反映模块职责，并与主类名对应。
- 禁止泛化命名：`utils.py`、`helper.py`、`common.py`、`manager.py`、`handler.py`。
- 禁止过长堆砌：`automation_db.py`、`task_service.py`、`browser_automation_service_handler.py`。

| 模块职责 | 推荐文件名 | 主类名 | 避免 |
|----------|-----------|--------|------|
| 赛狐登录 | `login.py` | `SaihuLogin` | `saihu_login.py`、`SaihuERPLogin.py`（过长） |
| FBA 货件索赔自动化 | `claim.py` | `FbaClaim` | `fba_claim_auto.py`、`automation_utils.py` |
| GUI 入口 | `run.py` | `RunGui` | `run_gui.py`、`main_window.py` |

新建文件前，先确认文件名是否简短、是否准确表达模块作用；不符合则换名再写代码。

---

## 2. 命名规范（强制）

### 2.1 总规则

1. **禁止**任何函数、方法名以 `_` 开头（包括 `_run`、`_emitLog` 等）。
2. 名称必须**简短、直观**，一眼能看懂用途。
3. 禁止堆砌过长命名，禁止无意义缩写。

### 2.2 命名长度建议

| 类型 | 建议长度 | 说明 |
|------|----------|------|
| 方法名 | 2～4 个英文单词 | 如 `getState`、`saveTask` |
| 变量名 | 1～3 个英文单词 | 如 `browserId`、`taskList` |
| 类名 | 1～3 个英文单词 | 如 `BossAuto`、`WebServer` |

### 2.3 类名 — `PascalCase`（大驼峰）

```python
class BossAuto: ...
class BossDb: ...
class WebServer: ...
class TaskService: ...
```

### 2.4 方法名 — `camelCase`（小驼峰）

```python
def initDb(self): ...
def getState(self): ...
def saveTask(self): ...
def runChat(self): ...
def stopWorker(self): ...
def isRunning(self): ...
```

**禁止写法：**

```python
def _run(self): ...                          # ❌ 禁止下划线开头
def _emitLog(self): ...                      # ❌ 禁止下划线开头
def submit_chat_task(self): ...              # ❌ 过长，改用 submitChat
def recommend_job_filters_snapshot(self): ... # ❌ 过长，改用 getJobFilters
```

**推荐写法：**

```python
def run(self): ...
def emitLog(self, msg): ...
def submitChat(self, payload): ...
def getJobFilters(self): ...
def createRule(self, data): ...
def updateRule(self, data): ...
def deleteRule(self, data): ...
```

### 2.5 变量与参数 — `camelCase`

```python
browserId = "xxx"
taskList = []
stopFlag = False
eventCallback = None
paramsJson = "{}"
```

### 2.6 文件名 — `snake_case.py`

文件名仍用下划线（Python 模块惯例），与类名分离：

| 文件名 | 类名 |
|--------|------|
| `auto.py` | `BossAuto` |
| `db.py` | `BossDb` |
| `server.py` | `WebServer` |
| `service.py` | `TaskService` |

### 2.7 方法命名参考表

| 动作 | 命名模式 | 示例 |
|------|----------|------|
| 获取 | `get` + 名词 | `getState`、`getRules`、`getWorkers` |
| 保存 | `save` + 名词 | `saveTask`、`saveSettings` |
| 创建 | `create` + 名词 | `createRule`、`createBrowser` |
| 更新 | `update` + 名词 | `updateRule`、`updateBrowser` |
| 删除 | `delete` + 名词 | `deleteRule`、`deleteBrowser` |
| 提交 | `submit` + 名词 | `submitChat`、`submitRecommend` |
| 启动 | `start` + 名词 / `run` + 名词 | `startServer`、`runChat` |
| 停止 | `stop` + 名词 | `stopWorker`、`stopTask` |
| 暂停 | `pause` + 名词 | `pauseWorker` |
| 恢复 | `resume` + 名词 | `resumeWorker` |
| 判断 | `is` + 形容词 | `isRunning`、`isPaused` |
| 连接 | `connect` + 名词 | `connectSocket`、`connectDb` |
| 发送 | `send` + 名词 | `sendMsg`、`sendRequest` |
| 处理 | `handle` + 名词 | `handleEvent`、`handleRequest` |

### 2.8 类型注解（强制）

- **禁止**在函数、方法上使用参数类型注解、返回值注解（`->`）。
- **禁止**在类上使用泛型、类属性类型注解（含 `__init__` 里的 `self.xxx: Type = ...`）。
- **禁止**为类型注解而 `from typing import ...`（如 `Any`、`Optional`、`Callable`）。
- 类型与约束用**中文 docstring + 关键步骤注释**说明，与第 3 节注释规范一致。

**禁止写法：**

```python
from typing import Any

def getCandidate(self, candidateKey: str) -> dict[str, Any] | None:
    ...

class BossTemplate:
    def __init__(self):
        self.words: dict[str, list[str]] = {}
```

**推荐写法：**

```python
def getCandidate(self, candidateKey):
    """按 candidateKey 查询候选人，未找到返回 None"""
    ...

class BossTemplate:
    def __init__(self):
        # 七类话术，key 与 message_templates.template_type 一致
        self.words = {}
```

---

## 3. 注释规范（强制）

### 3.1 总规则

- **每个函数/方法**：开头写一段中文说明「做什么」。
- **函数内每个关键操作步骤**：操作完成后紧跟一行中文注释，说明「这一步做了什么、结果是什么」。

### 3.2 函数级注释

```python
def submitChat(self, payload):
    """提交沟通任务到队列"""
    browserId = payload.get("browserId", "")
    # 校验浏览器 ID 是否为空
    if not browserId:
        raise ValueError("浏览器 ID 不能为空")

    # 组装任务参数
    task = {
        "type": "chat",
        "browserId": browserId,
        "times": payload.get("times", []),
    }

    # 写入任务队列
    self.taskQueue.put(task)
    # 通知前端刷新工人状态
    self.emitEvent({"type": "workerUpdate"})

    return {"ok": True, "taskId": task["id"]}
```

### 3.3 注释密度要求

| 场景 | 是否必须注释 |
|------|-------------|
| 条件判断后 | ✅ 说明判断目的 |
| 循环体内关键步骤 | ✅ 说明循环在做什么 |
| 数据库读写后 | ✅ 说明读写了什么 |
| API / WebSocket 调用后 | ✅ 说明发送/接收内容 |
| 异常捕获后 | ✅ 说明捕获原因和处理方式 |
| 简单赋值（一眼明了） | 可选 |

### 3.4 禁止的注释写法

```python
# 提交任务                          ❌ 太笼统，没说明提交什么
self.queue.put(task)               ❌ 关键操作无注释

# i += 1                           ❌ 无意义注释
i += 1
```

### 3.5 推荐的注释写法

```python
# 从 payload 取出浏览器 ID
browserId = payload.get("browserId", "")

# 浏览器 ID 为空则拒绝提交
if not browserId:
    raise ValueError("浏览器 ID 不能为空")

# 构造沟通任务并放入队列
self.taskQueue.put(task)

# 任务入队成功，推送状态更新事件
self.emitEvent({"type": "workerUpdate"})
```

---

## 4. 完整文件示例

```python
# service.py

import threading
from queue import Queue


class TaskService:
    """自动化任务调度服务"""

    def __init__(self):
        # 默认浏览器 ID
        self.defaultBrowserId = ""
        # 任务队列
        self.taskQueue = Queue()
        # 工人列表
        self.workers = []
        # 默认话术
        self.defaultWords = [
            "方便加一下微信吗？",
            "咱们的时间自由安排，期待回复哦",
        ]
        # 默认筛选条件
        self.filterDefaults = {
            "活跃度": "不限",
            "性别": "不限",
            "学历要求": "不限",
        }
        # 线程锁
        self.lock = threading.RLock()

    def submitChat(self, payload):
        """提交沟通任务"""
        # 取出浏览器 ID
        browserId = str(payload.get("browserId", "")).strip()
        # 校验浏览器 ID
        if not browserId:
            raise ValueError("浏览器 ID 不能为空")

        # 构造任务对象
        task = {
            "type": "chat",
            "browserId": browserId,
            "times": payload.get("times", []),
        }

        # 查找对应工人的 Worker
        worker = self.findWorker(browserId)
        # 未找到则新建 Worker
        if not worker:
            worker = self.createWorker(browserId)

        # 任务加入 Worker 队列
        worker.addTask(task)
        # 返回最新工人状态
        return worker.getSnapshot()

    def findWorker(self, browserId):
        """按浏览器 ID 查找 Worker"""
        # 遍历工人列表
        for worker in self.workers:
            # 匹配到相同 browserId 则返回
            if worker.browserId == browserId:
                return worker
        # 未找到返回 None
        return None

    def createWorker(self, browserId):
        """创建新 Worker"""
        # 实例化 Worker（此处省略具体类）
        worker = None  # Worker(browserId)
        # 加入工人列表
        self.workers.append(worker)
        # 启动 Worker 线程
        worker.start()
        return worker

    def getState(self):
        """获取全部状态供前端初始化"""
        # 收集所有工人快照
        workers = [w.getSnapshot() for w in self.workers]
        # 组装完整状态对象
        return {
            "workers": workers,
            "defaultWords": self.defaultWords,
            "filterDefaults": self.filterDefaults,
        }


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "browserId": "test-browser-001",
        "times": ["09:00", "14:00"],
    }

    # 创建服务实例
    service = TaskService()
    # 提交测试沟通任务
    result = service.submitChat(config)
    # 打印返回结果
    print(result)
```

---

## 5. 前端规范（HTML / JS）

前端同样遵循：**简短命名 + 中文注释**。

### 5.1 JavaScript

- 变量、函数：`camelCase`
- **禁止** `_` 开头命名
- 每个函数内关键步骤必须有中文注释

```javascript
function connectSocket() {
    // 创建 WebSocket 连接
    state.socket = new WebSocket(wsUrl());

    // 连接成功回调
    state.socket.onopen = () => {
        // 标记已连接
        state.connected = true;
        // 更新页面连接状态
        setConnStatus("已连接");
        // 开始定时同步
        startSync();
    };
}
```

### 5.2 DOM id — `camelCase`

```html
<button id="refreshBtn">刷新</button>
<select id="browserId"></select>
<div id="taskList"></div>
```

### 5.3 CSS 类名 — `kebab-case`

```html
<div class="panel editor-panel">
<button class="btn danger">
```

---

## 6. API 与 JSON 字段

### 6.1 GUI / 任务动作 — 简短点号或动词格式

```
config.load
config.save
claim.run
claim.preview
login.test
export.openDir
```

> 日志、按钮回调等命名保持简短；与 BOSS 子项目 WebSocket 的 `type` 字段规则相同，仅业务域不同。

### 6.2 JSON 字段 — `camelCase`

```json
{
  "username": "xxx",
  "password": "xxx",
  "exportDir": "C:\\path\\to\\export",
  "imgPath": "C:\\path\\to\\temp"
}
```

### 6.3 业务域中文字段（例外）

与赛狐 FBA 货件页面一致的筛选项，保留中文键名：

```python
self.filterDefaults = {
    "店铺": "全部",
    "时间范围": "最近30天",
    "差异类型": "全部",
}
```

---

## 7. 数据库命名

| 类型 | 规范 | 示例 |
|------|------|------|
| 表名 | `snake_case` | `task_runs` |
| 列名 | `snake_case` | `browser_id`、`started_at` |

> 数据库层保持 `snake_case`（SQLite 惯例），Python 读写时在代码内转换为 `camelCase` 变量。

---

## 8. 规范对照：旧写法 → 新写法

| 旧写法（禁止） | 新写法（推荐） |
|---------------|---------------|
| 模块级 `DEFAULT_WORDS = [...]` | `self.defaultWords = [...]` 写在 `__init__` |
| 单独 `automation_defaults.py` | 合并到对应类的 `__init__` |
| `def _run(self)` | `def run(self)` |
| `def _emit_log(self, msg)` | `def emitLog(self, msg)` |
| `def submit_chat_task(self)` | `def submitChat(self)` |
| `def keyword_rules_snapshot(self)` | `def getRules(self)` |
| `def recommend_job_filters_snapshot(self)` | `def getJobFilters(self)` |
| 无 main 块 | 必须有 `if __name__ == "__main__"` |
| 配置散落各处 | 调试配置统一放 main 块的 `config` |
| `def save(self, data: dict) -> int:` | `def save(self, data):` + docstring |
| `self.db: BossDb \| None = None` | `self.db = None` + 注释说明 |
| `from typing import Any` | 删除，用 docstring 描述结构 |

---

## 9. 检查清单（提交前自检）

- [ ] 文件名简短易读，且准确反映模块职责
- [ ] 文件内只有一个主类
- [ ] 所有常量、默认值在 `__init__` 里
- [ ] 有 `if __name__ == "__main__"` 且 config 在 main 块
- [ ] 无任何 `_` 开头的方法/函数名
- [ ] 方法名简短（不超过 4 个单词）
- [ ] 每个方法有中文 docstring
- [ ] 每个关键操作步骤后有中文注释
- [ ] 无参数/返回值/属性类型注解；无 `typing` 导入
- [ ] 可用 `python 文件名.py` 单独运行调试

---

## 10. 修订记录

| 日期 | 版本 | 说明 |
|------|------|------|
| 2026-06-28 | 2.0 | 按项目要求重写：一文件一类、常量进 init、配置进 main、禁止下划线命名、强制中文逐步注释 |
| 2026-06-29 | 2.1 | 新增 §2.8：禁止函数/类类型注解与 typing 导入 |
| 2026-07-02 | 2.1-fba | 自 BOSS 直聘子项目复制规范；更新 §1、§6 等项目介绍为 FBA 货件差异自动索赔，命名规则不变 |
