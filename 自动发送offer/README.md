# 企业微信入职申请-审批通过-自动发送offer系统

这是一个基于企业微信API的自动化入职管理系统，实现入职申请提交、审批流程监听、审批通过后自动发送Offer的功能。

## 功能特性

- ✅ 提交入职审批申请
- ✅ 审批状态实时监听
- ✅ 审批通过自动发送Offer
- ✅ 审批驳回自动通知
- ✅ 支持Markdown格式消息
- ✅ 支持文件附件发送
- ✅ 企业微信自动登录

## 系统架构

```
├── config.example.py      # 配置模板（复制为 config.py，勿提交真实密钥）
├── config.py              # 本地配置（.gitignore）
├── wecom_api.py           # 企业微信API基础模块
├── approval.py            # 审批申请模块
├── offer_sender.py        # Offer发送模块
├── approval_monitor.py    # 审批监听模块
├── QiYeVxLogin.py         # 企业微信登录管理器
├── main.py                # 主程序入口
├── requirements.txt       # 依赖文件
└── README.md             # 使用文档
```

## 安装步骤

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置企业微信信息

复制 `config.example.py` 为 `config.py`（`config.py` 已在 `.gitignore` 中，不会提交到 Git），填入您的企业微信配置信息：

```python
# 企业微信企业ID
CORP_ID = "your_corp_id"

# 自建应用Secret
AGENT_SECRET = "your_agent_secret"

# 自建应用AgentID
AGENT_ID = 1000000

# 审批模板ID（需要在企业微信管理后台创建）
APPROVAL_TEMPLATE_ID = "your_template_id"

# 回调URL（用于接收审批状态变化通知）
CALLBACK_URL = "https://your-domain.com/callback"

# 回调Token
CALLBACK_TOKEN = "your_callback_token"

# 回调EncodingAESKey
CALLBACK_ENCODING_AES_KEY = "your_encoding_aes_key"
```

### 3. 创建审批模板

在企业微信管理后台创建审批模板：

1. 登录企业微信管理后台
2. 进入"应用管理" -> "自建应用"
3. 选择您的应用，进入"审批接口"
4. 创建审批模板，添加以下字段：
   - 姓名（文本）
   - 部门（文本）
   - 职位（文本）
   - 入职日期（日期）
   - 联系电话（文本）
   - 邮箱（文本）
   - 学历（文本）
   - 薪资（文本）
   - 招聘负责人（成员）
   - 备注（文本）

5. 记录模板ID，填入 `config.py` 的 `APPROVAL_TEMPLATE_ID`

### 4. 配置回调（可选）

如果需要使用回调方式监听审批状态：

1. 在企业微信管理后台，进入自建应用的"设置API接收"
2. 填入回调URL、Token和EncodingAESKey
3. 开启"审批状态通知事件"

## 使用方法

### 1. 提交入职申请

```bash
python main.py --action submit \
  --user-id "zhangsan" \
  --name "张三" \
  --department "技术部" \
  --position "软件工程师" \
  --entry-date "2024-07-01" \
  --phone "13800138000" \
  --email "zhangsan@example.com" \
  --education "本科" \
  --salary "15000" \
  --recruiter "lisi" \
  --notes "应届毕业生"
```

提交后会自动启动审批监听，审批通过后自动发送Offer。

### 2. 启动审批监听

```bash
python main.py --action monitor --interval 60
```

- `--interval`: 轮询间隔（秒），默认60秒

### 3. 直接发送Offer（不经过审批）

```bash
python main.py --action send \
  --user-id "zhangsan" \
  --name "张三" \
  --department "技术部" \
  --position "软件工程师" \
  --entry-date "2024-07-01" \
  --salary "15000"
```

### 4. 确保企业微信登录

```bash
python main.py --action login --timeout 300
```

- `--timeout`: 登录超时时间（秒），默认300秒

## API说明

### WeComAPI

企业微信API基础类，提供认证、消息发送等功能。

**主要方法：**
- `get_access_token()`: 获取access_token
- `send_text_message(user_id, content)`: 发送文本消息
- `send_markdown_message(user_id, content)`: 发送Markdown消息
- `send_file_message(user_id, media_id)`: 发送文件消息
- `upload_file(file_path)`: 上传文件获取media_id
- `get_user_info(user_id)`: 获取用户信息
- `get_department_list()`: 获取部门列表

### ApprovalManager

审批管理类，提供审批申请提交、查询等功能。

**主要方法：**
- `submit_onboarding_approval(applicant_user_id, onboarding_data)`: 提交入职审批申请
- `get_approval_detail(sp_no)`: 获取审批详情
- `get_approval_status(sp_no)`: 获取审批状态

### OfferSender

Offer发送类，提供Offer生成和发送功能。

**主要方法：**
- `generate_offer_content(onboarding_data)`: 生成Offer内容
- `send_offer(user_id, onboarding_data, send_file)`: 发送Offer
- `send_approval_notification(user_id, approval_status, onboarding_data)`: 发送审批状态通知

### ApprovalMonitor

审批监听类，提供审批状态监听功能。

**主要方法：**
- `add_approval_record(sp_no, onboarding_data, applicant_user_id)`: 添加审批记录
- `start_monitoring(interval)`: 启动监听
- `stop_monitoring()`: 停止监听
- `get_monitoring_status()`: 获取监听状态

## 审批状态说明

- `1`: 审批中
- `2`: 已通过
- `3`: 已驳回
- `4`: 已撤销

## 日志

日志文件位于 `./logs/app.log`，包含系统运行日志和错误信息。

## 注意事项

1. **企业微信配置**: 确保config.py中的企业微信配置信息正确
2. **审批模板**: 必须先在企业微信管理后台创建审批模板
3. **权限配置**: 确保自建应用有足够的权限（审批、消息发送等）
4. **网络连接**: 确保服务器可以访问企业微信API（qyapi.weixin.qq.com）
5. **Token有效期**: access_token有效期为7200秒，系统会自动刷新
6. **回调配置**: 如果使用回调方式，需要配置公网可访问的回调URL

## 常见问题

### 1. 获取access_token失败

- 检查CORP_ID和AGENT_SECRET是否正确
- 检查网络连接是否正常
- 检查自建应用是否有权限

### 2. 提交审批申请失败

- 检查APPROVAL_TEMPLATE_ID是否正确
- 检查审批模板是否存在
- 检查申请人用户ID是否正确

### 3. 发送消息失败

- 检查AGENT_ID是否正确
- 检查接收人用户ID是否正确
- 检查应用是否有消息发送权限

### 4. 审批监听不工作

- 检查审批单号是否正确
- 检查网络连接是否正常
- 检查轮询间隔设置是否合理

## 扩展开发

### 添加数据库支持

可以修改代码添加数据库支持，用于存储审批记录和历史数据：

```python
import sqlite3

def save_approval_record(sp_no, onboarding_data, status):
    conn = sqlite3.connect('approvals.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO approvals (sp_no, data, status, create_time)
        VALUES (?, ?, ?, ?)
    ''', (sp_no, json.dumps(onboarding_data), status, datetime.now()))
    conn.commit()
    conn.close()
```

### 自定义Offer模板

修改 `offer_sender.py` 中的 `generate_offer_content` 方法，自定义Offer内容和格式。

### 添加Web界面

可以使用Flask或Django添加Web界面，方便用户提交入职申请和查看审批状态。

## 许可证

MIT License

## 联系方式

如有问题或建议，请联系开发团队。
