"""
企业微信配置文件示例
复制为 config.py 并填入真实值，config.py 已加入 .gitignore
"""

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

# Offer模板文件路径
OFFER_TEMPLATE_PATH = "./offer_template.html"

# 日志配置
LOG_LEVEL = "INFO"
LOG_FILE = "./logs/app.log"

# 数据库配置（可选，用于存储审批记录）
DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "wecom_approval",
}
