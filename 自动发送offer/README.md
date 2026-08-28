# 人事 Offer 制单工具

这是一个最基本的企业微信自建应用。人事从企业微信工作台打开页面，上传候选人简历图片、PDF或Word文件；系统提取信息并显示复核表，人事二次确认后生成正式 Offer PDF、留存邮件草稿并真实发送。候选人无需加入企业。

## 正确流程

```text
企业微信内部成员打开应用
  -> 企业微信静默确认成员身份
  -> 人事上传候选人简历
  -> 本地提取文本或 OCR 识别图片
  -> 自动填写候选人和岗位信息
  -> 人事修改、补齐并最终确认
  -> 点击“确认并发送 Offer”
  -> 弹窗二次确认收件人、抄送人和不可撤回提示
  -> 生成 Offer PDF 和本地 EML 草稿
  -> 同步到腾讯企业邮箱 Drafts 草稿箱留档
  -> 通过 SMTP 真实发送给候选人，并把负责人写入邮件 Cc 抄送
```

## 文件职责

```text
run.py              企业微信身份校验和三步制单流程入口
resume.py           图片、PDF、Word简历读取与字段提取
jobs.py             本地任务和识别结果存储
offer.py            正式 Offer PDF 生成
draft.py            EML生成与企业邮箱草稿箱同步
settings.py         本地配置读取
templates/index.html 上传、复核和完成共用的单页面
static/style.css    人事操作界面样式
static/vendor/tabler/ 本地 Tabler UI 组件样式、脚本与许可
tests/              不发送真实邮件的流程测试
```

## 支持格式

- 图片：PNG、JPG、JPEG、WEBP、BMP
- PDF：文本PDF直接提取，扫描PDF逐页OCR
- Word：DOCX正文和表格

自动识别姓名、邮箱、手机号、学历、院校、专业、所在城市和求职意向。识别结果只是建议值，必须由人事复核。

薪酬按试用期和转正分别填写基本工资、保密费和绩效，总薪酬自动合计。入职日期和报到起止时间均通过日期、时间选择控件填写。

页面使用本地打包的 Tabler UI 1.4.0，不依赖外部 CDN，也不需要单独的 Vue 或 Node.js 前端服务。static/style.css 只负责公司品牌色、薪酬结构布局和企业微信手机端适配。

试用期和转正薪酬均按“基本工资 + 保密费 + 绩效 = 总薪酬”显示，总薪酬由页面和后端自动合计，不能手动修改。真实发送采用二次确认弹窗；取消不会提交表单，确认后系统先留存草稿再通过 SMTP 发送。固定抄送人会写入邮件的 `Cc` 字段，页面只展示抄送人姓名，不展示邮箱；候选人点击“回复全部”时抄送人会收到回复。

## 安装与启动

```powershell
cd F:\Automation\自动发送offer
python -m pip install -r requirements.txt
python run.py
```

本机调试打开：

```text
http://127.0.0.1:8700
```

本机地址仅用于调试。正式使用时需要把服务部署到 HTTPS 域名，并把企业微信自建应用主页配置为：

```text
https://www.bonison.net/offer/
```

服务器配置需要补充：

```python
PUBLIC_URL = "https://www.bonison.net/offer"
FLASK_SECRET_KEY = "随机长字符串"
CORP_ID = "企业ID"
AGENT_ID = "自建应用AgentId"
AGENT_SECRET = "自建应用Secret"
OFFER_NOTIFY_USER_IDS = ["负责人企业微信UserID"]
OFFER_CC_RECIPIENTS = [
    {"name": "何倩怡", "email": "heqianyi@bonison.net"},
    {"name": "宁致远", "email": "ningzhiyuan@bonison.net"},
]
```

应用可见范围应只包含有权制作 Offer 的人事员工。原来的“API接收消息”回调不是本流程必需项，可以继续独立运行，但不参与 Offer 制单。

## 邮箱草稿

腾讯企业邮箱配置：

```python
SMTP_USERNAME = "wangxiao@bonison.net"
SMTP_PASSWORD = "客户端专用密码"
MAIL_FROM = "wangxiao@bonison.net"
IMAP_HOST = "imap.exmail.qq.com"
IMAP_PORT = 993
IMAP_USE_SSL = True
DRAFT_FOLDER = "Drafts"
```

确认发送后，系统会通过 IMAP 把带 PDF 附件的邮件写入草稿箱，再通过 SMTP 真实发送。草稿箱同步失败时不会继续发送；无论同步是否成功，都会生成本地 `.eml` 作为兜底。

## 文件位置

- 简历和识别结果：`data/jobs/<任务编号>/`
- Offer PDF和邮件草稿：`output/<任务编号>/`
- 本地配置：`config.py`

以上目录均已加入 `.gitignore`。候选人简历、识别文字、PDF、邮件草稿和邮箱密码不会提交到 Git。

## 测试

```powershell
python -m unittest discover -s tests -v
```

测试只使用模拟候选人数据，生成临时PDF和EML，不连接真实邮箱。

## 注意事项

- OCR可能识别错误，姓名、邮箱、岗位、薪资和入职日期必须人工复核。
- 简历和Offer包含个人信息，应定期清理 `data/` 与 `output/`。
- Offer中的劳动关系、薪酬、试用期和福利条款应由公司人事或法务最终确认。
- 页面二次确认后会真实发送邮件，确认前必须核对收件邮箱和全部录用信息。

## 宝塔 Linux 部署

生产环境使用 Gunicorn，不使用 Flask 自带调试服务器：

```bash
gunicorn --workers 2 --bind 127.0.0.1:8700 --timeout 180 run:app
```

Nginx 将 `/offer/` 反向代理到 `http://127.0.0.1:8700`。服务器需安装中文字体，并在 `config.py` 设置：

```python
PORTAL_HOST = "127.0.0.1"
PORTAL_PORT = 8700
PUBLIC_URL = "https://www.bonison.net/offer"
FONT_PATH = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
```
