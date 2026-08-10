"""日报生成器 — 组装今日/明日事项，分条 Markdown 推送企业微信群机器人。"""

import time
import requests
from datetime import datetime, timedelta

DEFAULT_TITLE = "工作日报"
DEFAULT_NUMBER = "18280194086"
DEFAULT_WEBHOOK = (
    "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=44a18a31-cd22-4c5c-984d-6896d610cac1"
)

STATUS_COLORS = {
    "开发中": "warning",
    "已完成": "info",
    "测试中": "comment",
}


class Comment:
    """工作日报：今日进展与明日待办分条 Markdown 推送。"""

    def __init__(self, config):
        self.title = config.get("title", DEFAULT_TITLE)
        self.report_date = config.get("report_date") or self.current_date()
        self.number = config.get("number", DEFAULT_NUMBER)
        self.send_wechat = config.get("send_wechat", True)
        self.wechat_webhook = config.get("wechat_webhook", DEFAULT_WEBHOOK)
        self.tomorrow_items = config.get("tomorrow_items") or []

    def current_date(self):
        return datetime.now().strftime("%Y-%m-%d")

    def tomorrow_date(self):
        return (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")

    def format_detail(self, value):
        """将详情规范为带序号的纯文本行。"""
        if isinstance(value, list):
            lines = [str(line).strip() for line in value if str(line).strip()]
            return "\n".join(f"{index}.{line}" for index, line in enumerate(lines, 1))
        text = str(value or "未填写").strip()
        if not text:
            return "未填写"
        parts = [line.strip() for line in text.splitlines() if line.strip()]
        if len(parts) <= 1:
            return parts[0] if parts else "未填写"
        return "\n".join(f"{index}.{line}" for index, line in enumerate(parts, 1))

    def format_detail_markdown(self, value):
        """详情转为 Markdown 引用块，多行逐条缩进。"""
        text = self.format_detail(value)
        lines = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            lines.append(f"> {line}")
        return "\n".join(lines) if lines else "> 未填写"

    def status_color(self, status):
        return STATUS_COLORS.get(str(status or "").strip(), "comment")

    def add_daily_item(self, data, title, owner, status, detail):
        data.append(
            {
                "事项": title,
                "负责人": owner,
                "状态": status,
                "详情": self.format_detail(detail),
            }
        )

    def add_tomorrow_item(self, data, title, owner, detail):
        data.append(
            {
                "事项": title,
                "负责人": owner,
                "详情": self.format_detail(detail),
            }
        )

    def build_items_markdown(self, heading, items, include_status=True):
        """将事项列表拼成 Markdown 正文。"""
        if not items:
            return ""

        lines = [f"## {heading}", ""]
        for index, item in enumerate(items, 1):
            title = item.get("事项", "未命名事项")
            owner = item.get("负责人", "未填写")
            lines.append(f"**{index}. {title}**")
            if include_status:
                status = item.get("状态", "开发中")
                color = self.status_color(status)
                lines.append(
                    f"> 负责人：{owner} | 状态：<font color=\"{color}\">{status}</font>"
                )
            else:
                lines.append(f"> 负责人：{owner}")
            lines.append("")
            lines.append("**详情**")
            lines.append(self.format_detail_markdown(item.get("详情", "")))
            if index < len(items):
                lines.extend(["", "---", ""])
        return "\n".join(lines).strip()

    # 自动化流程操作（CLI 示例数据）
    def main(self):
        data = []
        self.add_daily_item(
            data,
            "Shopify网站页面排版调整",
            "李薇",
            "开发中",
            "将当前网站从“产品陈列型页面”升级为“美国消费者信任的专业清洁用品品牌官网",
        )
        self.add_daily_item(
            data,
            "下载全站点汇总报告",
            "杨露",
            "开发中",
            [
                "用于下载美国、加拿大、墨西哥、巴西四个站点的月度 Summary PDF 报告。",
                "要求自动切换站点、选择上个月时间、生成并下载报告，文件按 默认文件名-店铺名-站点 命名。",
            ],
        )
        return data

    def excel(self, items=None):
        """将事项列表展平为带元数据的行（供日志或扩展导出）。"""
        items = self.main() if items is None else items
        if not items:
            return []

        data = []
        for index, item in enumerate(items, 1):
            data.append(
                {
                    "日报标题": self.title,
                    "日期": self.report_date,
                    "序号": index,
                    "事项": item.get("事项", "未命名事项"),
                    "负责人": item.get("负责人", "未填写"),
                    "状态": item.get("状态", "开发中"),
                    "详情": item.get("详情", "未填写"),
                }
            )
        print(data)
        return data

    def _post_payload(self, session, webhook, payload, label):
        """带重试的 webhook POST。"""
        for attempt in range(1, 4):
            try:
                res = session.post(
                    webhook,
                    json=payload,
                    headers={"Connection": "close"},
                    timeout=(5, 20),
                )
                result = res.json()
                print(f"{label} 企业微信发送结果: {result}")
                if result.get("errcode") == 0:
                    return True
                if result.get("errcode") == 45009:
                    time.sleep(10 * attempt)
                    continue
                return False
            except requests.exceptions.RequestException as exc:
                print(f"{label} 企业微信第 {attempt} 次发送失败: {exc}")
                if attempt < 3:
                    time.sleep(3 * attempt)
        return False

    def send_markdown(self, session, webhook, content, label):
        payload = {"msgtype": "markdown", "markdown": {"content": content}}
        return self._post_payload(session, webhook, payload, label)

    def send_mention_text(self, session, webhook, content):
        if not self.number:
            return True
        payload = {
            "msgtype": "text",
            "text": {
                "content": content,
                "mentioned_mobile_list": [str(self.number)],
            },
        }
        return self._post_payload(session, webhook, payload, "提醒通知")

    def message_send(self, daily_items=None, tomorrow_items=None):
        """今日日报与明日待办分两条 Markdown 发送；最后可选 @ 手机号。"""
        daily_items = self.main() if daily_items is None else list(daily_items or [])
        tomorrow_items = (
            self.tomorrow_items if tomorrow_items is None else list(tomorrow_items or [])
        )

        if not daily_items and not tomorrow_items:
            print("没有需要发送的日报或待办数据")
            return

        if not self.wechat_webhook or "YOUR_WEBHOOK_KEY" in self.wechat_webhook:
            print("未配置企业微信 webhook，跳过发送")
            return

        webhook = self.wechat_webhook
        session = requests.Session()
        session.trust_env = False

        if daily_items:
            heading = f"{self.title} {self.report_date}"
            content = self.build_items_markdown(heading, daily_items, include_status=True)
            self.send_markdown(session, webhook, content, heading)
            time.sleep(3)

        if tomorrow_items:
            heading = f"明日待办 {self.tomorrow_date()}"
            content = self.build_items_markdown(heading, tomorrow_items, include_status=False)
            self.send_markdown(session, webhook, content, heading)
            time.sleep(3)

        if self.number:
            parts = []
            if daily_items:
                parts.append(f"{self.title}（{self.report_date}）")
            if tomorrow_items:
                parts.append(f"明日待办（{self.tomorrow_date()}）")
            notice = "、".join(parts) + " 已推送，请查阅"
            self.send_mention_text(session, webhook, notice)

    def run(self):
        """入口：CLI 下采集示例数据并推送。"""
        daily_items = self.main()
        if self.send_wechat:
            self.message_send(daily_items=daily_items, tomorrow_items=self.tomorrow_items)
        return self.excel(daily_items)


if __name__ == "__main__":
    config = {
        "number": DEFAULT_NUMBER,
        "wechat_webhook": DEFAULT_WEBHOOK,
        "send_wechat": True,
        "title": DEFAULT_TITLE,
        "tomorrow_items": [],
    }
    comment = Comment(config)
    comment.run()
