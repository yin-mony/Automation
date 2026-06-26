import time
import requests
from datetime import datetime


class Comment:
    def __init__(self, config):
        self.title = config.get("title", "工作日报")
        self.report_date = self.current_date()
        self.number = config.get("number", "18280194086")
        self.send_wechat = config.get("send_wechat", True)
        self.wechat_webhook = config.get(
            "wechat_webhook",
            "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=44a18a31-cd22-4c5c-984d-6896d610cac1",
        )

    def current_date(self):
        return datetime.now().strftime("%Y-%m-%d")

    def format_detail(self, value):
        if isinstance(value, list):
            return "\n".join(
                f"{index}.{str(line).strip()}"
                for index, line in enumerate(value, 1)
            )
        return str(value or "未填写").strip()

    def add_daily_item(self, data, title, owner, status, detail):
        data.append(
            {
                "事项": title,
                "负责人": owner,
                "状态": status,
                "详情": self.format_detail(detail),
            }
        )

    # 自动化流程操作
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

    # 处理自动化流程生成的数据
    def excel(self):
        items = self.main()
        if not items:
            print("没有需要生成的日报数据")
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
        if not data:
            print("没有需要生成的日报数据")
            return []
        return data

    def message_send(self, data=None):
        if data is None:
            data = self.excel()
        if not data:
            print("没有需要发送的日报数据")
            return
        if not self.wechat_webhook or "YOUR_WEBHOOK_KEY" in self.wechat_webhook:
            print("未配置企业微信 webhook，跳过发送")
            return

        webhook = self.wechat_webhook
        session = requests.Session()
        session.trust_env = False
        report_map = {}
        for item in data:
            report_title = item.get("日报标题", self.title) or self.title
            report_date = item.get("日期", self.report_date) or self.report_date
            report_name = f"{report_title}-{report_date}"
            report_map.setdefault(report_name, []).append(item)

        for report_name, items in report_map.items():
            content = f"{report_name}\n"

            for index, item in enumerate(items, 1):
                content += (
                    f"{index}.{item.get('事项', '')}\n"
                    f"负责人：{item.get('负责人', '')}\n"
                    f"状态：{item.get('状态', '')}\n"
                    f"详情：{item.get('详情', '')}\n\n"
                )

            payload = {
                "msgtype": "text",
                "text": {
                    "content": content.rstrip(),
                    "mentioned_mobile_list": [f"{self.number}"]
                }
            }
            for attempt in range(1, 4):
                try:
                    res = session.post(
                        webhook,
                        json=payload,
                        headers={"Connection": "close"},
                        timeout=(5, 20)
                    )
                    result = res.json()
                    print(f"{report_name} 企业微信发送结果: {result}")
                    if result.get("errcode") == 0:
                        break
                    if result.get("errcode") == 45009:
                        time.sleep(10 * attempt)
                        continue
                    break
                except requests.exceptions.RequestException as e:
                    print(f"{report_name} 企业微信第 {attempt} 次发送失败: {e}")
                    if attempt < 3:
                        time.sleep(3 * attempt)
            time.sleep(3)

    # 启动
    def run(self):
        data = self.excel()
        if self.send_wechat:
            self.message_send(data)
        return data


if __name__ == '__main__':
    config = {
        # 企业微信使用的手机号
        "number": "18280194086",
        # 企业微信机器人 webhook
        "wechat_webhook": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=44a18a31-cd22-4c5c-984d-6896d610cac1",
        # 是否发送企业微信
        "send_wechat": True,
        # 日报标题
        "title": "工作日报",
    }
    comment = Comment(config)
    comment.run()
