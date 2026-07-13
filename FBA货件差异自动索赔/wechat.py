"""企业微信群机器人消息推送。"""

import json
import re
import urllib.request


class Wechat:
    """企业微信群机器人 CASE 结果汇总推送"""

    def __init__(self, config=None):
        # 默认群机器人 Webhook，可由 GUI 配置覆盖
        self.defaultWebhook = ""
        # 默认 @ 手机号，可由 GUI 配置覆盖
        self.defaultMobile = ""
        # 请求超时时间
        self.timeout = 15
        # 运行配置
        self.config = config or {}
        self.webhook = str(self.config.get("wechatWebhook") or self.defaultWebhook).strip()
        self.mobile = str(self.config.get("wechatMobile") or self.defaultMobile).strip()
        self.email = str(self.config.get("email") or "").strip()

    def getMobiles(self, mobileText):
        """将 GUI 填写的手机号文本整理为企业微信 @ 列表"""
        text = str(mobileText or "").strip()
        # 未填写手机号时不 @ 指定人员
        if not text:
            return []
        # 支持逗号、分号、空格、换行分隔多个手机号
        parts = re.split(r"[,，;；\s]+", text)
        mobiles = []
        for item in parts:
            item = str(item or "").strip()
            # 过滤空值并去重
            if item and item not in mobiles:
                mobiles.append(item)
        return mobiles

    def buildCaseText(self, data):
        """按 CASE 处理结果生成企业微信文本消息"""
        resultList = data.get("resultList") or []
        failList = data.get("failList") or []
        skipList = data.get("skipList") or []
        email = str(data.get("email") or self.email or "未配置").strip() or "未配置"
        caseResultPath = str(data.get("caseResultPath") or "").strip()

        lines = [
            "FBA货件索赔处理完成",
            f"成功提交：{len(resultList)} 票",
            f"处理失败：{len(failList)} 票",
            f"跳过处理：{len(skipList)} 票",
            f"文件结果：POP/POD 文件已上传，相关文件已发送至邮箱：{email}",
        ]
        if caseResultPath:
            lines.append(f"本地结果文件：{caseResultPath}")

        # 汇总成功提交的 CASE 结果
        if resultList:
            lines.append("")
            lines.append("CASE 结果：")
            for index, item in enumerate(resultList, start=1):
                shipmentId = str(item.get("shipmentId") or "").strip()
                caseId = str(item.get("caseId") or "").strip()
                popFile = str(item.get("popFile") or "").strip()
                if popFile:
                    lines.append(f"{index}. 货件编号：{shipmentId}，CASE问题编号：{caseId}，POP文件：{popFile}")
                else:
                    lines.append(f"{index}. 货件编号：{shipmentId}，CASE问题编号：{caseId}")

        # 汇总失败原因，方便人工回查
        if failList:
            lines.append("")
            lines.append("失败明细：")
            for index, item in enumerate(failList, start=1):
                shipmentId = str(item.get("shipmentId") or "").strip()
                reason = str(item.get("reason") or "").strip()
                lines.append(f"{index}. 货件编号：{shipmentId}，原因：{reason}")

        # 汇总跳过原因，方便区分无差异货件
        if skipList:
            lines.append("")
            lines.append("跳过明细：")
            for index, item in enumerate(skipList, start=1):
                shipmentId = str(item.get("shipmentId") or "").strip()
                reason = str(item.get("reason") or "").strip()
                lines.append(f"{index}. 货件编号：{shipmentId}，原因：{reason}")

        return "\n".join(lines)

    def sendText(self, content, mobileText=None):
        """发送企业微信群机器人 text 消息"""
        if not self.webhook:
            raise ValueError("企业微信 Webhook 不能为空")
        mobiles = self.getMobiles(mobileText if mobileText is not None else self.mobile)
        payload = {
            "msgtype": "text",
            "text": {
                "content": str(content or ""),
                "mentioned_mobile_list": mobiles,
            },
        }
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            self.webhook,
            data=data,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        # 调用企业微信机器人接口
        with urllib.request.urlopen(request, timeout=self.timeout) as response:
            body = response.read().decode("utf-8")
        result = json.loads(body)
        # 企业微信返回非 0 表示发送失败
        if result.get("errcode") != 0:
            raise RuntimeError(f"企业微信发送失败: {result}")
        return result

    def sendCase(self, data):
        """发送 CASE 结果汇总并 @ 指定手机号"""
        content = self.buildCaseText(data)
        mobileText = data.get("wechatMobile") or self.mobile
        # CASE 汇总消息已生成，开始发送
        return self.sendText(content, mobileText)


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "wechatWebhook": "",
        "wechatMobile": "",
        "email": "test@example.com",
    }
    data = {
        "resultList": [
            {
                "shipmentId": "FBA_TEST",
                "caseId": "19000000000",
                "popFile": "Lydia deal-US_FBA_TEST_POP.pdf",
            }
        ],
        "failList": [],
        "skipList": [],
        "email": config["email"],
        "wechatMobile": config["wechatMobile"],
    }
    service = Wechat(config)
    print(service.buildCaseText(data))
