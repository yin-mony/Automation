from datetime import datetime
import os
from pathlib import Path

import pandas as pd
import requests


class CommentAnalyzer:
    """亚马逊评论 AI 分析与报告生成"""

    def __init__(self, path="", apiKey=""):
        # 评论 Excel 文件路径
        self.path = Path(path) if path else Path("")
        # OpenAI API Key
        self.apiKey = apiKey
        # OpenAI Chat Completions 接口
        self.apiUrl = "https://api.openai.com/v1/chat/completions"
        # 默认分析模型
        self.model = "gpt-4o-mini"
        # 单次传入 AI 的评论上限，避免 token 过多
        self.sampleLimit = 50
        # 本轮 AI 调用失败次数
        self.apiErrorCount = 0
        # 本轮 AI 调用失败摘要
        self.apiErrorMessages = []

    def run(self):
        """读取评论 Excel 并按 ASIN 生成分析报告"""
        # 每次运行前重置接口错误状态
        self.apiErrorCount = 0
        self.apiErrorMessages = []
        # 读取评论 Excel
        df = pd.read_excel(self.path)
        # 收集每个 ASIN 的分析结果
        allResults = []

        for asin in df["ASIN"].unique():
            # 拆分当前 ASIN 的评论
            print(f"\n正在分析 ASIN: {asin}")
            asinDf = df[df["ASIN"] == asin]
            goodDf = asinDf[asinDf["星级"].astype(str).str.contains("5星|4星")]
            badDf = asinDf[asinDf["星级"].astype(str).str.contains("3星|2星|1星")]
            print(f"  总评论: {len(asinDf)}, 好评: {len(goodDf)}, 差评: {len(badDf)}")

            # 分别生成好评卖点、差评痛点和改进建议
            goodSummary = self.analyzeReview(goodDf["评论内容"].tolist(), "good", asin)
            badSummary = self.analyzeReview(badDf["评论内容"].tolist(), "bad", asin)
            suggestions = self.getSuggest(badDf["评论内容"].tolist(), asin)

            # 组装当前 ASIN 报告结果
            allResults.append({
                "ASIN": asin,
                "总评论": len(asinDf),
                "好评数": len(goodDf),
                "差评数": len(badDf),
                "好评率": f"{len(goodDf) / len(asinDf) * 100:.1f}%" if len(asinDf) > 0 else "0%",
                "好评卖点": goodSummary,
                "差评痛点": badSummary,
                "改进建议": suggestions,
            })

        # 保存完整报告
        self.saveReport(allResults)
        if self.apiErrorCount > 0:
            firstMessage = self.apiErrorMessages[0] if self.apiErrorMessages else "接口返回异常"
            raise RuntimeError(f"AI分析接口调用失败 {self.apiErrorCount} 次，请检查 API Key 或网络配置。{firstMessage}")
        print("\n分析完成")

    def analyzeReview(self, reviews, reviewType, asin):
        """调用 AI 分析好评卖点或差评痛点"""
        # 没有评论时直接返回空数据说明
        if not reviews:
            return f"暂无足够的{reviewType}数据"

        # 截取评论样本，避免单次请求过大
        reviewsSample = reviews[:self.sampleLimit]
        reviewsText = "\n".join([f"- {review}" for review in reviewsSample])

        if reviewType == "good":
            # 好评提示词，要求总结买家认可的卖点
            prompt = f"""请分析以下关于商品 {asin} 的好评（5星和4星评论），总结出买家最认可的3-5个卖点。

要求：
1. 每个卖点用一句话概括
2. 如果有具体数据支持更佳
3. 用中文回复，简洁明了

好评内容：
{reviewsText}

请输出格式（每个卖点一行）：
1. [卖点描述]
2. [卖点描述]
..."""
        else:
            # 差评提示词，要求总结买家抱怨的痛点
            prompt = f"""请分析以下关于商品 {asin} 的差评（3星、2星和1星评论），总结出买家最抱怨的3-5个痛点。

要求：
1. 每个痛点用一句话概括
2. 如果有具体数据支持更佳
3. 用中文回复，简洁明了

差评内容：
{reviewsText}

请输出格式（每个痛点一行）：
1. [痛点描述]
2. [痛点描述]
..."""

        try:
            # 发送请求并返回模型输出
            return self.callApi(prompt)
        except Exception as exc:
            # API 失败时写入日志并返回可读错误
            message = self.recordApiError(exc)
            print(f"  AI分析失败: {message}")
            return "AI分析失败，请检查API配置"

    def getSuggest(self, reviews, asin):
        """基于差评生成改进建议"""
        # 没有差评时直接返回说明
        if not reviews:
            return "暂无明显的差评问题"

        # 截取评论样本，避免单次请求过大
        reviewsSample = reviews[:self.sampleLimit]
        reviewsText = "\n".join([f"- {review}" for review in reviewsSample])

        # 改进建议提示词
        prompt = f"""基于以下关于商品 {asin} 的差评内容，请提出3-5条具体的改进建议。

要求：
1. 建议要具体、可执行
2. 按重要性排序
3. 用中文回复，每条一行

差评内容：
{reviewsText}

请输出格式：
1. [改进建议]
2. [改进建议]
..."""

        try:
            # 发送请求并返回模型输出
            return self.callApi(prompt)
        except Exception as exc:
            # API 失败时写入日志并返回可读错误
            message = self.recordApiError(exc)
            print(f"  生成建议失败: {message}")
            return "AI分析失败，请检查API配置"

    def recordApiError(self, exc):
        """记录并脱敏 API 错误信息"""
        # 累计失败次数，供 run 判断最终状态
        self.apiErrorCount += 1
        # 只保留可排查摘要，避免日志里带出 API Key 片段
        message = str(exc).strip() or type(exc).__name__
        if self.apiKey:
            message = message.replace(self.apiKey, "[API_KEY]")
            if len(self.apiKey) >= 12:
                message = message.replace(self.apiKey[:8], "[API_KEY]")
                message = message.replace(self.apiKey[-4:], "[API_KEY]")
        self.apiErrorMessages.append(message)
        return message

    def callApi(self, prompt):
        """调用 OpenAI Chat Completions API"""
        # 组装请求头
        headers = {
            "Authorization": f"Bearer {self.apiKey}",
            "Content-Type": "application/json",
        }

        # 组装请求体
        data = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": "你是一个专业的电商评论分析助手，擅长从用户评论中提取关键信息并给出有价值的分析。",
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.7,
            "max_tokens": 1000,
        }

        # 发送 API 请求
        try:
            response = requests.post(self.apiUrl, headers=headers, json=data, timeout=60)
        except requests.RequestException as exc:
            raise RuntimeError(f"API请求失败: {type(exc).__name__}") from exc
        if response.status_code != 200:
            raise RuntimeError(self.formatApiError(response))

        # 解析模型回复文本
        result = response.json()
        return result["choices"][0]["message"]["content"].strip()

    def formatApiError(self, response):
        """格式化 API 错误摘要，不输出接口返回全文"""
        # 解析 OpenAI 错误类型，避免打印可能包含 Key 片段的 message
        errorCode = ""
        errorType = ""
        try:
            payload = response.json()
            if isinstance(payload, dict):
                error = payload.get("error", {})
                if isinstance(error, dict):
                    errorCode = error.get("code") or ""
                    errorType = error.get("type") or ""
        except ValueError:
            pass

        if response.status_code == 401:
            reason = "API Key 无效或已失效"
        elif response.status_code == 429:
            reason = "请求频率或额度受限"
        elif response.status_code >= 500:
            reason = "OpenAI 服务暂时异常"
        else:
            reason = "接口返回异常"

        details = "，".join([item for item in [errorCode, errorType] if item])
        if details:
            return f"API调用失败: {response.status_code}，{reason}，错误类型: {details}"
        return f"API调用失败: {response.status_code}，{reason}"

    def saveReport(self, allResults):
        """保存 Excel 和文本分析报告"""
        # 创建报告目录
        outDir = self.path.parent / "分析报告"
        outDir.mkdir(parents=True, exist_ok=True)
        # 生成时间戳文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 保存 Excel 报告
        excelFile = outDir / f"评论分析报告_{timestamp}.xlsx"
        with pd.ExcelWriter(excelFile, engine="openpyxl") as writer:
            summaryData = []
            for result in allResults:
                # 提取整体统计字段
                summaryData.append({
                    "ASIN": result["ASIN"],
                    "总评论": result["总评论"],
                    "好评数": result["好评数"],
                    "差评数": result["差评数"],
                    "好评率": result["好评率"],
                })
            pd.DataFrame(summaryData).to_excel(writer, sheet_name="整体统计", index=False)

            reportData = []
            for result in allResults:
                # 提取完整总结字段
                reportData.append({
                    "ASIN": result["ASIN"],
                    "总评论": result["总评论"],
                    "好评数": result["好评数"],
                    "差评数": result["差评数"],
                    "好评率": result["好评率"],
                    "好评卖点": result["好评卖点"],
                    "差评痛点": result["差评痛点"],
                    "改进建议": result["改进建议"],
                })
            pd.DataFrame(reportData).to_excel(writer, sheet_name="总结报告", index=False)

        # 保存文本报告
        txtFile = outDir / f"分析报告_{timestamp}.txt"
        with open(txtFile, "w", encoding="utf-8") as file:
            file.write("=" * 60 + "\n")
            file.write("亚马逊评论AI分析报告\n")
            file.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            file.write("=" * 60 + "\n\n")

            for result in allResults:
                # 写入单个 ASIN 的完整分析段落
                file.write(f"\n{'=' * 50}\n")
                file.write(f"ASIN: {result['ASIN']}\n")
                file.write(f"{'=' * 50}\n")
                file.write(
                    f"总评论: {result['总评论']} | 好评: {result['好评数']} | "
                    f"差评: {result['差评数']} | 好评率: {result['好评率']}\n\n"
                )
                file.write(f"【好评卖点】\n{result['好评卖点']}\n\n")
                file.write(f"【差评痛点】\n{result['差评痛点']}\n\n")
                file.write(f"【改进建议】\n{result['改进建议']}\n\n")

        print(f"Excel报告: {excelFile}")
        print(f"文本报告: {txtFile}")


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "excelPath": r"C:\RPA流程\亚马逊评论分析\flie\亚马逊评论.xlsx",
        "apiKey": os.getenv("OPENAI_API_KEY", ""),
    }

    # API Key 为空时只提示调试方式，避免误发请求
    if not config["apiKey"]:
        print("请在 analysis.py 的 main 配置中填写 OpenAI API Key 后再调试。")
    else:
        analyzer = CommentAnalyzer(config["excelPath"], config["apiKey"])
        analyzer.run()
