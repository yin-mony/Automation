from pathlib import Path
import pandas as pd
from datetime import datetime
import requests
import json


class CommentAnalyzer:
    def __init__(self, path, api_key):
        self.path = Path(path)
        self.api_key = api_key
        self.api_url = "https://api.deepseek.com/v1/chat/completions"

    def run(self):
        """运行分析"""
        df = pd.read_excel(self.path)
        all_results = []

        for asin in df['ASIN'].unique():
            print(f"\n正在分析 ASIN: {asin}")

            asin_df = df[df['ASIN'] == asin]
            good = asin_df[asin_df['星级'].astype(str).str.contains('5星|4星')]
            bad = asin_df[asin_df['星级'].astype(str).str.contains('3星|2星|1星')]

            print(f"  总评论: {len(asin_df)}, 好评: {len(good)}, 差评: {len(bad)}")

            # 调用AI分析
            good_summary = self.analyze_with_ai(good['评论内容'].tolist(), 'good', asin)
            bad_summary = self.analyze_with_ai(bad['评论内容'].tolist(), 'bad', asin)
            suggestions = self.get_suggestions_from_ai(bad['评论内容'].tolist(), asin)

            all_results.append({
                'ASIN': asin,
                '总评论': len(asin_df),
                '好评数': len(good),
                '差评数': len(bad),
                '好评率': f"{len(good) / len(asin_df) * 100:.1f}%" if len(asin_df) > 0 else "0%",
                '好评卖点': good_summary,
                '差评痛点': bad_summary,
                '改进建议': suggestions
            })

        self.save_report(all_results)
        print(f"\n✅ 分析完成！")

    def analyze_with_ai(self, reviews, review_type, asin):
        """调用AI分析评论"""
        if not reviews:
            return f"暂无足够的{review_type}数据"

        # 限制评论数量（避免token过多）
        reviews_sample = reviews[:50]
        reviews_text = "\n".join([f"- {r}" for r in reviews_sample])

        if review_type == 'good':
            prompt = f"""请分析以下关于商品 {asin} 的好评（5星和4星评论），总结出买家最认可的3-5个卖点。

要求：
1. 每个卖点用一句话概括
2. 如果有具体数据支持更佳
3. 用中文回复，简洁明了

好评内容：
{reviews_text}

请输出格式（每个卖点一行）：
1. [卖点描述]
2. [卖点描述]
..."""
        else:
            prompt = f"""请分析以下关于商品 {asin} 的差评（3星、2星和1星评论），总结出买家最抱怨的3-5个痛点。

要求：
1. 每个痛点用一句话概括
2. 如果有具体数据支持更佳
3. 用中文回复，简洁明了

差评内容：
{reviews_text}

请输出格式（每个痛点一行）：
1. [痛点描述]
2. [痛点描述]
..."""

        try:
            response = self.call_deepseek_api(prompt)
            return response
        except Exception as e:
            print(f"  AI分析失败: {e}")
            return "AI分析失败，请检查API配置"

    def get_suggestions_from_ai(self, reviews, asin):
        """调用AI生成改进建议"""
        if not reviews:
            return "暂无明显的差评问题"

        reviews_sample = reviews[:50]
        reviews_text = "\n".join([f"- {r}" for r in reviews_sample])

        prompt = f"""基于以下关于商品 {asin} 的差评内容，请提出3-5条具体的改进建议。

要求：
1. 建议要具体、可执行
2. 按重要性排序
3. 用中文回复，每条一行

差评内容：
{reviews_text}

请输出格式：
1. [改进建议]
2. [改进建议]
..."""

        try:
            response = self.call_deepseek_api(prompt)
            return response
        except Exception as e:
            print(f"  生成建议失败: {e}")
            return "AI分析失败，请检查API配置"

    def call_deepseek_api(self, prompt):
        """调用DeepSeek API"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        data = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system",
                 "content": "你是一个专业的电商评论分析助手，擅长从用户评论中提取关键信息并给出有价值的分析。"},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.7,
            "max_tokens": 1000
        }

        response = requests.post(self.api_url, headers=headers, json=data, timeout=60)

        if response.status_code == 200:
            result = response.json()
            return result['choices'][0]['message']['content'].strip()
        else:
            raise Exception(f"API调用失败: {response.status_code} - {response.text}")

    def save_report(self, all_results):
        """保存报告"""
        out_dir = self.path.parent / "分析报告"
        out_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Excel报告
        excel_file = out_dir / f"评论分析报告_{timestamp}.xlsx"

        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Sheet1: 整体统计
            summary_data = []
            for r in all_results:
                summary_data.append({
                    'ASIN': r['ASIN'],
                    '总评论': r['总评论'],
                    '好评数': r['好评数'],
                    '差评数': r['差评数'],
                    '好评率': r['好评率']
                })
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='整体统计', index=False)

            # Sheet2: 总结报告
            report_data = []
            for r in all_results:
                report_data.append({
                    'ASIN': r['ASIN'],
                    '总评论': r['总评论'],
                    '好评数': r['好评数'],
                    '差评数': r['差评数'],
                    '好评率': r['好评率'],
                    '好评卖点': r['好评卖点'],
                    '差评痛点': r['差评痛点'],
                    '改进建议': r['改进建议']
                })
            pd.DataFrame(report_data).to_excel(writer, sheet_name='总结报告', index=False)

        # 文本报告
        txt_file = out_dir / f"分析报告_{timestamp}.txt"
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write("=" * 60 + "\n")
            f.write("亚马逊评论AI分析报告\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")

            for r in all_results:
                f.write(f"\n{'=' * 50}\n")
                f.write(f"ASIN: {r['ASIN']}\n")
                f.write(f"{'=' * 50}\n")
                f.write(
                    f"总评论: {r['总评论']} | 好评: {r['好评数']} | 差评: {r['差评数']} | 好评率: {r['好评率']}\n\n")
                f.write(f"【好评卖点】\n{r['好评卖点']}\n\n")
                f.write(f"【差评痛点】\n{r['差评痛点']}\n\n")
                f.write(f"【改进建议】\n{r['改进建议']}\n\n")

        print(f"✅ Excel报告: {excel_file}")
        print(f"✅ 文本报告: {txt_file}")


# 使用
if __name__ == "__main__":
    # 配置
    DEEPSEEK_API_KEY = "sk-c6110db8ead745e5bf1078a63c80a427"  # 替换为你的API Key
    excel_path = r"C:\RPA流程\亚马逊评论分析\flie\亚马逊评论.xlsx"

    analyzer = CommentAnalyzer(excel_path, DEEPSEEK_API_KEY)
    analyzer.run()