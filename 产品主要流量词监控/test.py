import time
from pathlib import Path
import re
import sys
from datetime import datetime

import pandas as pd
import requests


WECHAT_WEBHOOK = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_WEBHOOK_KEY"
MENTION_USERID = "18280194086"


def excel(path):
    path = Path(path)
    if path.is_dir():
        files = []
        for suffix in ("*.xlsx", "*.xls", "*.csv"):
            files.extend(path.glob(suffix))
    else:
        files = [path]
    if not files:
        print(f"当前路径下未找到下载文件: {path}")
        return []
    data = []
    for file_path in files:
        if not file_path.exists():
            print(f"文件不存在: {file_path}")
            continue
        if "自然排名提取结果" in file_path.stem:
            continue
        print(f"正在处理下载文件: {file_path}")
        asin_match = re.search(r"(?<![A-Z0-9])B[A-Z0-9]{9}(?![A-Z0-9])", file_path.stem.upper())
        asin = asin_match.group() if asin_match else ""
        if file_path.suffix.lower() == ".csv":
            df = pd.read_csv(file_path)
        else:
            try:
                df = pd.read_excel(file_path, sheet_name="关键词反查结果")
            except Exception:
                df = pd.read_excel(file_path)

        columns = list(df.columns)
        keyword_index = None
        nature_index = None
        sp_index = None
        time_index = None
        for index, col in enumerate(columns):
            col = str(col).strip()
            if keyword_index is None and "关键词" in col:
                keyword_index = index
            if nature_index is None and col == "自然排名":
                nature_index = index
            if sp_index is None and "SP广告排名" in col:
                sp_index = index
        if keyword_index is None:
            print(f"未找到关键词列: {file_path}")
            continue
        if nature_index is None:
            print(f"未找到自然排名列: {file_path}")
            continue
        end_index = sp_index if sp_index is not None else len(columns)
        for index in range(nature_index + 1, end_index):
            col = str(columns[index]).strip()
            if "抓取时间" in col:
                time_index = index
                break
        if time_index is None:
            print(f"未找到自然排名对应的抓取时间列: {file_path}")
            continue
        result = df.iloc[:, [keyword_index, nature_index, time_index]].copy()
        result.columns = ["关键词", "自然排名", "抓取时间"]
        result = result.dropna(subset=["关键词", "自然排名"])
        result["自然排名"] = result["自然排名"].astype(int)
        result["asin"] = asin
        data.extend(result[["asin", "关键词", "自然排名", "抓取时间"]].to_dict("records"))
    print(data)

    send_data = [item for item in data if item.get("自然排名", 0) > 10]
    print(send_data)

    if not send_data:
        print("没有自然排名大于10的数据，不发送企业微信消息")
        return data

    save_path = path / f"自然排名预警完整数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx" if path.is_dir() else path.parent / f"自然排名预警完整数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    pd.DataFrame(send_data).to_excel(save_path, index=False)
    print(f"完整预警数据已保存: {save_path}")

    asin_map = {}
    for item in send_data:
        asin = item.get("asin", "") or "未识别ASIN"
        asin_map.setdefault(asin, []).append(item)

    for asin, items in asin_map.items():
        content = (
            f"产品关键词监控预警\n"
            f"ASIN：{asin}\n"
            f"自然排名大于10：{len(items)} 条\n"
            f"仅展示前10条，完整数据已保存到本地：{save_path}\n"
        )

        for index, item in enumerate(items[:10], 1):
            content += (
                f"\n{index}. 关键词：{item.get('关键词', '')}\n"
                f"自然排名：{item.get('自然排名', '')}\n"
                f"抓取时间：{item.get('抓取时间', '')}\n"
            )
        session = requests.Session()
        session.trust_env = False
        res = requests.post(
            WECHAT_WEBHOOK,
            json={
                "msgtype": "text",
                "text": {
                    "content": content,
                    "mentioned_mobile_list": [MENTION_USERID]
                }
            },
            timeout=10
        )
        time.sleep(3)
        print(f"{asin} 发送结果:", res.json())

    return data



if __name__ == "__main__":
    if len(sys.argv) > 1:
        excel(sys.argv[1])
    else:
        excel(r"C:\Users\admin\Desktop\产品流量词监控")
