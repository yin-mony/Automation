# 纯新品列表创建商品并在线配对主逻辑实现
from pathlib import Path
import time
import argparse
import re

import pandas as pd

EXCEL_FILE_PATH = Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
EXCEL_SHEET_NAME = "新品sku配对自动提醒"


def _extract_xp_code(value):
    text = str(value or "").upper()
    match = re.search(r"XP\d+", text)
    return match.group(0) if match else ""


def read_Excel(excel_file_path=None, sheet_name=EXCEL_SHEET_NAME):
    file_path = Path(excel_file_path) if excel_file_path else EXCEL_FILE_PATH
    if not file_path.exists():
        raise FileNotFoundError(f"Excel 文件不存在: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    required_cols = ["情况", "赛狐新品开发编号", "sku", "ASIN", "人员"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"工作表缺少必要列：{', '.join(missing)}")

    filtered_df = df[df["情况"].astype(str).str.strip() == "未配对"].copy()
    xp_mask = (
        filtered_df["赛狐新品开发编号"]
        .astype(str)
        .str.upper()
        .str.contains(r"XP\d+", regex=True, na=False)
    )
    contains_xp_df = filtered_df[xp_mask].copy()

    records = []
    for index, (_, row) in enumerate(contains_xp_df.iterrows(), start=1):
        request_no = _extract_xp_code(row["赛狐新品开发编号"])
        new_prefix = str(row["sku"]).strip() if pd.notna(row["sku"]) else ""
        new_asin = str(row["ASIN"]).strip() if pd.notna(row["ASIN"]) else ""
        new_name = str(row["人员"]).strip() if pd.notna(row["人员"]) else ""
        if not request_no or not new_prefix or not new_asin or not new_name:
            raise ValueError(f"第 {index} 条记录缺少必要值（赛狐新品开发编号/sku/ASIN/人员）")
        records.append(
            {
                "request_no": request_no,
                "new_prefix": new_prefix,
                "new_name": new_name,
                "new_asin": new_asin,
            }
        )

    grouped_data = {
        "contains_xp_rows": contains_xp_df.values.tolist(),
        "records": records,
    }

    print("\n字典集合（包含 XP，每行独立数组）：")
    print(grouped_data)
    print(f"\n包含 XP 的数据行数量：{len(grouped_data['contains_xp_rows'])}")
    return grouped_data


class DewMainWorkflow:
    def __init__(self, excel_file_path=None, username=None, password=None, sheet_name=EXCEL_SHEET_NAME):
        self.excel_file_path = Path(excel_file_path) if excel_file_path else EXCEL_FILE_PATH
        self.sheet_name = sheet_name
        self.username = username
        self.password = password
        self.page = None
        self.success_count = 0
        self.failed_count = 0

    def _wait_online_pair_button(self, asin, timeout=12):
        """
        在线产品列表在筛选后有异步渲染延迟，避免 3 秒内误判“已配对”。
        """
        pair_xpath = f'x://tr[.//text()[contains(., "{asin}")]]//span[contains(text(), "配对")]'
        end_time = time.time() + max(timeout, 1)
        pair_btn = None
        while time.time() < end_time and not pair_btn:
            pair_btn = self.page.ele(pair_xpath, timeout=1)
            if pair_btn:
                break
            time.sleep(0.5)
        return pair_btn

    def _process_single_record(self, index, total_count, record):
        request_no = record["request_no"]
        new_prefix = record["new_prefix"]
        new_name = record["new_name"]
        new_asin = record["new_asin"]
        print(f"\n[{index}/{total_count}] 处理编号：{request_no}", flush=True)

        self.page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://a[text()="新品开发"]', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
        time.sleep(1)
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
            f"{request_no}\n", clear=True
        )
        time.sleep(1)
        self.page.ele('x://div/ul/li[contains(text(), "生成普通商品")]').click(by_js=True)
        time.sleep(1.5)

        safe_prefix = new_prefix.replace("'", "\\'").replace('"', '\\"')
        self.page.run_js(
            f"""
            const xpath = '//label[contains(text(), "品名")]/following::div/input';
            const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            const inputElement = result.singleNodeValue;

            if (inputElement) {{
                inputElement.value = '{safe_prefix}' + inputElement.value;
                inputElement.dispatchEvent(new Event('input', {{ bubbles: true }}));
            }}
            """
        )
        time.sleep(0.2)

        self.page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).clear()
        time.sleep(0.2)
        self.page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).input(new_prefix)
        time.sleep(0.2)

        self.page.ele('x://label[contains(text(), "查看人")]/following::div/input', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"]', timeout=8).input(
            f"{new_name}\n", clear=True
        )
        time.sleep(1)
        self.page.ele(f'x://div[@class="select-menu"]/div[2]//span[text()="{new_name}"]', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
        time.sleep(1)

        # self.page.ele('x://label[contains(text(), "商品备注")]/following::div/textarea', timeout=8).input(
        #     "自动化程序测试备注，请忽略不要作为正式商品", clear=True
        # )
        time.sleep(1)
        self.page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="保存"])[last()]', timeout=8).click()
        time.sleep(1)

        if self.page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
            self.page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="取消"])[last()]', timeout=8).click()
            time.sleep(1)
            print("商品SKU已存在，取消保存")
        else:
            print("商品SKU不存在，保存商品")

        self.page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://a[text()="在线产品"]', timeout=8).click()
        time.sleep(1)
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
            f"{new_asin}\n", clear=True
        )
        time.sleep(1)

        pair_btn = self._wait_online_pair_button(new_asin, timeout=12)
        if not pair_btn:
            print(f"[{index}/{total_count}] 未检测到“配对”按钮（可能已配对或 ASIN 查询结果未加载完成），跳过。", flush=True)
            self.success_count += 1
            return

        pair_btn.click()
        time.sleep(1)
        search_input = None
        wait_deadline = time.time() + 12
        while time.time() < wait_deadline and not search_input:
            search_input = self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=1)
            if not search_input:
                search_input = self.page.ele('x://input[@placeholder="搜索内容"]', timeout=1)
            if not search_input:
                time.sleep(0.5)
        if not search_input:
            print(f"[{index}/{total_count}] 等待 12 秒后仍未找到“搜索内容”输入框，跳过当前配对。", flush=True)
            return
        search_input.input(f"{new_prefix}\n", clear=True)
        time.sleep(1)
        confirm_pair_btn = None
        confirm_deadline = time.time() + 8
        while time.time() < confirm_deadline and not confirm_pair_btn:
            confirm_pair_btn = self.page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=1)
            if not confirm_pair_btn:
                time.sleep(0.5)
        if not confirm_pair_btn:
            print(f"[{index}/{total_count}] 等待 8 秒后仍未找到最终“配对”按钮，跳过当前配对。", flush=True)
            return
        confirm_pair_btn.click()
        time.sleep(1)
        self.success_count += 1

    def run(self, page):
        self.page = page
        if self.page is None:
            raise RuntimeError("DewMainWorkflow.run() 需要传入已登录的 page 对象。")
        grouped_result = read_Excel(excel_file_path=self.excel_file_path, sheet_name=self.sheet_name)
        records = grouped_result.get("records", [])
        if not records:
            raise RuntimeError("没有可执行数据：未筛选到包含 XP 的未配对记录。")

        print(f"\n统计：包含XP={len(records)}，将按该数量循环执行。")

        for index, record in enumerate(records, start=1):
            try:
                self._process_single_record(index, len(records), record)
            except Exception as exc:
                self.failed_count += 1
                print(f"[{index}/{len(records)}] 处理失败，跳过当前记录: {exc}", flush=True)
                continue

        print(
            f"流程执行完成，总数: {len(records)}，成功: {self.success_count}，失败: {self.failed_count}",
            flush=True,
        )
        return self.page


def main():
    from OnlyMain import SaiHuMain

    parser = argparse.ArgumentParser(description="在线产品配对SKU主流程")
    parser.add_argument("--excel", dest="excel_path", default=None, help="Excel 文件路径")
    parser.add_argument("--sheet", dest="sheet_name", default=EXCEL_SHEET_NAME, help="Excel 工作表名称")
    parser.add_argument("--username", dest="username", default=None, help="赛狐登录账号")
    parser.add_argument("--password", dest="password", default=None, help="赛狐登录密码")
    args = parser.parse_args()
    workflow = DewMainWorkflow(
        excel_file_path=args.excel_path,
        username=args.username,
        password=args.password,
        sheet_name=args.sheet_name,
    )
    page = SaiHuMain(username=args.username, password=args.password).login()
    workflow.run(page=page)


if __name__ == "__main__":
    main()











