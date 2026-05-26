from pathlib import Path
import time
import argparse

import pandas as pd
from pandas import Timestamp

EXCEL_FILE_PATH = Path(r"C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx")


def read_Excel(file_path):
    if not file_path.exists():
        raise FileNotFoundError(f"未找到表格文件: {file_path}")

    try:
        raw_df = pd.read_excel(file_path, header=None)
    except Exception as exc:
        print(f"pandas 读取失败，尝试使用 Excel COM 回退读取: {exc}")
        try:
            import win32com.client  # type: ignore
        except ImportError as import_exc:
            raise RuntimeError(
                "回退读取需要 pywin32（win32com），请先安装: pip install pywin32"
            ) from import_exc

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(
                str(file_path.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=1,
            )
            worksheet = workbook.Worksheets(1)
            used_range = worksheet.UsedRange.Value
            if not used_range:
                raw_df = pd.DataFrame()
            else:
                rows = list(used_range)
                data = [list(row) if isinstance(row, tuple) else [row] for row in rows]
                raw_df = pd.DataFrame(data)
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()

    if raw_df.empty:
        print("表格无可用数据")
        return {"最新日期": None, "记录总量": 0, "数据": []}

    header_idx = None
    for idx in raw_df.index:
        row_values = raw_df.loc[idx].fillna("").astype(str).str.strip().tolist()
        if "时间" in row_values and "SKU" in row_values:
            header_idx = idx
            break

    if header_idx is None:
        raise KeyError("未识别到包含“SKU”和“时间”的表头行")

    raw_header = raw_df.loc[header_idx].fillna("").astype(str).str.strip().tolist()
    df = raw_df.loc[header_idx + 1 :].copy()
    df.columns = raw_header

    non_empty_cols = [col for col in df.columns if str(col).strip()]
    df = df[non_empty_cols]
    col_counter = {}
    unique_cols = []
    for col in df.columns:
        name = str(col).strip()
        count = col_counter.get(name, 0) + 1
        col_counter[name] = count
        unique_cols.append(name if count == 1 else f"{name}_{count}")
    df.columns = unique_cols
    df = df.dropna(how="all")

    if "时间" not in df.columns:
        raise KeyError("表格中不存在“时间”列")
    if "备注" not in df.columns:
        df["备注"] = pd.NA

    df = df.dropna(how="all")
    if "SKU" in df.columns:
        df = df[df["SKU"].astype(str).str.strip().ne("SKU")]
    df = df[df["时间"].astype(str).str.strip().ne("时间")]

    if "序号" in df.columns:
        seq_series = pd.to_numeric(df["序号"], errors="coerce")
        df = df[seq_series.notna()]
    elif "SKU" in df.columns:
        sku_series = df["SKU"].astype(str).str.strip()
        df = df[sku_series.ne("") & sku_series.ne("nan")]

    time_series = df["时间"].copy()
    time_series = time_series.replace(r"^\s*$", pd.NA, regex=True).ffill()

    df["时间"] = pd.to_datetime(time_series, errors="coerce", utc=True)
    valid_df = df.dropna(subset=["时间"]).copy()
    if valid_df.empty:
        print("“时间”列没有可用的有效时间数据")
        return {"最新日期": None, "记录总量": 0, "数据": []}

    valid_df["时间"] = valid_df["时间"].dt.tz_convert("Asia/Shanghai").dt.tz_localize(None)
    latest_date = valid_df["时间"].dt.date.max()
    latest_rows = valid_df[valid_df["时间"].dt.date == latest_date]
    structured_rows = []

    print(f"最新日期: {latest_date}，共 {len(latest_rows)} 条")
    print("-" * 40)

    for row_index, (_, row) in enumerate(latest_rows.iterrows(), start=1):
        row_data = {}
        print(f"第 {row_index} 条数据:")
        for col in latest_rows.columns:
            value = row[col]
            if col == "时间" and pd.notna(value):
                if isinstance(value, Timestamp):
                    value = value.strftime("%Y-%m-%d")
                else:
                    value = str(value)[:10]
            if col == "备注":
                is_empty_remark = pd.isna(value) or str(value).strip() == ""
                if is_empty_remark:
                    value = "（空备注）"
                row_data["备注是否为空"] = is_empty_remark
            row_data[col] = value
            print(f"{col}: {value}")
        structured_rows.append(row_data)
        print("-" * 40)

    summary = {
        "最新日期": str(latest_date),
        "记录总量": len(structured_rows),
        "数据": structured_rows,
    }
    print(f"统计总量数据: {summary['记录总量']} 条")
    return summary


class LowMainWorkflow:
    def __init__(self, excel_file_path=None, username=None, password=None):
        self.excel_file_path = Path(excel_file_path) if excel_file_path else EXCEL_FILE_PATH
        self.username = username
        self.password = password
        self.page = None
        self.success_count = 0
        self.failed_count = 0

    def _click_tab(self, tab_name):
        tab_locators = [
            f'x://div[normalize-space()="{tab_name}"]',
            f'x://span[normalize-space()="{tab_name}"]',
            f'x://*[@role="tab" and normalize-space()="{tab_name}"]',
        ]
        for locator in tab_locators:
            ele = self.page.ele(locator, timeout=2)
            if ele:
                ele.click(by_js=True)
                return
        raise RuntimeError(f"未找到可点击的“{tab_name}”标签")

    def _close_popups(self):
        for _ in range(3):
            next_btn = self.page.ele('x://button/span[contains(text(),"下一条")]', timeout=1.2)
            if not next_btn:
                break
            try:
                next_btn.click()
                time.sleep(0.2)
            except Exception:
                break

        close_btn = self.page.ele('x://button[2]/span[contains(text(),"关闭")]', timeout=1.2)
        if close_btn:
            try:
                close_btn.click()
                time.sleep(0.2)
            except Exception:
                pass

    def _process_single_record(self, idx, total_count, row_data):
        new_pingming = str(row_data.get("品名", "")).strip()
        new_sku = str(row_data.get("SKU", "")).strip()
        new_name = str(row_data.get("负责人", "")).strip()
        new_cost = str(row_data.get("采购价（元）", "")).strip()
        new_asin = str(row_data.get("ASIN", "")).strip()
        new_length = str(row_data.get("长 包装规格（cm）", "")).strip()
        new_width = str(row_data.get("宽 包装规格（cm)", "")).strip()
        new_height = str(row_data.get("高 包装规格（cm）", "")).strip()
        new_weight = str(row_data.get("单品毛重（kg）", "")).strip()

        print(f"[{idx}/{total_count}] 开始处理 SKU: {new_sku}", flush=True)

        self.page.ele('x://div/ul/li/span[text()="商品"]', timeout=5).click()
        time.sleep(1)
        self.page.ele('x://a[text()="商品列表"]', timeout=5).click()
        time.sleep(1)
        self.page.ele('x://button//span[text()="添加商品"]', timeout=5).click()
        time.sleep(1)
        self.page.ele('x://span[text()="添加单个商品"]', timeout=3).click()
        time.sleep(1)

        self.page.ele('x://label[contains(text(), "品名")]/following::div/input[1]', timeout=3).input(f"{new_pingming}")
        time.sleep(1)

        self.page.ele('x://label[contains(text(), "SKU")]/following::div/input[1]', timeout=3).input(f"{new_sku}")
        time.sleep(1)

        self.page.ele('x://label[text()="查看人："]/following::div[@placeholder="请选择"]', timeout=5).click()
        time.sleep(0.5)
        self.page.ele('x:(//div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"])', timeout=5).input(
            f"{new_name}\n", clear=True
        )
        self.page.ele(f'x://span[@title="{new_name}"]').click(by_js=True)
        time.sleep(0.5)
        self.page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
        time.sleep(0.5)

        time.sleep(1)
        print(f"[{idx}/{total_count}] 切换到采购信息", flush=True)
        self._click_tab("采购信息")
        time.sleep(0.5)
        self.page.ele('x:(//span[text()="采购成本"]/following::div/input)[1]').input(f"{new_cost}")
        time.sleep(0.5)
        print(f"[{idx}/{total_count}] 切换到规格信息", flush=True)
        self._click_tab("规格信息")
        time.sleep(1)

        self.page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="长"]').input(f"{new_length}")
        time.sleep(0.5)
        self.page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="宽"]').input(f"{new_width}")
        time.sleep(0.5)
        self.page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="高"]').input(f"{new_height}")
        time.sleep(0.5)
        self.page.ele('x://span[text()="商品规格"]/following::div/input[@debounce="100"]').input(f"{new_weight}")
        time.sleep(0.5)

        self.page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="长"])[2]', timeout=8).input(f"{new_length}")
        time.sleep(0.5)
        self.page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="宽"])[2]').input(f"{new_width}")
        time.sleep(0.5)
        self.page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="高"])[2]').input(f"{new_height}")
        time.sleep(0.5)
        self.page.ele('x:(//span[text()="商品包装重量"]/following::td//input)[10]').input(f"{new_weight}")

        self.page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="保存"])[last()]', timeout=3).click()
        time.sleep(1)

        if self.page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=5):
            self.page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="取消"])[last()]', timeout=3).click()
            time.sleep(1)
            print("商品SKU已存在，取消保存")

        self.page.ele('x://div/ul/li/span[text()="销售"]', timeout=3).click()
        time.sleep(1)
        self.page.ele('x://a[text()="在线产品"]', timeout=3).click()
        time.sleep(1)
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=3).click()
        self.page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=3).input(
            f"{new_asin}\n", clear=True
        )
        time.sleep(1)

        if not self.page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=3):
            print("配对按钮不存在，跳过配对(该商品已配对过)")
            self.success_count += 1
            return

        self.page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=3).click()
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
            print(f"[{idx}/{total_count}] 等待 12 秒后仍未找到“搜索内容”输入框，跳过当前配对。", flush=True)
            return
        search_input.input(f"{new_sku}\n", clear=True)
        time.sleep(1)
        self.success_count += 1

    def run(self, page, skip_close_popups=False):
        self.page = page
        if self.page is None:
            raise RuntimeError("LowMainWorkflow.run() 需要传入已登录的 page 对象。")
        summary = read_Excel(self.excel_file_path)
        records = summary.get("数据", [])
        total_count = summary.get("记录总量", 0)
        if not records:
            print("没有可执行数据，结束页面流程。", flush=True)
            return self.page

        print(f"开始执行页面流程，总次数: {total_count}", flush=True)
        if not skip_close_popups:
            self._close_popups()
        else:
            print("检测到复用登录态，跳过关闭引导弹窗步骤。", flush=True)

        for idx, row_data in enumerate(records, start=1):
            try:
                self._process_single_record(idx, total_count, row_data)
            except Exception as exc:
                self.failed_count += 1
                print(f"[{idx}/{total_count}] 处理失败，跳过当前记录: {exc}", flush=True)
                continue

        print(f"流程执行完成，总数: {total_count}，成功: {self.success_count}，失败: {self.failed_count}", flush=True)
        return self.page


def main():
    from OnlyMain import SaiHuMain

    parser = argparse.ArgumentParser(description="低价商品列表创建商品主流程")
    parser.add_argument("--excel", dest="excel_path", default=None, help="Excel 文件路径")
    parser.add_argument("--username", dest="username", default=None, help="赛狐登录账号")
    parser.add_argument("--password", dest="password", default=None, help="赛狐登录密码")
    args = parser.parse_args()
    workflow = LowMainWorkflow(
        excel_file_path=args.excel_path,
        username=args.username,
        password=args.password,
    )
    login_client = SaiHuMain(username=args.username, password=args.password)
    page = login_client.login()
    workflow.run(page=page, skip_close_popups=login_client.reused_session)


if __name__ == "__main__":
    main()