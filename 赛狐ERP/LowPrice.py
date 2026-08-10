import os
import time
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin


# 模式二：低价商城列表（Excel文件）直接在商品列表创建 SKU 并在线配对
class LowPricePage:
    def __init__(self, config):
        self.config = config
        self.page = config["page"]
        self.username = config["username"]
        self.password = config["password"]
        self.excel_path = Path(config["excel_path"]) if config.get("excel_path") else None
        self.base_dir = Path(config.get("base_dir") or Path(__file__).resolve().parent)

    def excel_file(self):
        if self.excel_path and self.excel_path.exists():
            paths = self.excel_path
        else:
            paths = Path(r"C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx")
            print(f"使用默认路径: {paths}")

        try:
            df = pd.read_excel(paths, sheet_name='工作表1')
        except Exception as exc:
            print(f"pandas读取失败，尝试Excel COM回退读取: {exc}")
            try:
                import win32com.client
            except ImportError as import_exc:
                raise RuntimeError("回退读取需要 pywin32（win32com），请先安装: pip install pywin32") from import_exc

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = None
            try:
                workbook = excel.Workbooks.Open(
                    str(paths.resolve()),
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                    CorruptLoad=1,
                )
                worksheet = workbook.Worksheets("工作表1")
                used_range = worksheet.UsedRange.Value
                if not used_range:
                    return []
                rows = list(used_range)
                data = [list(row) if isinstance(row, tuple) else [row] for row in rows]
                raw_df = pd.DataFrame(data)

                header_idx = None
                for idx in raw_df.index:
                    row_values = raw_df.loc[idx].fillna("").astype(str).str.strip().tolist()
                    if "时间" in row_values and "SKU" in row_values:
                        header_idx = idx
                        break

                if header_idx is None:
                    raise RuntimeError("回退读取未识别到包含“时间”和“SKU”的表头行")

                raw_header = raw_df.loc[header_idx].fillna("").astype(str).str.strip().tolist()
                df = raw_df.loc[header_idx + 1:].copy()
                df.columns = raw_header
                df = df.dropna(how="all")
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
                excel.Quit()

        required_columns = [
            "时间",
            "品名",
            "SKU",
            "ASIN",
            "长 包装规格（cm）",
            "宽 包装规格（cm)",
            "高 包装规格（cm）",
            "单品毛重（kg）",
            "采购价（元）",
            "负责人",
        ]
        missing_columns = [column for column in required_columns if column not in df.columns]
        if missing_columns:
            raise RuntimeError(f"低价表缺少必要列: {', '.join(missing_columns)}")

        time_series = df["时间"].astype(str).str.strip()
        time_series = time_series.replace("时间", pd.NA)
        df["时间"] = pd.to_datetime(time_series, errors="coerce", format="mixed", utc=True)
        df = df[df["时间"].notna()]
        if df.empty:
            print("未找到可用的时间数据")
            return []
        df["时间"] = df["时间"].dt.tz_localize(None)
        latest_date = df["时间"].dt.date.max()
        print(f"自动检测到的最新日期: {latest_date}")
        result = df[df["时间"].dt.date == latest_date]

        selected_data = result[required_columns].astype(str)
        print(selected_data.to_string(index=False))
        dict_list = selected_data.to_dict('records')
        print(f"总共需要处理 {len(dict_list)} 条数据")

        data = []
        for item in dict_list:
            data.append({
                "品名": item["品名"],
                "SKU": item["SKU"],
                "ASIN": item["ASIN"],
                "长 包装规格（cm）": item["长 包装规格（cm）"],
                "宽 包装规格（cm)": item["宽 包装规格（cm)"],
                "高 包装规格（cm）": item["高 包装规格（cm）"],
                "单品毛重（kg）": item["单品毛重（kg）"],
                "采购价（元）": item["采购价（元）"],
                "负责人": item["负责人"],
                "sku": item["SKU"],
                "asin": item["ASIN"],
            })
        return data

    def main(self):
        page = self.page
        data = self.excel_file()
        login = SaiHuERPLogin({
            "page": page,
            "username": self.username,
            "password": self.password,
            "img_path": self.base_dir,
        })
        if not login.login():
            raise RuntimeError("赛狐登录失败，请检查账号、密码或验证码识别。")
        print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)

        for idx, item in enumerate(data, 1):
            print(f"\n处理第 {idx}/{len(data)} 条: {item['sku']}")
            try:
                page.ele('x://div/ul/li/span[text()="商品"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://a[text()="商品列表"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://button//span[text()="添加商品"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://span[text()="添加单个商品"]', timeout=3).click()
                time.sleep(1)

                data_pm = item['品名']
                page.ele('x://label[contains(text(), "品名")]/following::div/input[1]', timeout=3).input(f"{data_pm}")
                time.sleep(1)

                data_sku = item["SKU"]
                page.ele('x://label[contains(text(), "SKU")]/following::div/input[1]', timeout=3).input(f"{data_sku}")
                time.sleep(1)

                data_name = item["负责人"]
                page.ele('x://label[text()="查看人："]/following::div[@placeholder="请选择"]', timeout=5).click()
                time.sleep(1)
                page.ele('x:(//div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"])',
                         timeout=5).input(f"{data_name}\n", clear=True)
                page.ele(f'x://span[@title="{data_name}"]').click(by_js=True)
                time.sleep(1)
                page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]',
                         timeout=8).click()
                time.sleep(1)

                print("继续填写，切换到采购信息页面")
                page.ele('x://div[normalize-space()="采购信息"]').click()
                time.sleep(1)
                data_price = item["采购价（元）"]
                page.ele('x:(//span[text()="采购成本"]/following::div/input)[1]').input(f"{data_price}")
                time.sleep(1)

                print("继续填写，切换到规格信息页面")
                page.ele('x://div[normalize-space()="规格信息"]').click()
                time.sleep(1)
                data_chang = item["长 包装规格（cm）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="长"]').input(f"{data_chang}")
                time.sleep(1)
                data_kuan = item["宽 包装规格（cm)"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="宽"]').input(f"{data_kuan}")
                time.sleep(1)
                data_gao = item["高 包装规格（cm）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="高"]').input(f"{data_gao}")
                time.sleep(1)
                data_liang = item["单品毛重（kg）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@debounce="100"]').input(f"{data_liang}")
                time.sleep(1)

                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="长"])[2]',
                         timeout=8).input(f"{data_chang}")
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="宽"])[2]').input(
                    f"{data_kuan}"
                )
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="高"])[2]').input(
                    f"{data_gao}"
                )
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装重量"]/following::td//input)[10]').input(f"{data_liang}")

                page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="保存"])[last()]',
                         timeout=3).click()
                time.sleep(3)
                if page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
                    page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="取消"])[last()]',
                             timeout=8).click()
                    time.sleep(1)
                    print("商品SKU已存在，取消保存")
                else:
                    print("商品SKU不存在，保存商品")

                page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://a[text()="在线产品"]', timeout=8).click()
                time.sleep(1)
                data_asin = item["ASIN"]
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                    f"{data_asin}\n", clear=True
                )
                time.sleep(1)
                button = page.ele(
                    f'x://tr[.//text()[contains(., "{data_asin}")]]//span[contains(text(), "配对")]',
                    timeout=10,
                )
                if not button:
                    print("该商品已配对过ASIN")
                    time.sleep(1)
                    continue
                button.click(by_js=True)
                print("已点击配对ASIN，进入配对详情列表")
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=3).input(
                    f"{data_sku}\n", clear=True
                )
                time.sleep(1)
                print("已查找到最新创建且对应的sku")
                page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=3).click()
                print("已点击最终配对ASIN")
                time.sleep(1)
            except Exception as exc:
                print(f"处理第 {idx} 条记录时出错: {exc}")
                print(f"出错的数据: SKU={item.get('SKU')}, ASIN={item.get('ASIN')}")
                continue

        print(f"\n所有数据处理完成！共处理 {len(data)} 条记录")


if __name__ == '__main__':
    config = {
        "page": ChromiumPage(),
        "username": os.getenv("SAIHU_USERNAME", ""),
        "password": os.getenv("SAIHU_PASSWORD", ""),
        "excel_path": r"C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx",
        "base_dir": Path(__file__).resolve().parent,
    }
    run = LowPricePage(config)
    run.main()
