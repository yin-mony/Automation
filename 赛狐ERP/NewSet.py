import time
import json
import os
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin


# 模式一：纯新品列表（Excel文件）创建商品并在线配对
class NewSetPage:
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
            paths = Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
            print(f"使用默认路径: {paths}")

        df = pd.read_excel(paths, sheet_name='新品sku配对自动提醒')
        required_columns = ["情况", "赛狐新品开发编号", "sku", "ASIN", "人员"]
        missing_columns = [column for column in required_columns if column not in df.columns]
        if missing_columns:
            raise RuntimeError(f"工作计划表缺少必要列: {', '.join(missing_columns)}")

        re_number = r'XP\d+'
        result = df[
            (df["情况"] == '未配对')
            & (df['赛狐新品开发编号'].astype(str).str.contains(re_number, na=False, regex=True))
        ]
        print(f"找到 {len(result)} 条需要处理的记录")
        print(result)
        dict_list = result.to_dict('records')
        print(f"总共需要处理 {len(dict_list)} 条数据")

        data = []
        for idx, item in enumerate(dict_list, 1):
            print(f"\n{'=' * 50}")
            print(f"正在处理第 {idx}/{len(dict_list)} 条记录")
            print(
                f"SKU: {item['sku']}, ASIN: {item['ASIN']}, 开发编号: {item['赛狐新品开发编号']}, 人员: {item['人员']}"
            )
            print(f"{'=' * 50}\n")
            new_fnsku = item['赛狐新品开发编号']
            new_sku = item['sku']
            new_name = item['人员']
            new_asin = item['ASIN']

            data.append({
                '新品开发编号': new_fnsku,
                'sku': new_sku,
                'asin': new_asin,
                '开发编号': new_fnsku,
                '人员': new_name,
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
                page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://a[text()="新品开发"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                    f"{item['新品开发编号']}\n", clear=True
                )
                time.sleep(1)

                sp_page = page.ele('x://div/ul/li[contains(text(), "生成普通商品")]', timeout=8)
                if sp_page:
                    page.ele('x://div/ul/li[contains(text(), "生成普通商品")]').click(by_js=True)
                    print("查到“生成普通商品”按钮，当前新品开发编号，需要进行'生成普通商品'操作")
                    time.sleep(1.5)

                    new_sku = item['sku']
                    td_name = (
                        str(new_sku)
                        .replace("\r", " ")
                        .replace("\n", " ")
                        .replace("\u2028", " ")
                        .replace("\u2029", " ")
                    )
                    td_name_js = json.dumps(td_name)
                    page.run_js(
                        f"""
                        const xpath = '//label[contains(text(), "品名")]/following::div/input';
                        const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                        const inputElement = result.singleNodeValue;

                        if (inputElement) {{
                            inputElement.value = {td_name_js} + inputElement.value;
                            inputElement.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        }}
                        """
                    )
                    time.sleep(1)
                    page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).input(
                        f'{new_sku}\n', clear=True
                    )
                    time.sleep(1)

                    new_name = item['人员']
                    page.ele('x://label[contains(text(), "查看人")]/following::div/input', timeout=8).click()
                    time.sleep(1)
                    page.ele('x://div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"]',
                             timeout=8).input(f"{new_name}\n", clear=True)
                    time.sleep(1)
                    page.ele(f'x://div[@class="select-menu"]/div[2]//span[text()="{new_name}"]', timeout=8).click(
                        by_js=True
                    )
                    time.sleep(1)
                    page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]',
                             timeout=8).click()
                    time.sleep(1)

                    page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="保存"])[last()]',
                             timeout=8).click()
                    time.sleep(1)
                    if page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
                        page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="取消"])[last()]',
                                 timeout=8).click()
                        time.sleep(1)
                        print("商品SKU已存在，取消保存")
                    else:
                        print("商品SKU不存在，保存商品")
                else:
                    print("当前新品开发编号，还未完成相关流程，无法进行后续操作")
                    continue

                page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://a[text()="在线产品"]', timeout=8).click()
                time.sleep(1)
                new_asin = item['asin']
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                    f"{new_asin}\n", clear=True
                )
                time.sleep(1)
                button = page.ele(
                    f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]',
                    timeout=10,
                )
                if not button:
                    print("该商品已配对过ASIN")
                    time.sleep(1)
                    continue
                button.click(by_js=True)
                print("已点击配对ASIN，进入配对详情列表")
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=3).input(
                    f"{new_sku}\n", clear=True
                )
                time.sleep(1)
                print("已查找到最新创建且对应的sku")
                page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=3).click()
                print("已点击最终配对ASIN")
                time.sleep(1)
            except Exception as exc:
                print(f"处理第 {idx} 条记录时出错: {exc}")
                print(f"出错的数据: SKU={item.get('sku')}, ASIN={item.get('asin')}")
                continue

        print(f"\n所有数据处理完成！共处理 {len(data)} 条记录")


if __name__ == '__main__':
    config = {
        "page": ChromiumPage(),
        "username": os.getenv("SAIHU_USERNAME", ""),
        "password": os.getenv("SAIHU_PASSWORD", ""),
        "excel_path": r"C:\Users\admin\Desktop\工作计划表.xlsx",
        "base_dir": Path(__file__).resolve().parent,
    }
    run = NewSetPage(config)
    run.main()
