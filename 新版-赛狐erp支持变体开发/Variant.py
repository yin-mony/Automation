import os
import time
from pathlib import Path
import pandas as pd
from SaihuERPLogin import SaiHuERPLogin


from DrissionPage import ChromiumPage

# 工作流程主逻辑
# 模式二
# 低横向变体工作表（Excel文件）直接在商品列表创建SKU创建商品并在线配对

class VariantPage:
    def __init__(self,page,username,password,excel_path):
        self.page = page
        self.username = username
        self.password = password
        self.excel_path = Path(excel_path) if excel_path else None

    def excel_file(self, path):
        if path and Path(path).exists():
            paths = Path(path)
        else:
            paths = Path(r"C:\Users\admin\Desktop\新品sku配对+横向变体配对自动提醒.xlsx")
            print(f"使用默认路径: {paths}")
        # 读取 Excel 文件
        df = pd.read_excel(paths, sheet_name='横向变体')  # 指定工作表
        result = df[df["情况"] == '未配对']
        # 提取需要的列值
        selected_data = result[[
            "sku", "ASIN",
            "FNSKU",
            "包装-长（cm）",
            "包装-宽（cm）",
            "包装-高（cm）",
            "包装-重量（g）",
            "不含税成本价格",
            "人员",
            "情况"
        ]].astype(str)
        # 输出对应的行数据
        print(selected_data.to_string(index=False))
        dict_list = selected_data.to_dict('records')
        print(f"总共需要处理 {len(dict_list)} 条数据")

        data = []
        for idx, item in enumerate(dict_list, 1):
            # 定义常量
            # 品名与sku一致
            pm = item["sku"]
            sku = item["sku"]
            asin = item["ASIN"]
            chang = item["包装-长（cm）"]
            kuan = item["包装-宽（cm）"]
            gao = item["包装-高（cm）"]
            liang = item["包装-重量（g）"]
            price = item["不含税成本价格"]
            name = item["人员"]

            data.append({
                "品名": pm,
                "SKU": sku,
                "ASIN": asin,
                "包装-长（cm）": chang,
                "包装-宽（cm）": kuan,
                "包装-高（cm）": gao,
                "包装-重量（g）": liang,
                "不含税成本价格": price,
                "人员": name,
            })
        return data

    def main(self):
        page = self.page
        data = self.excel_file(self.excel_path)
        login = SaiHuERPLogin(
            page=page,
            username=self.username,
            password=self.password,
            img_path=Path(__file__).resolve().parent
        )
        login.login()
        print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)
        for idx, item in enumerate(data, 1):
            print(f"\n处理第 {idx}/{len(data)} 条: {item['SKU']}")
            try:
                # 开始流程
                page.ele('x://div/ul/li/span[text()="商品"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://a[text()="商品列表"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://button//span[text()="添加商品"]', timeout=5).click()
                time.sleep(1)
                page.ele('x://span[text()="添加单个商品"]', timeout=3).click()
                time.sleep(1)
                # 品名
                data_pm = item['品名']
                page.ele('x://label[contains(text(), "品名")]/following::div/input[1]', timeout=3).input(f"{data_pm}")
                time.sleep(1)
                # SKU
                data_sku = item["SKU"]
                page.ele('x://label[contains(text(), "SKU")]/following::div/input[1]', timeout=3).input(f"{data_sku}")
                time.sleep(1)
                # 查看人 对应列表文件里的 人员
                data_name = item["人员"]
                page.ele('x://label[text()="查看人："]/following::div[@placeholder="请选择"]', timeout=5).click()
                time.sleep(1)
                page.ele('x:(//div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"])',
                         timeout=5).input(
                    f"{data_name}\n", clear=True
                )
                page.ele(f'x://span[@title="{data_name}"]').click(by_js=True)
                time.sleep(1)
                page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
                time.sleep(1)
                print("继续填写，切换到采购信息页面")
                page.ele('x://div[normalize-space()="采购信息"]').click()
                time.sleep(1)
                # 采购成本 对应列表文件里的 不含税成本价格
                data_price = item["不含税成本价格"]
                page.ele('x:(//span[text()="采购成本"]/following::div/input)[1]').input(f"{data_price}")
                time.sleep(1)
                print("继续填写，切换到规格信息页面")
                page.ele('x://div[normalize-space()="规格信息"]').click()
                time.sleep(1)
                data_chang = item["包装-长（cm）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="长"]').input(f"{data_chang}")
                time.sleep(1)
                data_kuan = item["包装-宽（cm）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="宽"]').input(f"{data_kuan}")
                time.sleep(1)
                data_gao = item["包装-高（cm）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="高"]').input(f"{data_gao}")
                time.sleep(1)
                data_liang = item["包装-重量（g）"]
                page.ele('x://span[text()="商品规格"]/following::div/input[@debounce="100"]').input(f"{data_liang}")
                time.sleep(1)

                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="长"])[2]', timeout=8).input(
                    f"{data_chang}")
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="宽"])[2]').input(
                    f"{data_kuan}")
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装规格"]/following::div//input[@placeholder="高"])[2]').input(
                    f"{data_gao}")
                time.sleep(1)
                page.ele('x:(//span[text()="商品包装重量"]/following::td//input)[10]').input(f"{data_liang}")
                # 单次点击(等待并检查是否有弹窗)
                page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="保存"])[last()]', timeout=3).click()
                time.sleep(3)
                if page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
                    page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="取消"])[last()]', timeout=8).click()
                    time.sleep(1)
                    print("商品SKU已存在，取消保存")
                else:
                    print("商品SKU不存在，保存商品")

                # 销售页面
                page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://a[text()="在线产品"]', timeout=8).click()
                time.sleep(1)
                # ASIN
                data_asin = item["ASIN"]
                # page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                    f"{data_asin}\n", clear=True
                )
                time.sleep(1)
                # 查找确认该ASIN是否存在配对按钮
                button = page.ele(f'x://tr[.//text()[contains(., "{data_asin}")]]//span[contains(text(), "配对")]',
                                  timeout=10)
                if not button:
                    print("该商品已配对过ASIN")
                    time.sleep(1)
                    continue
                button.click(by_js=True)
                print("已点击配对ASIN，进入配对详情列表")
                # 输入对应的SKU查找
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=3).input(f"{data_sku}\n",
                                                                                                       clear=True)
                time.sleep(1)
                print("已查找到最新创建且对应的sku")
                page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=3).click()
                print("已点击最终配对ASIN")
                time.sleep(1)
            except Exception as e:
                print(f"处理第 {idx} 条记录时出错: {e}")
                print(f"出错的数据: SKU={item.get('SKU')}, ASIN={item.get('ASIN')}")
                # 可以选择继续处理下一条或中断
                continue
        print(f"\n所有数据处理完成！共处理 {len(data)} 条记录")


if __name__ == '__main__':
    page = ChromiumPage()
    run = VariantPage(
        page=page,
        username='zidonghua',
        password='',
        excel_path=r"C:\Users\admin\Desktop\新品sku配对+横向变体配对自动提醒.xlsx"
    )
    run.main()




