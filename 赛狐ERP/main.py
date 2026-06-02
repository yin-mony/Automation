from pathlib import Path
from DrissionPage import ChromiumPage,Chromium
# from test import SaihuERPLogin
from login import SaihuERPLogin
import time
import pandas as pd
import re
import argparse


# 开始流程

# 模式一
# 纯新品列表（Excel文件）创建商品并在线配对
def new_set_pairing(username, password,path):
    page = ChromiumPage()
    EXCEL_FILE_PATH = Path(path) if path else Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
    # 读取 Excel 文件
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='新品sku配对自动提醒')  # 指定工作表
    login = SaihuERPLogin(page, username=username, password=password,img_dir=r"C:\Users\admin\Desktop")
    login.login()
    print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)

    # 定义正则（XP+数字）
    re_number = r'XP\d+'
    result =df[(df["情况"] == '未配对') & (df['赛狐新品开发编号'].str.contains(re_number, na=False, regex=True))]
    print(f"找到 {len(result)} 条需要处理的记录")
    print(result)

    dict_list = result.to_dict('records')
    print(f"总共需要处理 {len(dict_list)} 条数据")
    # 提取sku,ASIN,赛狐新品开发编号,人员
    for idx,item in enumerate(dict_list,1):
        print(f"\n{'=' * 50}")
        print(f"正在处理第 {idx}/{len(dict_list)} 条记录")
        print(f"SKU: {item['sku']}, ASIN: {item['ASIN']}, 开发编号: {item['赛狐新品开发编号']}, 人员: {item['人员']}")
        print(f"{'=' * 50}\n")
        try:
            # 新品编号
            new_fnsku = item['赛狐新品开发编号']
            # sku
            new_sku = item['sku']
            # 查看人（人员）
            new_name = item['人员']
            # ASIN
            new_asin = item['ASIN']

            # 商品-新品开发页面
            page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click()
            time.sleep(1)
            page.ele('x://a[text()="新品开发"]', timeout=5).click()
            time.sleep(1)
            # page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
            # time.sleep(1)
            page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                f"{new_fnsku}\n", clear=True
            )
            time.sleep(1)
            # 判断是否存在"生成普通商品"按钮
            sp_page = page.ele('x://div/ul/li[contains(text(), "生成普通商品")]')
            if sp_page:
                page.ele('x://div/ul/li[contains(text(), "生成普通商品")]').click(by_js=True)
                print("查到“生成普通商品”按钮，当前新品开发编号，需要进行'生成普通商品'操作")
                time.sleep(1.5)
                # sku
                # new_sku = item['sku']
                # 品名（与sku同名）
                td_name = new_sku.replace("'", "\\'").replace('"', '\\"')
                page.run_js(
                    f"""
                                        const xpath = '//label[contains(text(), "品名")]/following::div/input';
                                        const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                                        const inputElement = result.singleNodeValue;

                                        if (inputElement) {{
                                            inputElement.value = '{td_name}' + inputElement.value;
                                            inputElement.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                        }}
                                        """
                )
                time.sleep(1)
                page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).input(f'{new_sku}\n',
                                                                                                     clear=True)
                time.sleep(1)
                # 查看人（人员）
                # new_name = item['人员']
                page.ele('x://label[contains(text(), "查看人")]/following::div/input', timeout=8).click()
                time.sleep(1)
                page.ele('x://div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"]',
                         timeout=8).input(
                    f"{new_name}\n", clear=True
                )
                time.sleep(1)
                page.ele(f'x://div[@class="select-menu"]/div[2]//span[text()="{new_name}"]', timeout=8).click(
                    by_js=True)
                time.sleep(1)
                page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
                time.sleep(1)
                # 备注输入框内容
                # page.ele('x://label[contains(text(), "商品备注")]/following::div/textarea', timeout=8).input(
                #     "自动化程序测试备注，请忽略不要作为正式商品", clear=True)
                # time.sleep(1)
                # 单次点击(等待并检查是否有弹窗)
                page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="保存"])[last()]', timeout=8).click()
                time.sleep(1)
                if page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
                    page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="取消"])[last()]',
                             timeout=8).click()
                    time.sleep(1)
                    print("商品SKU已存在，取消保存")
                else:
                    print("商品SKU不存在，保存商品")
            else:
                # print("当前新品开发编号，还未完成相关流程，无法进行'生成普通商品'操作")
                print("当前新品开发编号，还未完成相关流程，无法进行后续操作")
                continue

            # 销售页面
            page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
            time.sleep(1)
            page.ele('x://a[text()="在线产品"]', timeout=8).click()
            time.sleep(1)
            # ASIN
            # new_asin = item['ASIN']
            # page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
            page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                f"{new_asin}\n", clear=True
            )
            time.sleep(1)
            # 查找确认该ASIN是否存在配对按钮
            button = page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=10)
            if not button:
                print("该商品已配对过ASIN")
                time.sleep(1)
                continue
            button.click(by_js=True)
            print("已点击配对ASIN，进入配对详情列表")
            # 输入对应的SKU查找
            page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=3).input(f"{new_sku}\n", clear=True)
            time.sleep(1)
            print("已查找到最新创建且对应的sku")
            page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=3).click()
            print("已点击最终配对ASIN")
            time.sleep(1)
        except Exception as e:
            print(f"处理第 {idx} 条记录时出错: {e}")
            print(f"出错的数据: SKU={item.get('sku')}, ASIN={item.get('ASIN')}")
            # 可以选择继续处理下一条或中断
            continue
    print(f"\n所有数据处理完成！共处理 {len(dict_list)} 条记录")


# 模式二
# 低价商城列表（Excel文件）直接在商品列表创建SKU创建商品并在线配对
def low_price_pairing(username, password,path):
    page = ChromiumPage()
    login = SaihuERPLogin(page, username=username, password=password,img_dir=r"C:\Users\admin\Desktop")
    login.login()
    print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)
    EXCEL_FILE_PATH = Path(path) if path else Path(r"C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx")
    # 读取 Excel 文件
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='工作表1')  # 指定工作表
    # 转成 Pandas 日期对象
    target_date = pd.to_datetime("2026/5/22")
    # 筛选出时间等于 target_date 的行
    result = df[df["时间"] == target_date]
    # print(df.columns.tolist())
    # 提取需要的列值
    selected_data  = result[[
        "品名", "SKU", "ASIN",
        "长 包装规格（cm）",
        "宽 包装规格（cm)",
        "高 包装规格（cm）",
        "单品毛重（kg）",
        "采购价（元）",
        "负责人",
        "时间"
    ]].astype(str)
    # 输出对应的行数据
    print(selected_data.to_string(index=False))
    dict_list = selected_data.to_dict('records')
    print(f"总共需要处理 {len(dict_list)} 条数据")
    for idx, item in enumerate(dict_list, 1):
        print(f"\n{'=' * 50}")
        print(f"正在处理第 {idx}/{len(dict_list)} 条记录")
        print(f"SKU: {item["SKU"]}, ASIN: {item['ASIN']}, 负责人: {item["负责人"]}")
        print(f"{'=' * 50}\n")

        # 定义常量
        data_pm = item["品名"]
        data_sku = item["SKU"]
        data_asin = item["ASIN"]
        data_chang = item["长 包装规格（cm）"]
        data_kuan = item["宽 包装规格（cm)"]
        data_gao = item["高 包装规格（cm）"]
        data_liang = item["单品毛重（kg）"]
        data_price = item["采购价（元）"]
        data_name = item["负责人"]
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
            page.ele('x://label[contains(text(), "品名")]/following::div/input[1]', timeout=3).input(f"{data_pm}")
            time.sleep(1)
            # SKU
            page.ele('x://label[contains(text(), "SKU")]/following::div/input[1]', timeout=3).input(f"{data_sku}")
            time.sleep(1)
            # 查看人 对应列表文件里的 负责人
            page.ele('x://label[text()="查看人："]/following::div[@placeholder="请选择"]', timeout=5).click()
            time.sleep(1)
            page.ele('x:(//div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"])',timeout=5).input(
                f"{data_name}\n", clear=True
            )
            page.ele(f'x://span[@title="{data_name}"]').click(by_js=True)
            time.sleep(1)
            page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
            time.sleep(1)
            print("继续填写，切换到采购信息页面")
            page.ele('x://div[normalize-space()="采购信息"]').click()
            time.sleep(1)
            # 采购成本 对应列表文件里的 采购价（元）
            page.ele('x:(//span[text()="采购成本"]/following::div/input)[1]').input(f"{data_price}")
            time.sleep(1)
            print("继续填写，切换到规格信息页面")
            page.ele('x://div[normalize-space()="规格信息"]').click()
            time.sleep(1)
            page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="长"]').input(f"{data_chang}")
            time.sleep(1)
            page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="宽"]').input(f"{data_kuan}")
            time.sleep(1)
            page.ele('x://span[text()="商品规格"]/following::div/input[@placeholder="高"]').input(f"{data_gao}")
            time.sleep(1)
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
                page.ele('x:(//div[@aria-label="添加普通商品"]//button/span[text()="取消"])[last()]',timeout=8).click()
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
            # new_asin = item['ASIN']
            # page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
            page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(
                f"{data_asin}\n", clear=True
            )
            time.sleep(1)
            # 查找确认该ASIN是否存在配对按钮
            button = page.ele(f'x://tr[.//text()[contains(., "{data_asin}")]]//span[contains(text(), "配对")]', timeout=10)
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
    print(f"\n所有数据处理完成！共处理 {len(dict_list)} 条记录")













# 主流程
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="赛狐ERP主流程入口")
    parser.add_argument(
        "--mode",
        choices=["mode1", "mode2"],
        required=True,
        help="运行模式：mode1=纯新品，mode2=低价商城",
    )
    parser.add_argument("--username", required=True, help="赛狐账号")
    parser.add_argument("--password", default="", help="赛狐密码")
    parser.add_argument("--path", default="", help="Excel 文件路径")
    args = parser.parse_args()

    if args.mode == "mode1":
        new_set_pairing(username=args.username, password=args.password, path=args.path)
    else:
        low_price_pairing(username=args.username, password=args.password, path=args.path)