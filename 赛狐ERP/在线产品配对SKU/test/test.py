import time
from pathlib import Path

import pandas as pd
from SaihuERPLogin import SaihuERPLogin

EXCEL_FILE_PATH = Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
EXCEL_SHEET_NAME = "新品sku配对自动提醒"


def read_Excel():
    """读取固定工作表，筛选“未配对”，提取并结构化存放包含 XP 的数据。"""
    if not EXCEL_FILE_PATH.exists():
        raise FileNotFoundError(f"Excel 文件不存在: {EXCEL_FILE_PATH}")

    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EXCEL_SHEET_NAME)
    if "情况" not in df.columns:
        raise ValueError("工作表中不存在“情况”列，无法筛选“未配对”数据。")
    id_col = "赛狐新品开发编号"
    if id_col not in df.columns:
        raise ValueError(f"工作表中不存在“{id_col}”列，无法进行 XP 编号判断。")

    filtered_df = df[df["情况"].astype(str).str.strip() == "未配对"].copy()

    # 示例：XP2601050001，按“XP + 数字”进行匹配
    xp_mask = (
        filtered_df[id_col]
        .astype(str)
        .str.upper()
        .str.contains(r"XP\d+", regex=True, na=False)
    )
    contains_xp_df = filtered_df[xp_mask].copy()

    # 结构化存放：外部字典，内部数组按行独立存放
    grouped_data = {
        "contains_xp_rows": contains_xp_df.values.tolist()
    }

    print("\n字典集合（包含 XP，每行独立数组）：")
    print(grouped_data)
    print(f"\n包含 XP 的数据行数量：{len(grouped_data['contains_xp_rows'])}")

    return grouped_data


# 赛狐ERP网页操作
def SaiHuERP_WebPage():
    # 1) 先调用 SaihuERPLogin 类，完成浏览器启动 + 跳转赛狐 + 登录
    page = SaihuERPLogin.start_and_login(
        fresh_profile=True,
        force_relogin=True,
    )

    # 2) 再开始执行已知 XPath 的页面定位与点击
    # 定位一级商品
    page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click()
    time.sleep(1)
    # 定位新品开发
    page.ele('x://a[text()="新品开发"]', timeout=8).click()
    time.sleep(1)
    # 定位需求编号输入栏
    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
    # 先清空输入框历史内容
    # time.sleep(0.1)
    # page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).clear()
    time.sleep(0.2)
    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input('XP2604210004\n',clear=True)
    time.sleep(0.2)
    if page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input('XP2604210004\n',clear=True):
        print("该编号流程需先提交，继续执行")
    else:
        print("需求编号不存在，退出程序")
        return
    # time.sleep(1.5)
    page.ele('x://div/ul/li[contains(text(), "生成普通商品")]').click(by_js=True)
    time.sleep(1.5)

    # 在点击生成普通商品后，填写信息
    # 要插入的文本前缀
    new_prefix = "Dried Flowers-40pc"
    # 要搜索的人员姓名
    new_name = "尹开元"
    # 转义特殊字符
    safe_prefix = new_prefix.replace("'", "\\'").replace('"', '\\"')

    # 品名定位并插入前缀
    page.run_js(f"""
        const xpath = '//label[contains(text(), "品名")]/following::div/input';
        const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
        const inputElement = result.singleNodeValue;

        if (inputElement) {{
            inputElement.value = '{safe_prefix}' + inputElement.value;
            inputElement.dispatchEvent(new Event('input', {{ bubbles: true }}));
            console.log('插入成功:', '{safe_prefix}');
        }} else {{
            console.error('未找到文本框');
        }}
    """)
    time.sleep(0.2)

    # SKU定位并清空输入框
    page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).clear()
    time.sleep(0.2)

    # SKU定位输入框填写SKU
    page.ele('x://label[contains(text(), "SKU")]/following::div/input', timeout=8).input(new_prefix)
    time.sleep(0.2)

    # 查看人定位并选择
    page.ele('x://label[contains(text(), "查看人")]/following::div/input', timeout=8).click()
    time.sleep(1)
    # 查看人输入框填写
    page.ele('x://div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"]', timeout=8).input(new_name)
    time.sleep(1)
    page.ele(f'x://div[@class="select-menu"]/div[2]//span[text()="{new_name}"]', timeout=8).click()
    time.sleep(1)
    page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
    time.sleep(1)

    # 商品备注定位并填写
    page.ele('x://label[contains(text(), "商品备注")]/following::div/textarea', timeout=8).input('自动化程序测试备注，请忽略不要作为正式商品',clear=True)
    time.sleep(1)
    # 保存按钮定位并点击
    page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="保存"])[last()]', timeout=8).click()
    time.sleep(1)

    # 判断商品SKU是否已存在，如果存在则取消保存
    if page.ele('x://*[contains(text(), "商品SKU已存在")]', timeout=8):
        page.ele('x:(//div[@aria-label="生成普通商品"]//button/span[text()="取消"])[last()]', timeout=8).click()
        time.sleep(1)
        print("商品SKU已存在，取消保存")
    else:
        # 如果不存在则保存
        print("商品SKU不存在，保存商品")

    # 定位一级销售
    page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
    time.sleep(1)
    # 定位在线商品
    page.ele('x://a[text()="在线产品"]', timeout=8).click()
    time.sleep(1)
    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).click()
    # 填写ASIN
    new_asin = 'B0H29VN1DG'
    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(f'{new_asin}\n',clear=True)
    time.sleep(1)
    # 定位配对按钮并点击(单个)
    # page.ele('x:(//td//span[contains(text(), "配对")])[1]', timeout=10).click()
    # 动态定位对应行的「配对」按钮
    page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=10).click()
    time.sleep(1)
    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=10).input(f'{new_prefix}\n',clear=True)
    time.sleep(1)
    page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=10).click()
    time.sleep(1)

if __name__ == "__main__":
    # 如需读取表格再放开以下两行
    grouped_result = read_Excel()
    print(f"\n统计：包含XP={len(grouped_result['contains_xp_rows'])}")

    # SaiHuERP_WebPage()
    # 1) 启动易得客浏览器（不调用易得客账号登录）
    # Specification(username="", password="")
    # time.sleep(2)

    # 2) 连接已启动的易得客浏览器调试端口
    # page = ChromiumPage("127.0.0.1:9222")

    # 3) 调用赛狐 ERP 登录流程（内部会打开赛狐 ERP 页面）
    # saihu_login = SaihuERPLogin(page)
    # saihu_login.login()

    
