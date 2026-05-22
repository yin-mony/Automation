import re
import sys
import time
from pathlib import Path

import pandas as pd

CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from EdgeRun import EdgeBrowserRunner
from SaihuERPLogin import SaihuERPLogin

EXCEL_FILE_PATH = Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
EXCEL_SHEET_NAME = "新品sku配对自动提醒"


def _extract_xp_code(value):
    text = str(value or "").upper()
    match = re.search(r"XP\d+", text)
    return match.group(0) if match else ""


def read_Excel(excel_file_path=None, sheet_name=EXCEL_SHEET_NAME):
    """读取固定工作表，筛选未配对并结构化提取 sku/ASIN/人员。"""
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


def start_edge_and_login_saihu(username=None, password=None, fresh_profile=False, force_relogin=False, wait_seconds=3):
    print("步骤1：启动并连接本地 Edge 浏览器...", flush=True)
    page = EdgeBrowserRunner.start_edge_and_connect(
        debug_port=EdgeBrowserRunner.DEFAULT_DEBUG_PORT,
        start_url=SaihuERPLogin.ENTRY_URL,
        fresh_profile=fresh_profile,
        wait_seconds=wait_seconds,
    )
    print(f"当前页面: {page.url}", flush=True)

    print("步骤2：执行赛狐 ERP 登录流程...", flush=True)
    login_client = SaihuERPLogin(page, username=username, password=password)
    login_client.login(force_relogin=force_relogin)
    print("步骤完成：已完成 Edge 拉起并执行赛狐 ERP 登录。", flush=True)
    return page


# 赛狐ERP网页操作
def SaiHuERP_WebPage(excel_file_path=None, username=None, password=None, sheet_name=EXCEL_SHEET_NAME):
    grouped_result = read_Excel(excel_file_path=excel_file_path, sheet_name=sheet_name)
    records = grouped_result["records"]
    if not records:
        raise RuntimeError("没有可执行数据：未筛选到包含 XP 的未配对记录。")

    print(f"\n统计：包含XP={len(records)}，将按该数量循环执行。")
    page = start_edge_and_login_saihu(
        username=username,
        password=password,
        fresh_profile=False,
        force_relogin=False,
    )

    for index, record in enumerate(records, start=1):
        request_no = record["request_no"]
        new_prefix = record["new_prefix"]
        new_name = record["new_name"]
        new_asin = record["new_asin"]
        print(f"\n[{index}/{len(records)}] 处理编号：{request_no}", flush=True)

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
        time.sleep(1)
        page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(f'{request_no}\n', clear=True)
        time.sleep(1)
        # time.sleep(1.5)
        page.ele('x://div/ul/li[contains(text(), "生成普通商品")]').click(by_js=True)
        time.sleep(1.5)

        # 在点击生成普通商品后，填写信息
        # 要插入的文本前缀
        # 要搜索的人员姓名
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
        page.ele('x://div[@class="select-menu"]/div/input[@class="sf_select__filter__input is-small"]', timeout=8).input(f'{new_name}\n', clear=True)
        time.sleep(1)
        page.ele(f'x://div[@class="select-menu"]/div[2]//span[text()="{new_name}"]', timeout=8).click()
        time.sleep(1)
        page.ele('x://div[@class="sf_select__footer"]/div[2]/button[2]/span[text()="确定"]', timeout=8).click()
        time.sleep(1)

        # 商品备注定位并填写
        page.ele('x://label[contains(text(), "商品备注")]/following::div/textarea', timeout=8).input('自动化程序测试备注，请忽略不要作为正式商品', clear=True)
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
        page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]', timeout=10).input(f'{new_asin}\n', clear=True)
        time.sleep(1)
        # 定位配对按钮并点击(单个)
        # page.ele('x:(//td//span[contains(text(), "配对")])[1]', timeout=10).click()
        # 动态定位对应行的「配对」按钮
        if not page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=3):
            print("该数据，已配对")
            continue
        page.ele(f'x://tr[.//text()[contains(., "{new_asin}")]]//span[contains(text(), "配对")]', timeout=10).click()
        time.sleep(1)
        page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]', timeout=10).input(f'{new_prefix}\n', clear=True)
        time.sleep(1)
        page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=10).click()
        time.sleep(1)


if __name__ == "__main__":
    SaiHuERP_WebPage()
