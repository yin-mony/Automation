import argparse
import json
import os
import re
import time
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd
from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin


# 赛狐 ERP 测试流程：调试接管、接口监听、规格补填、完整创建配对
class SaihuDebugTest:
    def __init__(self, config):
        self.config = config
        self.run_mode = config.get("run_mode") or "debug"
        self.business_mode = config.get("business_mode") or "low_price"
        self.action = config.get("action") or "watch"
        self.username = config.get("username") or os.getenv("SAIHU_USERNAME", "")
        self.password = config.get("password") or os.getenv("SAIHU_PASSWORD", "")
        self.base_dir = Path(config.get("base_dir") or Path(__file__).resolve().parent)
        self.browser_port = config.get("browser_port") or ""
        self.wait_seconds = int(config.get("wait_seconds") or 120)
        self.excel_path = Path(config["excel_path"]) if config.get("excel_path") else None
        self.sku = str(config.get("sku") or "").strip()
        self.asin = str(config.get("asin") or "").strip()
        self.limit = int(config.get("limit") or 0)
        self.start_index = int(config.get("start_index") or 1)
        self.auto_save = bool(config.get("auto_save"))
        self.allow_any_status = bool(config.get("allow_any_status"))
        self.page = None
        self.listen_targets = [
            "https://www.sellfox.com/",
            "https://www.sellfox.com/api/",
            "https://www.sellfox.com/amzup-web-main/",
        ]

    def main(self):
        if self.browser_port:
            print(f"接管浏览器调试端口: {self.browser_port}", flush=True)
            self.page = ChromiumPage(int(self.browser_port))
        else:
            print("使用默认 ChromiumPage，优先复用当前可用浏览器上下文。", flush=True)
            self.page = ChromiumPage()

        if self.run_mode == "debug":
            try:
                latest_tab = self.page.latest_tab
                if latest_tab:
                    self.page = latest_tab
                    print("已接管当前最新标签页。", flush=True)
            except Exception:
                print("未切换 latest_tab，直接使用当前页面。", flush=True)
        print(f"当前页面: {getattr(self.page, 'url', '')}", flush=True)

        if self.run_mode == "formal":
            print("正式流程：先登录赛狐。", flush=True)
            login = SaiHuERPLogin({
                "page": self.page,
                "username": self.username,
                "password": self.password,
                "img_path": self.base_dir,
            })
            if not login.login():
                raise RuntimeError("赛狐登录失败，停止测试。")
        else:
            print("调试流程：不登录，不重跑登录态，只接管当前页面继续。", flush=True)

        if self.action == "watch":
            self.watch_packets(self.wait_seconds, None, "手动监听")
            return

        if self.action == "snapshot":
            self.fill_specs(None)
            return

        if self.action in ("fill_specs", "fill_and_watch"):
            rows = self.load_rows()
            item = rows[0]
            if self.action == "fill_and_watch":
                try:
                    self.page.listen.stop()
                except Exception:
                    pass
                try:
                    self.page.listen.start(self.listen_targets)
                except Exception:
                    self.page.listen.start(self.listen_targets[0])
            self.fill_specs(item)
            if self.auto_save:
                self.page.run_js("""
                const buttons = Array.from(document.querySelectorAll('button'))
                  .filter(btn => {
                    const style = window.getComputedStyle(btn);
                    const rect = btn.getBoundingClientRect();
                    return style.display !== 'none' && style.visibility !== 'hidden'
                      && rect.width > 0 && rect.height > 0
                      && (btn.textContent || '').includes('保存');
                  });
                const btn = buttons.length ? buttons[buttons.length - 1] : null;
                if (btn) btn.click();
                return !!btn;
                """)
            if self.action == "fill_and_watch":
                self.watch_packets(self.wait_seconds, item, "补填后监听")
            return

        rows = self.load_rows()
        print(f"完整流程准备处理 {len(rows)} 条。", flush=True)
        for idx, item in enumerate(rows, 1):
            print("\n" + "#" * 80, flush=True)
            print(f"完整流程 {idx}/{len(rows)} SKU={item.get('SKU')} ASIN={item.get('ASIN')}", flush=True)
            try:
                self.page.run_js("""
                function visible(el) {
                  if (!el) return false;
                  const style = getComputedStyle(el);
                  const rect = el.getBoundingClientRect();
                  return style.display !== 'none' && style.visibility !== 'hidden'
                    && rect.width > 0 && rect.height > 0;
                }
                const dialogs = Array.from(document.querySelectorAll('.el-dialog, [role="dialog"], div[aria-label]')).filter(visible);
                const dialog = dialogs.length ? dialogs[dialogs.length - 1] : null;
                if (dialog) {
                  const close = dialog.querySelector('.el-dialog__headerbtn, .el-dialog__close');
                  if (close) close.click();
                }
                return !!dialog;
                """)
                time.sleep(0.5)
                if self.business_mode in ("low_price", "mode_two", "mode2", "variant", "mode_three", "mode3"):
                    page = self.page
                    print("进入商品列表并打开添加单个商品。", flush=True)
                    page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    try:
                        page.ele('x://a[text()="商品列表"]', timeout=8).click(by_js=True)
                    except Exception:
                        print("未找到商品列表菜单，直接进入商品列表地址。", flush=True)
                        page.get("https://www.sellfox.com/amzup-web-main/web/commodity/index.html")
                    time.sleep(1)
                    page.ele('x://button//span[text()="添加商品"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    page.ele('x://span[text()="添加单个商品"]', timeout=5).click(by_js=True)
                    time.sleep(1)

                    print("填写基础信息和规格信息。", flush=True)
                    self.fill_specs(item)
                    owner_name = json.dumps(item["负责人"], ensure_ascii=False)
                    owner_clicked = page.run_js(f"""
                    const ownerName = {owner_name};
                    function visible(el) {{
                      if (!el) return false;
                      const style = getComputedStyle(el);
                      const rect = el.getBoundingClientRect();
                      return style.display !== 'none' && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0;
                    }}
                    function norm(text) {{
                      return String(text || '').replace(/\\s+/g, '');
                    }}
                    const dialogs = Array.from(document.querySelectorAll('.el-dialog, [role="dialog"], div[aria-label]')).filter(visible);
                    const root = dialogs.length ? dialogs[dialogs.length - 1] : document;
                    const formItem = Array.from(root.querySelectorAll('.el-form-item')).find(node => {{
                      const label = node.querySelector('.el-form-item__label, label');
                      return label && norm(label.textContent).includes('查看人');
                    }});
                    const componentNode = formItem && formItem.querySelector('.OrgStruStaffSelect');
                    const component = componentNode && componentNode.__vue__;
                    const parent = component && component.$parent;
                    const user = component && (component.baseUserList || []).find(option => {{
                      return norm(option.label) === norm(ownerName) || norm(option.label).includes(norm(ownerName));
                    }});
                    if (!user) return false;
                    const value = [user.value];
                    const option = [user];
                    if (parent && parent.setData) parent.setData(value);
                    if (parent && parent.handleChange) parent.handleChange(value);
                    if (component && component.handleSelectChange) component.handleSelectChange(value);
                    if (parent) {{
                      parent.innerValue = value;
                      parent.$emit('input', value);
                      parent.$emit('on-change', option);
                    }}
                    if (component) {{
                      component.$emit('input', value);
                      component.$emit('change', value);
                    }}
                    return true;
                    """)
                    time.sleep(0.5)
                    if not owner_clicked:
                        print(f"查看人未自动选中: {item['负责人']}", flush=True)
                    else:
                        print(f"查看人已选中: {item['负责人']}", flush=True)
                    page.run_js("""
                    function visible(el) {
                      if (!el) return false;
                      const style = getComputedStyle(el);
                      const rect = el.getBoundingClientRect();
                      return style.display !== 'none' && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0;
                    }
                    const buttons = Array.from(document.querySelectorAll('button')).filter(visible);
                    const btn = buttons.reverse().find(button => (button.textContent || '').trim().includes('确定'));
                    if (btn) btn.click();
                    return !!btn;
                    """)
                    time.sleep(1)
                    page.ele('x://div[normalize-space()="采购信息"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    self.fill_specs(item)
                    page.ele('x://div[normalize-space()="规格信息"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    self.fill_specs(item)
                else:
                    page = self.page
                    print("进入新品开发并打开生成普通商品。", flush=True)
                    page.ele('x://div/ul/li/span[text()="商品"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    try:
                        page.ele('x://a[text()="新品开发"]', timeout=8).click(by_js=True)
                    except Exception:
                        print("未找到新品开发菜单，尝试点击顶部新品开发标签。", flush=True)
                        page.ele('x://span[text()="新品开发"]', timeout=8).click(by_js=True)
                    time.sleep(1)
                    page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]',
                             timeout=10).input(f"{item['新品开发编号']}\n", clear=True)
                    time.sleep(1)
                    btn = page.ele('x://div/ul/li[contains(text(), "生成普通商品")]', timeout=8)
                    if not btn:
                        print("未出现生成普通商品，跳过。", flush=True)
                        continue
                    btn.click(by_js=True)
                    time.sleep(1.5)
                    print("填写新品生成普通商品信息。", flush=True)
                    self.fill_specs(item)
                    owner_name = json.dumps(item["负责人"], ensure_ascii=False)
                    owner_clicked = page.run_js(f"""
                    const ownerName = {owner_name};
                    function visible(el) {{
                      if (!el) return false;
                      const style = getComputedStyle(el);
                      const rect = el.getBoundingClientRect();
                      return style.display !== 'none' && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0;
                    }}
                    function norm(text) {{
                      return String(text || '').replace(/\\s+/g, '');
                    }}
                    const dialogs = Array.from(document.querySelectorAll('.el-dialog, [role="dialog"], div[aria-label]')).filter(visible);
                    const root = dialogs.length ? dialogs[dialogs.length - 1] : document;
                    const formItem = Array.from(root.querySelectorAll('.el-form-item')).find(node => {{
                      const label = node.querySelector('.el-form-item__label, label');
                      return label && norm(label.textContent).includes('查看人');
                    }});
                    const componentNode = formItem && formItem.querySelector('.OrgStruStaffSelect');
                    const component = componentNode && componentNode.__vue__;
                    const parent = component && component.$parent;
                    const user = component && (component.baseUserList || []).find(option => {{
                      return norm(option.label) === norm(ownerName) || norm(option.label).includes(norm(ownerName));
                    }});
                    if (!user) return false;
                    const value = [user.value];
                    const option = [user];
                    if (parent && parent.setData) parent.setData(value);
                    if (parent && parent.handleChange) parent.handleChange(value);
                    if (component && component.handleSelectChange) component.handleSelectChange(value);
                    if (parent) {{
                      parent.innerValue = value;
                      parent.$emit('input', value);
                      parent.$emit('on-change', option);
                    }}
                    if (component) {{
                      component.$emit('input', value);
                      component.$emit('change', value);
                    }}
                    return true;
                    """)
                    time.sleep(0.5)
                    if not owner_clicked:
                        print(f"查看人未自动选中: {item['负责人']}", flush=True)
                    else:
                        print(f"查看人已选中: {item['负责人']}", flush=True)
                    if not all(str(item.get(key) or "").strip() for key in [
                        "长 包装规格（cm）", "宽 包装规格（cm)", "高 包装规格（cm）", "单品毛重（kg）"
                    ]):
                        print("纯新品缺少长宽高重量，按原流程不读取不补填规格，继续保存后流程。", flush=True)

                self.fill_specs(item)
                if self.action == "prepare_flow":
                    print("已到保存前测试点：本次不点击保存，不触发配对。", flush=True)
                    continue
                try:
                    self.page.listen.stop()
                except Exception:
                    pass
                try:
                    self.page.listen.start(self.listen_targets)
                except Exception:
                    self.page.listen.start(self.listen_targets[0])
                print("准备点击保存并监听保存接口。", flush=True)
                self.page.run_js("""
                function visible(el) {
                  if (!el) return false;
                  const style = window.getComputedStyle(el);
                  const rect = el.getBoundingClientRect();
                  return style.display !== 'none' && style.visibility !== 'hidden'
                    && rect.width > 0 && rect.height > 0;
                }
                function fire(el, type) {
                  const rect = el.getBoundingClientRect();
                  el.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: rect.left + rect.width / 2,
                    clientY: rect.top + rect.height / 2
                  }));
                }
                const dialogs = Array.from(document.querySelectorAll('.el-dialog, [role="dialog"], div[aria-label]')).filter(visible);
                const root = dialogs.length ? dialogs[dialogs.length - 1] : document;
                const buttons = Array.from(root.querySelectorAll('.dialog-footer button, .el-dialog__footer button, button')).filter(btn => {
                  const style = window.getComputedStyle(btn);
                  const rect = btn.getBoundingClientRect();
                  return style.display !== 'none' && style.visibility !== 'hidden'
                    && rect.width > 0 && rect.height > 0
                    && (btn.textContent || '').includes('保存');
                });
                const btn = buttons.length ? buttons.reverse()[0] : null;
                if (btn) {
                  btn.scrollIntoView({block: 'center', inline: 'center'});
                  fire(btn, 'mouseover');
                  fire(btn, 'mousedown');
                  fire(btn, 'mouseup');
                  fire(btn, 'click');
                }
                return !!btn;
                """)
                save_result = self.watch_packets(15, item, "保存接口")
                if save_result.get("sku_exists"):
                    print("检测到 SKU 已存在，取消弹窗后继续配对。", flush=True)
                    self.page.run_js("""
                    const buttons = Array.from(document.querySelectorAll('button')).filter(btn => {
                      const style = window.getComputedStyle(btn);
                      const rect = btn.getBoundingClientRect();
                      return style.display !== 'none' && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0
                        && (btn.textContent || '').includes('取消');
                    });
                    const btn = buttons.length ? buttons[buttons.length - 1] : null;
                    if (btn) btn.click();
                    return !!btn;
                    """)
                    time.sleep(1)
                elif not save_result.get("success"):
                    print("保存接口未确认成功，跳过配对。", flush=True)
                    continue

                page = self.page
                page.ele('x://div/ul/li/span[text()="销售"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://a[text()="在线产品"]', timeout=8).click()
                time.sleep(1)
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="双击可批量搜索内容"]',
                         timeout=10).input(f"{item['ASIN']}\n", clear=True)
                time.sleep(1.5)
                pair_btn = page.ele(
                    f'x://tr[.//text()[contains(., "{item["ASIN"]}")]]//span[contains(text(), "配对")]',
                    timeout=10,
                )
                if not pair_btn:
                    print("该商品已配对过 ASIN，跳过配对。", flush=True)
                    continue
                pair_btn.click(by_js=True)
                time.sleep(1)
                page.ele('x://div[@class="sel_ipt"]//input[@placeholder="搜索内容"]',
                         timeout=5).input(f"{item['SKU']}\n", clear=True)
                time.sleep(1)
                try:
                    page.listen.stop()
                except Exception:
                    pass
                try:
                    page.listen.start(self.listen_targets)
                except Exception:
                    page.listen.start(self.listen_targets[0])
                final_btn = page.ele('x://div[@class="vxe-cell"]/button/span[contains(text(), "配对")]', timeout=8)
                if not final_btn:
                    print("未找到最终配对按钮。", flush=True)
                    continue
                final_btn.click()
                self.watch_packets(10, item, "配对接口")
            except Exception as exc:
                print(f"完整流程当前行失败: {exc}", flush=True)
                continue
        print("完整流程结束。", flush=True)

    def load_rows(self):
        if self.business_mode in ("low_price", "mode_two", "mode2"):
            excel_path = self.excel_path or Path(r"C:\Users\admin\Desktop\低价商城创建ERP-SKU.xlsx")
            if not excel_path.exists():
                raise RuntimeError(f"低价模式 Excel 不存在: {excel_path}")
            try:
                df = pd.read_excel(excel_path, sheet_name="工作表1")
            except Exception as exc:
                print(f"低价模式 pandas 读取失败，改用 xlsx XML 兜底读取: {exc}", flush=True)
                ns = {
                    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                }
                with zipfile.ZipFile(excel_path) as archive:
                    shared_strings = []
                    if "xl/sharedStrings.xml" in archive.namelist():
                        shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                        for item in shared_root.findall("a:si", ns):
                            shared_strings.append("".join(node.text or "" for node in item.findall(".//a:t", ns)))

                    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
                    rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                    rels = {node.attrib["Id"]: node.attrib["Target"] for node in rel_root}
                    sheet_path = ""
                    for sheet in workbook.find("a:sheets", ns):
                        if sheet.attrib.get("name") == "工作表1":
                            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                            target = rels.get(rel_id, "").lstrip("/")
                            sheet_path = "xl/" + target if target and not target.startswith("xl/") else target
                            break
                    if not sheet_path:
                        raise RuntimeError("低价模式 Excel 未找到工作表1")

                    sheet_root = ET.fromstring(archive.read(sheet_path))
                    rows_data = []
                    max_col = 0
                    for row_node in sheet_root.findall(".//a:sheetData/a:row", ns):
                        row_values = {}
                        for cell in row_node.findall("a:c", ns):
                            ref = cell.attrib.get("r", "")
                            col_text = re.sub(r"\d", "", ref)
                            col_num = 0
                            for char in col_text:
                                col_num = col_num * 26 + ord(char.upper()) - 64
                            if col_num <= 0:
                                continue
                            value = ""
                            cell_type = cell.attrib.get("t")
                            if cell_type == "inlineStr":
                                value = "".join(node.text or "" for node in cell.findall(".//a:t", ns))
                            else:
                                value_node = cell.find("a:v", ns)
                                if value_node is not None:
                                    value = value_node.text or ""
                                    if cell_type == "s" and value.isdigit() and int(value) < len(shared_strings):
                                        value = shared_strings[int(value)]
                            row_values[col_num - 1] = value
                            max_col = max(max_col, col_num)
                        if row_values:
                            rows_data.append(row_values)

                    table = []
                    for row_values in rows_data:
                        table.append([row_values.get(idx, "") for idx in range(max_col)])
                    if not table:
                        raise RuntimeError("低价模式 Excel 没有读取到数据")
                    header = [str(value).strip() for value in table[0]]
                    df = pd.DataFrame(table[1:], columns=header)
            required = [
                "时间", "品名", "SKU", "ASIN", "长 包装规格（cm）", "宽 包装规格（cm)",
                "高 包装规格（cm）", "单品毛重（kg）", "采购价（元）", "负责人",
            ]
            missing = [col for col in required if col not in df.columns]
            if missing:
                raise RuntimeError(f"低价表缺少必要列: {', '.join(missing)}")
            time_series = df["时间"].astype(str).str.strip().replace("时间", pd.NA)
            df["时间"] = pd.to_datetime(time_series, errors="coerce", format="mixed", utc=True)
            df = df[df["时间"].notna()]
            df["时间"] = df["时间"].dt.tz_localize(None)
            latest_date = df["时间"].dt.date.max()
            df = df[df["时间"].dt.date == latest_date]
            print(f"低价模式 Excel 自动筛选最新日期: {latest_date}", flush=True)
            if self.sku:
                df = df[df["SKU"].astype(str).str.strip() == self.sku]
            if self.asin:
                df = df[df["ASIN"].astype(str).str.strip() == self.asin]
            rows = []
            for item in df[required].astype(str).to_dict("records"):
                row = {}
                for key, value in item.items():
                    text = str(value).replace("\r", " ").replace("\n", " ").strip()
                    if text.lower() == "nan":
                        text = ""
                    row[key] = text
                row["单品毛重（kg）"] = row["单品毛重（kg）"].replace("KG", "").replace("kg", "").strip()
                rows.append(row)
        elif self.business_mode in ("variant", "mode_three", "mode3"):
            excel_path = self.excel_path or Path(r"C:\Users\admin\Desktop\新品sku配对+横向变体配对自动提醒.xlsx")
            if not excel_path.exists():
                raise RuntimeError(f"变体开发模式 Excel 不存在: {excel_path}")
            df = pd.read_excel(excel_path, sheet_name="横向变体")
            required = [
                "sku", "ASIN", "FNSKU", "包装-长（cm）", "包装-宽（cm）", "包装-高（cm）",
                "包装-重量（g）", "不含税成本价格", "人员", "情况",
            ]
            missing = [col for col in required if col not in df.columns]
            if missing:
                raise RuntimeError(f"变体开发表缺少必要列: {', '.join(missing)}")
            if not self.allow_any_status:
                df = df[df["情况"] == "未配对"]
            else:
                print("变体开发测试开启 allow-any-status：不按情况过滤，仅用于保存前页面验证。", flush=True)
            if self.sku:
                df = df[df["sku"].astype(str).str.strip() == self.sku]
            if self.asin:
                df = df[df["ASIN"].astype(str).str.strip() == self.asin]
            rows = []
            for _, source in df.iterrows():
                sku = str(source.get("sku") or "").replace("\r", " ").replace("\n", " ").strip()
                asin = str(source.get("ASIN") or "").replace("\r", " ").replace("\n", " ").strip()
                item = {
                    "品名": sku,
                    "SKU": sku,
                    "ASIN": asin,
                    "负责人": str(source.get("人员") or "").strip(),
                    "长 包装规格（cm）": str(source.get("包装-长（cm）") or "").strip(),
                    "宽 包装规格（cm)": str(source.get("包装-宽（cm）") or "").strip(),
                    "高 包装规格（cm）": str(source.get("包装-高（cm）") or "").strip(),
                    "单品毛重（kg）": str(source.get("包装-重量（g）") or "").strip(),
                    "采购价（元）": str(source.get("不含税成本价格") or "").strip(),
                }
                for key, value in item.items():
                    if str(value).lower() == "nan":
                        item[key] = ""
                rows.append(item)
        else:
            excel_path = self.excel_path or Path(r"C:\Users\admin\Desktop\工作计划表.xlsx")
            if not excel_path.exists():
                raise RuntimeError(f"纯新品模式 Excel 不存在: {excel_path}")
            df = pd.read_excel(excel_path, sheet_name="新品sku配对自动提醒")
            required = ["情况", "赛狐新品开发编号", "sku", "ASIN", "人员"]
            missing = [col for col in required if col not in df.columns]
            if missing:
                raise RuntimeError(f"工作计划表缺少必要列: {', '.join(missing)}")
            df = df[
                (df["情况"] == "未配对")
                & (df["赛狐新品开发编号"].astype(str).str.contains(r"XP\d+", na=False, regex=True))
            ]
            if self.sku:
                df = df[df["sku"].astype(str).str.strip() == self.sku]
            if self.asin:
                df = df[df["ASIN"].astype(str).str.strip() == self.asin]
            rows = []
            for _, row in df.iterrows():
                spec_text = str(row.get("如开发未知包装尺寸和重量请填写") or "").strip()
                spec_numbers = re.findall(r"\d+(?:\.\d+)?", spec_text)
                price_text = str(row.get("如开发未知报价请填写套装报价") or row.get("采购价（元）") or "").strip()
                item = {
                    "新品开发编号": str(row.get("赛狐新品开发编号") or "").strip(),
                    "品名": str(row.get("品名") or row.get("sku") or "").strip(),
                    "SKU": str(row.get("sku") or row.get("SKU") or "").strip(),
                    "ASIN": str(row.get("ASIN") or "").strip(),
                    "负责人": str(row.get("人员") or row.get("负责人") or "").strip(),
                    "长 包装规格（cm）": str(row.get("长 包装规格（cm）") or row.get("长 包装规格（cm)") or (spec_numbers[0] if len(spec_numbers) > 0 else "")).strip(),
                    "宽 包装规格（cm)": str(row.get("宽 包装规格（cm)") or row.get("宽 包装规格（cm）") or (spec_numbers[1] if len(spec_numbers) > 1 else "")).strip(),
                    "高 包装规格（cm）": str(row.get("高 包装规格（cm）") or row.get("高 包装规格（cm)") or (spec_numbers[2] if len(spec_numbers) > 2 else "")).strip(),
                    "单品毛重（kg）": str(row.get("单品毛重（kg）") or row.get("包装重量") or (spec_numbers[3] if len(spec_numbers) > 3 else "")).strip(),
                    "采购价（元）": price_text,
                }
                for key, value in item.items():
                    if str(value).lower() == "nan":
                        item[key] = ""
                rows.append(item)

        start = max(0, self.start_index - 1)
        rows = rows[start:]
        if self.limit > 0:
            rows = rows[:self.limit]
        if not rows:
            raise RuntimeError("没有筛选到待处理数据。")
        return rows

    def fill_specs(self, item):
        script = """
        const data = arguments[0] || {};
        function visible(el) {
          if (!el) return false;
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== 'none' && style.visibility !== 'hidden'
            && rect.width > 0 && rect.height > 0;
        }
        function rootNode() {
          const dialogs = Array.from(document.querySelectorAll('.el-dialog, [role="dialog"], div[aria-label]'))
            .filter(visible)
            .sort((a, b) => {
              const ar = a.getBoundingClientRect();
              const br = b.getBoundingClientRect();
              return (br.width * br.height) - (ar.width * ar.height);
            });
          return dialogs[0] || document;
        }
        const root = rootNode();
        function norm(text) {
          return String(text || '').replace(/\\s+/g, '');
        }
        function labelNode(text) {
          const target = norm(text);
          const nodes = Array.from(root.querySelectorAll('label, span, div, td, th, p'))
            .filter(visible)
            .filter(node => norm(node.textContent).includes(target))
            .filter(node => norm(node.textContent).length <= Math.max(target.length + 12, 24))
            .sort((a, b) => {
              const at = norm(a.textContent);
              const bt = norm(b.textContent);
              const ar = a.getBoundingClientRect();
              const br = b.getBoundingClientRect();
              return at.length - bt.length || (ar.width * ar.height) - (br.width * br.height);
            });
          return nodes[0] || null;
        }
        function inputsBetween(startText, endText) {
          const start = labelNode(startText);
          const end = endText ? labelNode(endText) : null;
          const inputs = Array.from(root.querySelectorAll('input, textarea')).filter(visible);
          if (!start) return [];
          return inputs.filter(input => {
            const afterStart = !!(start.compareDocumentPosition(input) & Node.DOCUMENT_POSITION_FOLLOWING);
            const beforeEnd = !end || !!(input.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING);
            return afterStart && beforeEnd;
          });
        }
        function setInput(input, value) {
          if (!input || value === undefined || value === null || String(value).trim() === '') return false;
          const text = String(value).trim();
          const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
          setter.call(input, text);
          input.dispatchEvent(new Event('input', {bubbles: true}));
          input.dispatchEvent(new Event('change', {bubbles: true}));
          input.dispatchEvent(new Event('blur', {bubbles: true}));
          return input.value === text;
        }
        function formInput(labelText) {
          const items = Array.from(root.querySelectorAll('.el-form-item')).filter(visible);
          const item = items.find(node => {
            const label = node.querySelector('.el-form-item__label, label');
            return label && norm(label.textContent).includes(norm(labelText));
          });
          if (!item) return null;
          const inputs = Array.from(item.querySelectorAll('input, textarea')).filter(visible);
          return inputs.find(input => !input.disabled && !input.readOnly) || inputs[0] || null;
        }
        function fillForm(labelText, value) {
          const input = formInput(labelText);
          return {label: labelText, placeholder: '', found: !!input, ok: setInput(input, value), value: input ? input.value : ''};
        }
        function fill(startText, value, placeholder, endText) {
          const inputs = inputsBetween(startText, endText);
          let input = null;
          if (placeholder) input = inputs.find(i => (i.getAttribute('placeholder') || '').trim() === placeholder);
          else input = inputs[0];
          return {label: startText, placeholder: placeholder || '', found: !!input, ok: setInput(input, value), value: input ? input.value : ''};
        }
        function fillPackingTable(lengthValue, widthValue, heightValue, weightValue) {
          const title = labelNode('包装规格') || labelNode('商品包装规格');
          const results = [];
          if (!title) {
            return [
              {label: '包装规格表-长', found: false, ok: false, value: ''},
              {label: '包装规格表-宽', found: false, ok: false, value: ''},
              {label: '包装规格表-高', found: false, ok: false, value: ''},
              {label: '包装规格表-重量', found: false, ok: false, value: ''}
            ];
          }
          const titleRect = title.getBoundingClientRect();
          const inputs = Array.from(root.querySelectorAll('input, textarea'))
            .filter(visible)
            .filter(input => {
              const rect = input.getBoundingClientRect();
              return rect.y > titleRect.y + 20 && !input.disabled && !input.readOnly;
            })
            .sort((a, b) => {
              const ar = a.getBoundingClientRect();
              const br = b.getBoundingClientRect();
              return ar.y - br.y || ar.x - br.x;
            });
          const longInputs = inputs.filter(input => (input.getAttribute('placeholder') || '').trim() === '长');
          const widthInputs = inputs.filter(input => (input.getAttribute('placeholder') || '').trim() === '宽');
          const heightInputs = inputs.filter(input => (input.getAttribute('placeholder') || '').trim() === '高');
          for (const input of longInputs) setInput(input, lengthValue);
          for (const input of widthInputs) setInput(input, widthValue);
          for (const input of heightInputs) setInput(input, heightValue);
          const otherInputs = inputs.filter(input => !['长', '宽', '高'].includes((input.getAttribute('placeholder') || '').trim()));
          const weightInput = otherInputs.length ? otherInputs[otherInputs.length - 1] : null;
          results.push({label: '包装规格表-长', found: longInputs.length > 0, ok: longInputs.every(input => input.value === String(lengthValue || '').trim()), value: longInputs.map(input => input.value).join('|')});
          results.push({label: '包装规格表-宽', found: widthInputs.length > 0, ok: widthInputs.every(input => input.value === String(widthValue || '').trim()), value: widthInputs.map(input => input.value).join('|')});
          results.push({label: '包装规格表-高', found: heightInputs.length > 0, ok: heightInputs.every(input => input.value === String(heightValue || '').trim()), value: heightInputs.map(input => input.value).join('|')});
          results.push({label: '包装规格表-重量', found: !!weightInput, ok: setInput(weightInput, weightValue), value: weightInput ? weightInput.value : ''});
          return results;
        }
        function snapshot() {
          const labels = ['品名', 'SKU', '采购成本', '商品规格', '商品包装规格', '商品包装重量'];
          return labels.map(label => ({
            label,
            found: !!labelNode(label),
            field: formInput(label) ? {
              placeholder: formInput(label).getAttribute('placeholder') || '',
              value: formInput(label).value || '',
              disabled: formInput(label).disabled,
              readonly: formInput(label).readOnly
            } : null,
            values: inputsBetween(label, null).slice(0, 8).map(input => ({
              placeholder: input.getAttribute('placeholder') || '',
              value: input.value || '',
              disabled: input.disabled,
              readonly: input.readOnly
            }))
          }));
        }
        const before = snapshot();
        const results = [];
        if (data['品名']) results.push(fillForm('品名', data['品名']));
        if (data['SKU']) results.push(fillForm('SKU', data['SKU']));
        if (data['采购价（元）']) results.push(fillForm('采购成本', data['采购价（元）']));
        results.push(fill('商品规格', data['长 包装规格（cm）'], '长', '商品包装规格'));
        results.push(fill('商品规格', data['宽 包装规格（cm)'], '宽', '商品包装规格'));
        results.push(fill('商品规格', data['高 包装规格（cm）'], '高', '商品包装规格'));
        const productInputs = inputsBetween('商品规格', '商品包装规格');
        const weightInput = productInputs.find(i => !['长', '宽', '高'].includes((i.getAttribute('placeholder') || '').trim()));
        results.push({label: '商品规格-重量', found: !!weightInput, ok: setInput(weightInput, data['单品毛重（kg）']), value: weightInput ? weightInput.value : ''});
        results.push(...fillPackingTable(
          data['长 包装规格（cm）'],
          data['宽 包装规格（cm)'],
          data['高 包装规格（cm）'],
          data['单品毛重（kg）']
        ));
        return {before, results, after: snapshot()};
        """
        result = self.page.run_js(script, item or {})
        print("页面字段/补填快照:", flush=True)
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
        return result

    def watch_packets(self, seconds, item, title):
        try:
            if not self.page.listen.listening:
                try:
                    self.page.listen.start(self.listen_targets)
                except Exception:
                    self.page.listen.start(self.listen_targets[0])
        except Exception:
            try:
                self.page.listen.start(self.listen_targets[0])
            except Exception:
                pass

        end_time = time.time() + int(seconds)
        result = {"success": False, "sku_exists": False, "specs_in_request": False, "messages": [], "urls": []}
        spec_values = []
        if item:
            spec_values = [
                str(item.get("长 包装规格（cm）") or "").strip(),
                str(item.get("宽 包装规格（cm)") or "").strip(),
                str(item.get("高 包装规格（cm）") or "").strip(),
                str(item.get("单品毛重（kg）") or "").strip(),
            ]

        print(f"开始监听：{title}，{seconds}s", flush=True)
        count = 0
        while time.time() <= end_time:
            packet = self.page.listen.wait(timeout=3, fit_count=False)
            if not packet:
                continue
            count += 1
            request = getattr(packet, "request", None)
            response = getattr(packet, "response", None)
            url = getattr(packet, "url", "") or getattr(request, "url", "")
            method = getattr(request, "method", "") if request else ""
            status = getattr(response, "status", "") or getattr(response, "status_code", "") if response else ""
            request_text = ""
            if request:
                request_text = str(
                    getattr(request, "postData", "")
                    or getattr(request, "post_data", "")
                    or getattr(request, "body", "")
                    or ""
                )
            response_body = getattr(response, "body", "") if response else ""
            response_text = json.dumps(response_body, ensure_ascii=False) if isinstance(response_body, (dict, list)) else str(response_body or "")
            print("\n" + "=" * 80, flush=True)
            print(f"[{title} #{count}] {method} {url} status={status}", flush=True)
            if request_text:
                print("[请求数据]", request_text[:3000], flush=True)
            if response_text:
                print("[响应数据]", response_text[:3000], flush=True)
            result["urls"].append(url)
            all_text = request_text + "\n" + response_text
            if "SKU已存在" in all_text or "sku已存在" in all_text or "已存在" in all_text:
                result["sku_exists"] = True
                result["success"] = True
            if spec_values and all(value and value in request_text for value in spec_values):
                result["specs_in_request"] = True
            try:
                parsed = json.loads(response_text)
            except Exception:
                parsed = None
            if isinstance(parsed, dict):
                msg = str(parsed.get("msg") or parsed.get("message") or parsed.get("errorMsg") or "")
                if msg:
                    result["messages"].append(msg)
                code = parsed.get("code")
                success = parsed.get("success")
                data = parsed.get("data")
                if success is True or str(code) in ("0", "200") or data is True:
                    result["success"] = True

        try:
            if self.page.listen.listening:
                self.page.listen.stop()
        except Exception:
            pass
        result["urls"] = list(dict.fromkeys(result["urls"]))
        result["messages"] = list(dict.fromkeys(result["messages"]))
        print(f"{title}监听结果:", flush=True)
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
        return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="赛狐 ERP 创建商品与配对接口监听调试")
    parser.add_argument(
        "--run-mode",
        choices=["debug", "formal"],
        default="debug",
        help="debug=接管当前页面；formal=先登录赛狐",
    )
    parser.add_argument(
        "--business-mode",
        choices=["low_price", "new_set", "variant", "mode_two", "mode_one", "mode_three", "mode2", "mode1", "mode3"],
        default="low_price",
        help="调试目标业务模式：low_price=低价商城；new_set=纯新品；variant=变体开发",
    )
    parser.add_argument(
        "--action",
        choices=["watch", "snapshot", "fill_specs", "fill_and_watch", "full_flow", "prepare_flow"],
        default="watch",
        help="watch=只监听；snapshot=字段快照；fill_specs=补填规格；fill_and_watch=补填后监听；prepare_flow=走到保存前停止；full_flow=完整创建配对",
    )
    parser.add_argument("--username", default=os.getenv("SAIHU_USERNAME", ""), help="赛狐账号")
    parser.add_argument("--password", default=os.getenv("SAIHU_PASSWORD", ""), help="赛狐密码")
    parser.add_argument("--browser-port", default="", help="已有浏览器调试端口，例如 9000 或 9222")
    parser.add_argument("--wait-seconds", default="120", help="监听等待秒数")
    parser.add_argument("--excel-path", default="", help="Excel 路径")
    parser.add_argument("--sku", default="", help="按 SKU 筛选")
    parser.add_argument("--asin", default="", help="按 ASIN 筛选")
    parser.add_argument("--limit", default="0", help="最多处理条数，0 表示不限制")
    parser.add_argument("--start-index", default="1", help="从第几条开始，1 表示第一条")
    parser.add_argument("--auto-save", action="store_true", help="fill_and_watch 模式补填后自动保存")
    parser.add_argument("--allow-any-status", action="store_true", help="测试用：不按未配对状态过滤")
    args = parser.parse_args()

    config = {
        "run_mode": args.run_mode,
        "business_mode": args.business_mode,
        "action": args.action,
        "username": args.username,
        "password": args.password,
        "browser_port": args.browser_port,
        "wait_seconds": args.wait_seconds,
        "excel_path": args.excel_path,
        "sku": args.sku,
        "asin": args.asin,
        "limit": args.limit,
        "start_index": args.start_index,
        "auto_save": args.auto_save,
        "allow_any_status": args.allow_any_status,
        "base_dir": Path(__file__).resolve().parent,
    }
    run = SaihuDebugTest(config)
    run.main()
