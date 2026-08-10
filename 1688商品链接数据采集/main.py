"""
1688 商品链接浏览器采集（方案 C）

从 Excel 读取「链接」列，用 DrissionPage 打开详情页并监听 XHR/JSONP 接口，
保存结构化 JSON。官方 Open API 见 api_1688.py（查自家店铺商品）。
"""

import json
import os
import re
import time
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage
from excel import ExcelDF
from mt import Mt
from email_util import PROJECT_NAME, deliver_outputs


class Ali1688:
    """1688 商品链接采集核心类。

    流程：Excel 读链接 → 浏览器监听接口抓 JSON → 按工作表落盘 → 汇总导出 Excel。
    page 为 None 时仅可调用 excel_df()（离线扫描已有 JSON）。
    """

    def __init__(self, page, config):
        """初始化采集配置；有 page 时一并创建多标签批采助手 Mt。

        Args:
            page: DrissionPage ChromiumPage，仅导出时可传 None。
            config: 含 file_path（输入 Excel）、path（JSON/汇总输出目录）、
                isOnline、sendEmail、email 等（与下载美国站子项目一致）。
        """
        self.page = page
        self.file_path = config['file_path']
        self.path = config['path']
        self.base_dir = Path(config['path'])
        self.is_online = bool(config.get('isOnline', False))
        self.send_email = bool(config.get('sendEmail', False))
        self.email = (config.get('email') or '').strip()
        self.sender_email = (config.get('sender_email') or '').strip()
        self.smtp_auth_code = (config.get('smtp_auth_code') or '').strip()
        # 详情页 XHR/JSONP 监听关键字（对应 detail_widget / sku_selector）
        self.LISTEN_URL_KEYWORDS = (
            'OfferDetailWidget.do',
            'WidgetOfferDetail.do',
            'queryofferskuselectormodel',
        )
        self.PAGE_WAIT = 3          # 打开页面后等待接口发出的秒数
        self.LISTEN_TIMEOUT = 30    # 单页监听最长等待
        self.LINK_GAP = 2           # 批次之间的间隔
        self.BATCH_SIZE = 2         # 每批并行打开的标签数
        self.mt = None
        if page is not None:
            self.mt = Mt(
                self.page,
                self.LISTEN_URL_KEYWORDS,
                page_wait=self.PAGE_WAIT,
                listen_timeout=self.LISTEN_TIMEOUT,
            )

    @property
    def output_excel_path(self):
        """规格汇总 Excel 的默认输出路径（{path}/规格汇总.xlsx）。"""
        return self.base_dir / '规格汇总.xlsx'

    def collect_output_files(self):
        """在输出目录中查找规格汇总 Excel 或文件名含项目名的文件。"""
        folder = self.base_dir
        if not folder.exists():
            print(f'输出目录不存在: {folder}')
            return []

        files = []
        if self.output_excel_path.is_file():
            files.append(self.output_excel_path)
        for p in folder.iterdir():
            if p.is_file() and PROJECT_NAME in p.name and p not in files:
                files.append(p)
        files = [str(p) for p in files]
        if files:
            print(f'在 {folder} 找到 {len(files)} 个可发送文件')
        else:
            print(f'在 {folder} 未找到规格汇总.xlsx 或含「{PROJECT_NAME}」的文件')
        return files

    def link_extract(self):
        """读取 Excel 并解析链接，返回可直接采集的条目列表。

        每条记录含：描述、工作表、原始链接、标准详情 link、offer_id、safe_sheet。
        offer_id 从 URL 的 offer/数字 或纯数字单元格提取。
        """
        rows = ExcelDF(self.file_path).read_excel()
        items = []
        total = len(rows)

        for i, row in enumerate(rows, 1):
            raw_link = row.get('链接', '').strip()
            sheet_name = row.get('工作表', '未命名')
            desc = row.get('描述', '')

            if not raw_link:
                print(f'[{i}/{total}] 跳过：链接为空（{sheet_name}）')
                continue

            text = str(raw_link).strip()
            match = re.search(r'offer/(\d+)', text)
            if match:
                offer_id = match.group(1)
            elif text.isdigit():
                offer_id = text
            else:
                print(f'[{i}/{total}] 跳过：无法从链接解析 offerId: {raw_link}')
                continue

            link = f'https://detail.1688.com/offer/{offer_id}.html'
            # 工作表名转安全目录名（去掉 Windows 非法字符）
            safe_sheet = re.sub(r'[<>:"/\\|?*]', '_', sheet_name).strip() or '未命名'

            items.append({
                '描述': desc,
                '工作表': sheet_name,
                '链接': raw_link,
                'link': link,
                'offer_id': offer_id,
                'safe_sheet': safe_sheet,
            })

        return items

    @staticmethod
    def error(msg=''):
        """采集失败时的空 payload 占位，可选附带 error 字段。"""
        payload = {
            'detail_widget': None,
            'sku_selector': None,
            'page_context': None,
            'api_urls': {},
        }
        if msg:
            payload['error'] = msg
        return payload

    def json_urls(self, link):
        """单标签顺序采集：打开详情页并监听接口。

        返回 detail_widget、sku_selector、page_context 及 api_urls 映射。
        收到 sku_selector 后由 Mt.captured_ready 提前结束监听；最后补充 window.context。
        """
        captured = {
            'detail_widget': None,
            'sku_selector': None,
            'page_context': None,
            'api_urls': {},
        }

        if self.page.listen.listening:
            self.page.listen.stop()

        self.page.listen.start(list(self.LISTEN_URL_KEYWORDS))
        self.page.get(link)
        time.sleep(self.PAGE_WAIT)

        deadline = time.time() + self.LISTEN_TIMEOUT
        while time.time() < deadline:
            packet = self.page.listen.wait(timeout=1, fit_count=False)
            if not packet:
                continue

            packets = packet if isinstance(packet, list) else [packet]
            for item in packets:
                url = getattr(item, 'url', '') or ''
                lower = url.lower()
                if 'offerdetailwidget.do' in lower or 'widgetofferdetail.do' in lower:
                    key = 'detail_widget'
                elif 'queryofferskuselectormodel' in lower:
                    key = 'sku_selector'
                else:
                    continue

                if captured[key] is not None:
                    continue

                body = item.response.body
                try:
                    if isinstance(body, (dict, list)):
                        parsed = body
                    else:
                        if isinstance(body, bytes):
                            text = body.decode('utf-8', errors='replace')
                        else:
                            text = str(body).strip()
                        if not text:
                            parsed = None
                        elif text.startswith('{') or text.startswith('['):
                            parsed = json.loads(text)
                        else:
                            # JSONP：callback({...});
                            match = re.search(r'^[^(]+\((.*)\)\s*;?\s*$', text, re.DOTALL)
                            parsed = json.loads(match.group(1)) if match else text
                    captured[key] = parsed
                    captured['api_urls'][key] = url
                except (json.JSONDecodeError, TypeError, ValueError) as exc:
                    captured['api_urls'][key] = f'{url} (parse error: {exc})'

            if Mt.captured_ready(captured):
                break

        if self.page.listen.listening:
            self.page.listen.stop()

        # page_context 常含 skuInfoMap，批采失败时可作为导出回退来源
        try:
            captured['page_context'] = self.page.run_js('return window.context || null')
        except Exception:
            captured['page_context'] = None
        return captured

    def output(self, item, payload, index):
        """组装 JSON 记录并写入 {path}/{safe_sheet}/{safe_sheet}{index}.json。"""
        safe_sheet = item['safe_sheet']
        save_dir = self.base_dir / safe_sheet
        save_dir.mkdir(parents=True, exist_ok=True)
        save_path = save_dir / f'{safe_sheet}{index}.json'

        record = {
            'meta': {
                'offer_id': item['offer_id'],
                'link': item['link'],
                '描述': item.get('描述', ''),
                '工作表': item.get('工作表', ''),
                'captured_at': datetime.now(timezone.utc).astimezone().isoformat(),
            },
            **payload,
        }

        with save_path.open('w', encoding='utf-8') as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        return save_path

    def data(self):
        """双标签批采；批采不可导出时回退单标签 json_urls 重试。

        成功判定以 Mt.payload_exportable 为准（含 page_context.skuInfoMap 回退）。
        """
        if self.mt is None:
            raise RuntimeError('采集需要浏览器页面，请先启动 Chromium')

        items = self.link_extract()
        if not items:
            print('Excel 中没有可采集的数据')
            return

        self.mt.prune_extra_tabs()
        counters = {}  # 各工作表内 JSON 序号
        total = len(items)
        run_success = 0
        run_fail = 0
        print(f'共 {total} 条链接待采集（每批 {self.BATCH_SIZE} 标签，失败自动单标签重试）')

        for start in range(0, total, self.BATCH_SIZE):
            batch = items[start:start + self.BATCH_SIZE]
            batch_end = start + len(batch)

            try:
                payloads = self.mt.json_urls_batch(batch)
            except Exception as exc:
                print(f'  本批采集失败: {exc}')
                payloads = {item['offer_id']: self.error(str(exc)) for item in batch}

            for j, item in enumerate(batch):
                i = start + j + 1
                oid = item['offer_id']
                sheet_name = item['工作表']
                counters[sheet_name] = counters.get(sheet_name, 0) + 1

                payload = payloads.get(oid, self.error('未采集到数据'))
                retried = False
                batch_exportable = Mt.payload_exportable(payload)
                # 批采未拿到可导出规格 → 关多余标签，单标签重试
                if not batch_exportable:
                    retried = True
                    print(f'[{i}/{total}] 批采未拿到规格，切换单标签 offer/{oid}')
                    self.mt.prune_extra_tabs()
                    self.mt.focus_main_tab()
                    try:
                        payload = self.json_urls(item['link'])
                    except Exception as exc:
                        payload = self.error(str(exc))
                elif not Mt.payload_complete(payload):
                    print(f'[{i}/{total}] 批采已拿到 page_context 规格 offer/{oid}')

                if Mt.payload_exportable(payload):
                    payload.pop('error', None)
                    run_success += 1
                    if retried:
                        if Mt.payload_complete(payload):
                            print('  单标签重试成功')
                        else:
                            print('  单标签重试成功（page_context 规格）')
                else:
                    err = payload.get('error') or '监听超时或数据不完整'
                    payload['error'] = err
                    run_fail += 1
                    print(f'  采集失败: {err}')

                print(
                    f'[{i}/{total}] 采集 offer/{oid} -> '
                    f'{item["safe_sheet"]}{counters[sheet_name]}.json'
                )
                save_path = self.output(item, payload, counters[sheet_name])
                print(f'  已保存: {save_path}')

            if batch_end < total:
                time.sleep(self.LINK_GAP)

        print(f'本次采集完成：成功 {run_success}/{total}，失败 {run_fail}/{total}')

    def excel_df(self):
        """扫描已保存 JSON，按 offer_id 去重后导出规格汇总 Excel。

        每个原始工作表对应一个 sheet；列宽等样式复用 ExcelDF.excel_columns。
        不需浏览器，page=None 时也可调用。
        """
        json_files = sorted(self.base_dir.glob('*/*.json'))
        if not json_files:
            print('未找到已采集的 JSON 文件')
            return

        rows_by_sheet = {}
        file_count = 0
        skip_count = 0
        by_offer = {}  # offer_id → 优先保留规格完整的那份 JSON

        for json_path in json_files:
            try:
                with json_path.open('r', encoding='utf-8') as f:
                    record = json.load(f)
            except (json.JSONDecodeError, OSError) as exc:
                print(f'跳过无法解析: {json_path} ({exc})')
                skip_count += 1
                continue

            meta = record.get('meta')
            if not meta:
                print(f'跳过缺少 meta: {json_path}')
                skip_count += 1
                continue

            offer_id = str(meta.get('offer_id', '')).strip()
            if not offer_id:
                print(f'跳过缺少 offer_id: {json_path}')
                skip_count += 1
                continue

            sku_map, sku_source = Mt.extract_sku_map(record)
            sku_data = Mt.selector_sku_data(record)
            complete = bool(sku_map)
            # 同一 offer 多文件时，优先保留 complete=True 的记录
            prev = by_offer.get(offer_id)
            if prev is None or (complete and not prev['complete']):
                by_offer[offer_id] = {
                    'path': json_path,
                    'record': record,
                    'meta': meta,
                    'sku_map': sku_map,
                    'sku_data': sku_data,
                    'sku_source': sku_source,
                    'complete': complete,
                }

        for offer_id, info in by_offer.items():
            if not info['complete']:
                print(f'跳过缺少规格数据: {info["path"]}')
                skip_count += 1
                continue

            file_count += 1
            record = info['record']
            meta = info['meta']
            sku_map = info['sku_map']
            sku_data = info['sku_data']

            # 从 page_context 取件重尺（pieceWeightScaleInfo）
            pack_index = {}
            page_ctx = record.get('page_context') or {}
            pack_list = (
                page_ctx.get('result', {})
                .get('data', {})
                .get('productPackInfo', {})
                .get('fields', {})
                .get('pieceWeightScale', {})
                .get('pieceWeightScaleInfo') or []
            )
            for pack_item in pack_list:
                sid = pack_item.get('skuId')
                if sid is not None:
                    pack_index[str(sid)] = pack_item

            freight = (sku_data.get('extraInfo') or {}).get('freightInfo') or {}
            sku_weight_map = freight.get('skuWeight') or {}  # 件重尺缺失时的重量回退

            sheet_name = meta.get('工作表', '未命名')
            safe_sheet = re.sub(r'[<>:"/\\|?*]', '_', sheet_name).strip() or '未命名'

            for sku_info in sku_map.values():
                sku_id = sku_info.get('skuId')
                sku_key = str(sku_id) if sku_id is not None else ''
                pack = pack_index.get(sku_key, {})

                length = pack.get('length')
                width = pack.get('width')
                height = pack.get('height')
                if length in (None, '', 0):
                    length = '无具体尺寸信息'
                if width in (None, '', 0):
                    width = '无具体尺寸信息'
                if height in (None, '', 0):
                    height = '无具体尺寸信息'

                weight_val = pack.get('weight')
                if weight_val not in (None, '', 0):
                    weight = f'{weight_val}g'
                else:
                    sw = sku_weight_map.get(sku_key)
                    if sw not in (None, '', 0):
                        weight = f'{sw}kg'
                    else:
                        weight = '无重量信息'

                rows_by_sheet.setdefault(safe_sheet, []).append({
                    'offer_id': meta.get('offer_id', ''),
                    '链接': meta.get('link', ''),
                    '描述': meta.get('描述', ''),
                    '规格信息': sku_info.get('specAttrs', ''),
                    '长': length,
                    '宽': width,
                    '高': height,
                    '重量': weight,
                    '原价': sku_info.get('price', ''),
                    '优惠价': sku_info.get('discountPrice', ''),
                })

        if not rows_by_sheet:
            print('没有可导出的规格数据')
            return

        output_path = self.output_excel_path
        total_rows = 0
        excel_fmt = ExcelDF(self.file_path)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for safe_sheet, rows in rows_by_sheet.items():
                sheet_label = safe_sheet[:31]  # Excel sheet 名最长 31 字符
                pd.DataFrame(rows).to_excel(
                    writer,
                    sheet_name=sheet_label,
                    index=False,
                )
                ws = writer.sheets[sheet_label]
                excel_fmt.excel_columns(ws)
                total_rows += len(rows)

        print(f'已扫描 {file_count} 个 JSON，跳过 {skip_count} 个，共 {total_rows} 行规格')
        print(f'已导出: {output_path}')

    def run(self):
        """一键执行：data() 采集 JSON，再 excel_df() 导出；可选发送邮件。"""
        env = '线上' if self.is_online else '线下'
        print(f'运行环境：{env}')
        print('=== 采集 1688 商品链接数据 ===')
        self.data()
        print('=== 导出规格汇总 ===')
        self.excel_df()
        if not self.send_email:
            print('未启用邮件发送，流程结束')
            return

        output_files = self.collect_output_files()
        deliver_outputs(
            {
                'sendEmail': True,
                'email': self.email,
                'sender_email': self.sender_email,
                'smtp_auth_code': self.smtp_auth_code,
            },
            output_files,
        )


if __name__ == '__main__':
    # CLI 入口：config 仅在此处定义，供直接 python main.py 使用
    config = {
        'file_path': r'C:\Users\admin\Desktop\压体积包装_分类汇总分享.xlsx',
        'path': r'C:\Users\admin\Desktop\1688商品链接数据采集',
        'isOnline': False,
        'sendEmail': False,
        'email': '',
        'sender_email': os.getenv('SMTP_SENDER', '1974419863@qq.com'),
        'smtp_auth_code': os.getenv('SMTP_AUTH_CODE', ''),
    }
    page = ChromiumPage()
    ali1688 = Ali1688(page=page, config=config)
    ali1688.run()
