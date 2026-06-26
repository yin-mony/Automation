"""1688 详情页多标签批采与 payload 解析工具。

供 main.Ali1688 使用：并行打开多个标签监听 XHR/JSONP，
并从 sku_selector 或 page_context 提取可导出的规格映射。
"""

import json
import re
import time


class Mt:
    """双标签批采：每标签独立 listen，规格判定与标签生命周期管理。"""

    def __init__(self, page, listen_url_keywords, page_wait=3, listen_timeout=30):
        """绑定浏览器页面对象与监听参数。

        Args:
            page: DrissionPage ChromiumPage。
            listen_url_keywords: 需监听的 URL 关键字列表。
            page_wait: 批量打开链接后等待接口发出的秒数。
            listen_timeout: 单标签监听超时。
        """
        self.page = page
        self.listen_url_keywords = listen_url_keywords
        self.page_wait = page_wait
        self.listen_timeout = listen_timeout

    @staticmethod
    def listen_wait_safe(listen_owner, timeout=1):
        """包装 listen.wait；规避 DrissionPage 连接断开时 fail 未定义的 bug。"""
        try:
            return listen_owner.listen.wait(timeout=timeout, fit_count=False)
        except UnboundLocalError:
            return False

    @staticmethod
    def empty_captured():
        """返回空的采集结果结构（与 main.json_urls 一致）。"""
        return {
            'detail_widget': None,
            'sku_selector': None,
            'page_context': None,
            'api_urls': {},
        }

    @staticmethod
    def parse_response_body(body):
        """将响应体解析为 dict/list，支持 JSON 与 JSONP callback(...)。"""
        if isinstance(body, (dict, list)):
            return body
        if isinstance(body, bytes):
            text = body.decode('utf-8', errors='replace')
        else:
            text = str(body).strip()
        if not text:
            return None
        if text.startswith('{') or text.startswith('['):
            return json.loads(text)
        # JSONP：去掉外层函数名与括号
        match = re.search(r'^[^(]+\((.*)\)\s*;?\s*$', text, re.DOTALL)
        return json.loads(match.group(1)) if match else text

    @staticmethod
    def packet_key(url):
        """根据 URL 判断包类型：detail_widget / sku_selector / None。"""
        lower = (url or '').lower()
        if 'offerdetailwidget.do' in lower or 'widgetofferdetail.do' in lower:
            return 'detail_widget'
        if 'queryofferskuselectormodel' in lower:
            return 'sku_selector'
        return None

    @staticmethod
    def captured_ready(captured):
        """收到 sku_selector 即可提前结束监听（不必等 detail_widget）。"""
        return captured.get('sku_selector') is not None

    @classmethod
    def payload_complete(cls, payload):
        """是否含 sku_selector.originalSkuInfoMap（最完整规格来源）。"""
        sku_selector = (payload or {}).get('sku_selector') or {}
        sku_data = (sku_selector.get('data') or {}).get('skuSelectorBizModel') or {}
        return bool(sku_data.get('originalSkuInfoMap'))

    @staticmethod
    def _collect_page_sku_maps(obj, found=None):
        """递归扫描 page_context，收集所有形似 skuInfoMap 的字典。"""
        if found is None:
            found = []
        if isinstance(obj, dict):
            sku_map = obj.get('skuInfoMap')
            if isinstance(sku_map, dict) and sku_map:
                sample = next(iter(sku_map.values()), None)
                # 通过 specAttrs/skuId 字段粗判是否为规格条目
                if isinstance(sample, dict) and (
                    sample.get('specAttrs') is not None or sample.get('skuId') is not None
                ):
                    found.append(sku_map)
            for val in obj.values():
                Mt._collect_page_sku_maps(val, found)
        elif isinstance(obj, list):
            for item in obj:
                Mt._collect_page_sku_maps(item, found)
        return found

    @classmethod
    def sku_map_from_selector(cls, record):
        """从 sku_selector 取 originalSkuInfoMap。"""
        sku_selector = (record or {}).get('sku_selector') or {}
        sku_data = (sku_selector.get('data') or {}).get('skuSelectorBizModel') or {}
        return sku_data.get('originalSkuInfoMap')

    @classmethod
    def sku_map_from_page_context(cls, record):
        """从 page_context 取最大的 skuInfoMap（批采缺 selector 时的回退）。"""
        page_ctx = (record or {}).get('page_context') or {}
        maps = cls._collect_page_sku_maps(page_ctx)
        if not maps:
            return None
        return max(maps, key=len)

    @classmethod
    def extract_sku_map(cls, record):
        """优先 sku_selector.originalSkuInfoMap，否则回退 page_context.skuInfoMap。

        Returns:
            (sku_map, source)：source 为 'sku_selector' | 'page_context' | None。
        """
        sku_map = cls.sku_map_from_selector(record)
        if sku_map:
            return sku_map, 'sku_selector'
        sku_map = cls.sku_map_from_page_context(record)
        if sku_map:
            return sku_map, 'page_context'
        return None, None

    @classmethod
    def payload_exportable(cls, payload):
        """任一路径能提取到 sku_map 即视为可导出（用于批采成功/回退判定）。"""
        sku_map, _ = cls.extract_sku_map(payload)
        return bool(sku_map)

    @classmethod
    def selector_sku_data(cls, record):
        """取 skuSelectorBizModel 节点（含运费等 extraInfo）。"""
        sku_selector = (record or {}).get('sku_selector') or {}
        return (sku_selector.get('data') or {}).get('skuSelectorBizModel') or {}

    @classmethod
    def apply_packet(cls, captured, url, body):
        """将单条网络包解析后写入 captured（每种 key 只保留首包）。"""
        key = cls.packet_key(url)
        if not key or captured[key] is not None:
            return
        try:
            parsed = cls.parse_response_body(body)
            captured[key] = parsed
            captured['api_urls'][key] = url
        except (json.JSONDecodeError, TypeError, ValueError) as exc:
            captured['api_urls'][key] = f'{url} (parse error: {exc})'

    def focus_main_tab(self):
        """批采结束后切回主标签，供单标签回退采集使用。"""
        try:
            self.page.get_tab(1)
        except Exception:
            try:
                self.page.activate_tab(1)
            except Exception:
                pass

    def prune_extra_tabs(self):
        """关闭除主标签外的所有标签（CDP Target.closeTarget）。"""
        main_id = getattr(self.page, 'tab_id', None)
        if not main_id:
            return
        for tid in list(self.page.tab_ids):
            if tid == main_id:
                continue
            try:
                self.page.browser._run_cdp('Target.closeTarget', targetId=tid)
            except Exception:
                pass

    def close_batch_tabs(self, tab_entries):
        """关闭本批新建的标签，先聚焦主标签再逐个 closeTarget。"""
        self.focus_main_tab()
        for _oid, tab, _link in tab_entries:
            tid = getattr(tab, 'tab_id', None)
            try:
                tab._run_cdp('Target.closeTarget', targetId=tid)
            except Exception:
                pass

    def collect_tab_listen(self, tab, captured):
        """在单个标签上循环 wait 监听，直到超时或 captured_ready。"""
        deadline = time.time() + self.listen_timeout
        while time.time() < deadline and not self.captured_ready(captured):
            packet = self.listen_wait_safe(tab, timeout=1)
            if not packet:
                continue
            packets = packet if isinstance(packet, list) else [packet]
            for pkt in packets:
                url = getattr(pkt, 'url', '') or ''
                if not self.packet_key(url):
                    continue
                self.apply_packet(captured, url, pkt.response.body)

    def json_urls_batch(self, batch_items):
        """双标签批采：每商品一新标签独立 listen，最后补 page_context。

        失败或不可导出的条目由 main.data() 回退单标签 json_urls 重试。
        """
        if not batch_items:
            return {}

        captured_map = {item['offer_id']: self.empty_captured() for item in batch_items}
        tab_entries = []

        if self.page.listen.listening:
            self.page.listen.stop()

        try:
            # 1. 建标签并启动各自监听
            for item in batch_items:
                tab = self.page.new_tab()
                tab.listen.start(list(self.listen_url_keywords))
                tab_entries.append((item['offer_id'], tab, item['link']))

            # 2. 并行导航到详情页
            for _oid, tab, link in tab_entries:
                tab.get(link)

            time.sleep(self.page_wait)

            # 3. 逐标签收包并读取 window.context
            for oid, tab, _link in tab_entries:
                self.collect_tab_listen(tab, captured_map[oid])
                if tab.listen.listening:
                    tab.listen.stop()
                try:
                    captured_map[oid]['page_context'] = tab.run_js(
                        'return window.context || null',
                    )
                except Exception:
                    captured_map[oid]['page_context'] = None
        finally:
            # 4. 无论成败都关闭批采标签，避免残留标签干扰下一批
            if tab_entries:
                self.close_batch_tabs(tab_entries)

        return captured_map
