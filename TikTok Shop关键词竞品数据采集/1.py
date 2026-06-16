import html
import json
import re
from http.cookies import SimpleCookie
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urljoin, urlparse

import requests


PRODUCT_URL = 'https://shop.tiktok.com/us/pdp/hand-sanitizer-power-mist-45ml-scented-hydrating-spray-for-kids-travel-size-for-adults-perfect-easte/1731009227005006008'

QUERY_PARAMS = {
    'source': 'ecommerce_searchresult',
    'enter_method': 'feed_list_search_word',
    'first_entrance': 'ecommerce_mall',
    'first_entrance_position': 'search',
    'first_entrance_tt_scene': 'seo',
}

# 如需复用浏览器 cookie，可以从 DevTools 复制后填到这里。
COOKIES = {}

COOKIE_JSON = Path(__file__).resolve().parent / 'cookies.json'
COOKIE_TXT = Path(__file__).resolve().parent / 'cookies.txt'

HEADERS = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'cache-control': 'max-age=0',
    'referer': 'https://shop.tiktok.com/us/s?q=Hand+Sanitizer+Spray&source=ecommerce_searchresult&enter_method=search&first_entrance=ecommerce_mall&first_entrance_position=search&first_entrance_tt_scene=seo',
    'sec-ch-ua': '"Google Chrome";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
}

OUT_DIR = Path(__file__).resolve().parent / 'request_analysis'
JSON_SCRIPT_DIR = OUT_DIR / 'json_scripts'
JS_DIR = OUT_DIR / 'js_files'

FIELD_KEYWORDS = (
    'shop',
    'seller',
    'store',
    'merchant',
    'product',
    'price',
    'title',
    'sold',
    'sale',
    'rating',
    'review',
    'sku',
)

API_KEYWORDS = (
    'api',
    'product',
    'pdp',
    'shop',
    'seller',
    'store',
    'merchant',
    'order',
    'review',
    'rating',
    'oec',
)


def build_product_url():
    separator = '&' if '?' in PRODUCT_URL else '?'
    return f'{PRODUCT_URL}{separator}{urlencode(QUERY_PARAMS)}'


def load_cookie_file():
    if COOKIE_JSON.exists():
        data = json.loads(COOKIE_JSON.read_text(encoding='utf-8'))
        if isinstance(data, dict):
            return data
        if isinstance(data, list):
            return {
                item.get('name'): item.get('value')
                for item in data
                if isinstance(item, dict) and item.get('name') and item.get('value') is not None
            }

    if COOKIE_TXT.exists():
        text = COOKIE_TXT.read_text(encoding='utf-8').strip()
        cookie = SimpleCookie()
        cookie.load(text)
        return {key: morsel.value for key, morsel in cookie.items()}

    return {}


def save_text(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8', errors='ignore')


def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def get_attr_map(tag_text):
    attrs = {}
    for key, _, value in re.findall(r'([\w:-]+)\s*=\s*(["\'])(.*?)\2', tag_text, re.S):
        attrs[key.lower()] = html.unescape(value)
    return attrs


def get_page_title(text):
    match = re.search(r'<title[^>]*>(.*?)</title>', text, re.S | re.I)
    if not match:
        return ''
    return html.unescape(re.sub(r'\s+', ' ', match.group(1)).strip())


def is_security_page(text):
    lower_text = text.lower()
    return any(
        keyword in lower_text
        for keyword in (
            'security check',
            'verify to continue',
            'drag the puzzle piece',
            'captcha',
        )
    )


def parse_product_id(url):
    path = urlparse(url).path.strip('/')
    parts = path.split('/')
    if not parts:
        return ''
    last = parts[-1]
    return last if last.isdigit() else ''


def request_product_html(url):
    session = requests.Session()
    session.headers.update(HEADERS)

    cookies = {}
    cookies.update(COOKIES)
    cookies.update(load_cookie_file())
    print(f'使用 cookie 数量: {len(cookies)}')

    for name, value in cookies.items():
        session.cookies.set(name, value)

    response = session.get(url, timeout=25, allow_redirects=True)
    response.encoding = response.encoding or response.apparent_encoding
    return session, response


def extract_meta_tags(text):
    result = []
    for match in re.finditer(r'<meta\b[^>]*>', text, re.S | re.I):
        attrs = get_attr_map(match.group(0))
        name = attrs.get('name') or attrs.get('property') or attrs.get('itemprop')
        content = attrs.get('content')
        if name or content:
            result.append({
                'name': name,
                'content': content,
                'attrs': attrs,
            })
    return result


def try_load_json(text):
    text = html.unescape(text).strip()
    if not text or text[0] not in '[{':
        return None
    try:
        return json.loads(text)
    except Exception:
        return None


def preview(value, max_len=180):
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    else:
        value = str(value)
    return value[:max_len]


def find_keyword_fields(obj, path=''):
    hits = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            current = f'{path}.{key}' if path else str(key)
            if any(word in str(key).lower() for word in FIELD_KEYWORDS):
                hits.append({
                    'path': current,
                    'key': key,
                    'type': type(value).__name__,
                    'preview': preview(value),
                })
            hits.extend(find_keyword_fields(value, current))
    elif isinstance(obj, list):
        for index, value in enumerate(obj):
            hits.extend(find_keyword_fields(value, f'{path}[{index}]'))
    return hits


def extract_script_tags(text):
    scripts = []
    for index, match in enumerate(re.finditer(r'<script\b([^>]*)>(.*?)</script>', text, re.S | re.I), 1):
        attrs = get_attr_map(match.group(1))
        content = match.group(2).strip()
        src = attrs.get('src', '')
        script_type = attrs.get('type', '')
        script_id = attrs.get('id', '')
        json_data = None
        json_file = ''
        field_hits = []

        if content and (
            'json' in script_type.lower()
            or content[:1] in '[{'
            or script_id == '__MODERN_SSR_DATA__'
        ):
            json_data = try_load_json(content)

        if json_data is not None:
            json_file = f'script_{index:03d}_{script_id or "json"}.json'
            save_json(JSON_SCRIPT_DIR / json_file, json_data)
            field_hits = find_keyword_fields(json_data)

        scripts.append({
            'index': index,
            'id': script_id,
            'type': script_type,
            'src': src,
            'content_length': len(content),
            'json_file': json_file,
            'field_hits': field_hits[:80],
        })
    return scripts


def extract_urls(text, base_url):
    urls = set()

    for item in re.findall(r'https?://[^"\'<>)\\\s]+', text):
        urls.add(html.unescape(item))

    for item in re.findall(r'["\']((?:/|//)[^"\']{8,})["\']', text):
        item = html.unescape(item)
        if item.startswith('//'):
            urls.add(f'https:{item}')
        else:
            urls.add(urljoin(base_url, item))

    return sorted(urls)


def filter_api_hints(urls):
    result = []
    for url in urls:
        lower = url.lower()
        if any(word in lower for word in API_KEYWORDS):
            result.append(url)
    return sorted(set(result))


def fetch_and_analyze_js(session, scripts, base_url, max_files=20):
    hints = []
    fetched = 0

    for script in scripts:
        src = script.get('src')
        if not src:
            continue

        script_url = urljoin(base_url, src)
        if not script_url.startswith(('http://', 'https://')):
            continue

        try:
            response = session.get(script_url, timeout=15)
            response.raise_for_status()
        except Exception as e:
            hints.append({
                'script_url': script_url,
                'error': str(e),
                'api_hints': [],
            })
            continue

        fetched += 1
        js_text = response.text
        script_name = f'script_{fetched:03d}.js'
        save_text(JS_DIR / script_name, js_text)

        urls = extract_urls(js_text, script_url)
        api_hints = filter_api_hints(urls)
        hints.append({
            'script_url': script_url,
            'saved_file': script_name,
            'api_hints': api_hints[:200],
        })

        if fetched >= max_files:
            break

    return hints


def analyze_html_response(session, response):
    url = response.url
    text = response.text
    save_text(OUT_DIR / 'raw_response.html', text)

    parsed = urlparse(url)
    meta_tags = extract_meta_tags(text)
    scripts = extract_script_tags(text)
    urls = extract_urls(text, url)
    api_hints = filter_api_hints(urls)
    js_hints = fetch_and_analyze_js(session, scripts, url)

    summary = {
        'request_url': build_product_url(),
        'final_url': url,
        'status_code': response.status_code,
        'content_type': response.headers.get('content-type', ''),
        'encoding': response.encoding,
        'body_length': len(text),
        'title': get_page_title(text),
        'is_security_page': is_security_page(text),
        'product_id': parse_product_id(url),
        'path': parsed.path,
        'query': parse_qs(parsed.query),
        'meta_count': len(meta_tags),
        'script_count': len(scripts),
        'json_script_count': sum(1 for item in scripts if item.get('json_file')),
        'url_count': len(urls),
        'api_hint_count': len(api_hints),
        'js_file_hint_count': sum(len(item.get('api_hints', [])) for item in js_hints),
    }

    save_json(OUT_DIR / 'summary.json', summary)
    save_json(OUT_DIR / 'meta_tags.json', meta_tags)
    save_json(OUT_DIR / 'script_tags.json', scripts)
    save_json(OUT_DIR / 'discovered_urls.json', urls)
    save_json(OUT_DIR / 'api_hints_from_html.json', api_hints)
    save_json(OUT_DIR / 'api_hints_from_js.json', js_hints)

    return summary


def main():
    OUT_DIR.mkdir(exist_ok=True)
    url = build_product_url()
    print(f'请求商品地址: {url}')

    session, response = request_product_html(url)
    summary = analyze_html_response(session, response)

    print(f"状态码: {summary['status_code']}")
    print(f"最终地址: {summary['final_url']}")
    print(f"页面标题: {summary['title']}")
    print(f"是否验证码页: {summary['is_security_page']}")
    if summary['is_security_page']:
        print('当前 requests.get() 仍未拿到真实商品页，请先在浏览器完成验证，并把验证后的 cookie 写入 cookies.json 或 cookies.txt 后重试。')
    print(f"HTML 内疑似接口: {summary['api_hint_count']} 条")
    print(f"JS 内疑似接口: {summary['js_file_hint_count']} 条")
    print(f'分析结果目录: {OUT_DIR}')


if __name__ == '__main__':
    main()
