#!/usr/bin/env python3
"""
续跑脚本 — 读取 broker_company_contact_queue.csv，对 pending_research 条目
执行 Google 搜索，提取邮箱/电话，原地更新 CSV。

用法:
    python resume.py                                         # 默认处理 100 条
    python resume.py --limit 500                             # 处理 500 条
    python resume.py --limit 0                               # 处理全部
    python resume.py --proxy-file proxies.txt                # 加载代理文件
    python resume.py --output-dir data                       # 指定输出目录
    python resume.py --resume                                # 仅跳过已完成的
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set
from urllib.parse import quote, urljoin, urlparse

from ippool import ProxyPool

try:
    from DrissionPage import ChromiumOptions, ChromiumPage
except ImportError:
    ChromiumOptions = None
    ChromiumPage = None


# ── 常量 ──────────────────────────────────────────────────────

DEFAULT_OUTPUT_DIR = Path("data")
DEFAULT_LIMIT = 100
CSV_FILENAME = "broker_company_contact_queue.csv"

DEEP_CRAWL_MAX_PAGES = 3

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/126.0.0.0 Safari/537.36"
)

BLOCKED_EMAIL_DOMAINS: Set[str] = {
    "example.com", "test.com", "domain.com", "yourdomain.com",
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com",
    "aol.com", "icloud.com", "protonmail.com", "mail.com",
}

BLOCKED_EMAIL_TLDS: Set[str] = {
    "png", "jpg", "jpeg", "gif", "svg", "webp",
    "pdf", "doc", "docx", "xls", "xlsx",
}

GENERIC_EMAIL_PREFIXES: Set[str] = {
    "info", "contact", "hello", "office", "support", "admin",
    "noreply", "no-reply", "sales", "service", "team",
}


# ── 邮箱/电话提取 ────────────────────────────────────────────

def extract_emails_and_phones(text: str) -> Dict[str, Any]:
    """从页面文本提取邮箱和电话。"""
    emails_raw = re.findall(
        r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text
    )
    phones_raw = re.findall(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text)

    seen: Set[str] = set()
    company_emails: List[str] = []
    broker_emails: List[str] = []

    for email in emails_raw:
        email_lower = email.lower().strip()
        if email_lower in seen:
            continue

        domain = email_lower.split("@")[-1] if "@" in email_lower else ""
        tld = domain.split(".")[-1] if "." in domain else ""

        if domain in BLOCKED_EMAIL_DOMAINS:
            continue
        if tld in BLOCKED_EMAIL_TLDS:
            continue

        seen.add(email_lower)

        prefix = email_lower.split("@")[0]
        if prefix in GENERIC_EMAIL_PREFIXES or _is_likely_personal(email_lower):
            company_emails.append(email_lower)
        else:
            broker_emails.append(email_lower)

    phones: List[str] = []
    seen_phones: Set[str] = set()
    for phone in phones_raw:
        cleaned = re.sub(r"[^\d]", "", phone)
        if len(cleaned) == 10 and cleaned not in seen_phones:
            seen_phones.add(cleaned)
            phones.append(cleaned)

    return {
        "emails": list(seen),
        "company_emails": company_emails,
        "broker_emails": broker_emails,
        "phones": phones,
    }


def _is_likely_personal(email: str) -> bool:
    local = email.split("@")[0].lower() if "@" in email else ""
    if re.match(r"^(info|contact|hello|office|support|admin)", local):
        return True
    if re.match(r"^\d{3,}", local):
        return True
    return False


def company_matches(text: str, company_name: str) -> bool:
    """检查公司名是否出现在页面文本中（大小写不敏感，容差匹配）。

    这是置信度核心逻辑：只要 TREC 公司名在页面文本中出现，
    就认为该页面与目标公司相关，不再依赖 broker 人名。
    """
    if not company_name or not text:
        return False
    company_norm = re.sub(r'\s+', ' ', company_name.lower().strip())
    text_norm = re.sub(r'\s+', ' ', text.lower())
    return company_norm in text_norm


# ── Google 搜索（带 DrissionPage） ────────────────────────────

def search_google(page: ChromiumPage, query: str) -> str:
    """在 Google 搜索关键词，返回页面纯文本。"""
    url = "https://www.google.com/search?q=" + quote(query) + "&hl=en"
    page.get(url)
    time.sleep(random.uniform(3.5, 6.5))

    # 等待可能的 loading 消失
    try:
        page.ele("x://div[@role='progressbar']", timeout=3)
        time.sleep(2)
    except Exception:
        pass

    body = page("x:/html/body")
    return body.text if body else ""



def _extract_search_result_links(page: ChromiumPage) -> List[str]:
    """从 Google 搜索结果页提取自然搜索结果的 URL。"""
    urls = []
    try:
        for a in page.eles("tag:a"):
            href = a.attr("href")
            if href and href.startswith("http") and "google.com" not in href:
                text = (a.text or "").strip()
                if len(text) > 3:
                    urls.append(href)
    except Exception:
        return []
    seen: Set[str] = set()
    unique = []
    for u in urls:
        norm = u.rstrip("/").lower()
        if norm not in seen:
            seen.add(norm)
            unique.append(u)
    return unique[:3]


def _find_contact_links(page: ChromiumPage, current_url: str) -> List[str]:
    """在当前页面找 Contact / About 等子页面链接（文本 + URL 路径双匹配）。"""
    parsed = urlparse(current_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    # 链接文本和 URL 路径都检测的关键词
    keywords = {"contact", "about", "location", "team", "support", "help"}
    url_keywords = {"contact", "about", "team", "location", "office"}
    links = []
    seen = set()
    try:
        for a in page.eles("tag:a"):
            text = (a.text or "").strip().lower()
            href = a.attr("href")
            if not href or href.startswith("javascript") or href.startswith("#"):
                continue
            full = urljoin(base, href).rstrip("/")
            if full in seen:
                continue
            seen.add(full)
            parsed_full = urlparse(full)
            if parsed_full.netloc != parsed.netloc:
                continue
            if full == current_url.rstrip("/"):
                continue
            # 匹配条件：链接文本含关键词 OR URL 路径含关键词
            path_match = any(kw in parsed_full.path.lower() for kw in url_keywords)
            text_match = any(kw in text for kw in keywords)
            if text_match or path_match:
                links.append(full)
    except Exception:
        pass
    return links[:5]


def crawl_website_for_contacts(
    page: ChromiumPage,
    company_name: str,
    max_pages: int = DEEP_CRAWL_MAX_PAGES,
) -> Dict[str, Any]:
    """从 Google 结果页提取链接 → 深度访问网站页面提取联系方式。

    访问每个页面时先用 company_matches 确认页面与目标公司相关，
    不匹配的页面直接跳过，避免污染数据。
    返回值加入 source_type: 'serp' | 'deep' | 'subpage'，
    方便调用者区分数据来源层级。
    """
    search_urls = _extract_search_result_links(page)
    if not search_urls:
        print(f"  深度爬取: 未找到搜索结果链接", flush=True)
        return {"emails": [], "phones": [], "source_type": "serp"}

    print(f"  深度爬取: 找到 {len(search_urls)} 个候选 URL", flush=True)

    all_emails: Set[str] = set()
    all_phones: Set[str] = set()
    visited: Set[str] = set()
    pending = list(search_urls)
    # 记录邮箱是从哪种页面找到的（deep=搜索结果直访, subpage=contact子页）
    source_from_deep_page = False

    while pending and len(visited) < max_pages:
        url = pending.pop(0)
        if url in visited:
            continue
        visited.add(url)

        try:
            page.set.timeouts(page_load=15)
            page.get(url)
            time.sleep(random.uniform(2, 4))

            body = page("x:/html/body")
            text = body.text if body else ""
            if not text:
                continue

            # ★ 公司名校验：不匹配直接跳过，不浪费时间找联系方式
            if not company_matches(text, company_name):
                print(f"  深度爬取 跳过 {url}: 公司名不匹配", flush=True)
                continue

            result = extract_emails_and_phones(text)
            if result.get("emails"):
                source_from_deep_page = True
            for e in result.get("emails", []):
                all_emails.add(e)
            for p in result.get("phones", []):
                all_phones.add(p)

            # 扩展爬取 contact/about 子页面并标记来源
            contact_links = _find_contact_links(page, url)
            for link in contact_links:
                if link not in visited and link not in pending:
                    pending.append(link)
        except Exception as exc:
            print(f"  深度爬取 跳过 {url}: {type(exc).__name__}", flush=True)
            continue

    source_type = "subpage" if (source_from_deep_page and len(visited) > 1) else "deep" if source_from_deep_page else "serp"
    return {
        "emails": list(all_emails),
        "phones": list(all_phones),
        "source_type": source_type,
    }

def broker_search(page: ChromiumPage, broker_name: str, company_name: str) -> Dict[str, Any]:
    """组合 broker/company 搜索，降级为公司名兜底 + 深度网站爬取。"""
    if broker_name and company_name:
        query = f'"{broker_name}" "{company_name}" email phone'
    elif broker_name:
        query = f'"{broker_name}" real estate broker email phone Texas'
    else:
        return {"emails": [], "phones": [], "company_emails": [], "broker_emails": []}

    # ── 阶段 1：Google 搜索 ──
    text = search_google(page, query)
    result = extract_emails_and_phones(text)
    print(f"  broker search  emails={len(result['emails'])} phones={len(result['phones'])}",
          flush=True)

    # 无结果时降级为公司名搜索
    if not result["emails"] and company_name:
        fallback_query = f'"{company_name}" "office email"'
        print(f"  降级搜索: {fallback_query}", flush=True)
        text2 = search_google(page, fallback_query)
        result2 = extract_emails_and_phones(text2)

        result["emails"] = result["emails"] or result2["emails"]
        result["phones"] = result["phones"] or result2["phones"]
        result["company_emails"] = result["company_emails"] or result2["company_emails"]
        result["broker_emails"] = result["broker_emails"] or result2["broker_emails"]

    # ── 阶段 2：深度网站爬取（仅 company_name 存在时） ──
    if company_name:
        deep = crawl_website_for_contacts(page, company_name)
        existing_emails = set(result["emails"])
        existing_phones = set(result["phones"])
        new_emails = [e for e in deep.get("emails", []) if e not in existing_emails]
        new_phones = [p for p in deep.get("phones", []) if p not in existing_phones]
        result["emails"].extend(new_emails)
        result["phones"].extend(new_phones)
        print(f"  深度爬取 新增 emails={len(new_emails)} phones={len(new_phones)}", flush=True)

    return result


# ── CSV 读写 ──────────────────────────────────────────────────

def read_queue(path: Path) -> List[Dict[str, str]]:
    if not path.exists():
        print(f"[错误] CSV 不存在: {path}", flush=True)
        sys.exit(1)
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_queue(path: Path, rows: List[Dict[str, str]], fieldnames: List[str]) -> None:
    tmp = path.with_suffix(".csv.tmp")
    with tmp.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    tmp.replace(path)


# ── 核心流水线 ───────────────────────────────────────────────

def process_batch(
    rows: List[Dict[str, str]],
    fieldnames: List[str],
    path: Path,
    limit: int,
    pool: ProxyPool,
) -> Dict[str, int]:
    """处理一批待搜索条目，原地更新 CSV。"""
    total_pending = len(rows)
    to_process = rows[:limit] if limit > 0 else rows[:]
    print(f"待搜索 {total_pending} 条，本轮处理 {len(to_process)} 条", flush=True)

    if not to_process:
        return {"processed": 0, "found_email": 0, "found_phone": 0, "errors": 0}

    # ── 启动浏览器 ──
    browser_opts = ChromiumOptions()
    browser_opts.set_argument("--incognito")
    browser_opts.set_user_agent(USER_AGENT)

    # 尝试设置代理
    proxy_dict = pool.get_proxy()
    applied_proxy = None
    if proxy_dict is not None:
        proxy_url = proxy_dict.get("http", "")
        if proxy_url:
            try:
                browser_opts.set_proxy(proxy_url)
                applied_proxy = proxy_url
                print(f"使用代理: {proxy_url}", flush=True)
            except Exception:
                print("代理设置失败，使用直连", flush=True)

    page = ChromiumPage(browser_opts)
    page.get("https://www.google.com")
    time.sleep(2)

    processed = 0
    found_email = 0
    found_phone = 0
    errors = 0

    try:
        for idx, row in enumerate(to_process):
            company = row.get("company_name", "")
            broker = row.get("broker_name", "")
            company_license = row.get("company_license_number", "")

            print(f"\n[{idx + 1}/{len(to_process)}] {company} | {broker}", flush=True)

            try:
                result = broker_search(page, broker, company)
            except Exception as e:
                print(f"  [错误] {e}", flush=True)
                pool.mark_bad(proxy_dict)
                errors += 1
                # 尝试换代理重新启动浏览器
                proxy_dict = pool.get_proxy()
                if proxy_dict:
                    try:
                        new_proxy = proxy_dict.get("http", "")
                        browser_opts.set_proxy(new_proxy)
                        page.quit()
                        page = ChromiumPage(browser_opts)
                        page.get("https://www.google.com")
                        time.sleep(2)
                        applied_proxy = new_proxy
                        print(f"更换代理: {new_proxy}", flush=True)
                    except Exception:
                        pass
                continue

            # 更新行数据
            if result["emails"]:
                row["office_email"] = "; ".join(result["company_emails"]) if result["company_emails"] else ""
                row["broker_email"] = "; ".join(result["broker_emails"]) if result["broker_emails"] else ""
                row["source_url"] = f"https://www.google.com/search?q={quote(f'{broker} {company} email phone')}"
                found_email += 1

            if result["phones"]:
                row["phone"] = "; ".join(result["phones"])
                found_phone += 1

            if result["emails"] or result["phones"]:
                row["contact_status"] = "contact_found"
            else:
                row["contact_status"] = "no_contact_found"

            # 每处理完一条就写回 CSV（原地续跑安全）
            write_queue(path, rows, fieldnames)

            processed += 1
            pool.mark_ok(proxy_dict)

            if idx < len(to_process) - 1:
                delay = random.uniform(3, 6)
                print(f"  等待 {delay:.1f}s ...", flush=True)
                time.sleep(delay)

    finally:
        page.quit()

    return {
        "processed": processed,
        "found_email": found_email,
        "found_phone": found_phone,
        "errors": errors,
    }


# ── CLI ───────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="续跑脚本 — 搜索 broker/company 联系方式")
    p.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR,
                   help="CSV 输出目录（默认 data）")
    p.add_argument("--limit", type=int, default=DEFAULT_LIMIT,
                   help="处理条数，0=全部（默认 100）")
    p.add_argument("--proxy-file", type=Path, default=None,
                   help="代理文件路径，每行一条")
    p.add_argument("--resume", action="store_true",
                   help="跳过已处理条目（contact_status != pending_research）")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    output_dir: Path = args.output_dir
    csv_path = output_dir / CSV_FILENAME

    if not csv_path.exists():
        print(f"[错误] 找不到 {csv_path}，请先运行 main.py", flush=True)
        return 1

    # 1. 读取 CSV
    rows = read_queue(csv_path)
    fieldnames = list(rows[0].keys()) if rows else []
    print(f"读取 {len(rows)} 条记录", flush=True)

    # 2. 筛选待处理
    if args.resume:
        pending = [r for r in rows if r.get("contact_status", "").strip() == "pending_research"]
        print(f"跳过已完成条目，剩余待处理 {len(pending)} 条", flush=True)
    else:
        pending = rows

    if not pending:
        print("没有待处理条目", flush=True)
        return 0

    # 3. 加载代理池
    pool = ProxyPool()
    if args.proxy_file:
        loaded = pool.load_from_file(args.proxy_file)
        print(f"加载代理 {loaded} 条", flush=True)
        if loaded == 0:
            print("[警告] 代理文件为空或格式无效，将使用直连", flush=True)
    else:
        print("未指定代理文件，将使用直连", flush=True)

    # 4. 根据 resume 情况决定 limit
    #    如果 --resume，limit 仍限制本次处理数
    limit = args.limit if args.limit > 0 else len(pending)

    # 5. 执行
    summary = process_batch(pending, fieldnames, csv_path, limit, pool)

    # 6. 输出汇总
    print("\n" + "=" * 50, flush=True)
    print("执行汇总:", flush=True)
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    print("=" * 50, flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
