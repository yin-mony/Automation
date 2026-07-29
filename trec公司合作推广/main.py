from __future__ import annotations

import argparse
import csv
import json
import random
import ssl
import sys
import time
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

TREC_TYPESENSE_SEARCH_URL = (
    "https://www.trec.texas.gov/ts/collections/licenses/documents/search"
)
TREC_DETAIL_URL = "https://www.trec.texas.gov/license-search/?detail_id={detail_id}"
TREC_TYPESENSE_API_KEY = "HvqEl9eBZY6YjQBAU8uW4e9KBGHRvqrd"

TEXAS_OPEN_DATA_URL = "https://data.texas.gov/resource/s7ft-44qi.json"

EMAIL_SUBJECT = "Quick question about your agents' CE renewals"
EMAIL_BODY = """Hi,

I hope you're having a good week.

I'm QIANYI Ho from EVERTIX LLC. We just launched a Texas CE course package, and I wanted to reach out because I think it could be a good fit for your agents - especially with renewal season coming up.

Here's the short version:

We offer the full 18-hour Texas CE package for $86.99 - which I genuinely believe is the lowest price you'll find right now. It covers all the required courses including Legal Update I & II, contract forms, and the elective hours.

The reason I'm reaching out to you specifically is simple - I'd love to partner with you and your brokerage. For every agent in your office who uses our package, I'll give you a 20% referral fee back on each purchase. No complicated system, no hoops - just a straightforward split.

For you, it's an easy way to help your team save on CE costs while putting a little back in your pocket. For your agents, they get a solid course at a price that's hard to beat.

Happy to set you up with a demo or a free test course so you can see the quality before recommending it to anyone. No pressure at all - just figured it was worth asking.

Let me know if you'd be open to a quick chat about it.

Thanks,
QIANYI Ho"""


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
)


HREF_FIELDS = [
    "detail_id",
    "href",
    "primary_license_number",
    "primary_license_type",
    "primary_status",
    "display_name",
]

LICENSE_ROW_FIELDS = [
    "detail_id",
    "href",
    "license_number",
    "license_type",
    "status",
    "display_name",
    "first_name",
    "middle_name",
    "last_name",
    "organization_name",
]

RELATIONSHIP_FIELDS = [
    "record_type",
    "license_type",
    "license_number",
    "full_name",
    "status",
    "license_expiration_date",
    "related_license_type",
    "related_license_number",
    "related_license_full_name",
    "href",
    "related_href",
    "updated",
]

AGENT_SPONSOR_FIELDS = [
    "agent_name",
    "agent_license_number",
    "agent_href",
    "sponsoring_broker_type",
    "sponsoring_broker_name",
    "sponsoring_broker_license_number",
    "sponsoring_broker_href",
    "license_expiration_date",
    "updated",
]

BROKER_COMPANY_QUEUE_FIELDS = [
    "company_name",
    "company_license_number",
    "company_href",
    "broker_name",
    "broker_license_number",
    "broker_href",
    "google_company_query",
    "google_broker_query",
    "office_email",
    "broker_email",
    "phone",
    "source_url",
    "contact_status",
    "outreach_subject",
    "outreach_body",
]


def build_ssl_context(insecure: bool) -> Optional[ssl.SSLContext]:
    if insecure:
        return ssl._create_unverified_context()
    return None


def request_json(
    url: str,
    params: Optional[Dict[str, object]] = None,
    headers: Optional[Dict[str, str]] = None,
    *,
    timeout: int = 30,
    retries: int = 5,
    insecure: bool = False,
) -> object:
    query_url = url
    if params:
        query_url = f"{url}?{urlencode(params)}"

    request_headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    if headers:
        request_headers.update(headers)

    context = build_ssl_context(insecure)
    last_error: Optional[BaseException] = None

    for attempt in range(1, retries + 1):
        try:
            req = Request(query_url, headers=request_headers)
            with urlopen(req, timeout=timeout, context=context) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except HTTPError as exc:
            last_error = exc
            if exc.code not in {429, 500, 502, 503, 504}:
                body = exc.read().decode("utf-8", errors="replace")[:500]
                raise RuntimeError(
                    f"HTTP {exc.code} for {query_url}: {body}"
                ) from exc
        except (URLError, TimeoutError) as exc:
            last_error = exc

        sleep_seconds = min(60, (2 ** (attempt - 1)) + random.random())
        print(
            f"Request failed, retry {attempt}/{retries} in "
            f"{sleep_seconds:.1f}s: {last_error}",
            flush=True,
        )
        time.sleep(sleep_seconds)

    raise RuntimeError(f"Request failed after {retries} retries: {query_url}") from last_error


def append_csv(path: Path, fieldnames: List[str], rows: Iterable[Mapping[str, object]]) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    exists = path.exists() and path.stat().st_size > 0
    count = 0
    with path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        if not exists:
            writer.writeheader()
        for row in rows:
            writer.writerow(row)
            count += 1
    return count


def load_existing_values(path: Path, field: str) -> set:
    values = set()
    if not path.exists():
        return values
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            value = (row.get(field) or "").strip()
            if value:
                values.add(value)
    return values


def load_existing_license_row_keys(path: Path) -> set:
    values = set()
    if not path.exists():
        return values
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            license_number = (row.get("license_number") or "").strip()
            detail_id = (row.get("detail_id") or "").strip()
            if license_number or detail_id:
                values.add(f"{license_number}|{detail_id}")
    return values


def load_checkpoint(path: Path) -> Dict[str, int]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {key: int(value) for key, value in data.items() if isinstance(value, int)}


def write_checkpoint(path: Path, **values: int) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")


def reset_files(paths: Iterable[Path]) -> None:
    for path in paths:
        if path.exists():
            path.unlink()


def load_license_href_map(path: Path) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if not path.exists():
        return mapping
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            license_number = (row.get("license_number") or "").strip()
            href = (row.get("href") or "").strip()
            if license_number and href:
                mapping[license_number] = href
    return mapping


def nested_value(value: object, key: str) -> str:
    if isinstance(value, dict):
        return str(value.get(key) or "")
    return ""


def normalize_spaces(value: object) -> str:
    return " ".join(str(value or "").split())


def title_case_name(value: str) -> str:
    value = normalize_spaces(value)
    if not value:
        return ""
    return value.title()


def display_name_from_typesense_doc(doc: Dict[str, object]) -> str:
    organization_name = normalize_spaces(doc.get("organizationName"))
    if organization_name:
        return title_case_name(organization_name)

    last_name = normalize_spaces(doc.get("lastName"))
    first_name = normalize_spaces(doc.get("firstName"))
    middle_name = normalize_spaces(doc.get("middleName"))
    parts = [first_name, middle_name, last_name]
    return title_case_name(" ".join(part for part in parts if part))


def search_status_filter(mode: str) -> str:
    if mode == "active":
        return "status.value:=[Active,Inactive - Expired]"
    if mode == "strict-active":
        return "status.value:=[Active]"
    if mode == "all":
        return ""
    raise ValueError(f"Unsupported status mode: {mode}")


def typesense_search_params(page: int, page_size: int, status_mode: str) -> Dict[str, object]:
    filter_by = "type.alias:Real Estate && type.subType:!=BENT && status.value:!Upgraded"
    status_filter = search_status_filter(status_mode)
    if status_filter:
        filter_by += f" && {status_filter}"

    return {
        "q": "*",
        "query_by": (
            "customId,lastName,firstName,middleName,organizationName,"
            "teamNames,dbas,alternateNames"
        ),
        "group_by": "detailId",
        "sort_by": "_text_match:desc,lastName:asc,firstName:asc",
        "filter_by": filter_by,
        "highlight_fields": (
            "firstName, lastName, middleName, organizationName, teamNames, "
            "dbas, alternateNames"
        ),
        "page": page,
        "per_page": page_size,
        "max_candidates": 400000,
        "drop_tokens_threshold": 0,
    }


def row_from_typesense_doc(doc: Dict[str, object]) -> Dict[str, str]:
    detail_id = normalize_spaces(doc.get("detailId"))
    href = TREC_DETAIL_URL.format(detail_id=detail_id) if detail_id else ""
    license_type = nested_value(doc.get("type"), "subType")
    status = nested_value(doc.get("status"), "value")
    return {
        "detail_id": detail_id,
        "href": href,
        "license_number": normalize_spaces(doc.get("customId")),
        "license_type": license_type,
        "status": status,
        "display_name": display_name_from_typesense_doc(doc),
        "first_name": normalize_spaces(doc.get("firstName")),
        "middle_name": normalize_spaces(doc.get("middleName")),
        "last_name": normalize_spaces(doc.get("lastName")),
        "organization_name": normalize_spaces(doc.get("organizationName")),
    }


def collect_hrefs(
    output_dir: Path,
    *,
    page_size: int,
    start_page: Optional[int],
    max_pages: Optional[int],
    max_hrefs: Optional[int],
    status_mode: str,
    sleep_seconds: float,
    insecure: bool,
    resume: bool,
) -> Dict[str, int]:
    href_path = output_dir / "trec_unique_hrefs.csv"
    license_rows_path = output_dir / "trec_license_rows.csv"
    checkpoint_path = output_dir / "trec_hrefs_checkpoint.json"

    if not resume:
        reset_files([href_path, license_rows_path, checkpoint_path])

    seen_detail_ids = load_existing_values(href_path, "detail_id") if resume else set()
    seen_license_rows = (
        load_existing_license_row_keys(license_rows_path) if resume else set()
    )

    total_found = None
    total_unique_written = 0
    total_license_rows_written = 0
    checkpoint = load_checkpoint(checkpoint_path) if resume else {}
    if resume and start_page is None and checkpoint.get("last_page"):
        page = checkpoint["last_page"] + 1
    else:
        page = start_page or 1
    pages_done = 0

    while True:
        params = typesense_search_params(page, page_size, status_mode)
        data = request_json(
            TREC_TYPESENSE_SEARCH_URL,
            params,
            headers={"X-TYPESENSE-API-KEY": TREC_TYPESENSE_API_KEY},
            timeout=45,
            insecure=insecure,
        )
        if not isinstance(data, dict):
            raise RuntimeError("Unexpected Typesense response")

        total_found = int(data.get("found") or 0)
        grouped_hits = data.get("grouped_hits") or []
        if not grouped_hits:
            break

        href_rows: List[Dict[str, object]] = []
        license_rows: List[Dict[str, str]] = []

        for group in grouped_hits:
            hits = group.get("hits") or []
            if not hits:
                continue

            primary_doc = hits[0].get("document") or {}
            primary_row = row_from_typesense_doc(primary_doc)
            detail_id = primary_row["detail_id"]

            if detail_id and detail_id not in seen_detail_ids:
                href_rows.append(
                    {
                        "detail_id": detail_id,
                        "href": primary_row["href"],
                        "primary_license_number": primary_row["license_number"],
                        "primary_license_type": primary_row["license_type"],
                        "primary_status": primary_row["status"],
                        "display_name": primary_row["display_name"],
                    }
                )
                seen_detail_ids.add(detail_id)

            for hit in hits:
                doc = hit.get("document") or {}
                row = row_from_typesense_doc(doc)
                license_number = row["license_number"]
                row_key = f"{license_number}|{row['detail_id']}"
                if row_key not in seen_license_rows:
                    license_rows.append(row)
                    seen_license_rows.add(row_key)

            if max_hrefs and len(seen_detail_ids) >= max_hrefs:
                break

        total_unique_written += append_csv(href_path, HREF_FIELDS, href_rows)
        total_license_rows_written += append_csv(
            license_rows_path, LICENSE_ROW_FIELDS, license_rows
        )

        pages_done += 1
        print(
            f"href page={page} total_found={total_found} "
            f"unique_seen={len(seen_detail_ids)} "
            f"new_unique={len(href_rows)} new_license_rows={len(license_rows)}",
            flush=True,
        )
        write_checkpoint(
            checkpoint_path,
            last_page=page,
            total_found=total_found,
            unique_href_rows_seen=len(seen_detail_ids),
            license_rows_seen=len(seen_license_rows),
        )

        if max_hrefs and len(seen_detail_ids) >= max_hrefs:
            break
        if max_pages and pages_done >= max_pages:
            break
        if page * page_size >= total_found:
            break

        page += 1
        time.sleep(sleep_seconds)

    return {
        "typesense_total_found": total_found or 0,
        "unique_href_rows_written": total_unique_written,
        "license_rows_written": total_license_rows_written,
        "unique_href_rows_seen": len(seen_detail_ids),
    }


def open_data_params(limit: int, offset: int) -> Dict[str, object]:
    return {
        "$select": (
            "license_type,license_number,full_name,status,"
            "license_expiration_date,related_license_type,"
            "related_license_number,related_license_full_name,updated"
        ),
        "$where": (
            "status='Active' AND "
            "(license_type='Sales Agent' OR license_type='Broker Company') AND "
            "related_license_full_name IS NOT NULL AND related_license_full_name != ''"
        ),
        "$order": "license_number",
        "$limit": limit,
        "$offset": offset,
    }


def relationship_record_type(license_type: str) -> str:
    if license_type == "Sales Agent":
        return "agent_to_sponsor"
    if license_type == "Broker Company":
        return "company_to_designated_broker"
    return "other"


def google_company_query(company_name: str) -> str:
    return f'"{company_name}" "office email"'


def google_broker_query(broker_name: str, company_name: str) -> str:
    if broker_name and company_name:
        return f'"{broker_name}" "{company_name}" email phone'
    if broker_name:
        return f'"{broker_name}" real estate broker email phone Texas'
    return ""


def collect_relationships(
    output_dir: Path,
    *,
    batch_size: int,
    max_rows: Optional[int],
    sleep_seconds: float,
    insecure: bool,
) -> Dict[str, int]:
    relationship_path = output_dir / "trec_relationships.csv"
    agent_path = output_dir / "agent_sponsor_map.csv"
    company_queue_path = output_dir / "broker_company_contact_queue.csv"
    license_href_map = load_license_href_map(output_dir / "trec_license_rows.csv")

    reset_files([relationship_path, agent_path, company_queue_path])

    offset = 0
    total_rows = 0
    total_relationship_rows = 0
    total_agent_rows = 0
    company_queue: Dict[str, Dict[str, object]] = {}

    while True:
        limit = batch_size
        if max_rows:
            limit = min(limit, max_rows - total_rows)
            if limit <= 0:
                break

        rows = request_json(
            TEXAS_OPEN_DATA_URL,
            open_data_params(limit, offset),
            timeout=45,
            insecure=insecure,
        )
        if not isinstance(rows, list):
            raise RuntimeError("Unexpected Texas Open Data response")
        if not rows:
            break

        relationship_rows: List[Dict[str, object]] = []
        agent_rows: List[Dict[str, object]] = []

        for item in rows:
            license_type = normalize_spaces(item.get("license_type"))
            license_number = normalize_spaces(item.get("license_number"))
            related_license_number = normalize_spaces(item.get("related_license_number"))
            full_name = title_case_name(item.get("full_name") or "")
            related_full_name = title_case_name(item.get("related_license_full_name") or "")
            href = license_href_map.get(license_number, "")
            related_href = license_href_map.get(related_license_number, "")

            relationship_rows.append(
                {
                    "record_type": relationship_record_type(license_type),
                    "license_type": license_type,
                    "license_number": license_number,
                    "full_name": full_name,
                    "status": normalize_spaces(item.get("status")),
                    "license_expiration_date": normalize_spaces(
                        item.get("license_expiration_date")
                    ),
                    "related_license_type": normalize_spaces(
                        item.get("related_license_type")
                    ),
                    "related_license_number": related_license_number,
                    "related_license_full_name": related_full_name,
                    "href": href,
                    "related_href": related_href,
                    "updated": normalize_spaces(item.get("updated")),
                }
            )

            if license_type == "Sales Agent":
                agent_rows.append(
                    {
                        "agent_name": full_name,
                        "agent_license_number": license_number,
                        "agent_href": href,
                        "sponsoring_broker_type": normalize_spaces(
                            item.get("related_license_type")
                        ),
                        "sponsoring_broker_name": related_full_name,
                        "sponsoring_broker_license_number": related_license_number,
                        "sponsoring_broker_href": related_href,
                        "license_expiration_date": normalize_spaces(
                            item.get("license_expiration_date")
                        ),
                        "updated": normalize_spaces(item.get("updated")),
                    }
                )

            if license_type == "Broker Company":
                key = license_number or full_name.upper()
                company_queue[key] = {
                    "company_name": full_name,
                    "company_license_number": license_number,
                    "company_href": href,
                    "broker_name": related_full_name,
                    "broker_license_number": related_license_number,
                    "broker_href": related_href,
                    "google_company_query": google_company_query(full_name),
                    "google_broker_query": google_broker_query(
                        related_full_name, full_name
                    ),
                    "office_email": "",
                    "broker_email": "",
                    "phone": "",
                    "source_url": "",
                    "contact_status": "pending_research",
                    "outreach_subject": EMAIL_SUBJECT,
                    "outreach_body": EMAIL_BODY,
                }

        total_relationship_rows += append_csv(
            relationship_path, RELATIONSHIP_FIELDS, relationship_rows
        )
        total_agent_rows += append_csv(agent_path, AGENT_SPONSOR_FIELDS, agent_rows)

        total_rows += len(rows)
        print(
            f"relationship offset={offset} received={len(rows)} "
            f"total_rows={total_rows}",
            flush=True,
        )

        if len(rows) < limit:
            break
        offset += len(rows)
        time.sleep(sleep_seconds)

    append_csv(company_queue_path, BROKER_COMPANY_QUEUE_FIELDS, company_queue.values())

    return {
        "open_data_rows_read": total_rows,
        "relationship_rows_written": total_relationship_rows,
        "agent_rows_written": total_agent_rows,
        "broker_company_queue_rows": len(company_queue),
    }


def write_email_template(output_dir: Path) -> Path:
    path = output_dir / "email_template.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    content = f"Subject: {EMAIL_SUBJECT}\n\n{EMAIL_BODY}\n"
    path.write_text(content, encoding="utf-8")
    return path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Collect TREC license hrefs and sponsor/company relationship tables."
        )
    )
    parser.add_argument(
        "--mode",
        choices=["all", "hrefs", "relationships", "template"],
        default="all",
        help="Which step to run.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for CSV outputs.",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=250,
        help="Typesense page size. TREC UI uses 10; 250 is faster and stable.",
    )
    parser.add_argument("--start-page", type=int, default=None)
    parser.add_argument("--max-pages", type=int, default=None)
    parser.add_argument("--max-hrefs", type=int, default=None)
    parser.add_argument("--max-relationships", type=int, default=None)
    parser.add_argument(
        "--status-mode",
        choices=["active", "strict-active", "all"],
        default="active",
        help=(
            "active matches TREC UI Active filter; strict-active only keeps exact "
            "status.value Active in Typesense."
        ),
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=50000,
        help="Texas Open Data batch size.",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.35,
        help="Delay between remote requests.",
    )
    parser.add_argument(
        "--insecure",
        action="store_true",
        help=(
            "Disable SSL verification. Use only if this machine has a local "
            "certificate-chain issue when calling trec.texas.gov."
        ),
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Skip href/detail rows already present in output CSVs.",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    output_dir: Path = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    summary: Dict[str, int] = {}

    if args.mode in {"all", "hrefs"}:
        summary.update(
            collect_hrefs(
                output_dir,
                page_size=args.page_size,
                start_page=args.start_page,
                max_pages=args.max_pages,
                max_hrefs=args.max_hrefs,
                status_mode=args.status_mode,
                sleep_seconds=args.sleep,
                insecure=args.insecure,
                resume=args.resume,
            )
        )

    if args.mode in {"all", "relationships"}:
        summary.update(
            collect_relationships(
                output_dir,
                batch_size=args.batch_size,
                max_rows=args.max_relationships,
                sleep_seconds=args.sleep,
                insecure=args.insecure,
            )
        )

    if args.mode in {"all", "template"}:
        template_path = write_email_template(output_dir)
        print(f"email_template={template_path}", flush=True)

    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
