#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
nocodb_client.py - upserts one extracted tender record into the NocoDB table
"Реестр конкурсов" (TenderDataExtractor / StrokTenderDataExtractor project).

IMPORTANT - where this must run:
    Network access to app.nocodb.com is only allowed from Claude's cloud
    execution environment (added to the account's domain allowlist under
    Settings -> Capabilities -> Allow network egress -> Additional allowed
    domains). It is NOT reachable from the local device shell that mounts
    this project folder (separate, more restrictive network sandbox). So
    this script must be executed in the cloud environment, not via the
    local device shell - even though the file itself lives in the project
    folder for documentation/versioning purposes.

Usage:
    python nocodb_client.py --json /path/to/extracted_fields.json \
        --token-file /path/to/API_key_NocoDB.txt

    extracted_fields.json uses the same flat field names as save_tender.py
    (see FIELD_MAP / CSV_COLUMNS in save_tender.py) - the LLM only needs to
    produce ONE json dict and pass it to both save_tender.py and this script.

Optionally, pass --xlsx to also sync the "Позиции и цены" specification table
(Шаг 5.1 в CLAUDE.md) into its OWN NocoDB table - one table per tender, named
"Спецификация — <номер закупки>" - and write a link to it into the main
record's "Ссылка на спецификацию" column, so the full item-level breakdown can
always be found later from the main tender row.

Prints a JSON result to stdout:
    {"action": "created"|"updated", "record_id": <int>, "purchase_number": "...",
     "spec_table_url": "..." (only if --xlsx was given)}
"""

import argparse
import json
import sys
import urllib.request
import urllib.error
import urllib.parse

WORKSPACE_ID = "wzs963xn"
BASE_ID = "pf3o1felxit8s5b"
BASE_URL = "https://app.nocodb.com/api/v2/tables"
META_URL = "https://app.nocodb.com/api/v2/meta"
TABLE_ID = "mdbebnsirecunr5"
RECORDS_URL = f"{BASE_URL}/{TABLE_ID}/records"
SPEC_LINK_COLUMN = "Ссылка на спецификацию"

# Column order for the per-tender specification table - must match the
# "Позиции и цены" template in CLAUDE.md, Шаг 5.1. All text, same reasoning
# as FIELD_MAP below (values can be "Не указано в документах" / "Не удалось
# определить" placeholders instead of numbers).
SPEC_COLUMNS = [
    ("№", "SingleLineText"),
    ("Наименование позиции", "LongText"),
    ("Марка/модель", "SingleLineText"),
    ("Кол-во", "SingleLineText"),
    ("Ед. изм.", "SingleLineText"),
    ("НМЦК за позицию", "SingleLineText"),
    ("Рыночная цена (найдено в интернете)", "SingleLineText"),
    ("Источник (ссылка)", "SingleLineText"),
    ("Примечание", "LongText"),
]

# Flat field name (matches save_tender.py / all_tenders.csv columns) -> NocoDB column title.
# "purchase_number" and "tender_url" map onto the two columns that already existed
# in the table before this integration; everything else was added for this integration.
# NOTE: "Есть регистрация на площадке" and "Аккредитация" are intentionally NOT mapped -
# they are manual-only tracking checkboxes and are never written by extraction.
FIELD_MAP = {
    "source_file": "Файл-источник",
    "extracted_date": "Дата извлечения",
    "product_name": "Наименование предмета закупки",
    "quantity": "Количество",
    "status": "Статус",
    "delivery_place": "Место поставки",
    "delivery_term": "Срок поставки",
    "procedure_type": "Тип процедуры",
    "price_reduction": "Шаг снижения цены",
    "price_reduction_date": "Дата снижения цены",
    "start_price": "Начальная цена",
    "currency": "Валюта",
    "payment_terms": "Условия оплаты",
    "bid_security": "Обеспечение заявки",
    "contract_security": "Обеспечение контракта",
    "publication_date": "Дата публикации",
    "submission_deadline": "Срок подачи заявок",
    "auction_date": "Дата аукциона",
    "results_date": "Дата подведения итогов",
    "contract_date": "Дата заключения контракта",
    "last_modified": "Дата последнего изменения",
    "organizer": "Заказчик",
    "inn": "ИНН заказчика",
    "kpp": "КПП заказчика",
    "ogrn": "ОГРН заказчика",
    "contact": "Контактное лицо заказчика",
    "address": "Адрес заказчика",
    "procurement_method": "Способ закупки",
    "platform": "Площадка",
    "platform_url": "Ссылка на площадку",
    "purchase_number": "Номер конкурса",
    "tender_url": "Ссылка на процедуру",
    "stage": "Стадия",
    "sme": "СМП/СОНО",
    "national_regime_applied": "Нацрежим - применяется",
    "national_regime_measure": "Нацрежим - мера",
    "national_regime_items": "Нацрежим - товары",
    "national_regime_basis": "Нацрежим - основание",
}

PURCHASE_NUMBER_COLUMN = "Номер конкурса"
SOURCE_FILE_COLUMN = "Файл-источник"
NOT_SPECIFIED = "Не указано в документах"


def _request(method, url, token, body=None):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("xc-token", token)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
            return resp.status, (json.loads(raw) if raw else {})
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        return e.code, {"error": raw}


def find_record(token, column, value):
    where = urllib.parse.quote(f"({column},eq,{value})")
    url = f"{RECORDS_URL}?where={where}&limit=1"
    status, body = _request("GET", url, token)
    if status != 200:
        raise RuntimeError(f"lookup failed: HTTP {status} {body}")
    rows = body.get("list", [])
    return rows[0] if rows else None


def upsert(token, flat_data: dict, spec_xlsx_path: str = None, fallback_id: str = None):
    """fallback_id: used to name/find the spec table when the tender has no real
    purchase_number (placeholder "Не указано в документах") - e.g. the tender's
    local folder name/id, so the spec table can still be created and found later
    even without an official notice number. Not required if the tender does have
    a real purchase_number."""
    purchase_number = flat_data.get("purchase_number")
    if not purchase_number:
        raise ValueError("flat_data must include 'purchase_number' for upsert matching")

    row = {}
    for flat_name, value in flat_data.items():
        col = FIELD_MAP.get(flat_name)
        if col is None:
            continue  # unknown field, skip rather than fail
        row[col] = value

    spec_table_url = None
    if spec_xlsx_path:
        # Prefer the real purchase number to name the spec table; fall back to
        # fallback_id (e.g. the tender's local folder name) when there isn't one -
        # matching on the NOT_SPECIFIED placeholder itself would collide across
        # every other tender that also lacks a purchase number.
        spec_id = purchase_number if purchase_number != NOT_SPECIFIED else fallback_id
        if spec_id:
            spec_rows = read_spec_xlsx(spec_xlsx_path)
            spec_table_url = sync_specification(token, spec_id, spec_rows)
            row[SPEC_LINK_COLUMN] = spec_table_url

    # When the purchase number itself wasn't found in the documents, matching on it
    # would collide across unrelated tenders that share the same placeholder text -
    # fall back to matching on the source file name instead in that case.
    if purchase_number == NOT_SPECIFIED:
        source_file = flat_data.get("source_file")
        existing = find_record(token, SOURCE_FILE_COLUMN, source_file) if source_file else None
    else:
        existing = find_record(token, PURCHASE_NUMBER_COLUMN, purchase_number)

    if existing:
        record_id = existing["Id"]
        status, body = _request("PATCH", RECORDS_URL, token, {"Id": record_id, **row})
        if status != 200:
            raise RuntimeError(f"update failed: HTTP {status} {body}")
        result = {"action": "updated", "record_id": record_id, "purchase_number": purchase_number}
    else:
        status, body = _request("POST", RECORDS_URL, token, row)
        if status != 200:
            raise RuntimeError(f"create failed: HTTP {status} {body}")
        result = {"action": "created", "record_id": body.get("Id"), "purchase_number": purchase_number}

    if spec_table_url:
        result["spec_table_url"] = spec_table_url
    return result


def _spec_table_title(purchase_number: str) -> str:
    # "/" (and a few other characters) are rejected by NocoDB in table titles.
    safe_number = purchase_number.replace("/", "-")
    return f"Спецификация — {safe_number}"


def find_table_by_title(token, title):
    status, body = _request("GET", f"{META_URL}/bases/{BASE_ID}/tables", token)
    if status != 200:
        raise RuntimeError(f"list tables failed: HTTP {status} {body}")
    for t in body.get("list", []):
        if t.get("title") == title:
            return t
    return None


def create_spec_table(token, title):
    columns = [{"title": name, "uidt": uidt} for name, uidt in SPEC_COLUMNS]
    status, body = _request("POST", f"{META_URL}/bases/{BASE_ID}/tables", token,
                             {"title": title, "columns": columns})
    if status != 200:
        raise RuntimeError(f"create spec table failed: HTTP {status} {body}")
    return body


def get_default_view_id(token, table_id):
    status, body = _request("GET", f"{META_URL}/tables/{table_id}/views", token)
    if status != 200:
        raise RuntimeError(f"list views failed: HTTP {status} {body}")
    views = body.get("list", [])
    return views[0]["id"] if views else None


def clear_table_records(token, table_id):
    status, body = _request("GET", f"{BASE_URL}/{table_id}/records?limit=1000", token)
    if status != 200:
        raise RuntimeError(f"list records failed: HTTP {status} {body}")
    ids = [{"Id": r["Id"]} for r in body.get("list", [])]
    if ids:
        status, body = _request("DELETE", f"{BASE_URL}/{table_id}/records", token, ids)
        if status != 200:
            raise RuntimeError(f"clear records failed: HTTP {status} {body}")


def sync_specification(token, purchase_number: str, spec_rows: list):
    """Ensure a per-tender spec table exists, (re)fill it with spec_rows, return its URL.

    spec_rows: list of dicts keyed by the SPEC_COLUMNS titles (e.g. one dict per
    row of Позиции_и_цены.xlsx, as read by openpyxl with the header row as keys).
    Re-running this for the same purchase_number replaces the table's contents
    rather than duplicating rows - safe to re-run after a correction.
    """
    title = _spec_table_title(purchase_number)
    table = find_table_by_title(token, title)
    if table is None:
        table = create_spec_table(token, title)
    else:
        clear_table_records(token, table["id"])

    table_id = table["id"]
    if spec_rows:
        status, body = _request("POST", f"{BASE_URL}/{table_id}/records", token, spec_rows)
        if status != 200:
            raise RuntimeError(f"insert spec rows failed: HTTP {status} {body}")

    view_id = get_default_view_id(token, table_id)
    url = f"https://app.nocodb.com/{WORKSPACE_ID}/{BASE_ID}/{table_id}"
    if view_id:
        url += f"/{view_id}"
    return url


def read_spec_xlsx(xlsx_path: str):
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    records = []
    for r in rows[1:]:
        rec = {h: v for h, v in zip(headers, r) if v is not None and h is not None}
        if rec:
            records.append(rec)
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", required=True, help="Path to the flat extracted-fields JSON")
    parser.add_argument("--token-file", required=True, help="Path to the file containing the NocoDB API token")
    parser.add_argument("--xlsx", required=False, help="Path to Позиции_и_цены.xlsx - if given, also syncs it into its own NocoDB table and links it from the main record")
    parser.add_argument("--fallback-id", required=False, help="Identifier to name the spec table by when the tender has no real purchase_number (e.g. the tender's local folder name)")
    args = parser.parse_args()

    with open(args.json, "r", encoding="utf-8") as f:
        flat_data = json.load(f)
    with open(args.token_file, "r", encoding="utf-8") as f:
        token = f.read().strip()

    result = upsert(token, flat_data, spec_xlsx_path=args.xlsx, fallback_id=args.fallback_id)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
