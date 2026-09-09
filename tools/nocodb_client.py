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
    This includes "internal_number" (Шаг 1a в CLAUDE.md) - save_tender.py is
    normally run first and its result JSON carries the resolved internal_number
    back, which should be merged into this JSON before calling this script, so
    the same auto-assigned/reused value ends up in NocoDB too instead of this
    script trying to assign its own.

Optionally, pass --xlsx to also sync the "Позиции и цены" specification table
(Шаг 5.1 в CLAUDE.md) into its OWN NocoDB table - one table per tender, named
"Спецификация — <номер закупки>" - and write a link to it into the main
record's "Ссылка на спецификацию" column, so the full item-level breakdown can
always be found later from the main tender row.

Prints a JSON result to stdout:
    {"action": "created"|"updated", "record_id": <int>, "purchase_number": "...",
     "verification": {"ok": bool, "checked_fields": int, "issues": [...]},
     "spec_table_url": "..." (only if --xlsx was given),
     "spec_verification": {"ok": bool, "expected_rows": int, "actual_rows": int,
                            "issues": [...]} (only if --xlsx was given)}

verification / spec_verification (added 2026-09-09): right after writing, this
script re-reads the record (and, with --xlsx, every row of the spec table)
back from NocoDB and compares it against what THIS run intended to send -
this is what catches the class of bug already seen twice on this project
("Файл-источник" and "Номер конкурса" both went empty in NocoDB while the
script itself printed a success result) automatically, instead of relying on
someone noticing by chance later. "ok": false means at least one field/row
came back different from what was sent (most often empty) - "issues" lists
exactly which ones. This never raises or changes the exit code by itself
(the create/update already succeeded by the time verification runs) -
callers must check "ok"/"issues" and tell the user, rather than assume
success just because the process exited cleanly.
"""

import argparse
import json
import re
import sys
import urllib.request
import urllib.error
import urllib.parse
from datetime import date

WORKSPACE_ID = "wzs963xn"
BASE_ID = "pf3o1felxit8s5b"
BASE_URL = "https://app.nocodb.com/api/v2/tables"
META_URL = "https://app.nocodb.com/api/v2/meta"
TABLE_ID = "mdbebnsirecunr5"
RECORDS_URL = f"{BASE_URL}/{TABLE_ID}/records"
SPEC_LINK_COLUMN = "Ссылка на спецификацию"

# Column order for the per-tender specification table - must match the
# "Позиции и цены" template in CLAUDE.md, Шаг 5.1.
# (title, uidt, extra column props for the create-table API call or None).
# Кол-во / НМЦК за позицию / Рыночная цена, руб. are typed Decimal/Currency
# (RU locale, comma decimal) instead of plain text; Источник (ссылка) is
# typed URL. Verified empirically against a scratch table on this base
# (2026-09-04, then deleted) before wiring this in: Decimal/Currency reject
# any non-numeric value and fail the WHOLE bulk insert (not just that
# field/row) - see NUMERIC_SPEC_COLUMNS/sanitize_spec_row below for how
# placeholders like "Не указано в документах" are handled so that stays
# safe. URL, by contrast, does NOT validate its value at the API level at
# all (an invalid or even "url1; url2" string is accepted as-is), so
# "Источник (ссылка)" needs no special handling despite sometimes holding
# two links joined by "; " (Шаг 5.1 в CLAUDE.md).
#
# "Рыночная цена (найдено в интернете)" stays plain LongText on purpose
# (reverted 2026-09-04, tender RAD260041848 exposed this): it routinely
# holds mixed-currency, approximate, free-text values straight from web
# search - e.g. "≈3 500 € за шт. (ориентировочно; ...)" or "$12 940,80
# (зарубежный дистрибьютор, ...)" - not just RUB numbers with commentary.
# Coercing THAT text into a numeric column would silently grab the first
# digit run regardless of currency (reading "3500" out of a €-quote and
# storing it as if it were 3500 RUB - very wrong). The clean, always-RUB,
# comment-free number the user wants for actual calculations goes in the
# separate "Рыночная цена, руб." column instead (added 2026-09-04) - the
# caller is responsible for converting non-RUB quotes to RUB before putting
# a value there (see Шаг 5.1 в CLAUDE.md for where that conversion happens);
# this module does not do currency conversion itself.
# isLocaleString: best-effort attempt at a RU-style comma decimal display
# for a plain Decimal column (Currency has a dedicated currency_locale meta
# that forces ru-RU for any viewer; Decimal has no confirmed equivalent -
# this key comes from a NocoDB PR referencing `meta.value.isLocaleString`
# -> `.toLocaleString(undefined, ...)`, which follows the viewer's OWN
# browser locale rather than forcing one. Not independently confirmed in
# the NocoDB UI - a check via browser was interrupted by a tool glitch
# 2026-09-04. If it doesn't show a comma for the user, drop this meta key
# without worry, the stored value/precision (dtxs=2) is correct either way.
# separator: the NocoDB UI's "Thousands and decimal separators" dropdown (screenshot from
# the user, 2026-09-07: field "НМЦК за позицию" set to "None and comma (1000000,00)" - no
# thousands separator, comma decimal point). Confirmed empirically by comparing that field's
# meta via GET /meta/tables/{id} after the user set it through the NocoDB UI and clicked
# "Update Field" - the previously-undocumented key is meta.separator, value "none_comma".
# (The isLocaleString key above does NOT control this dropdown - see the comment on it; it's
# a different, apparently cosmetic key kept only because it's harmless.) Applied to every
# existing "НМЦК за позицию"/"Рыночная цена, руб." Decimal column across all spec tables via
# a one-off PATCH /meta/columns/{id} on 2026-09-07 in addition to being baked in here for any
# spec table created from now on. "Кол-во" (QTY_DECIMAL_META) is a quantity, not a price, and
# was deliberately left on the default separator - only prices were in scope for this change.
# precision: the NocoDB UI's Decimal "Precision" dropdown (how many decimal
# digits are shown, e.g. "1.00") is driven by meta.precision, which is a
# SEPARATE setting from dtxs (the actual Postgres column scale). Setting
# dtxs alone stores the right value but the UI still showed "1.0" (1 digit)
# until meta.precision was added too - confirmed empirically 2026-09-04
# (user screenshot) by probing a scratch table and comparing to the live
# "Рыночная цена, руб." column, then fixing it via a PATCH to
# /meta/columns/{id}. Both must always be set together and kept in sync.
RU_DECIMAL_META = {"dtxp": 15, "dtxs": 2, "meta": {"isLocaleString": True, "precision": 2, "separator": "none_comma"}}
QTY_DECIMAL_META = {"dtxp": 15, "dtxs": 3, "meta": {"precision": 3}}
SPEC_COLUMNS = [
    ("№", "SingleLineText", None),
    ("Наименование позиции", "LongText", None),
    ("Марка/модель", "SingleLineText", None),
    ("Кол-во", "Decimal", QTY_DECIMAL_META),
    ("Ед. изм.", "SingleLineText", None),
    ("НМЦК за позицию", "Decimal", RU_DECIMAL_META),
    ("Рыночная цена (найдено в интернете)", "LongText", None),
    ("Рыночная цена, руб.", "Decimal", RU_DECIMAL_META),
    ("Источник (ссылка)", "URL", None),
    ("Примечание", "LongText", None),
]

# Columns that must be a real number (or null) for NocoDB to accept them -
# see the comment above SPEC_COLUMNS for why this matters.
NUMERIC_SPEC_COLUMNS = ("Кол-во", "НМЦК за позицию", "Рыночная цена, руб.")

# "НМЦК за позицию" is sometimes written as "77250.00 за ед. (154500.00 за
# 2 шт.)" (per-unit price, with the actual line total in parentheses) -
# confirmed on a real tender (32616325015, 2026-09-04). A plain first-number
# parse would silently grab the per-unit figure instead of the total the
# column name promises ("за позицию" = for the whole line). This pattern
# pulls out the parenthetical total when present.
POSITION_TOTAL_RE = re.compile(r"\(([\d\s.,]+)\s*за\s*\d+\s*(?:шт|ед)\.?\)", re.IGNORECASE)


def parse_position_total(raw):
    """Parse "НМЦК за позицию" specifically - prefers the "(Y за N шт.)"
    total over a leading per-unit figure when both are present in the text,
    otherwise behaves like parse_ru_number()."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw)
    match = POSITION_TOTAL_RE.search(text)
    if match:
        return parse_ru_number(match.group(1))
    return parse_ru_number(text)


def parse_ru_number(raw):
    """Best-effort parse of one spec-table numeric cell into a float.

    Accepts an already-numeric openpyxl cell (int/float) as well as RU
    -formatted number text ("1 500 000,55", "1500000.55", "12 шт." - a
    trailing non-numeric unit is ignored - same digit-token approach as
    normalize_number() in save_tender.py). Returns None when no numeric
    token can be found at all (e.g. "Не указано в документах", "Не удалось
    определить") - the caller must make sure that original text isn't just
    silently lost, see sanitize_spec_row().
    """
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    raw = str(raw).strip()
    if not raw:
        return None

    match = re.search(r"\d(?:[\d\s.,]*\d)?", raw)
    if not match:
        return None
    cleaned = re.sub(r"\s", "", match.group(0))

    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        parts = cleaned.split(",")
        if len(parts[-1]) == 2:
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif cleaned.count(".") > 1:
        cleaned = cleaned.replace(".", "")
    elif cleaned.count(".") == 1:
        integer_part, _, frac_part = cleaned.partition(".")
        if len(frac_part) == 3 and len(integer_part) <= 3:
            cleaned = cleaned.replace(".", "")

    try:
        return float(cleaned)
    except ValueError:
        return None


def sanitize_spec_row(row: dict) -> dict:
    """Turn one Позиции_и_цены.xlsx row into a record NocoDB will accept.

    NUMERIC_SPEC_COLUMNS have to be real numbers or null - a placeholder
    string in any one of them would make NocoDB reject the entire bulk
    insert (confirmed empirically, see the SPEC_COLUMNS comment). When a
    value can't be parsed as a number, the field is sent as null and the
    original text is preserved by appending it to "Примечание" instead of
    being silently dropped.
    """
    out = dict(row)
    extra_notes = []
    for col in NUMERIC_SPEC_COLUMNS:
        if col not in out:
            continue
        raw = out[col]
        parser = parse_position_total if col == "НМЦК за позицию" else parse_ru_number
        value = parser(raw)
        if value is None:
            if raw is not None and str(raw).strip():
                extra_notes.append(f"{col}: {raw}")
            out[col] = None
        else:
            out[col] = value
    if extra_notes:
        existing = str(out.get("Примечание") or "").strip()
        addition = "; ".join(extra_notes)
        out["Примечание"] = f"{existing}; {addition}" if existing else addition
    return out

# Flat field name (matches save_tender.py / all_tenders.csv columns) -> NocoDB column title.
# "purchase_number" and "tender_url" map onto the two columns that already existed
# in the table before this integration; everything else was added for this integration.
# "internal_number" (added Шаг 1a в CLAUDE.md) maps onto "Внутренний номер" - unlike the
# other columns, that one is auto-created by ensure_column() below if it's missing,
# instead of requiring the user to add it by hand in the NocoDB UI first.
# NOTE: "Есть регистрация на площадке" и "Аккредитация" are intentionally NOT mapped -
# they are manual-only tracking checkboxes and are never written by extraction.
FIELD_MAP = {
    # "payment_days_summary" is a COMPUTED, NocoDB-only convenience field (not one
    # of the 7 extraction blocks, not written to local XML/CSV by save_tender.py).
    # The LLM fills it when building the JSON passed to this script - see
    # CLAUDE.md, Шаг 7.1a "Краткий срок оплаты и цветовая метка" for the exact
    # extraction/formatting rules (MSP-specific terms, worst-case day count,
    # "N дней" / "N-M дней" format, appended " МСП" label, leading colour
    # marker: "🟢 " when an advance payment (аванс) is provided (favourable
    # for the supplier, regardless of day count - takes priority), else
    # "🔴 " when postpayment-only and > 10 days, else no marker. The marker
    # is what a manually-configured NocoDB "is like 🟢%" / "is like 🔴%"
    # colouring rule keys off of; the API cannot set cell colour itself,
    # only write the marker text).
    "internal_number": "Внутренний номер",
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
    "payment_days_summary": "Срок оплаты (кратко)",
}

# Columns this script will auto-create on the main "Реестр конкурсов" table if
# they're missing, instead of requiring the user to add them by hand first.
# title -> (uidt, extra column props or None). Keep this list small/deliberate -
# most FIELD_MAP columns were added by hand in the NocoDB UI when the
# integration was first built; this is only for columns introduced later that
# should just work out of the box.
#
# "Номер конкурса" (PURCHASE_NUMBER_COLUMN below) is included here because a
# live check against the production table (2026-09-07, testing the new
# internal_number feature) found it missing entirely - find_record() against
# it failed with a hard HTTP 422 ("Column alias 'Номер конкурса' not found"),
# which meant upsert() couldn't look up ANY tender by purchase number and the
# whole NocoDB sync step was broken, not just for internal_number. Confirmed
# with the user (2026-09-07) that recreating this column is the fix, rather
# than matching on some other existing column. Note: "Внутренний номер" is
# also already present in the live table as its primary/title column (pv:
# True) and already held one hand-entered value ("К897", Cyrillic К, no
# dash) plus, on other rows, old purchase/folder numbers used as makeshift
# identifiers before this column existed for that purpose - confirmed with
# the user that new auto-assigned values should still use the "K-983" format
# (Latin K + dash) from Шаг 1a в CLAUDE.md rather than matching that legacy style.
AUTO_CREATE_COLUMNS = {
    "Внутренний номер": ("SingleLineText", None),
    "Номер конкурса": ("SingleLineText", None),
}

# Column title -> title of the column it should sit immediately after in the
# default grid view, applied only right after ensure_columns() actually
# CREATES that column (a column the user later reorders by hand is never
# touched again). Without this, the "create column" API just appends new
# columns at the very end of the table, which is where "Номер конкурса"
# landed when it was auto-created (see AUTO_CREATE_COLUMNS above) - the user
# asked (2026-09-07) for it to sit between "Внутренний номер" and "Ссылка на
# процедуру" instead (its natural spot, next to the other tender-identifying
# columns), so this makes that positioning automatic/self-healing if the
# column is ever deleted and auto-recreated again, instead of requiring the
# same manual reorder through the API a second time.
AUTO_CREATE_COLUMN_POSITION_AFTER = {
    "Номер конкурса": "Внутренний номер",
}

PURCHASE_NUMBER_COLUMN = "Номер конкурса"
SOURCE_FILE_COLUMN = "Файл-источник"
NOT_SPECIFIED = "Не указано в документах"

# Fields where a missing value should read as a plain dash "—" rather than
# NOT_SPECIFIED - see the matching DASH_FIELDS in save_tender.py for the
# full rationale (user request, 2026-09-07: no link to the procurement
# documentation -> "—", not "Не указано в документах"). Claude already
# writes "—" directly when assembling the flat JSON (same file feeds both
# save_tender.py and this script), so this is a safety net here too, in
# case that JSON still has the old placeholder or an empty value.
DASH = "—"
DASH_FIELDS = ("tender_url",)


def apply_dash_placeholders(flat_data):
    """Normalizes DASH_FIELDS in-place: missing/empty/NOT_SPECIFIED -> "—"."""
    for flat_name in DASH_FIELDS:
        value = str(flat_data.get(flat_name) or "").strip()
        if not value or value == NOT_SPECIFIED:
            flat_data[flat_name] = DASH
    return flat_data


# Fields whose NocoDB target column is a "Ссылка на ..." URL field that (unlike
# the spec table's "Источник (ссылка)" column - see the big comment near
# SPEC_COLUMNS above) DOES enforce strict isURL validation at the API level -
# confirmed empirically 2026-09-09 on a real tender (ЗК-715-2024) whose
# Извещение had no direct link: sending tender_url = "—" (the DASH placeholder
# above, per Блок 6 в CLAUDE.md) into "Ссылка на процедуру" made the WHOLE
# create/update request fail with HTTP 400 {"msg":"Validation failed : isURL"}
# - these two were most likely hand-created in the NocoDB UI (unlike the
# API-created spec table columns), where "Validate as URL" can be ticked on.
# "platform_url" -> "Ссылка на площадку" maps the same way and would fail the
# same way if it is ever missing/placeholder text, so it gets the same guard
# even though only tender_url has hit this in practice so far.
STRICT_URL_FIELDS = ("tender_url", "platform_url")
_URL_RE = re.compile(r"^https?://", re.IGNORECASE)


def _looks_like_url(value) -> bool:
    return bool(_URL_RE.match(str(value or "").strip()))

def format_source_files(raw):
    """Mirrors save_tender.py's format_source_files(): the same " + "-joined
    string built for save_tender.py's --source-file (Шаг 7, п.3 в CLAUDE.md)
    -> one file per line ("a.docx\nb.pdf") - "Файл-источник" is a rich-text
    column here too (see save_tender.py's own docstring on this, 2026-09-07).
    Added 2026-09-08 after a user report that this column kept coming out
    empty in NocoDB: save_tender.py takes the file list via --source-file
    (a CLI flag, not a JSON key) and never needed "source_file" in the JSON
    at all, so past sessions had no reason to put it there - but this script
    only reads flat_data from --json (no equivalent CLI flag exists here),
    so the field must actually be present in that JSON for this column to
    get anything. CLAUDE.md Шаг 7 now says to include it. A single file (no
    " + " found) is returned unchanged."""
    if raw is None:
        return raw
    parts = [p.strip() for p in str(raw).split(" + ") if p.strip()]
    if len(parts) <= 1:
        return raw
    return "\n".join(parts)




# "Начальная цена" (main table, FIELD_MAP["start_price"]) is plain SingleLineText,
# not Decimal - it has to stay text because it also holds the NOT_SPECIFIED
# placeholder and occasional multi-number text ("3805861.62 руб. включая НДС
# (3119558.70 руб. без НДС)", tender 905-U310-26/3333, K-985) that a real
# Decimal column would reject outright. save_tender.py's normalize_number()
# already guarantees any actual price number inside that text is a clean
# "<digits>.<exactly two digits>" token (no thousands separator) before this
# script ever sees it - the only thing left to fix for the NocoDB copy of the
# value is the decimal point itself: the user asked (2026-09-07, same request
# that produced the Decimal "separator": "none_comma" meta above) for every
# price actually showing up in NocoDB to use a comma there instead, with no
# other dots/commas in the number. This only touches the copy of the string
# sent to NocoDB - save_tender.py's own XML/CSV output (and the on-disk JSON)
# are untouched, since nothing here mutates the source JSON file itself.
PRICE_NUMBER_RE = re.compile(r"(?<!\d)(\d+)\.(\d{2})(?!\d)")


def to_comma_decimal(text):
    """Replace every clean "<digits>.<exactly 2 digits>" token in `text` with
    "<digits>,<2 digits>" - i.e. swap the decimal point for a comma without
    touching anything else (placeholder text, ✅/🔺 markers, "руб. включая
    НДС" wording, a second number in parentheses, etc. all pass through
    unchanged). Safe to call on any string, including NOT_SPECIFIED/empty."""
    if not text:
        return text
    return PRICE_NUMBER_RE.sub(lambda m: f"{m.group(1)},{m.group(2)}", text)


# "Адрес заказчика" - user request (2026-09-08): the address itself should
# render in bold in NocoDB, but the postal index (6-digit ЯЯЯЯЯЯ, wherever it
# sits in the string - leading "123456, г. Москва..." and trailing "...,
# 123456" both occur in real tenders) must stay un-bolded. The column was
# plain LongText with no rich-text formatting until now - PATCHed via the
# Meta API (2026-09-08) to add meta.richMode=True/smartMode=False, matching
# "Наименование предмета закупки" (see product_name_formatting.md in project
# memory), so Markdown **bold** actually renders instead of showing as
# literal asterisks.
POSTAL_INDEX_RE = re.compile(r"(?<!\d)(\d{6})(?!\d)")


def bold_address_except_index(address):
    """Wraps `address` in Markdown bold, carving out the first standalone
    6-digit postal index (if any) so it stays plain text. Safe to call on
    any string, including NOT_SPECIFIED/DASH/empty (returned unchanged)."""
    if not address or address in (NOT_SPECIFIED, DASH):
        return address
    match = POSTAL_INDEX_RE.search(address)
    if not match:
        return f"**{address}**"
    before, index, after = address[:match.start()], match.group(1), address[match.end():]
    parts = []
    parts.append(f"**{before}**" if before.strip(" ,") else before)
    parts.append(index)
    parts.append(f"**{after}**" if after.strip(" ,") else after)
    return "".join(parts)


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


def get_table_columns(token, table_id):
    """Returns {column_title: column_meta_dict} for an existing table."""
    status, body = _request("GET", f"{META_URL}/tables/{table_id}", token)
    if status != 200:
        raise RuntimeError(f"get table failed: HTTP {status} {body}")
    return {c["title"]: c for c in body.get("columns", [])}


def position_column_after(token, table_id, column_id, anchor_title):
    """Moves column_id to sit immediately after the column titled
    anchor_title in the table's default grid view (see
    AUTO_CREATE_COLUMN_POSITION_AFTER). Best-effort: any failure here is
    swallowed (the column still exists and works, just wherever the create
    -column API put it) rather than failing the whole sync over cosmetics.

    Mechanics: a grid view's left-to-right column order is controlled by
    each column's OWN per-view record (GET/PATCH
    /meta/views/{viewId}/columns/{viewColumnId} - viewColumnId, not the
    column's own id, confirmed empirically 2026-09-07: PATCHing by the
    column id itself 404s with ERR_FIELD_NOT_FOUND), not by the base
    column's table-level `order` field. Orders are plain floats and don't
    need to be contiguous, so the new column is just given the midpoint
    between the anchor's order and whatever comes right after it (or
    anchor_order + 1 if the anchor is currently last).
    """
    try:
        view_id = get_default_view_id(token, table_id)
        if not view_id:
            return
        status, body = _request("GET", f"{META_URL}/views/{view_id}/columns", token)
        if status != 200:
            return
        view_columns = body.get("list", [])

        id_to_title = get_table_columns(token, table_id)  # title -> column meta
        title_to_id = {meta["id"]: title for title, meta in id_to_title.items()}

        by_column_id = {vc["fk_column_id"]: vc for vc in view_columns}
        target_vc = by_column_id.get(column_id)
        anchor_id = id_to_title.get(anchor_title, {}).get("id")
        anchor_vc = by_column_id.get(anchor_id) if anchor_id else None
        if not target_vc or not anchor_vc:
            return

        anchor_order = anchor_vc["order"]
        later_orders = [vc["order"] for vc in view_columns if vc["order"] > anchor_order]
        new_order = (anchor_order + min(later_orders)) / 2 if later_orders else anchor_order + 1

        _request("PATCH", f"{META_URL}/views/{view_id}/columns/{target_vc['id']}", token, {"order": new_order})
    except Exception:
        pass


def ensure_columns(token, table_id, titles):
    """Auto-creates any of AUTO_CREATE_COLUMNS listed in `titles` that don't
    already exist on the table - e.g. "Внутренний номер" the first time this
    script runs after Шаг 1a was added, so the user doesn't have to go add it
    by hand in the NocoDB UI first. Safe to call every run (no-op once the
    column exists). Newly-created columns listed in
    AUTO_CREATE_COLUMN_POSITION_AFTER are also moved into place (see
    position_column_after) instead of being left at the end of the table."""
    wanted = [t for t in titles if t in AUTO_CREATE_COLUMNS]
    if not wanted:
        return
    existing = get_table_columns(token, table_id)
    for title in wanted:
        if title in existing:
            continue
        uidt, extra = AUTO_CREATE_COLUMNS[title]
        col = {"title": title, "uidt": uidt}
        if extra:
            col.update(extra)
        status, body = _request("POST", f"{META_URL}/tables/{table_id}/columns", token, col)
        if status != 200:
            raise RuntimeError(f"create column '{title}' failed: HTTP {status} {body}")
        anchor_title = AUTO_CREATE_COLUMN_POSITION_AFTER.get(title)
        if anchor_title:
            # The create-column response is the TABLE's meta, not the new
            # column's (confirmed empirically 2026-09-07 on a scratch column -
            # its "id" is fk_model_id, i.e. the table id) - re-fetch columns to
            # find the new column's own id by title.
            new_column_id = get_table_columns(token, table_id).get(title, {}).get("id")
            if new_column_id:
                position_column_after(token, table_id, new_column_id, anchor_title)


def upsert(token, flat_data: dict, spec_xlsx_path: str = None, fallback_id: str = None, fx_rates: dict = None):
    """fallback_id: used to name/find the spec table when the tender has no real
    purchase_number (placeholder "Не указано в документах") - e.g. the tender's
    local folder name/id, so the spec table can still be created and found later
    even without an official notice number. Not required if the tender does have
    a real purchase_number.

    fx_rates: passed straight through to enrich_market_price_rub() - see
    convert_market_price_to_rub() for the expected shape ({"€": ..., "$": ...})."""
    purchase_number = flat_data.get("purchase_number")
    if not purchase_number:
        raise ValueError("flat_data must include 'purchase_number' for upsert matching")

    apply_dash_placeholders(flat_data)
    if flat_data.get("start_price"):
        flat_data["start_price"] = to_comma_decimal(str(flat_data["start_price"]))
    if flat_data.get("address"):
        flat_data["address"] = bold_address_except_index(str(flat_data["address"]))
    if flat_data.get("source_file"):
        flat_data["source_file"] = format_source_files(str(flat_data["source_file"]))

    ensure_columns(token, TABLE_ID, [FIELD_MAP.get(k) for k in flat_data.keys()])

    row = {}
    for flat_name, value in flat_data.items():
        col = FIELD_MAP.get(flat_name)
        if col is None:
            continue  # unknown field, skip rather than fail
        if flat_name in STRICT_URL_FIELDS and not _looks_like_url(value):
            # Don't send "—"/NOT_SPECIFIED/anything else into a strictly-
            # validating URL column - see STRICT_URL_FIELDS comment above.
            # Omitting the key just leaves that column blank in NocoDB,
            # instead of failing the entire record create/update.
            continue
        row[col] = value

    spec_table_url = None
    spec_table_id = None
    spec_safe_rows = None
    if spec_xlsx_path:
        # Prefer the real purchase number to name the spec table; fall back to
        # fallback_id (e.g. the tender's local folder name) when there isn't one -
        # matching on the NOT_SPECIFIED placeholder itself would collide across
        # every other tender that also lacks a purchase number.
        spec_id = purchase_number if purchase_number != NOT_SPECIFIED else fallback_id
        if spec_id:
            spec_rows = read_spec_xlsx(spec_xlsx_path)
            spec_rows = enrich_market_price_rub(spec_rows, fx_rates)
            spec_table_url, spec_table_id, spec_safe_rows = sync_specification(token, spec_id, spec_rows)
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

    # Post-write verification (added 2026-09-09, user request - this project's
    # own memory already logged two silent-failure incidents where the script
    # printed success but a column ("Файл-источник", then "Номер конкурса")
    # actually stayed empty in NocoDB): re-read what was actually written and
    # compare it against what this run intended to send, instead of trusting
    # the create/update HTTP 200 alone. Best-effort - a failure of the
    # verification GET itself is reported as its own issue rather than raised,
    # since the record write already succeeded by this point.
    result["verification"] = verify_main_record(token, result["record_id"], row)

    if spec_table_url:
        result["spec_table_url"] = spec_table_url
        result["spec_verification"] = verify_spec_table(token, spec_table_id, spec_safe_rows or [])
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
    columns = []
    for name, uidt, extra in SPEC_COLUMNS:
        col = {"title": name, "uidt": uidt}
        if extra:
            col.update(extra)
        columns.append(col)
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
    """Ensure a per-tender spec table exists, (re)fill it with spec_rows.

    spec_rows: list of dicts keyed by the SPEC_COLUMNS titles (e.g. one dict per
    row of Позиции_и_цены.xlsx, as read by openpyxl with the header row as keys).
    Re-running this for the same purchase_number replaces the table's contents
    rather than duplicating rows - safe to re-run after a correction.

    Returns (url, table_id, safe_rows) - table_id and the sanitized rows that
    were actually sent (see sanitize_spec_row) are returned alongside the url
    so the caller (upsert(), see verify_spec_table() above) can re-read the
    table afterwards and verify it landed correctly, without re-deriving what
    was sent from scratch.
    """
    title = _spec_table_title(purchase_number)
    table = find_table_by_title(token, title)
    if table is None:
        table = create_spec_table(token, title)
    else:
        clear_table_records(token, table["id"])

    table_id = table["id"]
    safe_rows = [sanitize_spec_row(r) for r in spec_rows] if spec_rows else []
    if safe_rows:
        # NocoDB's bulk-insert endpoint caps a single request at 100 records
        # (HTTP 422 ERR_MAX_PAYLOAD_LIMIT_EXCEEDED) - found 2026-09-09 on a
        # real 116-position tender. Split into batches to stay under that cap.
        NOCODB_MAX_BATCH = 100
        for i in range(0, len(safe_rows), NOCODB_MAX_BATCH):
            batch = safe_rows[i:i + NOCODB_MAX_BATCH]
            status, body = _request("POST", f"{BASE_URL}/{table_id}/records", token, batch)
            if status != 200:
                raise RuntimeError(f"insert spec rows failed: HTTP {status} {body}")

    view_id = get_default_view_id(token, table_id)
    url = f"https://app.nocodb.com/{WORKSPACE_ID}/{BASE_ID}/{table_id}"
    if view_id:
        url += f"/{view_id}"
    return url, table_id, safe_rows


def read_spec_xlsx(xlsx_path: str):
    """Read Позиции_и_цены.xlsx into one dict per actual item row.

    Skips the trailing "Итого НМЦК:" totals row and any other non-item row
    (a blank separator row, or a sheet-level note row that only fills the
    "№"/"Примечание" cells with free text) - a real position always has
    "Наименование позиции" filled, so that's what's used to tell a real row
    from a summary/note row (confirmed against real tenders' xlsx files
    2026-09-04: RAD260041848's "Итого НМЦК:" row, 9594163's blank row and
    its trailing "Примечание: ..." note row all lack it). NocoDB's per
    -tender spec table is meant to be an item list, not a place for the
    tender-level total or free-standing notes.
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    records = []
    for r in rows[1:]:
        rec = {h: v for h, v in zip(headers, r) if v is not None and h is not None}
        if rec and str(rec.get("Наименование позиции") or "").strip():
            records.append(rec)
    return records


# Currency symbol -> what it's called in the conversion note. Extend this
# (and pass the matching --<code>-rate on the CLI, lower-cased) if a tender
# ever quotes a market price in some other currency.
MARKET_PRICE_CURRENCIES = {"€": "EUR", "$": "USD"}


def convert_market_price_to_rub(raw_text, rates: dict = None):
    """Derive the clean "Рыночная цена, руб." value + a conversion note (or
    None, None) from the free-text "Рыночная цена (найдено в интернете)"
    cell.

    "Рыночная цена (найдено в интернете)" is deliberately left as plain text
    (see the SPEC_COLUMNS comment) because it routinely mixes currencies and
    commentary - e.g. "≈3 500 € за шт. (ориентировочно; ...)" or "$12 940,80
    (зарубежный дистрибьютор, ...)", not just RUB numbers with a comment.
    This function is what actually does the currency-aware parsing so that
    work doesn't have to be redone by hand for every tender (added
    2026-09-04, after doing it as a one-off for tender RAD260041848).

    rates: {"€": <RUB per EUR>, "$": <RUB per USD>, ...} - today's CBR rate
    for whichever foreign currencies actually show up in this tender's data
    (look them up with a web search, e.g. "курс доллара евро ЦБ РФ сегодня",
    and pass them via --usd-rate/--eur-rate). A currency symbol found in the
    text with no matching rate in `rates` is left unconverted (returns None
    for the value, with a note explaining why) rather than guessed at -
    silently treating "$12 940,80" as 12940.80 RUB would be badly wrong.

    Returns (rub_value_or_None, note_or_None).
    """
    if raw_text is None:
        return None, None
    text = str(raw_text).strip()
    if not text:
        return None, None
    lowered = text.lower()
    if lowered.startswith("не удалось определить") or lowered.startswith("не указано"):
        return None, None

    rates = rates or {}
    currency_symbol = None
    for sym in MARKET_PRICE_CURRENCIES:
        if sym in text:
            currency_symbol = sym
            break

    value = parse_ru_number(text)
    if value is None:
        return None, None

    if currency_symbol is None:
        # No foreign-currency symbol found - already RUB (₽/руб. or a bare number).
        return round(value, 2), None

    rate = rates.get(currency_symbol)
    if rate is None:
        return None, (
            f"Рыночная цена, руб.: не переведено — курс {currency_symbol}/руб не передан "
            f"(--{MARKET_PRICE_CURRENCIES[currency_symbol].lower()}-rate) для \"{text}\""
        )

    rub = round(value * rate, 2)
    note = (
        f"Рыночная цена, руб.: переведено из {value:.2f} {currency_symbol} по курсу ЦБ РФ "
        f"на {date.today().strftime('%d.%m.%Y')} (1 {currency_symbol} = {rate} руб.)"
    )
    return rub, note


def enrich_market_price_rub(records: list, rates: dict = None) -> list:
    """Fill in "Рыночная цена, руб." (and a conversion note in "Примечание")
    on each spec row, derived from its "Рыночная цена (найдено в интернете)"
    text - see convert_market_price_to_rub(). Does not mutate the input."""
    out = []
    for rec in records:
        rec = dict(rec)
        rub_value, note = convert_market_price_to_rub(rec.get("Рыночная цена (найдено в интернете)"), rates)
        rec["Рыночная цена, руб."] = rub_value
        if note:
            existing = str(rec.get("Примечание") or "").strip()
            rec["Примечание"] = f"{existing}; {note}" if existing else note
        out.append(rec)
    return out


def _normalize_compare(value):
    """Normalize a value (from either the row we sent or NocoDB's response) for
    equality comparison in verify_main_record/verify_spec_table below -
    stringifies, normalizes newlines, strips surrounding whitespace. None and
    empty string both collapse to "" so "field wasn't sent" and "field is
    literally blank" compare the same way."""
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def _values_match(expected, actual, numeric=False):
    """True if `expected` (what this run intended to send) and `actual` (what
    NocoDB now holds) should be considered the same value. numeric=True
    compares as floats with a small tolerance (Decimal columns round-trip
    through JSON as numbers, not strings) instead of exact text."""
    if numeric:
        try:
            e = float(expected) if expected not in (None, "") else None
        except (TypeError, ValueError):
            e = None
        try:
            a = float(actual) if actual not in (None, "") else None
        except (TypeError, ValueError):
            a = None
        if e is None and a is None:
            return True
        if e is None or a is None:
            return False
        return abs(e - a) < 0.005
    return _normalize_compare(expected) == _normalize_compare(actual)


def verify_main_record(token, record_id, expected_row):
    """Re-fetches the just-written record from "Реестр конкурсов" (Шаг 7.1 в
    CLAUDE.md) and compares every column this run actually sent (`expected_row`
    - the same dict passed to the create/update request) against what NocoDB
    now holds. Added 2026-09-09 after this project's memory already recorded
    two incidents where a column silently stayed empty in NocoDB even though
    the script itself reported success ("Файл-источник" on 2026-09-07,
    "Номер конкурса" missing as a column entirely, also found 2026-09-07) -
    this makes that class of bug visible automatically on every run instead of
    waiting for someone to notice by chance.

    Best-effort: if the verification GET itself fails (network hiccup, etc.)
    that's reported as a single issue rather than raised - the record write
    already succeeded by the time this runs, so a failed *check* shouldn't be
    reported as a failed *write*.

    Returns {"ok": bool, "checked_fields": int, "issues": [...]}. Each issue
    is {"column": ..., "expected": ..., "actual": ..., "problem": "пусто в
    NocoDB" | "значение не совпадает"}.
    """
    try:
        status, actual_record = _request("GET", f"{RECORDS_URL}/{record_id}", token)
    except Exception as e:
        return {"ok": False, "checked_fields": 0, "issues": [
            {"column": None, "expected": None, "actual": None,
             "problem": f"не удалось перечитать запись для проверки: {e}"}
        ]}
    if status != 200:
        return {"ok": False, "checked_fields": 0, "issues": [
            {"column": None, "expected": None, "actual": None,
             "problem": f"не удалось перечитать запись для проверки: HTTP {status} {actual_record}"}
        ]}

    issues = []
    for col, expected_value in expected_row.items():
        actual_value = actual_record.get(col)
        if not _values_match(expected_value, actual_value):
            issues.append({
                "column": col,
                "expected": expected_value,
                "actual": actual_value,
                "problem": "пусто в NocoDB" if _normalize_compare(actual_value) == "" else "значение не совпадает",
            })
    return {"ok": not issues, "checked_fields": len(expected_row), "issues": issues}


def verify_spec_table(token, table_id, expected_rows):
    """Re-fetches every row just written into a "Спецификация — ..." table
    (Шаг 7.1 в CLAUDE.md) and compares it against `expected_rows` (the
    already-sanitized rows sync_specification() actually sent - see
    sanitize_spec_row()): row count, "Наименование позиции", and the numeric
    columns (Кол-во / НМЦК за позицию / Рыночная цена, руб.). Matched by
    insertion order (NocoDB assigns Id sequentially within a bulk insert -
    the same assumption clear_table_records() already relies on for cleanup).
    Added 2026-09-09 alongside verify_main_record() - see its docstring for
    why.

    Returns {"ok": bool, "expected_rows": int, "actual_rows": int or None,
    "issues": [...]}. A row-count mismatch is reported on its own (row-by-row
    comparison is skipped in that case, since positions can no longer be
    trusted to line up). Each per-row issue is {"row": <1-based index>,
    "item": <"Наименование позиции">, "column": ..., "expected": ...,
    "actual": ..., "problem": ...}.
    """
    try:
        status, body = _request("GET", f"{BASE_URL}/{table_id}/records?limit=1000", token)
    except Exception as e:
        return {"ok": False, "expected_rows": len(expected_rows), "actual_rows": None, "issues": [
            {"row": None, "item": None, "column": None, "expected": None, "actual": None,
             "problem": f"не удалось перечитать таблицу спецификации для проверки: {e}"}
        ]}
    if status != 200:
        return {"ok": False, "expected_rows": len(expected_rows), "actual_rows": None, "issues": [
            {"row": None, "item": None, "column": None, "expected": None, "actual": None,
             "problem": f"не удалось перечитать таблицу спецификации для проверки: HTTP {status} {body}"}
        ]}

    actual_rows = sorted(body.get("list", []), key=lambda r: r.get("Id", 0))
    issues = []
    if len(actual_rows) != len(expected_rows):
        issues.append({
            "row": None, "item": None, "column": None, "expected": None, "actual": None,
            "problem": f"количество строк не совпадает: отправлено {len(expected_rows)}, найдено в NocoDB {len(actual_rows)}",
        })
    else:
        checked_columns = ("Наименование позиции",) + NUMERIC_SPEC_COLUMNS
        for idx, (expected, actual) in enumerate(zip(expected_rows, actual_rows), start=1):
            item_name = expected.get("Наименование позиции") or f"строка {idx}"
            for col in checked_columns:
                expected_value = expected.get(col)
                actual_value = actual.get(col)
                numeric = col in NUMERIC_SPEC_COLUMNS
                if not _values_match(expected_value, actual_value, numeric=numeric):
                    issues.append({
                        "row": idx,
                        "item": item_name,
                        "column": col,
                        "expected": expected_value,
                        "actual": actual_value,
                        "problem": "пусто в NocoDB" if _normalize_compare(actual_value) == "" else "значение не совпадает",
                    })
    return {"ok": not issues, "expected_rows": len(expected_rows), "actual_rows": len(actual_rows), "issues": issues}
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", required=True, help="Path to the flat extracted-fields JSON")
    parser.add_argument("--token-file", required=True, help="Path to the file containing the NocoDB API token")
    parser.add_argument("--xlsx", required=False, help="Path to Позиции_и_цены.xlsx - if given, also syncs it into its own NocoDB table and links it from the main record")
    parser.add_argument("--fallback-id", required=False, help="Identifier to name the spec table by when the tender has no real purchase_number (e.g. the tender's local folder name)")
    parser.add_argument("--usd-rate", required=False, type=float, help="Today's CBR RUB-per-USD rate (look it up first) - needed only if some position's Рыночная цена is quoted in $")
    parser.add_argument("--eur-rate", required=False, type=float, help="Today's CBR RUB-per-EUR rate (look it up first) - needed only if some position's Рыночная цена is quoted in €")
    args = parser.parse_args()

    with open(args.json, "r", encoding="utf-8") as f:
        flat_data = json.load(f)
    with open(args.token_file, "r", encoding="utf-8") as f:
        token = f.read().strip()

    fx_rates = {}
    if args.usd_rate:
        fx_rates["$"] = args.usd_rate
    if args.eur_rate:
        fx_rates["€"] = args.eur_rate

    result = upsert(token, flat_data, spec_xlsx_path=args.xlsx, fallback_id=args.fallback_id, fx_rates=fx_rates or None)

    verification = result.get("verification") or {}
    if not verification.get("ok", True):
        print(
            f"WARNING: post-write verification found {len(verification.get('issues', []))} "
            "issue(s) in the main record - see \"verification\" in the JSON output below.",
            file=sys.stderr,
        )
    spec_verification = result.get("spec_verification")
    if spec_verification and not spec_verification.get("ok", True):
        print(
            f"WARNING: post-write verification found {len(spec_verification.get('issues', []))} "
            "issue(s) in the specification table - see \"spec_verification\" in the JSON output below.",
            file=sys.stderr,
        )

    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
