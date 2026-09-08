#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
save_tender.py - saves data for one extracted tender (TenderDataExtractor project).

Handles the mechanical work that used to be done by hand by the LLM:
  - builds valid XML (with correct escaping of special characters);
  - normalizes numeric fields (price etc.) per the Block 2 rules in CLAUDE.md;
  - transliterates / sanitizes the output filename;
  - assigns/reconciles the internal_number (Шаг 1a в CLAUDE.md) - either uses
    a user-supplied value as-is, reuses the value already on record for this
    purchase_number, or auto-assigns the next "K-<N>" number (starting at
    K-983) and persists the counter so the next tender continues from there;
  - appends a row to the summary all_tenders.csv, or updates an existing row
    by purchase number (upsert) instead of creating duplicates;
  - validates the resulting XML before writing it to disk.

The LLM only needs to extract the data from the documents and pass it as a
flat JSON dict (field_name -> value); field names match the CSV column names.
If a field is not found in the documentation, its value should be the
"Не указано в документах" placeholder (NOT_SPECIFIED below), matching the
CLAUDE.md convention.

Usage:
    python save_tender.py \
        --json /path/to/extracted_fields.json \
        --extract-dir "/path/to/AIExtract" \
        --csv-path "/path/to/all_tenders.csv" \
        --source-file "file1.docx + file2.pdf" \
        [--internal-number-state /path/to/internal_number_state.txt]

Prints a JSON result to stdout:
    {"xml_path": "...", "csv_path": "...", "csv_row_action": "appended"|"updated",
     "internal_number": "K-983"}
"""

import argparse
import csv
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date

NOT_SPECIFIED = "Не указано в документах"

BLOCKS = {
    "block1_product": ["name", "quantity", "status", "delivery_place", "delivery_term",
                        "procedure_type", "price_reduction", "price_reduction_date"],
    "block2_price": ["start_price", "currency", "payment_terms"],
    "block3_security": ["bid_security", "contract_security"],
    "block4_dates": ["publication_date", "submission_deadline", "auction_date",
                      "results_date", "contract_date", "last_modified"],
    "block5_organizer": ["name", "inn", "kpp", "ogrn", "contact", "address"],
    "block6_procedure": ["procurement_method", "platform", "platform_url",
                          "purchase_number", "tender_url", "stage", "sme"],
    "block7_national_regime": ["applied", "measure_type", "restricted_items", "basis"],
}

# Flat field name (used in the LLM's JSON and in CSV columns) -> (block, xml tag)
FIELD_MAP = {
    "product_name": ("block1_product", "name"),
    "quantity": ("block1_product", "quantity"),
    "status": ("block1_product", "status"),
    "delivery_place": ("block1_product", "delivery_place"),
    "delivery_term": ("block1_product", "delivery_term"),
    "procedure_type": ("block1_product", "procedure_type"),
    "price_reduction": ("block1_product", "price_reduction"),
    "price_reduction_date": ("block1_product", "price_reduction_date"),

    "start_price": ("block2_price", "start_price"),
    "currency": ("block2_price", "currency"),
    "payment_terms": ("block2_price", "payment_terms"),

    "bid_security": ("block3_security", "bid_security"),
    "contract_security": ("block3_security", "contract_security"),

    "publication_date": ("block4_dates", "publication_date"),
    "submission_deadline": ("block4_dates", "submission_deadline"),
    "auction_date": ("block4_dates", "auction_date"),
    "results_date": ("block4_dates", "results_date"),
    "contract_date": ("block4_dates", "contract_date"),
    "last_modified": ("block4_dates", "last_modified"),

    "organizer": ("block5_organizer", "name"),
    "inn": ("block5_organizer", "inn"),
    "kpp": ("block5_organizer", "kpp"),
    "ogrn": ("block5_organizer", "ogrn"),
    "contact": ("block5_organizer", "contact"),
    "address": ("block5_organizer", "address"),

    "procurement_method": ("block6_procedure", "procurement_method"),
    "platform": ("block6_procedure", "platform"),
    "platform_url": ("block6_procedure", "platform_url"),
    "purchase_number": ("block6_procedure", "purchase_number"),
    "tender_url": ("block6_procedure", "tender_url"),
    "stage": ("block6_procedure", "stage"),
    "sme": ("block6_procedure", "sme"),

    "national_regime_applied": ("block7_national_regime", "applied"),
    "national_regime_measure": ("block7_national_regime", "measure_type"),
    "national_regime_items": ("block7_national_regime", "restricted_items"),
    "national_regime_basis": ("block7_national_regime", "basis"),
}

# Summary CSV column order - must match the template in CLAUDE.md.
# "internal_number" is not one of the 7 extraction blocks (it's assigned by
# this script, not read from the tender documents - see Шаг 1a в CLAUDE.md)
# but lives as the first column since it's the user's own record identifier.
# "purchase_number" and "tender_url" (normally Block 6 fields, alongside
# procurement_method/platform/platform_url/stage/sme) are pulled out of block
# order and placed right after "internal_number" instead - by user request
# (2026-09-07): the three tender-identifying columns (internal number,
# извлечённый номер закупки, ссылка на тендер) read better grouped together
# at the front, ahead of all the extracted detail fields. This mirrors the
# same positioning done for NocoDB's "Номер конкурса" column - see
# AUTO_CREATE_COLUMN_POSITION_AFTER in tools/nocodb_client.py.
CSV_COLUMNS = [
    "internal_number", "purchase_number", "tender_url",
    "source_file", "extracted_date", "product_name", "quantity", "status",
    "delivery_place", "delivery_term", "procedure_type", "price_reduction",
    "price_reduction_date", "start_price", "currency", "payment_terms",
    "bid_security", "contract_security", "publication_date", "submission_deadline",
    "auction_date", "results_date", "contract_date", "last_modified", "organizer",
    "inn", "kpp", "ogrn", "contact", "address", "procurement_method", "platform",
    "platform_url", "stage", "sme",
    "national_regime_applied", "national_regime_measure", "national_regime_items",
    "national_regime_basis",
]

NUMERIC_FIELDS = ("start_price",)

# Fields where a missing value is written as a plain dash "—" instead of the
# usual NOT_SPECIFIED placeholder text - by user request (2026-09-07): a
# tender with no link to its procurement documentation (tender_url -
# "Ссылка на тендер" / "Ссылка на закупочную документацию", Блок 6) should
# show "—" in the table, not the full "Не указано в документах" phrase.
# Claude is instructed (CLAUDE.md) to write "—" directly when assembling the
# JSON for Шаг 7/7.1, so this is mainly a safety net for XML/CSV in case that
# value arrives missing, empty, or as the old NOT_SPECIFIED placeholder text.
DASH = "—"
DASH_FIELDS = ("tender_url",)


def apply_dash_placeholders(fields):
    """Normalizes DASH_FIELDS in-place: missing/empty/NOT_SPECIFIED -> "—".
    Returns `fields` for convenience. Must run before build_xml()/the CSV row
    is assembled, so both pick up the normalized value the same way."""
    for flat_name in DASH_FIELDS:
        value = str(fields.get(flat_name) or "").strip()
        if not value or value == NOT_SPECIFIED:
            fields[flat_name] = DASH
    return fields


# --- internal_number (Шаг 1a в CLAUDE.md) -----------------------------------
# Auto-assigned numbers look like "K-983", "K-984", ... The counter is
# persisted in a plain text file (just the last-issued integer) so it
# survives across separate runs/sessions.
INTERNAL_NUMBER_PREFIX = "K-"
INTERNAL_NUMBER_START = 983  # first auto-assigned number, when no state file exists yet
INTERNAL_NUMBER_TRAILING_DIGITS_RE = re.compile(r"(\d+)\s*$")


def _internal_number_state_path(explicit_path, csv_path):
    if explicit_path:
        return explicit_path
    base_dir = os.path.dirname(os.path.abspath(csv_path)) or "."
    return os.path.join(base_dir, "internal_number_state.txt")


def _read_counter(state_path):
    """Returns the last-issued integer, or (START - 1) if there's no state
    file yet / it's unreadable - so the very first auto-assignment is START."""
    if not os.path.isfile(state_path):
        return INTERNAL_NUMBER_START - 1
    try:
        with open(state_path, "r", encoding="utf-8") as f:
            text = f.read().strip()
        return int(text)
    except (ValueError, OSError):
        return INTERNAL_NUMBER_START - 1


def _write_counter(state_path, value):
    parent = os.path.dirname(state_path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(state_path, "w", encoding="utf-8") as f:
        f.write(str(value))


def find_existing_row(csv_path, purchase_number):
    """Looks up an already-saved CSV row by purchase_number (same matching
    key as upsert_csv_row) - used to reuse its internal_number on a re-save
    instead of minting a new one every time the same tender is re-processed."""
    if not os.path.isfile(csv_path):
        return None
    if not purchase_number or purchase_number == NOT_SPECIFIED:
        return None
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if str(row.get("purchase_number") or "").strip() == purchase_number:
                return row
    return None


def resolve_internal_number(candidate, state_path):
    """Decides the final internal_number to save and updates the counter
    state file accordingly.

    candidate: whatever the caller already has for this tender - either a
    value the user explicitly chose, a value reused from an existing CSV row
    (see find_existing_row), or empty/NOT_SPECIFIED when neither applies.

    - candidate given (non-empty, not NOT_SPECIFIED) -> used as-is. If its
      trailing digits are >= the counter, the counter is bumped to match so
      a later auto-assignment never collides with it.
    - candidate empty -> auto-assign "K-<counter + 1>" and persist the new
      counter value.
    """
    candidate = str(candidate or "").strip()
    last = _read_counter(state_path)
    if candidate and candidate != NOT_SPECIFIED:
        match = INTERNAL_NUMBER_TRAILING_DIGITS_RE.search(candidate)
        if match:
            n = int(match.group(1))
            if n > last:
                _write_counter(state_path, n)
        return candidate
    next_n = last + 1
    _write_counter(state_path, next_n)
    return f"{INTERNAL_NUMBER_PREFIX}{next_n}"


TRANSLIT_TABLE = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d",
    "е": "e", "ё": "e", "ж": "zh", "з": "z", "и": "i",
    "й": "y", "к": "k", "л": "l", "м": "m", "н": "n",
    "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
    "у": "u", "ф": "f", "х": "h", "ц": "ts", "ч": "ch",
    "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "",
    "э": "e", "ю": "yu", "я": "ya",
}


def normalize_number(raw):
    """
    "1 500 000,00 руб." -> "1500000.00"
    "1500000" -> "1500000.00"
    "Не указано в документах" / empty / not a number -> returned unchanged.

    A leading DeepSeek cross-check marker from Шаг 5.2 in CLAUDE.md (`✅ ` for
    a confirmed match, `🔺 ` for an unresolved discrepancy - the latter
    normally followed by a "(расхождение: ...)" note) is preserved around the
    normalized number instead of being silently dropped. Шаг 5.2 requires the
    marker to survive into XML/CSV/NocoDB, and Block 4 dates (never numeric
    -normalized) already keep it correctly - price should behave the same way.
    """
    if raw is None:
        return NOT_SPECIFIED
    raw = str(raw).strip()
    if not raw or raw == NOT_SPECIFIED:
        return NOT_SPECIFIED

    marker_match = re.match(r"^(✅|🔺)\s*", raw)
    marker = marker_match.group(0) if marker_match else ""
    body = raw[marker_match.end():] if marker_match else raw

    # Find a CONTIGUOUS numeric token (starts and ends with a digit) so we don't
    # accidentally grab a stray period from an abbreviation like "руб." - that
    # bug used to multiply the amount by 100 (e.g. "1 500 000,00 руб." -> 150000000.00).
    match = re.search(r"\d(?:[\d\s.,]*\d)?", body)
    if not match:
        return raw  # doesn't look like a number at all - leave as is

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
            # looks like "1.500" (thousands separator), not a decimal number
            cleaned = cleaned.replace(".", "")

    try:
        value = float(cleaned)
    except ValueError:
        return raw

    normalized = f"{value:.2f}"
    if marker:
        # Keep the marker and whatever text follows the number (currency word,
        # or the 🔺 discrepancy explanation) - only the number itself gets
        # normalized, matching what Шаг 5.2 requires for a marked value.
        tail = body[match.end():]
        return f"{marker}{normalized}{tail}"
    return normalized


def transliterate(text):
    """Cyrillic -> Latin. Used for the organizer's short name when the
    purchase number is unknown."""
    out = []
    for ch in text:
        lower = ch.lower()
        if lower in TRANSLIT_TABLE:
            t = TRANSLIT_TABLE[lower]
            out.append(t.upper() if ch.isupper() and t else t)
        else:
            out.append(ch)
    return "".join(out)


def sanitize_filename_part(text):
    """General filename-part cleanup: strip special chars, spaces -> _.
    Does NOT touch Cyrillic - the purchase number is used as-is per CLAUDE.md."""
    text = re.sub(r'[«»(),.:;"\']', "", text)
    text = re.sub(r"\s+", "_", text.strip())
    text = re.sub(r"[^\w\-]", "", text, flags=re.UNICODE)
    return text


def format_source_files(raw):
    """"файл1.docx + файл2.pdf + файл3.docx" (--source-file, Шаг 7 в
    CLAUDE.md) -> "файл1.docx\\nфайл2.pdf\\nфайл3.docx" - one file per line
    instead of a single " + "-joined run-on string. By user request
    (2026-09-07): "Файл-источник" in NocoDB is already a rich-text column
    (richMode, like "Наименование предмета закупки" - see Шаг 1 formatting
    rule в CLAUDE.md) so the same "list, not one line" treatment applies
    here too. A single file (no " + " found) is returned unchanged - no
    point turning one filename into a one-item "list"."""
    if raw is None:
        return raw
    parts = [p.strip() for p in str(raw).split(" + ") if p.strip()]
    if len(parts) <= 1:
        return raw
    return "\n".join(parts)


def build_filename(fields, extracted_date):
    purchase_number = str(fields.get("purchase_number") or "").strip()
    if purchase_number and purchase_number != NOT_SPECIFIED:
        base = sanitize_filename_part(purchase_number)
    else:
        organizer = str(fields.get("organizer") or "").strip()
        base = transliterate(organizer) if organizer else "TENDER"
        base = sanitize_filename_part(base)[:40]
        if not base:
            base = "TENDER"
    return f"{base}_{extracted_date}.xml"


def build_xml(fields, source_file, extracted_date):
    root = ET.Element("tender")
    meta = ET.SubElement(root, "meta")
    ET.SubElement(meta, "source_file").text = source_file
    ET.SubElement(meta, "extracted_date").text = extracted_date
    ET.SubElement(meta, "internal_number").text = str(fields.get("internal_number") or NOT_SPECIFIED)

    reverse_map = {(block, tag): flat for flat, (block, tag) in FIELD_MAP.items()}

    for block_name, tags in BLOCKS.items():
        block_el = ET.SubElement(root, block_name)
        for tag in tags:
            flat_name = reverse_map.get((block_name, tag))
            value = fields.get(flat_name) if flat_name else None
            if value is None or str(value).strip() == "":
                value = NOT_SPECIFIED
            ET.SubElement(block_el, tag).text = str(value)

    # ET.indent() (Python 3.9+) only touches whitespace BETWEEN elements
    # (each leaf element's own .text is left exactly as set above) - unlike
    # the previous minidom.toprettyxml() + "drop blank lines" approach this
    # replaced, which silently ate blank lines that were part of a value's
    # own text (found 2026-09-07: a multi-line "Наименование предмета
    # закупки" value with a deliberate blank line before "Список позиций:"
    # came back on a validate_xml() round-trip with that blank line gone -
    # the document was still well-formed XML, just silently missing data).
    # Confirmed empirically that ET.indent() round-trips a multi-line value
    # (including embedded blank lines) exactly - see Шаг 1 (Наименование
    # продукции) formatting rule в CLAUDE.md, which relies on this.
    ET.indent(root, space="  ")
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding="unicode") + "\n"


def validate_xml(xml_text):
    try:
        ET.fromstring(xml_text)
        return True, None
    except ET.ParseError as e:
        return False, str(e)


def upsert_csv_row(csv_path, row):
    """Appends a row to the CSV, or updates an existing one by purchase_number."""
    rows = []
    if os.path.isfile(csv_path):
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

    key = str(row.get("purchase_number") or "").strip()
    updated = False
    if key and key != NOT_SPECIFIED:
        for i, existing in enumerate(rows):
            if str(existing.get("purchase_number") or "").strip() == key:
                rows[i] = row
                updated = True
                break
    if not updated:
        rows.append(row)

    parent_dir = os.path.dirname(csv_path)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for r in rows:
            writer.writerow({col: (r.get(col) or NOT_SPECIFIED) for col in CSV_COLUMNS})

    return updated


def main():
    parser = argparse.ArgumentParser(description="Save tender data to XML + update the summary CSV")
    parser.add_argument("--json", required=True, help="Path to a JSON file with the extracted fields")
    parser.add_argument("--extract-dir", required=True, help="AIExtract folder to save the XML into")
    parser.add_argument("--csv-path", required=True, help="Path to the summary all_tenders.csv")
    parser.add_argument("--source-file", required=True, help="Processed source files joined by ' + '")
    parser.add_argument("--internal-number-state", required=False,
                         help="Path to the internal_number counter state file "
                              "(default: internal_number_state.txt next to --csv-path)")
    args = parser.parse_args()

    with open(args.json, "r", encoding="utf-8") as f:
        fields = json.load(f)

    apply_dash_placeholders(fields)

    for numeric_field in NUMERIC_FIELDS:
        if numeric_field in fields:
            fields[numeric_field] = normalize_number(fields[numeric_field])

    extracted_date = fields.get("extracted_date") or date.today().isoformat()

    # internal_number (Шаг 1a в CLAUDE.md): priority is
    # (1) a value the user explicitly supplied in this JSON,
    # (2) the value already on record for this purchase_number (re-save/update
    #     must not mint a new number every time),
    # (3) auto-assign the next "K-<N>" from the persisted counter.
    purchase_number = str(fields.get("purchase_number") or "").strip()
    provided_internal = str(fields.get("internal_number") or "").strip()
    if not provided_internal or provided_internal == NOT_SPECIFIED:
        existing_row = find_existing_row(args.csv_path, purchase_number)
        if existing_row:
            existing_internal = str(existing_row.get("internal_number") or "").strip()
            if existing_internal and existing_internal != NOT_SPECIFIED:
                provided_internal = existing_internal

    state_path = _internal_number_state_path(args.internal_number_state, args.csv_path)
    fields["internal_number"] = resolve_internal_number(provided_internal, state_path)

    source_file = format_source_files(args.source_file)

    xml_text = build_xml(fields, source_file, extracted_date)
    ok, err = validate_xml(xml_text)
    if not ok:
        print(f"Error: generated XML is not valid ({err}), file not saved.", file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.extract_dir, exist_ok=True)
    xml_filename = build_filename(fields, extracted_date)
    xml_path = os.path.join(args.extract_dir, xml_filename)
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_text)

    csv_row = {
        "internal_number": fields["internal_number"],
        "source_file": source_file,
        "extracted_date": extracted_date,
    }
    for flat_name in FIELD_MAP:
        value = fields.get(flat_name, NOT_SPECIFIED)
        csv_row[flat_name] = value if value not in (None, "") else NOT_SPECIFIED

    updated = upsert_csv_row(args.csv_path, csv_row)

    result = {
        "xml_path": xml_path,
        "csv_path": args.csv_path,
        "csv_row_action": "updated" if updated else "appended",
        "internal_number": fields["internal_number"],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
