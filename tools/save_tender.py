#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
save_tender.py - saves data for one extracted tender (TenderDataExtractor project).

Handles the mechanical work that used to be done by hand by the LLM:
  - builds valid XML (with correct escaping of special characters);
  - normalizes numeric fields (price etc.) per the Block 2 rules in CLAUDE.md;
  - transliterates / sanitizes the output filename;
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
        --source-file "file1.docx + file2.pdf"

Prints a JSON result to stdout:
    {"xml_path": "...", "csv_path": "...", "csv_row_action": "appended"|"updated"}
"""

import argparse
import csv
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from xml.dom import minidom
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

# Summary CSV column order - must match the template in CLAUDE.md
CSV_COLUMNS = [
    "source_file", "extracted_date", "product_name", "quantity", "status",
    "delivery_place", "delivery_term", "procedure_type", "price_reduction",
    "price_reduction_date", "start_price", "currency", "payment_terms",
    "bid_security", "contract_security", "publication_date", "submission_deadline",
    "auction_date", "results_date", "contract_date", "last_modified", "organizer",
    "inn", "kpp", "ogrn", "contact", "address", "procurement_method", "platform",
    "platform_url", "purchase_number", "tender_url", "stage", "sme",
    "national_regime_applied", "national_regime_measure", "national_regime_items",
    "national_regime_basis",
]

NUMERIC_FIELDS = ("start_price",)

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
    """
    if raw is None:
        return NOT_SPECIFIED
    raw = str(raw).strip()
    if not raw or raw == NOT_SPECIFIED:
        return NOT_SPECIFIED

    # Find a CONTIGUOUS numeric token (starts and ends with a digit) so we don't
    # accidentally grab a stray period from an abbreviation like "руб." - that
    # bug used to multiply the amount by 100 (e.g. "1 500 000,00 руб." -> 150000000.00).
    match = re.search(r"\d(?:[\d\s.,]*\d)?", raw)
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

    return f"{value:.2f}"


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

    reverse_map = {(block, tag): flat for flat, (block, tag) in FIELD_MAP.items()}

    for block_name, tags in BLOCKS.items():
        block_el = ET.SubElement(root, block_name)
        for tag in tags:
            flat_name = reverse_map.get((block_name, tag))
            value = fields.get(flat_name) if flat_name else None
            if value is None or str(value).strip() == "":
                value = NOT_SPECIFIED
            ET.SubElement(block_el, tag).text = str(value)

    rough = ET.tostring(root, encoding="unicode")
    pretty = minidom.parseString(rough).toprettyxml(indent="  ")
    lines = [line for line in pretty.split("\n") if line.strip()]
    lines[0] = '<?xml version="1.0" encoding="UTF-8"?>'
    return "\n".join(lines) + "\n"


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
    args = parser.parse_args()

    with open(args.json, "r", encoding="utf-8") as f:
        fields = json.load(f)

    for numeric_field in NUMERIC_FIELDS:
        if numeric_field in fields:
            fields[numeric_field] = normalize_number(fields[numeric_field])

    extracted_date = fields.get("extracted_date") or date.today().isoformat()

    xml_text = build_xml(fields, args.source_file, extracted_date)
    ok, err = validate_xml(xml_text)
    if not ok:
        print(f"Error: generated XML is not valid ({err}), file not saved.", file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.extract_dir, exist_ok=True)
    xml_filename = build_filename(fields, extracted_date)
    xml_path = os.path.join(args.extract_dir, xml_filename)
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_text)

    csv_row = {"source_file": args.source_file, "extracted_date": extracted_date}
    for flat_name in FIELD_MAP:
        value = fields.get(flat_name, NOT_SPECIFIED)
        csv_row[flat_name] = value if value not in (None, "") else NOT_SPECIFIED

    updated = upsert_csv_row(args.csv_path, csv_row)

    result = {
        "xml_path": xml_path,
        "csv_path": args.csv_path,
        "csv_row_action": "updated" if updated else "appended",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
