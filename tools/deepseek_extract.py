#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deepseek_extract.py - sends tender document text to the DeepSeek API and asks
it to extract RAW, pre-formalization values for a fixed list of CRITICAL
fields, independently of Claude's own reading of the same documents.

WHY "RAW / pre-formalization":
    Claude's own extraction into tender_data.json already applies CLAUDE.md's
    formatting rules (date format, condensed payment-terms wording, etc.).
    Asking DeepSeek to do the same reformatting would make two independently
    -worded-but-equivalent answers look like a mismatch (e.g. "30 календарных
    дней" vs "30 дней"), which is noise, not a real discrepancy. So this
    script explicitly asks DeepSeek to quote/report values as literally
    written in the source text, NOT to normalize or reformat them - Claude
    compares those raw quotes against what Claude itself read from the same
    passages, and only THEN decides match/mismatch and writes the final,
    formalized value (with a ✅ or ⚠ marker) into tender_data.json.

WHAT THIS SCRIPT DOES NOT DO:
    It does not compare anything itself and does not talk to NocoDB or write
    any files. It only calls the DeepSeek API once and prints its JSON answer
    to stdout. Claude reads that JSON and does the comparison in its own
    reasoning (field by field), because judging whether two differently
    -worded quotes describe the same fact needs understanding, not a string
    diff.

WHERE THIS MUST RUN:
    Same constraint as tools/nocodb_client.py: network access to
    api.deepseek.com is only allowed from Claude's cloud execution
    environment once the domain has been added to the account's egress
    allowlist (Settings -> Capabilities -> Allow network egress ->
    Additional allowed domains). Not reachable from the local device shell.

Usage:
    python deepseek_extract.py --text /path/to/combined_tender_text.txt \
        --api-key-file /path/to/API_key_DeepSeek.txt \
        [--model deepseek-v4-pro] [--out /path/to/deepseek_raw.json]

    "combined_tender_text.txt" should be the plain-text (pandoc markdown is
    fine) content of the tender documents that are actually relevant to the
    critical fields below - typically: the main ДоЗ/Извещение document, the
    Технические требования appendix, and the Проект договора appendix.
    Very large, irrelevant appendices (e.g. a multi-thousand-line "Порядок
    оценки" scoring-methodology appendix) should be left out to save tokens
    and cost - they don't affect price/dates/quantities/payment/delivery.

Prints DeepSeek's JSON response to stdout: a flat dict
    {field_name: "raw quote/value as written in the source text"}
covering exactly CRITICAL_FIELDS below (some values may legitimately be
"не найдено в переданном тексте" if DeepSeek can't locate that field in
what it was given).
"""

import argparse
import json
import sys
import urllib.request
import urllib.error

API_URL = "https://api.deepseek.com/chat/completions"

# Critical-field allowlist agreed with the user (2026-09-02): price, all
# Block 4 dates, quantity/positions, payment terms, delivery term. Customer
# details (ИНН/КПП/ОГРН) were explicitly excluded as not critical enough to
# spend a cross-check on. Extend this list later if the user asks.
CRITICAL_FIELDS = {
    "start_price": "Начальная (максимальная) цена договора / лота, с валютой, ровно как написано в тексте (не пересчитывать, не переводить в другой формат числа).",
    "publication_date": "Дата официального размещения Извещения/Документации о закупке, как написана в тексте.",
    "submission_deadline": "Дата и время окончания срока подачи заявок, как написаны в тексте (включая часовой пояс, если указан).",
    "auction_date": "Дата проведения аукциона/переторжки/редукциона (или явное указание, что не предусмотрено), как в тексте.",
    "results_date": "Дата подведения итогов закупки / окончания рассмотрения заявок, как в тексте.",
    "contract_date": "Срок/порядок заключения договора (дата, либо правило вида 'не ранее X и не позднее Y дней с даты...'), как в тексте.",
    "quantity": "Перечень и количество закупаемых позиций (наименование, марка/артикул, количество, единица измерения) - точно как в таблице перечня/объёма товара.",
    "payment_terms": "Условия оплаты по договору: наличие и размер аванса, сроки платежей, от какого события отсчитываются сроки - точно как в тексте (не сокращать, не пересказывать своими словами).",
    "delivery_term": "Срок и/или место поставки товара/выполнения работ/оказания услуг, точно как в тексте.",
}

SYSTEM_PROMPT = (
    "Ты выполняешь техническую задачу извлечения текста из документации по закупке (тендеру). "
    "Твоя задача - НЕ формализовать, НЕ нормализовать и НЕ пересказывать своими словами, а "
    "процитировать/указать соответствующий фрагмент текста максимально близко к оригиналу, как он "
    "написан в предоставленном документе. Если для какого-то поля в тексте несколько фрагментов "
    "(например, аванс и окончательный платёж отдельно) - включи оба, каждый как написано. Если "
    "поле в переданном тексте не найдено - так и напиши: 'не найдено в переданном тексте' (не выдумывай "
    "и не подставляй значение из общих знаний о таких закупках). "
    "Ответь СТРОГО в виде JSON-объекта (без markdown, без пояснений вне JSON) с ключами ровно "
    "как в списке полей ниже."
)


def build_user_prompt(text: str) -> str:
    fields_desc = "\n".join(f"- \"{k}\": {v}" for k, v in CRITICAL_FIELDS.items())
    return (
        "Ниже текст документации о закупке (может включать несколько документов подряд, разделённых "
        "заголовками '=====FILE: ...====='). Извлеки из него дословные фрагменты для следующих полей:\n\n"
        f"{fields_desc}\n\n"
        "Текст документации:\n\n"
        "-----BEGIN DOCUMENT TEXT-----\n"
        f"{text}\n"
        "-----END DOCUMENT TEXT-----"
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--text", required=True, help="Path to combined plain-text/markdown tender documents")
    parser.add_argument("--api-key-file", required=True, help="Path to file containing the DeepSeek API key")
    parser.add_argument("--model", default="deepseek-v4-pro", help="DeepSeek model name (default: deepseek-v4-pro)")
    parser.add_argument("--out", default=None, help="Optional path to also save the JSON result to a file")
    args = parser.parse_args()

    with open(args.text, "r", encoding="utf-8") as f:
        text = f.read()

    with open(args.api_key_file, "r", encoding="utf-8") as f:
        api_key = f.read().strip()

    body = {
        "model": args.model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(text)},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }

    req = urllib.request.Request(
        API_URL,
        data=json.dumps(body).encode("utf-8"),
        method="POST",
    )
    req.add_header("Authorization", f"Bearer {api_key}")
    req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        print(json.dumps({"error": f"HTTP {e.code}", "body": err_body}), file=sys.stderr)
        sys.exit(1)
    except urllib.error.URLError as e:
        print(json.dumps({"error": "connection_failed", "detail": str(e)}), file=sys.stderr)
        sys.exit(1)

    response = json.loads(raw)
    try:
        content = response["choices"][0]["message"]["content"]
    except (KeyError, IndexError) as e:
        print(json.dumps({"error": "unexpected_response_shape", "raw": response}), file=sys.stderr)
        sys.exit(1)

    try:
        extracted = json.loads(content)
    except json.JSONDecodeError:
        # Model didn't return clean JSON despite response_format - surface raw text instead of failing silently.
        extracted = {"_raw_non_json_content": content}

    result = {
        "model": args.model,
        "usage": response.get("usage", {}),
        "fields": extracted,
    }

    output_json = json.dumps(result, ensure_ascii=False, indent=2)
    print(output_json)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(output_json)


if __name__ == "__main__":
    main()
