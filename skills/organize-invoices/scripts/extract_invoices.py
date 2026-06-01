#!/usr/bin/env python3
"""Read-only extractor for Chinese electronic invoice PDFs.

The script prints a JSON summary that another Codex step can review before any
Excel creation, PDF rename, or deletion happens.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

try:
    from pypdf import PdfReader
except Exception as exc:  # pragma: no cover - environment setup issue
    PdfReader = None
    PYPDF_IMPORT_ERROR = exc
else:
    PYPDF_IMPORT_ERROR = None


DATE_CN_RE = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
DATE_ISO_RE = re.compile(r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})")
TAX_ID_RE = re.compile(r"(?<![A-Z0-9])([0-9A-Z]{18})(?![A-Z0-9])")
COMPANY_RE = re.compile(
    r"(?<![\u4e00-\u9fffA-Za-z0-9])"
    r"([A-Za-z0-9\u4e00-\u9fff（）()·\-]{2,80}"
    r"(?:有限责任公司|股份有限公司|集团公司|有限公司|公司|加油站|火锅店|店))"
    r"(?![\u4e00-\u9fffA-Za-z0-9])"
)
LEADING_CURRENCY_RE = re.compile(r"[¥￥]\s*(-?\d+(?:,\d{3})*(?:\.\d{1,2})?)")
TRAILING_CURRENCY_RE = re.compile(r"(-?\d+(?:,\d{3})*(?:\.\d{1,2})?)\s*[¥￥]")
INVOICE_NO_RE = re.compile(r"发票号码\s*[:：]?\s*([0-9]{8,30})")
INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\|?*\r\n\t]+')


def decimal_to_string(value: Decimal | None) -> str | None:
    if value is None:
        return None
    return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def extract_text(path: Path) -> tuple[str, int, str | None]:
    if PdfReader is None:
        return "", 0, f"pypdf import failed: {PYPDF_IMPORT_ERROR}"
    try:
        reader = PdfReader(str(path))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        return text, len(reader.pages), None
    except Exception as exc:
        return "", 0, f"{type(exc).__name__}: {exc}"


def find_invoice_date(text: str) -> str | None:
    match = DATE_CN_RE.search(text) or DATE_ISO_RE.search(text)
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    return f"{year:04d}-{month:02d}-{day:02d}"


def find_invoice_number(text: str) -> str | None:
    match = INVOICE_NO_RE.search(normalize_spaces(text))
    if match:
        return match.group(1)
    numbers = re.findall(r"(?<!\d)(2[0-9]{19})(?!\d)", text)
    return numbers[0] if numbers else None


def find_currency_amounts(text: str) -> list[dict[str, Any]]:
    amounts: list[dict[str, Any]] = []
    seen: set[tuple[int, str]] = set()
    for pattern in (LEADING_CURRENCY_RE, TRAILING_CURRENCY_RE):
        for match in pattern.finditer(text):
            raw = match.group(1).replace(",", "")
            key = (match.start(), raw)
            if key in seen:
                continue
            seen.add(key)
            try:
                value = Decimal(raw)
            except Exception:
                continue
            amounts.append({"value": value, "position": match.start(), "raw": match.group(0)})
    return amounts


def choose_total_amount(text: str) -> tuple[Decimal | None, list[str]]:
    warnings: list[str] = []

    for marker in re.finditer("价税合计", text):
        tail = text[marker.start() : marker.start() + 260]
        nearby_positive = [item["value"] for item in find_currency_amounts(tail) if item["value"] > 0]
        if nearby_positive:
            # PDF text order varies. Near the total marker, the price-tax total is
            # normally the largest positive currency amount in the local window.
            return max(nearby_positive), warnings

    compact = re.sub(r"\s+", "", text)
    marker = compact.find("价税合计")
    if marker >= 0:
        tail = compact[marker : marker + 180]
        nums = [Decimal(raw) for raw in re.findall(r"-?\d+\.\d{2}", tail)]
        positive_nums = [value for value in nums if value > 0]
        if positive_nums:
            warnings.append("amount inferred from numbers near 价税合计 without currency symbol")
            return max(positive_nums), warnings

    positive = [item["value"] for item in find_currency_amounts(text) if item["value"] > 0]
    if positive:
        warnings.append("amount inferred from largest currency value because 价税合计 marker was not usable")
        return max(positive), warnings

    return None, ["amount not found"]


def find_entities(text: str) -> list[dict[str, str]]:
    flat = normalize_spaces(text)
    names = []
    for match in COMPANY_RE.finditer(flat):
        name = match.group(1)
        if name.startswith(("国家税务总局", "全国统一发票")):
            continue
        names.append({"name": name, "start": match.start(1), "end": match.end(1)})

    taxes = [{"tax_id": match.group(1), "start": match.start(1)} for match in TAX_ID_RE.finditer(flat)]

    entities: list[dict[str, str]] = []
    used_tax_indexes: set[int] = set()
    for name in names:
        selected_index = None
        for index, tax in enumerate(taxes):
            if index in used_tax_indexes:
                continue
            if tax["start"] >= name["end"] and tax["start"] - name["end"] <= 180:
                selected_index = index
                break
        if selected_index is None:
            continue
        used_tax_indexes.add(selected_index)
        entities.append({"name": name["name"], "tax_id": taxes[selected_index]["tax_id"]})

    if not entities and len(names) == len(taxes):
        entities = [{"name": name["name"], "tax_id": tax["tax_id"]} for name, tax in zip(names, taxes)]

    deduped = []
    seen_pairs = set()
    for entity in entities:
        key = (entity["name"], entity["tax_id"])
        if key not in seen_pairs:
            seen_pairs.add(key)
            deduped.append(entity)
    return deduped


def infer_batch_buyer_tax(records: list[dict[str, Any]]) -> str | None:
    counts: Counter[str] = Counter()
    for record in records:
        for entity in record.get("entities", []):
            counts[entity["tax_id"]] += 1
    if not counts:
        return None
    tax_id, count = counts.most_common(1)[0]
    if count >= 2 and count > len(records) / 2:
        return tax_id
    return None


def choose_seller(record: dict[str, Any], buyer_tax: str | None) -> tuple[dict[str, str] | None, list[str]]:
    entities = record.get("entities", [])
    warnings: list[str] = []
    if buyer_tax:
        sellers = [entity for entity in entities if entity["tax_id"] != buyer_tax]
        if len(sellers) == 1:
            return sellers[0], warnings
        if len(sellers) > 1:
            warnings.append("multiple non-buyer entities found; seller requires review")
            return sellers[-1], warnings

    if len(entities) == 2:
        warnings.append("seller inferred as second entity because no repeated buyer tax was detected")
        return entities[1], warnings
    if len(entities) == 1:
        warnings.append("only one entity found; treating it as seller requires review")
        return entities[0], warnings
    return None, ["seller entity not found"]


def sanitize_filename_part(value: str) -> str:
    value = INVALID_FILENAME_CHARS_RE.sub("_", value)
    value = re.sub(r"\s+", "", value)
    return value.strip(" .")


def build_records(folder: Path) -> dict[str, Any]:
    pdfs = sorted(folder.glob("*.pdf"), key=lambda path: path.name.casefold())
    raw_records: list[dict[str, Any]] = []

    for pdf in pdfs:
        text, pages, error = extract_text(pdf)
        amount, amount_warnings = choose_total_amount(text)
        record: dict[str, Any] = {
            "source_file": pdf.name,
            "pages": pages,
            "invoice_number": find_invoice_number(text),
            "invoice_date": find_invoice_date(text),
            "amount": decimal_to_string(amount),
            "entities": find_entities(text),
            "warnings": [],
        }
        if error:
            record["warnings"].append(error)
        record["warnings"].extend(amount_warnings)
        raw_records.append(record)

    buyer_tax = infer_batch_buyer_tax(raw_records)
    buyer_entity = None
    if buyer_tax:
        for record in raw_records:
            for entity in record.get("entities", []):
                if entity["tax_id"] == buyer_tax:
                    buyer_entity = entity
                    break
            if buyer_entity:
                break

    for record in raw_records:
        seller, seller_warnings = choose_seller(record, buyer_tax)
        record["warnings"].extend(seller_warnings)
        record["seller_name"] = seller["name"] if seller else None
        record["seller_tax_id"] = seller["tax_id"] if seller else None
        if record["invoice_date"] and record["seller_name"]:
            record["proposed_pdf_name"] = f"{record['invoice_date']}-{sanitize_filename_part(record['seller_name'])}.pdf"
        else:
            record["proposed_pdf_name"] = None

    records = sorted(raw_records, key=lambda item: (item.get("invoice_date") or "9999-99-99", item["source_file"]))

    missing_required = []
    for record in records:
        missing = [
            field
            for field in ("invoice_date", "seller_tax_id", "seller_name", "amount")
            if not record.get(field)
        ]
        if missing:
            missing_required.append({"source_file": record["source_file"], "missing": missing})

    duplicate_groups = []
    by_seller_tax: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        if record.get("seller_tax_id"):
            by_seller_tax[record["seller_tax_id"]].append(record)
    for tax_id, group in sorted(by_seller_tax.items()):
        if len(group) > 1:
            duplicate_groups.append(
                {
                    "seller_tax_id": tax_id,
                    "records": [
                        {
                            "source_file": item["source_file"],
                            "invoice_date": item.get("invoice_date"),
                            "seller_name": item.get("seller_name"),
                            "amount": item.get("amount"),
                        }
                        for item in group
                    ],
                }
            )

    rename_collisions = []
    by_target: defaultdict[str, list[str]] = defaultdict(list)
    for record in records:
        target = record.get("proposed_pdf_name")
        if target:
            by_target[target].append(record["source_file"])
    for target, sources in sorted(by_target.items()):
        if len(sources) > 1:
            rename_collisions.append({"target": target, "sources": sources})

    total = sum((Decimal(record["amount"]) for record in records if record.get("amount")), Decimal("0.00"))

    return {
        "folder": str(folder),
        "pdf_count": len(pdfs),
        "buyer_inferred": buyer_entity,
        "records": records,
        "total_amount": decimal_to_string(total),
        "missing_required": missing_required,
        "duplicate_seller_tax_groups": duplicate_groups,
        "rename_collisions": rename_collisions,
        "warnings": [
            {"source_file": record["source_file"], "warnings": record["warnings"]}
            for record in records
            if record.get("warnings")
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract invoice fields from a folder of Chinese invoice PDFs.")
    parser.add_argument("folder", help="Folder containing invoice PDFs")
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output")
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        print(json.dumps({"error": f"Folder not found: {folder}"}, ensure_ascii=False), file=sys.stderr)
        return 2

    result = build_records(folder)
    print(json.dumps(result, ensure_ascii=False, indent=2 if args.pretty else None))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
