#!/usr/bin/env python3
"""Patch selected Excel columns to sharedStrings text cells.

Use after workbook export for invoice identifiers such as tax IDs and invoice
numbers. This avoids scientific notation in Excel-compatible viewers.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"x": MAIN_NS, "ct": CONTENT_TYPES_NS}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("", CONTENT_TYPES_NS)

DEFAULT_HEADERS = [
    "税号",
    "销售方税号",
    "购买方税号",
    "统一社会信用代码",
    "纳税人识别号",
    "发票号码",
]


def split_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def column_letters(cell_ref: str) -> str:
    match = re.match(r"([A-Z]+)", cell_ref)
    return match.group(1) if match else ""


def row_number(cell_ref: str) -> int | None:
    match = re.search(r"(\d+)$", cell_ref)
    return int(match.group(1)) if match else None


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        value = cell.find("x:v", NS)
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    if cell_type == "inlineStr":
        return "".join((node.text or "") for node in cell.findall(".//x:t", NS))
    value = cell.find("x:v", NS)
    return value.text if value is not None and value.text is not None else ""


def load_shared_strings(data: dict[str, bytes]) -> tuple[ET.Element, list[str]]:
    filename = "xl/sharedStrings.xml"
    if filename in data and data[filename].strip():
        root = ET.fromstring(data[filename])
    else:
        root = ET.Element(f"{{{MAIN_NS}}}sst")

    values: list[str] = []
    for item in root.findall("x:si", NS):
        values.append("".join((node.text or "") for node in item.findall(".//x:t", NS)))
    return root, values


def shared_string_index(root: ET.Element, values: list[str], text: str) -> int:
    try:
        return values.index(text)
    except ValueError:
        item = ET.SubElement(root, f"{{{MAIN_NS}}}si")
        text_node = ET.SubElement(item, f"{{{MAIN_NS}}}t")
        text_node.text = text
        values.append(text)
        return len(values) - 1


def set_cell_shared_string(cell: ET.Element, text: str, index: int) -> None:
    ref = cell.attrib.get("r")
    style = cell.attrib.get("s")
    cell.clear()
    if ref:
        cell.set("r", ref)
    if style is not None:
        cell.set("s", style)
    cell.set("t", "s")
    value = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value.text = str(index)


def patch_content_types(data: dict[str, bytes]) -> None:
    filename = "[Content_Types].xml"
    if filename not in data:
        return
    root = ET.fromstring(data[filename])
    exists = any(
        node.attrib.get("PartName") == "/xl/sharedStrings.xml"
        for node in root.findall("ct:Override", NS)
    )
    if not exists:
        override = ET.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override")
        override.set("PartName", "/xl/sharedStrings.xml")
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )
    data[filename] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def patch_workbook(path: Path, headers: set[str], columns: set[str], header_row: int) -> int:
    with zipfile.ZipFile(path, "r") as zin:
        items = zin.infolist()
        data = {item.filename: zin.read(item.filename) for item in items}

    shared_root, shared_values = load_shared_strings(data)
    patched = 0

    for filename in list(data):
        if not filename.startswith("xl/worksheets/") or not filename.endswith(".xml"):
            continue

        root = ET.fromstring(data[filename])
        target_columns = set(columns)

        for cell in root.findall(".//x:c", NS):
            ref = cell.attrib.get("r", "")
            if row_number(ref) != header_row:
                continue
            if cell_text(cell, shared_values).strip() in headers:
                target_columns.add(column_letters(ref))

        if not target_columns:
            continue

        for cell in root.findall(".//x:c", NS):
            ref = cell.attrib.get("r", "")
            row = row_number(ref)
            if row is None or row <= header_row:
                continue
            if column_letters(ref) not in target_columns:
                continue
            text = cell_text(cell, shared_values).strip()
            if not text:
                continue
            index = shared_string_index(shared_root, shared_values, text)
            set_cell_shared_string(cell, text, index)
            patched += 1

        data[filename] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    shared_root.set("count", str(len(shared_values)))
    shared_root.set("uniqueCount", str(len(shared_values)))
    data["xl/sharedStrings.xml"] = ET.tostring(shared_root, encoding="utf-8", xml_declaration=True)
    patch_content_types(data)

    fd, temp_name = tempfile.mkstemp(prefix="xlsx_text_patch_", suffix=".xlsx", dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for item in items:
                zout.writestr(item, data[item.filename])
                written.add(item.filename)
            for filename, content in data.items():
                if filename not in written:
                    zout.writestr(filename, content)
        shutil.move(str(temp_path), str(path))
    except Exception:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
        raise

    return patched


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch XLSX identifier columns to sharedStrings text cells.")
    parser.add_argument("xlsx", help="Path to the XLSX workbook to patch in place")
    parser.add_argument("--headers", default=",".join(DEFAULT_HEADERS), help="Comma-separated header names to patch")
    parser.add_argument("--columns", default="", help="Comma-separated Excel column letters to patch, such as B,D")
    parser.add_argument("--header-row", type=int, default=1, help="1-based header row number")
    args = parser.parse_args()

    path = Path(args.xlsx).expanduser().resolve()
    if not path.exists() or path.suffix.lower() != ".xlsx":
        raise SystemExit(f"XLSX file not found: {path}")

    headers = set(split_csv(args.headers))
    columns = {column.upper() for column in split_csv(args.columns)}
    patched = patch_workbook(path, headers, columns, args.header_row)
    print(f"patched_cells={patched}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
