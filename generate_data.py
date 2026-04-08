from __future__ import annotations

import datetime as dt
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DAYS = ["mon", "tue", "wed", "thu", "fri"]
DAY_LABELS = ["周一", "周二", "周三", "周四", "周五"]


def col_idx(ref: str) -> int:
    letters = re.match(r"([A-Z]+)", ref)
    if not letters:
        raise ValueError(f"Invalid cell ref: {ref}")
    value = 0
    for ch in letters.group(1):
        value = value * 26 + ord(ch) - 64
    return value


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    result = []
    for item in root.findall("a:si", NS):
        text = "".join(node.text or "" for node in item.iterfind(".//a:t", NS))
        result.append(text)
    return result


def read_sheet(archive: zipfile.ZipFile, sheet_name: str, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(archive.read(sheet_name))
    cells: dict[tuple[int, int], str] = {}

    for row in root.findall("a:sheetData/a:row", NS):
        row_idx = int(row.attrib["r"])
        for cell in row.findall("a:c", NS):
            ref = cell.attrib["r"]
            col = col_idx(ref)
            cell_type = cell.attrib.get("t")
            value_node = cell.find("a:v", NS)

            if cell_type == "s" and value_node is not None:
                value = shared_strings[int(value_node.text)]
            elif cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.iterfind(".//a:t", NS))
            elif value_node is not None:
                value = value_node.text or ""
            else:
                value = ""

            cells[(row_idx, col)] = value.strip()

    merge_cells = root.find("a:mergeCells", NS)
    if merge_cells is not None:
        for merge in merge_cells.findall("a:mergeCell", NS):
            start, end = merge.attrib["ref"].split(":")
            start_row = int(re.search(r"\d+", start).group())
            start_col = col_idx(start)
            end_row = int(re.search(r"\d+", end).group())
            end_col = col_idx(end)
            top_value = cells.get((start_row, start_col), "")
            for row_idx in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cells.setdefault((row_idx, col), top_value)

    max_row = max(row for row, _ in cells)
    max_col = max(col for _, col in cells)
    rows = []
    for row_idx in range(1, max_row + 1):
        rows.append([cells.get((row_idx, col), "") for col in range(1, max_col + 1)])
    return rows


def normalize_start(raw: str) -> str:
    if not raw:
        return ""
    try:
        base = dt.datetime(1899, 12, 30)
        return (base + dt.timedelta(days=float(raw))).date().isoformat()
    except ValueError:
        return raw


def parse_workbook(path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = load_shared_strings(archive)
        rows = read_sheet(archive, "xl/worksheets/sheet2.xml", shared_strings)

    records = []
    for row in rows[3:]:
        province, city, district = row[0], row[1], row[2]
        if not province or province.startswith("备注"):
            continue

        schedule = {}
        available_days = []
        dispatch_notes = []
        for key, label, value in zip(DAYS, DAY_LABELS, row[3:8]):
            clean = value.replace(" ", "")
            schedule[key] = clean
            if clean and clean not in {"—", "-"}:
                available_days.append(key)
                if clean != "发车":
                    dispatch_notes.append(f"{label}:{clean}")

        if not available_days:
            continue

        records.append(
            {
                "province": province,
                "city": city,
                "district": district,
                "route": row[9],
                "frequency": row[8],
                "schedule": schedule,
                "availableDays": available_days,
                "dispatchNotes": "；".join(dispatch_notes),
                "remark": row[10],
                "startDate": normalize_start(row[11]),
            }
        )

    return records


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python3 generate_data.py <xlsx-path>")
        return 1

    source = Path(sys.argv[1]).expanduser().resolve()
    data = parse_workbook(source)
    print(json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
