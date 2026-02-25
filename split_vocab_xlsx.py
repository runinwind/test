#!/usr/bin/env python3
"""Split vocabulary entries in column B of an .xlsx into one word per row.

Output is a UTF-8 CSV with 3 columns:
1) Original column A value
2) Word
3) Remaining text (part of speech + meaning + usage)
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
WORD_RE = re.compile(r"^[A-Za-z][A-Za-z0-9'’._-]*$")
POS_RE = re.compile(
    r"^(?:"
    r"n|v|vt|vi|adj|adv|prep|pron|conj|int|num|art|det|aux|modal|abbr|"
    r"phr|phrase|idiom|interj|pl"
    r")(?:\.|/|\b)",
    re.IGNORECASE,
)


@dataclass
class Entry:
    word: str
    detail: str


def col_letters(cell_ref: str) -> str:
    letters = []
    for ch in cell_ref:
        if ch.isalpha():
            letters.append(ch)
        else:
            break
    return "".join(letters)


def read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    name = "xl/sharedStrings.xml"
    if name not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(name))
    values: List[str] = []
    for si in root.findall(f"{NS_MAIN}si"):
        text_parts: List[str] = []
        for t in si.iter(f"{NS_MAIN}t"):
            text_parts.append(t.text or "")
        values.append("".join(text_parts))
    return values


def first_worksheet_path(zf: zipfile.ZipFile) -> str:
    candidates = sorted(n for n in zf.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    if not candidates:
        raise ValueError("No worksheet XML found in xlsx")
    return candidates[0]


def cell_text(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        is_node = cell.find(f"{NS_MAIN}is")
        if is_node is None:
            return ""
        return "".join((t.text or "") for t in is_node.iter(f"{NS_MAIN}t"))

    v = cell.find(f"{NS_MAIN}v")
    if v is None or v.text is None:
        return ""

    if cell_type == "s":
        idx = int(v.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""

    return v.text


def read_a_b_columns(xlsx_path: Path) -> List[Tuple[str, str]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared = read_shared_strings(zf)
        ws = ET.fromstring(zf.read(first_worksheet_path(zf)))

    rows: List[Tuple[str, str]] = []
    sheet_data = ws.find(f"{NS_MAIN}sheetData")
    if sheet_data is None:
        return rows

    for row in sheet_data.findall(f"{NS_MAIN}row"):
        a_val = ""
        b_val = ""
        for c in row.findall(f"{NS_MAIN}c"):
            ref = c.attrib.get("r", "")
            col = col_letters(ref)
            text = cell_text(c, shared)
            if col == "A":
                a_val = text
            elif col == "B":
                b_val = text
        rows.append((a_val, b_val))

    return rows


def is_probable_new_entry_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False

    if "\t" in stripped:
        first, _ = stripped.split("\t", 1)
        return bool(WORD_RE.match(first))

    parts = stripped.split(maxsplit=1)
    if not parts:
        return False
    word = parts[0]
    if not WORD_RE.match(word):
        return False

    if len(parts) == 1:
        return True

    remainder = parts[1].lstrip()
    if not remainder:
        return True
    if POS_RE.match(remainder):
        return True
    if remainder[0] in "[/(【［":
        return True

    return False


def split_entries(cell_text_value: str) -> List[Entry]:
    lines = [line.rstrip() for line in cell_text_value.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    entries: List[Entry] = []
    current_word: Optional[str] = None
    current_detail_lines: List[str] = []

    def flush() -> None:
        nonlocal current_word, current_detail_lines
        if current_word is not None:
            entries.append(Entry(current_word, "\n".join(current_detail_lines).strip()))
        current_word = None
        current_detail_lines = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            if current_word is not None and current_detail_lines:
                current_detail_lines.append("")
            continue

        if is_probable_new_entry_line(line):
            flush()
            if "\t" in line:
                word, detail = line.split("\t", 1)
                current_word = word.strip()
                current_detail_lines = [detail.strip()]
            else:
                parts = line.split(maxsplit=1)
                current_word = parts[0].strip()
                current_detail_lines = [parts[1].strip()] if len(parts) > 1 else [""]
        else:
            if current_word is None:
                current_word = line
                current_detail_lines = [""]
            else:
                current_detail_lines.append(line)

    flush()
    return entries


def transform_rows(rows: Sequence[Tuple[str, str]]) -> List[Tuple[str, str, str]]:
    output: List[Tuple[str, str, str]] = []
    for col1, col2 in rows:
        if not col2.strip():
            continue
        entries = split_entries(col2)
        for e in entries:
            output.append((col1, e.word, e.detail))
    return output


def run_self_test() -> None:
    sample = """abandon vt. 放弃\n常见搭配：abandon doing\nability n. 能力\nable adj. 能够的\n用法：be able to"""
    parsed = split_entries(sample)
    assert len(parsed) == 3, parsed
    assert parsed[0].word == "abandon"
    assert "常见搭配" in parsed[0].detail
    assert parsed[2].word == "able"
    assert "be able to" in parsed[2].detail
    print("self-test passed")


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Split Excel column B vocabulary cells into one word per row")
    parser.add_argument("input", nargs="?", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("单词_拆分结果.csv"), help="Output CSV path")
    parser.add_argument("--self-test", action="store_true", help="Run internal parser tests")
    args = parser.parse_args(argv)

    if args.self_test:
        run_self_test()
        return 0

    input_path = args.input
    if input_path is None:
        candidates = sorted(Path.cwd().glob("*.xlsx"))
        if len(candidates) == 1:
            input_path = candidates[0]
            print(f"Auto-detected input file: {input_path}")
        elif not candidates:
            parser.error("No .xlsx found in current directory. Please place your file here or pass the input path.")
        else:
            names = ", ".join(str(c) for c in candidates)
            parser.error(f"Multiple .xlsx files found: {names}. Please pass the input path explicitly.")

    rows = read_a_b_columns(input_path)
    out_rows = transform_rows(rows)

    with args.output.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["第1列", "单词", "释义"])
        writer.writerows(out_rows)

    print(f"Done. Wrote {len(out_rows)} rows to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
