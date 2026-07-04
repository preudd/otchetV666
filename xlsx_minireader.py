from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from typing import Iterator, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET


_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def _col_to_index(col_letters: str) -> int:
    col_letters = col_letters.upper()
    idx = 0
    for ch in col_letters:
        if not ("A" <= ch <= "Z"):
            break
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def _split_cell_ref(cell_ref: str) -> Tuple[int, int]:
    m = _CELL_REF_RE.match(cell_ref)
    if not m:
        raise ValueError(f"Bad cell reference: {cell_ref!r}")
    col_letters, row_digits = m.groups()
    return int(row_digits) - 1, _col_to_index(col_letters)


def _read_shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        xml_bytes = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml_bytes)
    out: List[str] = []
    for si in root.findall("main:si", _NS):
        # A shared string can be in <t> or rich text runs <r><t>
        parts: List[str] = []
        t = si.find("main:t", _NS)
        if t is not None and t.text is not None:
            parts.append(t.text)
        else:
            for r in si.findall("main:r", _NS):
                rt = r.find("main:t", _NS)
                if rt is not None and rt.text is not None:
                    parts.append(rt.text)
        out.append("".join(parts))
    return out


def _cell_text(cell: ET.Element, shared: Sequence[str]) -> Optional[str]:
    t = cell.get("t")  # type
    if t == "s":  # shared string index
        v = cell.find("main:v", _NS)
        if v is None or v.text is None:
            return None
        try:
            i = int(v.text)
            return shared[i] if 0 <= i < len(shared) else None
        except ValueError:
            return None
    if t == "inlineStr":
        is_el = cell.find("main:is", _NS)
        if is_el is None:
            return None
        t_el = is_el.find("main:t", _NS)
        return t_el.text if t_el is not None else None

    v = cell.find("main:v", _NS)
    if v is None or v.text is None:
        return None
    return v.text


def iter_sheet_rows(
    xlsx_path: str,
    sheet_xml_path: str = "xl/worksheets/sheet1.xml",
    *,
    max_rows: Optional[int] = None,
) -> Iterator[List[Optional[str]]]:
    """
    Minimal .xlsx reader (no external deps).
    Yields rows as 0-based list of cell values (strings/numbers as text).
    """
    with zipfile.ZipFile(xlsx_path) as z:
        shared = _read_shared_strings(z)
        xml_bytes = z.read(sheet_xml_path)
        root = ET.fromstring(xml_bytes)

        produced = 0
        for row in root.findall(".//main:sheetData/main:row", _NS):
            values: List[Optional[str]] = []
            for cell in row.findall("main:c", _NS):
                ref = cell.get("r")
                if not ref:
                    continue
                _, col_idx = _split_cell_ref(ref)
                if col_idx >= len(values):
                    values.extend([None] * (col_idx + 1 - len(values)))
                values[col_idx] = _cell_text(cell, shared)

            yield values
            produced += 1
            if max_rows is not None and produced >= max_rows:
                return

