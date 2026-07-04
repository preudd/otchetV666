# -*- coding: utf-8 -*-
"""Разбить B64 на две части для хостинга с лимитом длины переменных."""
from __future__ import annotations

import math
import sys
from pathlib import Path


def main() -> None:
    if len(sys.argv) != 2:
        print("Использование: python split_b64.py b64_for_host.txt")
        raise SystemExit(1)

    b64 = Path(sys.argv[1]).read_text(encoding="utf-8-sig").strip()
    mid = math.ceil(len(b64) / 2)
    part1, part2 = b64[:mid], b64[mid:]
    print(f"total_len={len(b64)} part1={len(part1)} part2={len(part2)}")
    print()
    print("GOOGLE_SERVICE_ACCOUNT_JSON_B64_1")
    print(part1)
    print()
    print("GOOGLE_SERVICE_ACCOUNT_JSON_B64_2")
    print(part2)


if __name__ == "__main__":
    main()
