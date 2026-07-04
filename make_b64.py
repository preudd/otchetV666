# -*- coding: utf-8 -*-
"""Сгенерировать GOOGLE_SERVICE_ACCOUNT_JSON_B64 для хостинга."""
from __future__ import annotations

import base64
import json
import sys

from google.oauth2.service_account import Credentials

from google_sheets import SCOPES, _parse_service_account_info


def main() -> None:
    if len(sys.argv) != 2:
        print("Использование: python make_b64.py service_account.json")
        raise SystemExit(1)

    path = sys.argv[1]
    raw = open(path, encoding="utf-8").read()
    info = _parse_service_account_info(raw)
    Credentials.from_service_account_info(info, scopes=SCOPES)

    b64 = base64.b64encode(raw.encode("utf-8")).decode("ascii")
    print("OK: ключ валидный")
    print(f"client_email: {info.get('client_email')}")
    print(f"project_id: {info.get('project_id')}")
    print()
    print("Скопируйте ВСЮ строку ниже в GOOGLE_SERVICE_ACCOUNT_JSON_B64 (одной строкой):")
    print(b64)


if __name__ == "__main__":
    main()
