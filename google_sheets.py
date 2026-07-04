# -*- coding: utf-8 -*-
"""
Подключение к Google Таблице и запись вечернего отчёта.
Таблица и листы месяцев должны уже существовать — бот их не создаёт.
"""
from __future__ import annotations

import base64
import json
import os
import re
from datetime import datetime
from typing import Any

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:  # pragma: no cover
    gspread = None  # type: ignore
    Credentials = None  # type: ignore

SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)

MONTH_SHEET_NAMES = {
    1: "январь",
    2: "февраль",
    3: "март",
    4: "апрель",
    5: "май",
    6: "июнь",
    7: "июль",
    8: "август",
    9: "сентябрь",
    10: "октябрь",
    11: "ноябрь",
    12: "декабрь",
}

# Пока лист не переименован (например, «Лист1» вместо «июль»)
MONTH_SHEET_ALIASES: dict[int, list[str]] = {
    7: ["Лист1", "лист1"],
}

EMPLOYEES: dict[str, dict[str, Any]] = {
    "alina": {
        "label": "Алина",
        "date_col": 23,  # W
        "values_start_col": 24,  # X (Выручка) .. AI (Расчет др); AJ-AL — касса
    },
    "ilya": {
        "label": "Илья",
        "date_col": 40,  # AN
        "values_start_col": 41,  # AO .. AZ; BA-BC — касса
    },
    "kira": {
        "label": "Кира",
        "date_col": 58,  # BF
        "values_start_col": 59,  # BG .. BR; BS-BU — касса
    },
}

# W/X..AI, AN/AO..AZ, BF/BG..BR — 12 полей из Excel, без кассовых колонок
VALUES_COUNT = 12

DEFAULT_SA_FILENAME = "service_account.json"


def _service_account_file_path() -> str:
    explicit = _load_setting("GOOGLE_SERVICE_ACCOUNT_FILE")
    if explicit:
        return explicit
    data_dir = os.environ.get("DATA_DIR", "/app/data").strip() or "/app/data"
    return os.path.join(data_dir, DEFAULT_SA_FILENAME)


def _get_b64_from_env() -> str:
    single = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_B64", "").strip()
    if single:
        return single
    parts: list[str] = []
    for index in range(1, 10):
        part = os.environ.get(f"GOOGLE_SERVICE_ACCOUNT_JSON_B64_{index}", "").strip()
        if not part:
            break
        parts.append(part)
    return "".join(parts)


def bootstrap_service_account_file() -> str | None:
    """Сохраняет JSON из env в файл на диске — надёжнее для хостинга."""
    path = _service_account_file_path()
    if os.path.isfile(path):
        return path

    json_raw = ""
    b64 = _get_b64_from_env()
    if b64:
        json_raw = _decode_b64_env(b64)
    if not json_raw:
        json_raw = _load_setting("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not json_raw:
        return None

    info = _parse_service_account_info(json_raw)
    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(info, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return path


def get_credentials_diagnostics() -> dict[str, str]:
    """Диагностика credentials для логов (без секретов)."""
    path = _service_account_file_path()
    b64 = _get_b64_from_env()
    diag = {
        "file_path": path,
        "file_exists": "yes" if os.path.isfile(path) else "no",
        "b64_len": str(len(re.sub(r"\s+", "", b64.strip().strip('"').strip("'")))) if b64 else "0",
    }
    try:
        json_raw = _load_service_account_json_raw()
        info = _parse_service_account_info(json_raw)
        diag["client_email"] = str(info.get("client_email") or "?")
        diag["private_key_id"] = str(info.get("private_key_id") or "?")
        diag["private_key_len"] = str(len(info.get("private_key") or ""))
    except Exception as exc:
        diag["parse_error"] = str(exc)[:160]
    return diag


def _load_setting(name: str, default: str = "") -> str:
    val = os.environ.get(name, "").strip()
    if val:
        return val
    try:
        import config  # type: ignore

        return str(getattr(config, name, default) or "").strip()
    except ImportError:
        return default


def _normalize_private_key(value: str) -> str:
    """Исправляет private_key после копирования JSON в переменные хостинга."""
    pk = value.strip().strip('"').strip("'")
    pk = pk.replace("\\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    if "BEGIN PRIVATE KEY" not in pk:
        return pk

    lines: list[str] = []
    for line in pk.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("-----"):
            lines.append(line)
        else:
            lines.append(re.sub(r"\s+", "", line))
    pk = "\n".join(lines)
    if not pk.endswith("\n"):
        pk += "\n"
    return pk


def _decode_b64_env(value: str) -> str:
    """Декодирует base64 из переменной окружения (убирает пробелы и кавычки)."""
    b64 = value.strip().strip('"').strip("'")
    b64 = re.sub(r"\s+", "", b64)
    if not b64:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON_B64 пустой.")
    try:
        return base64.b64decode(b64, validate=True).decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        pass
    pad = (-len(b64)) % 4
    try:
        return base64.b64decode(b64 + ("=" * pad)).decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON_B64: невалидный base64. "
            "Сгенерируйте заново: python make_b64.py service_account.json"
        ) from exc


def _parse_service_account_info(json_raw: str) -> dict:
    raw = json_raw.strip()
    if (raw.startswith("'") and raw.endswith("'")) or (raw.startswith('"') and raw.endswith('"') and "{" not in raw):
        raw = raw[1:-1].strip()

    try:
        info = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON содержит невалидный JSON. "
            "На хостинге используйте GOOGLE_SERVICE_ACCOUNT_JSON_B64."
        ) from exc

    pk = info.get("private_key")
    if not isinstance(pk, str) or not pk.strip():
        raise RuntimeError("В JSON сервисного аккаунта отсутствует private_key.")

    info = dict(info)
    info["private_key"] = _normalize_private_key(pk)
    if "BEGIN PRIVATE KEY" not in info["private_key"]:
        raise RuntimeError(
            "private_key повреждён или обрезан. "
            "Удалите GOOGLE_SERVICE_ACCOUNT_JSON и задайте только GOOGLE_SERVICE_ACCOUNT_JSON_B64."
        )
    return info


def _load_service_account_json_raw() -> str:
    """JSON сервисного аккаунта из env (обычный или base64) или config."""
    path = _service_account_file_path()
    if os.path.isfile(path):
        with open(path, encoding="utf-8") as fh:
            return fh.read()

    b64 = _get_b64_from_env()
    if b64:
        return _decode_b64_env(b64)

    return _load_setting("GOOGLE_SERVICE_ACCOUNT_JSON")


def get_config_status() -> dict[str, str]:
    """Статус настроек для логов при старте (без секретов)."""
    spreadsheet_id = _load_setting("GOOGLE_SPREADSHEET_ID")
    json_raw = ""
    if spreadsheet_id:
        try:
            json_raw = _load_service_account_json_raw()
        except RuntimeError:
            json_raw = ""
    file_path = _service_account_file_path()
    file_exists = bool(file_path and os.path.isfile(file_path))

    b64 = _get_b64_from_env()
    if file_exists:
        sa_source = "file"
    elif b64:
        sa_source = "env_b64"
    elif json_raw:
        sa_source = "env_json"
    elif file_path and os.path.isfile(_load_setting("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")):
        sa_source = "file"
    else:
        sa_source = "missing"

    gspread_ok = "yes" if gspread is not None else "no"

    return {
        "spreadsheet_id": "set" if spreadsheet_id else "MISSING",
        "service_account": sa_source,
        "gspread": gspread_ok,
        "configured": "yes" if is_configured() else "no",
    }


def is_configured() -> bool:
    spreadsheet_id = _load_setting("GOOGLE_SPREADSHEET_ID")
    if not spreadsheet_id:
        return False
    try:
        if _load_service_account_json_raw():
            return True
    except RuntimeError:
        return False
    file_path = _service_account_file_path()
    return bool(file_path and os.path.isfile(file_path))


def _load_credentials():
    if Credentials is None or gspread is None:
        raise RuntimeError(
            "Не установлены библиотеки для Google Sheets. Выполните: pip install -r requirements.txt"
        )

    try:
        bootstrap_service_account_file()
    except RuntimeError as exc:
        print("WARNING: failed to bootstrap service account file:", exc)
    file_path = _service_account_file_path()
    if file_path and os.path.isfile(file_path):
        return Credentials.from_service_account_file(file_path, scopes=SCOPES)

    json_raw = _load_service_account_json_raw()
    if json_raw:
        info = _parse_service_account_info(json_raw)
        try:
            return Credentials.from_service_account_info(info, scopes=SCOPES)
        except Exception as exc:
            msg = str(exc)
            if "invalid" in msg.lower() and "key" in msg.lower():
                email = info.get("client_email", "?")
                pk = info.get("private_key", "")
                hint = (
                    f"Invalid private key (аккаунт {email}, длина ключа {len(pk)}). "
                    "Загрузите service_account.json в /app/data/ на хостинге "
                    "или разбейте B64 на GOOGLE_SERVICE_ACCOUNT_JSON_B64_1 и _2."
                )
                raise RuntimeError(hint) from exc
            raise

    file_path = _load_setting("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
    if not file_path or not os.path.isfile(file_path):
        raise RuntimeError(
            "Не задан доступ к Google Sheets. Укажите GOOGLE_SERVICE_ACCOUNT_FILE "
            "или переменную окружения GOOGLE_SERVICE_ACCOUNT_JSON."
        )
    return Credentials.from_service_account_file(file_path, scopes=SCOPES)


def get_service_account_email() -> str | None:
    json_raw = _load_service_account_json_raw()
    if json_raw:
        try:
            return _parse_service_account_info(json_raw).get("client_email")
        except (json.JSONDecodeError, RuntimeError):
            return None

    file_path = _load_setting("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
    if file_path and os.path.isfile(file_path):
        try:
            with open(file_path, encoding="utf-8") as fh:
                return json.load(fh).get("client_email")
        except (OSError, json.JSONDecodeError):
            return None
    return None


def get_client():
    return gspread.authorize(_load_credentials())


def get_spreadsheet():
    spreadsheet_id = _load_setting("GOOGLE_SPREADSHEET_ID")
    if not spreadsheet_id:
        raise RuntimeError("Не задан GOOGLE_SPREADSHEET_ID.")
    return get_client().open_by_key(spreadsheet_id)


def get_employee_label(employee_key: str) -> str:
    emp = EMPLOYEES.get(employee_key)
    if not emp:
        raise ValueError(f"Неизвестный сотрудник: {employee_key}")
    return str(emp["label"])


def spreadsheet_url() -> str | None:
    spreadsheet_id = _load_setting("GOOGLE_SPREADSHEET_ID")
    if not spreadsheet_id:
        return None
    custom = _load_setting("GOOGLE_SPREADSHEET_URL")
    if custom:
        return custom
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"


def _parse_report_date(report_date: str) -> tuple[int, int, int]:
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            dt = datetime.strptime(report_date.strip(), fmt)
            year = dt.year if dt.year >= 2000 else dt.year + 2000
            return dt.day, dt.month, year
        except ValueError:
            continue
    raise ValueError(f"Неверный формат даты отчёта: {report_date}")


def _month_sheet_name(month: int) -> str:
    return MONTH_SHEET_NAMES[month]


def _resolve_month_worksheet(spreadsheet, month: int):
    expected = _month_sheet_name(month)
    worksheets = spreadsheet.worksheets()
    by_title = {ws.title.strip().lower(): ws for ws in worksheets}

    for candidate in (expected, expected.capitalize(), f"{month:02d}", str(month)):
        ws = by_title.get(candidate.lower())
        if ws:
            return ws

    for alias in MONTH_SHEET_ALIASES.get(month, []):
        ws = by_title.get(alias.lower())
        if ws:
            return ws

    # Фоллбек: поддержка названий вида "отчет июль", "июль 2026", и т.п.
    expected_l = expected.lower()
    report_like = [
        ws
        for ws in worksheets
        if expected_l in ws.title.strip().lower() and "отчет" in ws.title.strip().lower()
    ]
    if report_like:
        return report_like[0]

    contains_month = [ws for ws in worksheets if expected_l in ws.title.strip().lower()]
    if contains_month:
        return contains_month[0]

    available = ", ".join(ws.title for ws in worksheets) or "нет"
    raise RuntimeError(
        f"Лист месяца «{expected}» не найден. Доступные листы: {available}"
    )


def get_worksheet_for_date(report_date: str):
    _day, month, _year = _parse_report_date(report_date)
    return _resolve_month_worksheet(get_spreadsheet(), month)


def test_connection() -> dict[str, Any]:
    spreadsheet = get_spreadsheet()
    worksheets = spreadsheet.worksheets()
    preview_sheet = worksheets[0] if worksheets else None

    return {
        "spreadsheet_title": spreadsheet.title,
        "spreadsheet_id": spreadsheet.id,
        "sheet_name": preview_sheet.title if preview_sheet else "—",
        "sheet_count": len(worksheets),
        "sheet_names": [ws.title for ws in worksheets],
        "service_account_email": get_service_account_email(),
        "url": spreadsheet_url(),
        "employees": [emp["label"] for emp in EMPLOYEES.values()],
    }


def _num(value: Any) -> int | float:
    if value is None:
        return 0
    try:
        num = float(value)
    except (TypeError, ValueError):
        return 0
    if abs(num - round(num)) < 0.000001:
        return int(round(num))
    return num


def build_employee_values(data: dict) -> list[int | float]:
    cs = data.get("category_sum") or {}
    return [
        _num(data.get("revenue")),
        _num(data.get("terminal")),
        _num(data.get("cash")),
        _num(data.get("receipt_count")),
        _num(cs.get("Билет 1час")),
        _num(cs.get("Вход безлимит")),
        _num(cs.get("Акция счастливые часы")),
        _num(cs.get("Аквагрим")),
        _num(cs.get("Шары")),
        _num(cs.get("Виар")),
        _num(data.get("advance_dr")),
        _num(cs.get("Комбо на ДР 3 часа")),
    ]


def _col_letter(col: int) -> str:
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _normalize_day_label(value: str) -> str | None:
    text = str(value or "").strip()
    m = re.fullmatch(r"(\d{1,2})[./](\d{1,2})(?:[./]\d{2,4})?", text)
    if not m:
        return None
    return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}"


def _find_day_row(worksheet, day: int, month: int) -> int:
    target = f"{day:02d}.{month:02d}"
    dates_col = worksheet.col_values(1)
    for row_idx, cell in enumerate(dates_col, start=1):
        if row_idx < 3:
            continue
        if _normalize_day_label(cell) == target:
            return row_idx
    raise RuntimeError(f"Строка для даты {target} не найдена на листе «{worksheet.title}».")


def write_report(data: dict, report_date: str, employee_key: str) -> dict[str, Any]:
    if not is_configured():
        raise RuntimeError(
            "Google Таблица не настроена. Задайте GOOGLE_SPREADSHEET_ID и ключ сервисного аккаунта."
        )

    emp = EMPLOYEES.get(employee_key)
    if not emp:
        raise ValueError(f"Неизвестный сотрудник: {employee_key}")

    day, month, year = _parse_report_date(report_date)
    worksheet = _resolve_month_worksheet(get_spreadsheet(), month)
    row = _find_day_row(worksheet, day, month)
    values = build_employee_values(data)

    if len(values) != VALUES_COUNT:
        raise RuntimeError("Внутренняя ошибка: неверное число полей для записи.")

    updates: list[dict[str, Any]] = []

    date_col = int(emp["date_col"])
    start_col = int(emp["values_start_col"])
    # Защитный барьер: не записывать ничего в левую часть таблицы (A-P).
    if date_col <= 16 or start_col <= 16:
        raise RuntimeError(
            f"Неверный маппинг колонок для {emp['label']}: запись в A-P запрещена."
        )

    date_cell = f"{_col_letter(date_col)}{row}"
    updates.append(
        {
            "range": date_cell,
            "values": [[f"{day:02d}.{month:02d}.{year}"]],
        }
    )

    end_col = start_col + VALUES_COUNT - 1
    values_range = f"{_col_letter(start_col)}{row}:{_col_letter(end_col)}{row}"
    updates.append({"range": values_range, "values": [values]})

    worksheet.batch_update(updates, value_input_option="USER_ENTERED")

    return {
        "action": "updated",
        "date": report_date,
        "day_label": f"{day:02d}.{month:02d}",
        "employee": emp["label"],
        "row": row,
        "url": spreadsheet_url(),
        "sheet_name": worksheet.title,
        "spreadsheet_title": get_spreadsheet().title,
        "message": (
            f"Отчёт за {report_date} записан для {emp['label']} "
            f"на лист «{worksheet.title}», строка {row}."
        ),
    }
