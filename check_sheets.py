# -*- coding: utf-8 -*-
"""Проверка подключения и маппинга Google Таблицы."""
from google_sheets import (
    EMPLOYEES,
    build_employee_values,
    is_configured,
    test_connection,
    _parse_report_date,
    _month_sheet_name,
    _col_letter,
)


def main() -> None:
    day, month, year = _parse_report_date("17.03.2026")
    assert day == 17 and month == 3 and year == 2026
    assert _month_sheet_name(7) == "июль"
    assert len(build_employee_values({"revenue": 100, "category_sum": {}})) == 12
    assert set(EMPLOYEES) == {"alina", "ilya", "kira"}

    assert EMPLOYEES["alina"]["date_col"] == 23 and _col_letter(23) == "W"
    assert _col_letter(38) == "AL"
    assert EMPLOYEES["ilya"]["date_col"] == 40 and _col_letter(40) == "AN"
    assert _col_letter(55) == "BC"
    assert EMPLOYEES["kira"]["date_col"] == 58 and _col_letter(58) == "BF"
    assert _col_letter(73) == "BU"

    if not is_configured():
        print("Mapping check passed (Google Sheets not configured locally).")
        return

    info = test_connection()
    print("Подключение успешно.")
    print(f"Таблица: {info['spreadsheet_title']}")
    print(f"Листы: {', '.join(info['sheet_names'])}")
    print(f"Сотрудники: {', '.join(info['employees'])}")
    if info.get("service_account_email"):
        print(f"Сервисный аккаунт: {info['service_account_email']}")
    if info.get("url"):
        print(f"URL: {info['url']}")


if __name__ == "__main__":
    main()
