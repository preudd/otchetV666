from report_parser import format_report


def main() -> None:
    data = {
        "revenue": 19600,
        "terminal": 18800,
        "cash": 800,
        "receipt_count": 11,
        "category_qty": {
            "Вход безлимит": 8,
            "Билет 1час": 0,
            "Акция счастливые часы": 1,
            "Акция последний час": 0,
            "Комбо на ДР 3 часа": 1,
        },
        "category_sum": {
            "Вход безлимит": 5800,
            "Билет 1час": 0,
            "Акция счастливые часы": 500,
            "Акция последний час": 0,
            "Аквагрим": 0,
            "Виар": 0,
            "Шары": 300,
            "Сопровождающий": 0,
            "Комбо на ДР 3 часа": 7500,
        },
        "return_total": 0,
        "combo_ag": (0, 0, 0),
        "combo_vr": (0, 0, 0),
        "combo_all": (0, 0, 0, 0),
        "advance_dr": 4500,
        "advance_animator": 0,
        "rent": 0,
        "advance_graduation": 0,
    }

    report = format_report(data, "25.03.2025")
    required = [
        "За сегодня:",
        "Выручка -",
        "Терминал -",
        "Наличка приход -",
        "Количество чеков -",
        "Вход безлимит -",
        "Билет 1час",
        "Акция счастливые часы",
        "Акция последний час",
        "Комбо(Билет+Аквагрим):",
        "Комбо(Билет+VR):",
        "Комбо все включено :",
        "Возврат -",
        "Аванс ДР -",
        "Аренда комнаты -",
        "Аванс выпускной -",
        "Комбо на ДР 3 часа -",
        "Касса:",
        "Инкассация:",
        "Остаток в кассе:",
    ]
    missing = [item for item in required if item not in report]
    if missing:
        raise SystemExit(f"Smoke check failed, missing fragments: {missing}")

    print("Smoke check passed.")


if __name__ == "__main__":
    main()
