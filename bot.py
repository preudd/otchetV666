# -*- coding: utf-8 -*-
"""
Телеграм-бот для формирования вечернего отчёта из Excel файла с экспортом чеков.
Отправьте боту файл .xlsx — он обработает данные и запишет их в Google Таблицу.
"""
import asyncio
import os
import sys
import tempfile
import re
from datetime import datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from report_parser import parse_excel_report
from rules_manager import UI_CATEGORIES, add_keyword, load_rules, remove_keyword
from google_sheets import (
    EMPLOYEES,
    get_config_status,
    get_credentials_diagnostics,
    get_employee_label,
    is_configured,
    test_connection,
    write_report,
    get_service_account_email,
)

STATE_PICK_CATEGORY, STATE_WAIT_KEYWORD, STATE_PICK_DELETE_CATEGORY, STATE_PICK_DELETE_KEYWORD = range(4)
APP_BUILD = "otchetV666-v2"


async def _on_startup(application: Application) -> None:
    me = await application.bot.get_me()
    print(f"Telegram bot online: @{me.username} (id={me.id})")

EMPLOYEE_KEYBOARD = InlineKeyboardMarkup(
    [
        [InlineKeyboardButton("Алина", callback_data="emp:alina")],
        [InlineKeyboardButton("Илья", callback_data="emp:ilya")],
        [InlineKeyboardButton("Кира", callback_data="emp:kira")],
    ]
)


def _load_bot_token() -> str | None:
    """
    Токен (по приоритету):
    1) Переменные окружения — удобно на хостинге (config.py в Git не кладём).
       Имена: REPORT_BOT_TOKEN, BOT_TOKEN, TELEGRAM_BOT_TOKEN
    2) Локальный файл config.py (копия из config.example.py)
    """
    for env_name in ('REPORT_BOT_TOKEN', 'BOT_TOKEN', 'TELEGRAM_BOT_TOKEN'):
        val = os.environ.get(env_name, '').strip()
        if val:
            return val
    try:
        from config import BOT_TOKEN as cfg_token
        t = (cfg_token or '').strip()
        if t and t != 'YOUR_BOT_TOKEN_HERE':
            return t
    except ImportError:
        pass
    return None


BOT_TOKEN = _load_bot_token() or 'YOUR_BOT_TOKEN_HERE'


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик команды /start"""
    keyboard = [
        [InlineKeyboardButton("Кому делаем отчёт?", callback_data="ui:pick_employee")],
        [InlineKeyboardButton("Добавить правило", callback_data="ui:add_rule")],
        [InlineKeyboardButton("Удалить правило", callback_data="ui:delete_rule")],
        [InlineKeyboardButton("Показать правила", callback_data="ui:show_rules")],
    ]
    selected = context.user_data.get("employee")
    selected_line = ""
    if selected and selected in EMPLOYEES:
        selected_line = f"\n\n👤 Сейчас выбрано: {get_employee_label(selected)}"
    await update.message.reply_text(
        "Привет! Я бот для вечернего отчёта.\n\n"
        "📤 Отправь мне файл Excel (.xlsx) с экспортом чеков — я обработаю его "
        "и запишу данные в Google Таблицу.\n\n"
        "👤 Перед отправкой файла можно выбрать сотрудника кнопкой "
        "«Кому делаем отчёт?» — или выбрать после загрузки файла.\n\n"
        "📊 Команда /sheets — проверить подключение к таблице.\n\n"
        "Дата в отчёте берётся из названия файла (например, "
        "'Экспорт чеков от 17-01-2026.xlsx' → 17.01.2026), "
        "или сегодняшняя, если дату не удастся определить."
        f"{selected_line}",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


def _format_money(value: float | int) -> str:
    return f"{float(value):,.0f}".replace(",", " ")


async def _send_report_result(
    update: Update,
    *,
    data: dict,
    report_date: str,
    result: dict,
) -> None:
    message = update.callback_query.message if update.callback_query else update.message
    if not message:
        return
    url = result.get("url") or "—"
    returns = _format_money(data.get("return_total", 0))
    revenue = _format_money(data.get("revenue", 0))
    cash = _format_money(data.get("cash", 0))
    copy_block = (
        f"Дата {report_date}\n\n"
        f"За сегодня:\n"
        f"Выручка - {revenue} руб.\n"
        f"Наличка приход - {cash} руб.\n"
        f"Возвраты - {returns} руб.\n"
        f"Аквагрим ДР - (внести вручную)"
    )
    await message.reply_text(
        f"✅ Отчёт записан для: {result.get('employee')}\n\n"
        f"<pre>{copy_block}</pre>\n\n"
        f"Лист: {result.get('sheet_name')}, строка {result.get('row')}\n"
        f"Ссылка: {url}",
        parse_mode="HTML",
    )


async def _process_pending_report(update: Update, context: ContextTypes.DEFAULT_TYPE, employee_key: str) -> None:
    pending = context.user_data.pop("pending_report", None)
    if not pending:
        label = get_employee_label(employee_key)
        text = f"✅ Отчёты будут записываться для: {label}"
        if update.callback_query:
            await update.callback_query.edit_message_text(text)
        elif update.message:
            await update.message.reply_text(text)
        return

    data = pending["data"]
    report_date = pending["report_date"]
    result = write_report(data, report_date, employee_key)
    await _send_report_result(update, data=data, report_date=report_date, result=result)


async def ui_pick_employee(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return
    await query.answer()
    await query.edit_message_text(
        "Кому делаем отчёт?",
        reply_markup=EMPLOYEE_KEYBOARD,
    )


async def ui_select_employee(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return
    await query.answer()
    employee_key = (query.data or "").split(":", 1)[1] if query.data else ""
    if employee_key not in EMPLOYEES:
        await query.edit_message_text("Не удалось выбрать сотрудника. Попробуй ещё раз /start.")
        return

    context.user_data["employee"] = employee_key
    if context.user_data.get("pending_report"):
        await _process_pending_report(update, context, employee_key)
    else:
        await query.edit_message_text(
            f"✅ Отчёты будут записываться для: {get_employee_label(employee_key)}"
        )


async def ui_add_rule(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Показывает список категорий, куда можно добавить ключевое слово."""
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    keyboard: list[list[InlineKeyboardButton]] = []
    for i, cat in enumerate(UI_CATEGORIES):
        # 2 колонки
        if i % 2 == 0:
            keyboard.append([])
        keyboard[-1].append(
            InlineKeyboardButton(cat["label"], callback_data=f"cat:{cat['rule_key']}")
        )
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])

    await query.edit_message_text(
        "Выбери категорию. Потом пришли ключевое слово/фрагмент, который встречается в названии позиции.",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_CATEGORY


async def ui_choose_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    data = query.data or ""
    # cat:<rule_key>
    rule_key = data.split(":", 1)[1] if ":" in data else ""
    if not rule_key:
        await query.edit_message_text("Не удалось выбрать категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END

    context.user_data["rule_key"] = rule_key
    await query.edit_message_text(
        "Пришли ключевое слово/фрагмент.\nНапример: `аванс выпускной` или `LION`",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Отмена", callback_data="cancel")]]),
    )
    return STATE_WAIT_KEYWORD


async def ui_receive_keyword(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    keyword = (update.message.text or "").strip() if update.message else ""
    rule_key = context.user_data.get("rule_key")
    if not rule_key:
        if update.message:
            await update.message.reply_text("Не удалось определить категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END
    if not keyword:
        if update.message:
            await update.message.reply_text("Ключевое слово пустое. Попробуй ещё раз.")
        return STATE_WAIT_KEYWORD

    try:
        add_keyword(rule_key, keyword)
    except Exception as e:
        if update.message:
            await update.message.reply_text(f"Ошибка при добавлении правила: {e}")
        return STATE_WAIT_KEYWORD

    label = next((c["label"] for c in UI_CATEGORIES if c["rule_key"] == rule_key), rule_key)
    await update.message.reply_text(
        f"✅ Добавлено!\nКатегория: {label}\nКлючевое слово: {keyword}"
    )
    context.user_data.pop("rule_key", None)
    return ConversationHandler.END


async def ui_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        try:
            await query.edit_message_text("Отменено.")
        except Exception:
            # иногда сообщение уже не редактируется (если, например, уже ответили)
            pass
    context.user_data.pop("rule_key", None)
    return ConversationHandler.END


async def ui_show_rules(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    rules = load_rules()
    lines = ["Текущие правила:"]
    for cat in UI_CATEGORIES:
        rk = cat["rule_key"]
        r = rules.get(rk)
        if not r:
            continue
        lines.append(f"- {cat['label']}: " + ", ".join(r.keywords))

    await query.edit_message_text("\n".join(lines))
    return ConversationHandler.END


async def ui_delete_rule(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    keyboard: list[list[InlineKeyboardButton]] = []
    for i, cat in enumerate(UI_CATEGORIES):
        if i % 2 == 0:
            keyboard.append([])
        keyboard[-1].append(
            InlineKeyboardButton(cat["label"], callback_data=f"delcat:{cat['rule_key']}")
        )
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])
    await query.edit_message_text(
        "Выбери категорию, из которой нужно удалить ключевое слово.",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_DELETE_CATEGORY


async def ui_choose_delete_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()
    data = query.data or ""
    rule_key = data.split(":", 1)[1] if ":" in data else ""
    if not rule_key:
        await query.edit_message_text("Не удалось выбрать категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END

    rules = load_rules()
    rule = rules.get(rule_key)
    if not rule:
        await query.edit_message_text("Категория не найдена.")
        return ConversationHandler.END

    context.user_data["delete_rule_key"] = rule_key
    keyboard = [[InlineKeyboardButton(kw, callback_data=f"delkw:{kw}")] for kw in rule.keywords]
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])
    await query.edit_message_text(
        "Выбери ключевое слово для удаления:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_DELETE_KEYWORD


async def ui_delete_keyword(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()
    data = query.data or ""
    keyword = data.split(":", 1)[1] if ":" in data else ""
    rule_key = context.user_data.get("delete_rule_key")
    if not rule_key or not keyword:
        await query.edit_message_text("Не удалось удалить ключевое слово. Повтори через /start.")
        return ConversationHandler.END
    try:
        remove_keyword(rule_key, keyword)
    except Exception as e:
        await query.edit_message_text(f"Ошибка удаления: {e}")
        return ConversationHandler.END

    label = next((c["label"] for c in UI_CATEGORIES if c["rule_key"] == rule_key), rule_key)
    await query.edit_message_text(f"✅ Удалено из '{label}': {keyword}")
    context.user_data.pop("delete_rule_key", None)
    return ConversationHandler.END


def extract_date_from_filename(filename: str) -> str | None:
    """Извлекает дату из названия файла. Пример: 'Экспорт чеков от 17-01-2026.xlsx' -> '17.01.2026'"""
    # Паттерны: 17-01-2026, 17.01.2026, 2026-01-17
    for pattern in [
        r'(\d{2})[\-\.](\d{2})[\-\.](\d{4})',  # 17-01-2026 или 17.01.2026
        r'(\d{4})[\-\.](\d{2})[\-\.](\d{2})',  # 2026-01-17
    ]:
        m = re.search(pattern, filename)
        if m:
            g = m.groups()
            if len(g[0]) == 4:  # год первым
                return f"{g[2]}.{g[1]}.{g[0]}"
            return f"{g[0]}.{g[1]}.{g[2]}"
    return None


async def cmd_sheets(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Проверка подключения к Google Таблице."""
    if not is_configured():
        email_hint = ""
        sa_email = get_service_account_email()
        if sa_email:
            email_hint = f"\n\nСервисный аккаунт: `{sa_email}`"
        await update.message.reply_text(
            "⚠️ Google Таблица не настроена.\n\n"
            "Нужно указать в `config.py` или переменных окружения:\n"
            "• `GOOGLE_SPREADSHEET_ID`\n"
            "• `GOOGLE_SERVICE_ACCOUNT_FILE` или `GOOGLE_SERVICE_ACCOUNT_JSON`\n\n"
            "Листы месяцев (`январь`, `февраль`, …) должны уже существовать в таблице.\n"
            "Таблицу нужно заранее расшарить на email сервисного аккаунта "
            "(роль «Редактор»)."
            f"{email_hint}",
            parse_mode="Markdown",
        )
        return

    try:
        info = test_connection()
        sheets_list = ", ".join(info.get("sheet_names") or [])
        sa_email = info.get("service_account_email") or "—"
        url = info.get("url") or "—"
        await update.message.reply_text(
            "✅ Подключение к Google Таблице работает.\n\n"
            f"Таблица: {info.get('spreadsheet_title')}\n"
            f"Листы: {sheets_list}\n"
            f"Сотрудники: {', '.join(info.get('employees') or [])}\n"
            f"Сервисный аккаунт: {sa_email}\n"
            f"Ссылка: {url}"
        )
    except Exception as e:
        await update.message.reply_text(
            f"❌ Не удалось подключиться к Google Таблице:\n{e}\n\n"
            "Проверьте ID таблицы, имя листа и доступ сервисного аккаунта."
        )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик загрузки документа (Excel файла)"""
    document = update.message.document
    filename = document.file_name or ""

    if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
        await update.message.reply_text(
            "⚠️ Пожалуйста, отправьте файл Excel (.xlsx или .xls)."
        )
        return

    await update.message.reply_text("⏳ Обрабатываю файл...")

    try:
        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            await file.download_to_drive(tmp.name)
            tmp_path = tmp.name

        try:
            data = parse_excel_report(tmp_path)
            report_date = extract_date_from_filename(filename)
            if not report_date:
                report_date = datetime.now().strftime('%d.%m.%Y')

            if not is_configured():
                await update.message.reply_text(
                    f"✅ Файл за {report_date} обработан.\n\n"
                    f"Выручка: {_format_money(data['revenue'])} руб.\n"
                    f"Чеков: {data['receipt_count']}\n\n"
                    "⚠️ Google Таблица не настроена — данные пока никуда не записаны.\n"
                    "Настройте подключение и проверьте командой /sheets."
                )
                return

            employee_key = context.user_data.get("employee")
            if employee_key in EMPLOYEES:
                result = write_report(data, report_date, employee_key)
                await _send_report_result(
                    update,
                    data=data,
                    report_date=report_date,
                    result=result,
                )
                return

            context.user_data["pending_report"] = {
                "data": data,
                "report_date": report_date,
            }
            await update.message.reply_text(
                f"✅ Файл за {report_date} обработан.\n\n"
                f"Выручка: {_format_money(data['revenue'])} руб.\n"
                f"Чеков: {data['receipt_count']}\n\n"
                "Кому делаем отчёт?",
                reply_markup=EMPLOYEE_KEYBOARD,
            )
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    except Exception as e:
        await update.message.reply_text(
            f"❌ Ошибка при обработке файла:\n{str(e)}\n\n"
            "Проверьте структуру Excel (столбцы I/O) и настройки Google Таблицы "
            "(доступ сервисного аккаунта, имя листа месяца)."
        )
        raise


def main() -> None:
    """Запуск бота"""
    # Создаём event loop для Python 3.10+ (иначе RuntimeError на MainThread)
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    if BOT_TOKEN == 'YOUR_BOT_TOKEN_HERE':
        print("=" * 50)
        print("ОШИБКА: не задан токен бота.")
        print("На хостинге: Environment Variables / Secrets, добавьте одну из:")
        print("  REPORT_BOT_TOKEN=<токен от @BotFather>")
        print("  или BOT_TOKEN=<токен>")
        print("  или TELEGRAM_BOT_TOKEN=<токен>")
        print("Локально: скопируйте config.example.py -> config.py и укажите BOT_TOKEN.")
        print("=" * 50)
        sys.exit(1)

    app = Application.builder().token(BOT_TOKEN).post_init(_on_startup).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("sheets", cmd_sheets))
    app.add_handler(CallbackQueryHandler(ui_pick_employee, pattern=r"^ui:pick_employee$"))
    app.add_handler(CallbackQueryHandler(ui_select_employee, pattern=r"^emp:"))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(ui_add_rule, pattern=r"^ui:add_rule$"),
            CallbackQueryHandler(ui_delete_rule, pattern=r"^ui:delete_rule$"),
            CallbackQueryHandler(ui_show_rules, pattern=r"^ui:show_rules$"),
        ],
        states={
            STATE_PICK_CATEGORY: [CallbackQueryHandler(ui_choose_category, pattern=r"^cat:")],
            STATE_WAIT_KEYWORD: [MessageHandler(filters.TEXT & ~filters.COMMAND, ui_receive_keyword)],
            STATE_PICK_DELETE_CATEGORY: [
                CallbackQueryHandler(ui_choose_delete_category, pattern=r"^delcat:")
            ],
            STATE_PICK_DELETE_KEYWORD: [
                CallbackQueryHandler(ui_delete_keyword, pattern=r"^delkw:")
            ],
        },
        fallbacks=[
            CallbackQueryHandler(ui_cancel, pattern=r"^cancel$"),
            CommandHandler("start", start),
        ],
        allow_reentry=True,
    )
    app.add_handler(conv)

    try:
        rules = load_rules()
        print(f"BUILD: {APP_BUILD}")
        print("RULES: action_happy_hours =", rules["action_happy_hours"].keywords)
        print("RULES: action_last_hour  =", rules["action_last_hour"].keywords)
        print("RULES: advance_dr        =", rules["advance_dr"].keywords)
    except Exception as e:
        print("WARNING: failed to load rules at startup:", e)

    token_status = "set" if BOT_TOKEN and BOT_TOKEN != "YOUR_BOT_TOKEN_HERE" else "MISSING"
    print(f"BOT_TOKEN: {token_status}")

    cfg = get_config_status()
    print(
        "Google Sheets config:",
        f"spreadsheet_id={cfg['spreadsheet_id']},",
        f"service_account={cfg['service_account']},",
        f"gspread={cfg['gspread']},",
        f"configured={cfg['configured']}",
    )
    diag = get_credentials_diagnostics()
    print(
        "Google Sheets credentials:",
        f"file={diag.get('file_path')},",
        f"exists={diag.get('file_exists')},",
        f"b64_len={diag.get('b64_len')},",
        f"email={diag.get('client_email', '?')},",
        f"key_id={diag.get('private_key_id', '?')},",
        f"key_len={diag.get('private_key_len', '?')}",
    )
    if diag.get("parse_error"):
        print("Google Sheets parse error:", diag["parse_error"])

    if is_configured():
        try:
            info = test_connection()
            print(f"Google Sheets: OK — {info.get('spreadsheet_title')} / {', '.join(info.get('sheet_names') or [])}")
        except Exception as e:
            print("Google Sheets: connection failed:", e)
    else:
        print(
            "Google Sheets: not configured. Set GOOGLE_SPREADSHEET_ID and "
            "GOOGLE_SERVICE_ACCOUNT_JSON (or GOOGLE_SERVICE_ACCOUNT_JSON_B64)."
        )

    print("Бот запущен. Отправьте файл Excel — данные будут записаны в Google Таблицу.")
    # Не выходим при кратковременных сетевых сбоях Telegram API.
    app.run_polling(
        allowed_updates=Update.ALL_TYPES,
        bootstrap_retries=-1,
        drop_pending_updates=True,
    )


if __name__ == "__main__":
    main()
