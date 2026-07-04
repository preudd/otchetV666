
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"

# --- Google Таблица (уже существующая) ---
# ID из URL: https://docs.google.com/spreadsheets/d/<GOOGLE_SPREADSHEET_ID>/edit
GOOGLE_SPREADSHEET_ID = "your_spreadsheet_id_here"

# Локально: путь к JSON-ключу сервисного аккаунта Google Cloud
GOOGLE_SERVICE_ACCOUNT_FILE = "service_account.json"

# На хостинге (Bothost): GOOGLE_SERVICE_ACCOUNT_JSON_B64_1 + _2
# Сгенерировать: python make_b64.py service_account.json && python split_b64.py b64_for_host.txt

# Необязательно: прямая ссылка на таблицу для сообщений бота
# GOOGLE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"

# Листы месяцев: январь, февраль, ... или отчет июль, отчет август
