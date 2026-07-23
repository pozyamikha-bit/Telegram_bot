"""
Диагностический скрипт: сравнивает реальные заголовки (первую строку)
каждого листа вашей Google Таблицы с тем, что ожидает код бота.

Запускать из той же папки, что и bot.py, с тем же окружением/venv и
тем же заполненным .env (или переменными окружения, если проверяете
на хостинге) — скрипт использует те же config.py и google_sheets.py.

Запуск:
    python check_headers.py
"""

import config
import google_sheets

CHECKS = [
    (config.GOOGLE_SHEET_WORKSHEET_REG, google_sheets.REGISTRATION_HEADER),
    (config.GOOGLE_SHEET_WORKSHEET_RECEIPTS, google_sheets.RECEIPTS_HEADER),
    (config.GOOGLE_SHEET_WORKSHEET_MODERATORS, google_sheets.MODERATORS_HEADER),
    (config.GOOGLE_SHEET_WORKSHEET_TEXTS, google_sheets.TEXTS_HEADER),
]


def main():
    client = google_sheets._get_client()
    spreadsheet = client.open_by_key(config.GOOGLE_SHEET_ID)

    any_problem = False

    for sheet_name, expected in CHECKS:
        print(f"\n=== Лист «{sheet_name}» ===")
        try:
            ws = spreadsheet.worksheet(sheet_name)
        except Exception as exc:
            print(f"  Лист не найден или недоступен: {exc}")
            any_problem = True
            continue

        actual = ws.row_values(1)
        print(f"  Ожидается: {expected}")
        print(f"  Сейчас:    {actual}")

        if actual == expected:
            print("  ОК — совпадает полностью.")
            continue

        any_problem = True

        blanks = [i + 1 for i, v in enumerate(actual) if v.strip() == ""]
        if blanks:
            print(f"  ПРОБЛЕМА: пустые ячейки в колонках {blanks} — это ломает чтение таблицы (ошибка gspread про 'not unique').")

        if len(actual) != len(expected):
            print(f"  ПРОБЛЕМА: разное количество колонок (сейчас {len(actual)}, должно быть {len(expected)}).")

        for i, exp_val in enumerate(expected):
            act_val = actual[i] if i < len(actual) else "<колонки нет вообще>"
            if act_val != exp_val:
                print(f"  Колонка {i + 1}: сейчас «{act_val}», должно быть «{exp_val}»")

    print("\n" + ("Найдены расхождения — см. выше, поправьте вручную в Google Таблице." if any_problem else "Все листы в порядке, заголовки совпадают!"))


if __name__ == "__main__":
    main()
