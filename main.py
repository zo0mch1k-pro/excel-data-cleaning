"""Очистка raw_orders.xlsx -> clean_orders.xlsx + before_after.png.

Запуск:  python main.py
"""

import json
import sys
from pathlib import Path

import pandas as pd

from before_after import make_before_after
from cleaner import clean_orders
from excel_report import BIG_ORDER, save_report
from generate_raw import generate, save_raw

PROJECT_DIR = Path(__file__).resolve().parent
RAW = PROJECT_DIR / "raw_orders.xlsx"
CLEAN = PROJECT_DIR / "clean_orders.xlsx"
PICTURE = PROJECT_DIR / "before_after.png"
STATS = PROJECT_DIR / "cleaning_stats.json"


def print_report(s: dict, big_orders: int) -> None:
    print(f"""
Было строк:                         {s['rows_raw']}
  пустых строк удалено:             {s['empty_removed']}
  точных дублей удалено:            {s['exact_duplicates']}
  скрытых дублей удалено:           {s['hidden_duplicates']}  (тот же заказ, записанный в другом формате)
Стало строк:                        {s['rows_clean']}

Исправлено строк:                   {s['rows_fixed']} из {s['rows_clean']}  (ячеек: {s['cells_fixed']})
  даты (текст -> дата):             {s['fixed_date']}
  телефоны (-> +7XXXXXXXXXX):       {s['fixed_phone']}
  суммы (текст -> число):           {s['fixed_amount']}
  количество (текст -> число):      {s['fixed_qty']}
  имена клиентов:                   {s['fixed_name']}
  города:                           {s['fixed_city']}
  названия товаров:                 {s['fixed_product']}
  номера заказов:                   {s['fixed_order_id']}

Проверки:
  телефон не указан или некорректен: {s['phones_invalid_or_missing']}  (оставлены пустыми)
  нераспознанных дат / сумм:         {s['dates_unparsed']} / {s['amounts_unparsed']}
  неизвестных городов / товаров:     {s['unknown_cities']} / {s['unknown_products']}
  сумма не равна кол-во x цена:      {s['amount_price_mismatch']}
  один номер заказа с разными данными: {s['order_id_conflicts']}
  крупных заказов (от {BIG_ORDER} ₽):     {big_orders}""")


def main() -> int:
    if not RAW.exists():
        print("raw_orders.xlsx не найден, генерирую новый...")
        rows, _ = generate()
        save_raw(rows, RAW)

    print(f"Читаю {RAW.name}")
    raw = pd.read_excel(RAW, dtype=object)  # dtype=object: не даём pandas угадывать типы
    clean, stats = clean_orders(raw)

    try:
        save_report(clean, CLEAN)
        make_before_after(raw, clean, PICTURE)
    except PermissionError as exc:
        print(f"Не могу записать файл: {exc.filename}. Закройте его в Excel и запустите снова.")
        return 1
    STATS.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")

    print_report(stats, int((clean["amount"] >= BIG_ORDER).sum()))
    print(f"\nСохранено: {CLEAN.name}, {PICTURE.name}, {STATS.name}")
    problems = stats["dates_unparsed"] + stats["amounts_unparsed"] + stats["unknown_products"]
    return 0 if problems == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
