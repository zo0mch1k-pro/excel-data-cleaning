"""Проверка очистки по эталону: генератор помнит, какими данные были до «загрязнения».

Запуск:  python test_cleaning.py
"""

import tempfile
from pathlib import Path

import pandas as pd

from cleaner import clean_orders, clean_phone
from generate_raw import BASE_ORDERS, generate, save_raw


def test_against_truth() -> None:
    rows, truth = generate()
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "raw.xlsx"
        save_raw(rows, path)  # проверяем через настоящий xlsx, как в main.py
        raw = pd.read_excel(path, dtype=object)
    clean, stats = clean_orders(raw)

    assert len(clean) == BASE_ORDERS, f"строк {len(clean)}, ожидалось {BASE_ORDERS}"
    expected = {o["id"]: o for o in truth}
    errors: dict[str, list] = {}
    for rec in clean.itertuples(index=False):
        o = expected[rec.order_id]
        phone = "" if o["phone"] in (None, "INVALID") else "+7" + o["phone"]
        checks = {"date": (rec.date, o["date"]), "name": (rec.name, o["name"]), "phone": (rec.phone, phone),
                  "city": (rec.city, o["city"]), "product": (rec.product, o["product"]),
                  "qty": (rec.qty, o["qty"]), "amount": (rec.amount, o["amount"])}
        for field, (got, want) in checks.items():
            if got != want:
                errors.setdefault(field, []).append((rec.order_id, got, want))

    for field, items in errors.items():
        print(f"  {field}: {len(items)} расхождений, например {items[:3]}")
    assert not errors, "очистка восстановила не все значения"
    print(f"OK: все {len(clean)} заказов совпали с эталоном по всем полям")


def test_phone_formats() -> None:
    for raw in ["+7 (912) 345-67-89", "8 912 345 67 89", "89123456789", "+79123456789",
                "7(912)345-67-89", "912-345-67-89", 89123456789, 89123456789.0]:
        assert clean_phone(raw) == "+79123456789", raw
    for raw in ["12-34-56", "нет", "+7 (999) 12", None]:
        assert clean_phone(raw) == "", raw
    print("OK: форматы телефонов")


if __name__ == "__main__":
    test_phone_formats()
    test_against_truth()
