"""Генерация вымышленной «грязной» таблицы заказов raw_orders.xlsx.

Все данные выдуманы библиотекой Faker, реальных людей и телефонов здесь нет.
Сид фиксированный, поэтому при каждом запуске получается одна и та же таблица.
"""

import random
from datetime import date, datetime, timedelta
from pathlib import Path

from faker import Faker
from openpyxl import Workbook
from openpyxl.styles import Font

from catalog import CITIES, PRODUCTS

SEED = 42
BASE_ORDERS = 1355      # уникальных заказов
EXACT_DUPLICATES = 70   # точные копии строк
HIDDEN_DUPLICATES = 50  # тот же заказ, но записан в другом формате
EMPTY_ROWS = 25         # пустые и «пустые на вид» строки (одни пробелы)
MALE_NAMES = ("Александр Алексей Андрей Антон Артём Борис Вадим Валерий Василий Виктор Виталий Владимир "
              "Владислав Вячеслав Георгий Глеб Григорий Даниил Денис Дмитрий Евгений Егор Иван Игорь "
              "Илья Кирилл Константин Леонид Максим Михаил Никита Николай Олег Павел Пётр Роман Руслан "
              "Сергей Станислав Степан Тимур Фёдор Юрий Ярослав").split()
FEMALE_NAMES = ("Алина Алёна Анастасия Анна Валентина Валерия Вера Виктория Галина Дарья Диана Евгения "
                "Екатерина Елена Елизавета Зоя Ирина Карина Кристина Ксения Лариса Любовь Людмила "
                "Маргарита Марина Мария Надежда Наталья Нина Оксана Ольга Полина Светлана София Тамара "
                "Татьяна Юлия Яна").split()
# основа отчества: + «ич» для мужчин, + «на» для женщин
PATRONYMIC_BASES = ("Александров Алексеев Андреев Анатольев Борисов Валерьев Васильев Викторов Владимиров "
                    "Геннадьев Григорьев Дмитриев Евгеньев Иванов Игорев Константинов Максимов Михайлов "
                    "Николаев Олегов Павлов Петров Романов Сергеев Юрьев").split()
COLUMNS = ["№ заказа", "Дата заказа", "Клиент", "Телефон", "Город", "Товар", "Количество", "Сумма"]

# сокращения и разговорные названия городов
CITY_ALIASES = {
    "Москва": ["г. Москва", "Мск"],
    "Санкт-Петербург": ["СПб", "Питер", "С-Петербург", "Санкт Петербург", "г. Санкт-Петербург"],
    "Екатеринбург": ["Екб", "г. Екатеринбург"],
    "Нижний Новгород": ["Н. Новгород", "Н.Новгород", "Нижний новгород"],
    "Ростов-на-Дону": ["Ростов на Дону", "Ростов-на-дону", "г. Ростов-на-Дону"],
    "Новосибирск": ["г. Новосибирск"],
}


def make_truth(rng: random.Random, fake: Faker) -> list[dict]:
    """Правильные (эталонные) данные заказов."""
    orders = []
    start = date(2025, 1, 1)
    city_names = [c for c, _ in CITIES]
    city_weights = [w for _, w in CITIES]
    for i in range(BASE_ORDERS):
        # фамилии от Faker, имена и отчества из списка распространённых (у Faker много архаичных)
        father = rng.choice(PATRONYMIC_BASES)
        if rng.random() < 0.5:
            name = f"{fake.last_name_male()} {rng.choice(MALE_NAMES)} {father}ич"
        else:
            name = f"{fake.last_name_female()} {rng.choice(FEMALE_NAMES)} {father}на"
        roll = rng.random()
        if roll < 0.02:
            phone = None                     # телефон не указан
        elif roll < 0.035:
            phone = "INVALID"                # мусор вместо телефона
        else:
            phone = "9" + "".join(str(rng.randint(0, 9)) for _ in range(9))
        product = rng.choice(list(PRODUCTS))
        qty = rng.choices([1, 2, 3, 4, 5], weights=[55, 25, 10, 6, 4])[0]
        orders.append({
            "id": f"ORD-{10001 + i}",
            "date": start + timedelta(days=rng.randint(0, 364)),
            "name": name,
            "phone": phone,
            "city": rng.choices(city_names, weights=city_weights)[0],
            "product": product,
            "qty": qty,
            "amount": qty * PRODUCTS[product][0],
        })
    return orders


def messy_spaces(rng: random.Random, text: str) -> str:
    if rng.random() < 0.5:
        text = text.replace(" ", "  ", 1)
    if rng.random() < 0.5:
        text = " " + text
    if rng.random() < 0.6:
        text = text + "  "
    return text


def messy_case(rng: random.Random, text: str) -> str:
    roll = rng.random()
    if roll < 0.2:
        return text.upper()
    if roll < 0.4:
        return text.lower()
    return text


def typo(rng: random.Random, word: str) -> str:
    """Одна опечатка: переставить, пропустить или удвоить букву (не первую)."""
    i = rng.randint(1, len(word) - 2)
    kind = rng.choice(["swap", "drop", "double"])
    if kind == "swap" and word[i] != word[i + 1]:
        return word[:i] + word[i + 1] + word[i] + word[i + 2:]
    if kind == "drop":
        return word[:i] + word[i + 1:]
    return word[:i] + word[i] + word[i:]


def dirty_row(rng: random.Random, o: dict) -> list:
    # дата: текст в трёх форматах или настоящая дата Excel
    fmt = rng.choices(["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%y", "excel"], weights=[35, 25, 20, 20])[0]
    d = datetime.combine(o["date"], datetime.min.time()) if fmt == "excel" else o["date"].strftime(fmt)

    # телефон
    p = o["phone"]
    if p is None:
        phone = None
    elif p == "INVALID":
        phone = rng.choice(["12-34-56", "нет", "уточнить", "+7 (999) 12"])
    else:
        a, b, c, e = p[:3], p[3:6], p[6:8], p[8:]
        phone = rng.choice([
            f"+7 ({a}) {b}-{c}-{e}", f"8 {a} {b} {c} {e}", f"8{p}", f"+7{p}",
            f"7({a}){b}-{c}-{e}", f"{a}-{b}-{c}-{e}", f"8({a}){b} {c} {e}", int(f"8{p}"),
        ])

    # город
    city = o["city"]
    roll = rng.random()
    if roll < 0.15 and city in CITY_ALIASES:
        city = rng.choice(CITY_ALIASES[city])
    elif roll < 0.35 and len(city) >= 6:
        city = typo(rng, city)
    elif roll < 0.55:
        city = messy_spaces(rng, messy_case(rng, city))

    # сумма с текстом
    amount = o["amount"]
    spaced = f"{amount:,}".replace(",", " ")
    amount_cell = rng.choice([
        amount, f"{spaced} руб.", f"{amount}", f"{spaced},00 ₽", f"{amount}р",
        f"{amount}.00", f"{spaced.replace(' ', chr(160))} р.", f"{spaced} руб",
    ])

    qty = o["qty"]
    qty_cell = rng.choices([qty, f"{qty} шт.", f"{qty}шт", f" {qty} "], weights=[80, 10, 5, 5])[0]

    order_id = o["id"] if rng.random() > 0.05 else rng.choice([f" {o['id']} ", o["id"].lower()])

    return [
        order_id, d,
        messy_spaces(rng, messy_case(rng, o["name"])) if rng.random() < 0.6 else o["name"],
        phone, city,
        messy_spaces(rng, messy_case(rng, o["product"])) if rng.random() < 0.4 else o["product"],
        qty_cell, amount_cell,
    ]


def generate(seed: int = SEED) -> tuple[list[list], list[dict]]:
    """Возвращает (строки грязной таблицы, эталонные заказы)."""
    rng = random.Random(seed)
    fake = Faker("ru_RU")
    fake.seed_instance(seed)

    truth = make_truth(rng, fake)
    rows = [dirty_row(rng, o) for o in truth]

    # дубли вставляем недалеко от оригинала, как это бывает при повторной выгрузке
    for _ in range(EXACT_DUPLICATES):
        i = rng.randrange(len(rows))
        rows.insert(min(i + rng.randint(1, 30), len(rows)), list(rows[i]))
    for _ in range(HIDDEN_DUPLICATES):
        o = rng.choice(truth)
        pos = next(k for k, r in enumerate(rows) if str(r[0]).strip().upper() == o["id"])
        rows.insert(min(pos + rng.randint(1, 30), len(rows)), dirty_row(rng, o))
    for _ in range(EMPTY_ROWS):
        empty = [None] * len(COLUMNS) if rng.random() < 0.5 else [" "] * 3 + [None] * 5
        rows.insert(rng.randrange(len(rows)), empty)
    return rows, truth


def save_raw(rows: list[list], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Заказы"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for letter, width in zip("ABCDEFGH", [13, 13, 38, 20, 20, 30, 12, 16]):
        ws.column_dimensions[letter].width = width
    wb.save(path)


if __name__ == "__main__":
    out = Path(__file__).resolve().parent / "raw_orders.xlsx"
    rows, _ = generate()
    save_raw(rows, out)
    print(f"Сохранено {len(rows)} строк в {out}")
