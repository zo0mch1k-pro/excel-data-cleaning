"""Очистка таблицы заказов: каждая функция приводит к единому виду одну колонку."""

import re
from datetime import date, datetime
from difflib import get_close_matches

import pandas as pd

from catalog import CITIES, PRODUCTS

CANON_CITIES = [c for c, _ in CITIES]
# сокращения, которые по похожести букв не угадать
CITY_ALIASES = {
    "мск": "Москва",
    "спб": "Санкт-Петербург", "питер": "Санкт-Петербург", "с-петербург": "Санкт-Петербург",
    "екб": "Екатеринбург",
    "н. новгород": "Нижний Новгород", "н.новгород": "Нижний Новгород",
}
DATE_FORMATS = ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%y", "%d.%m.%y", "%d/%m/%Y"]
COLUMN_MAP = {  # исходное название колонки -> внутреннее имя
    "№ заказа": "order_id", "Дата заказа": "date", "Клиент": "name", "Телефон": "phone",
    "Город": "city", "Товар": "product", "Количество": "qty", "Сумма": "amount",
}


def squeeze(value) -> str:
    """Текст без пробелов по краям и без двойных пробелов (в т.ч. неразрывных)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def clean_order_id(value) -> str:
    return squeeze(value).upper()


def clean_date(value) -> date | None:
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()
    if isinstance(value, date):
        return value
    text = squeeze(value)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def clean_name(value) -> str:
    # «иванов  иван» / «ИВАНОВ ИВАН» -> «Иванов Иван»; двойные фамилии через дефис тоже работают
    return squeeze(value).title()


def clean_phone(value) -> str:
    """Любая запись российского мобильного -> +7XXXXXXXXXX. Мусор -> пустая строка."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    digits = re.sub(r"\D", "", squeeze(value))
    if len(digits) == 11 and digits[0] in "78":
        digits = digits[1:]
    if len(digits) == 10 and digits[0] == "9":
        return "+7" + digits
    return ""


def clean_city(value) -> str:
    text = squeeze(value)
    text = re.sub(r"^г\.\s*", "", text, flags=re.IGNORECASE)  # «г. Москва»
    key = text.lower()
    if key in CITY_ALIASES:
        return CITY_ALIASES[key]
    # Ростов на Дону / Санкт Петербург: сравниваем без дефисов и пробелов
    plain = re.sub(r"[\s-]", "", key)
    for city in CANON_CITIES:
        if re.sub(r"[\s-]", "", city.lower()) == plain:
            return city
    # опечатки: ищем самый похожий город из справочника
    match = get_close_matches(key, [c.lower() for c in CANON_CITIES], n=1, cutoff=0.8)
    if match:
        return CANON_CITIES[[c.lower() for c in CANON_CITIES].index(match[0])]
    return text  # неизвестный город оставляем как есть, он попадёт в отчёт


def clean_product(value) -> str:
    key = squeeze(value).lower()
    for product in PRODUCTS:
        if product.lower() == key:
            return product
    return squeeze(value)


def clean_number(value) -> float | None:
    """«1 500,00 ₽», «1500 руб.», «2 шт.» -> число."""
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    text = squeeze(value).replace(" ", "")
    text = re.sub(r"[^\d,.]", "", text).replace(",", ".")
    text = text.rstrip(".")  # точка от «руб.» или «шт.»
    try:
        return float(text)
    except ValueError:
        return None


CLEANERS = {
    "order_id": clean_order_id, "date": clean_date, "name": clean_name, "phone": clean_phone,
    "city": clean_city, "product": clean_product, "qty": clean_number, "amount": clean_number,
}


def _changed(column: str, raw, clean) -> bool:
    """Отличается ли исправленное значение от исходного (для статистики)."""
    if column == "date":
        return not isinstance(raw, (datetime, pd.Timestamp))  # текст стал настоящей датой
    if column in ("qty", "amount"):
        return not isinstance(raw, (int, float))  # текст стал числом
    missing = raw is None or (isinstance(raw, float) and pd.isna(raw))
    return ("" if missing else str(raw)) != str(clean)


def clean_orders(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Возвращает (чистая таблица, статистика по исправлениям)."""
    stats: dict[str, int] = {"rows_raw": len(raw)}
    df = raw.rename(columns=lambda c: COLUMN_MAP.get(squeeze(c), squeeze(c)))

    # 1. пустые строки: всё пусто или одни пробелы
    is_empty = df.apply(lambda row: all(squeeze(v) == "" for v in row), axis=1)
    df = df[~is_empty].reset_index(drop=True)
    stats["empty_removed"] = int(is_empty.sum())

    # 2. точные копии строк (совпадает всё, вплоть до пробела)
    exact = df.astype(str).duplicated()
    stats["exact_duplicates"] = int(exact.sum())

    # 3. приводим каждую колонку к единому виду и считаем, сколько значений поменялось
    clean = pd.DataFrame({col: df[col].map(func) for col, func in CLEANERS.items()})
    changed = pd.DataFrame({col: [_changed(col, r, c) for r, c in zip(df[col], clean[col])]
                            for col in CLEANERS})

    # 4. дубли после нормализации: тот же заказ, записанный по-разному
    all_dups = clean.astype(str).duplicated()
    stats["hidden_duplicates"] = int(all_dups.sum()) - stats["exact_duplicates"]
    keep = ~all_dups
    clean, changed = clean[keep].reset_index(drop=True), changed[keep].reset_index(drop=True)
    for col in CLEANERS:
        stats[f"fixed_{col}"] = int(changed[col].sum())
    stats["rows_fixed"] = int(changed.any(axis=1).sum())
    stats["cells_fixed"] = int(changed.values.sum())

    # 5. проверки качества
    stats["phones_invalid_or_missing"] = int((clean["phone"] == "").sum())
    stats["dates_unparsed"] = int(clean["date"].isna().sum())
    stats["amounts_unparsed"] = int(clean["amount"].isna().sum())
    stats["unknown_cities"] = int((~clean["city"].isin(CANON_CITIES)).sum())
    stats["unknown_products"] = int((~clean["product"].isin(PRODUCTS)).sum())
    stats["order_id_conflicts"] = int(clean["order_id"].duplicated().sum())  # один номер, разные данные
    price = clean["product"].map(lambda p: PRODUCTS.get(p, (None,))[0])
    stats["amount_price_mismatch"] = int((clean["qty"] * price != clean["amount"]).sum())

    clean["qty"] = clean["qty"].astype(int)
    clean = clean.sort_values(["date", "order_id"]).reset_index(drop=True)
    stats["rows_clean"] = len(clean)
    return clean, stats
