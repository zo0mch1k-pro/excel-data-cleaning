"""Картинка before_after.png: фрагмент грязной и чистой таблицы рядом."""

import re
from datetime import date, datetime
from pathlib import Path
from textwrap import shorten

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from cleaner import clean_order_id, squeeze

BAD = "#FBDADA"     # ошибка в исходных данных
GOOD = "#DDF2E3"    # исправленное значение
EMPTY = "#EDEDED"
HEADER = "#1F4E78"
COLS = ["order_id", "date", "name", "phone", "city", "amount"]
TITLES = ["№ заказа", "Дата", "Клиент", "Телефон", "Город", "Сумма"]


def _raw_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d %H:%M:%S")  # так Excel показывает дату без формата
    return str(value).replace(" ", " ")


def _show_spaces(text: str) -> str:
    """Лишние пробелы (по краям и двойные) делаем видимыми: «·»."""
    return re.sub(r"^ +| +$| {2,}", lambda m: "·" * len(m.group()), text)


def _clean_text(col: str, value) -> str:
    if col == "date":
        return value.strftime("%d.%m.%Y") if isinstance(value, date) else ""
    if col == "amount":
        return f"{value:,.0f} ₽".replace(",", " ")
    return str(value)


def _pick_window(flags: pd.DataFrame, size: int) -> int:
    """Начало фрагмента, где есть и пустая строка, и дубль, и побольше ошибок."""
    best, best_score = 0, -1
    for start in range(0, len(flags) - size):
        part = flags.iloc[start:start + size]
        if part["empty"].any() and part["dup"].any():
            score = part["dup"].sum() * 3 + part["dirty"].sum()
            if score > best_score:
                best, best_score = start, score
    return best


def make_before_after(raw: pd.DataFrame, clean: pd.DataFrame, path: Path, size: int = 14) -> None:
    raw = raw.copy()
    raw.columns = COLS[:4] + ["city", "product", "qty", "amount"]
    ids = raw["order_id"].map(clean_order_id)
    empty = raw.apply(lambda r: all(squeeze(v) == "" for v in r), axis=1)
    by_id = clean.set_index("order_id", drop=False)

    def dirty_cells(i: int) -> list[bool]:
        if empty[i] or ids[i] not in by_id.index:
            return [False] * len(COLS)
        rec = by_id.loc[ids[i]]
        return [_raw_text(raw.at[i, c]) != _clean_text(c, rec[c]) for c in COLS]

    flags = pd.DataFrame({"empty": empty, "dup": ids.duplicated() & ~empty,
                          "dirty": [sum(dirty_cells(i)) for i in range(len(raw))]})
    start = _pick_window(flags, size)
    window = range(start, start + size)

    # --- «до»
    before_rows, before_colors = [], []
    for i in window:
        mark = "пусто" if empty[i] else ("дубль" if flags.at[i, "dup"] else "")
        cells = [mark] + [_show_spaces(_raw_text(raw.at[i, c]))[:32] for c in COLS]
        if empty[i]:
            colors = [EMPTY] * len(cells)
        elif mark:
            colors = [BAD] * len(cells)
        else:
            colors = ["white"] + [BAD if d else "white" for d in dirty_cells(i)]
        before_rows.append(cells)
        before_colors.append(colors)

    # --- «после»: те же заказы без пустых строк и дублей
    after_rows, after_colors = [], []
    for i in window:
        if empty[i] or flags.at[i, "dup"]:
            continue
        rec = by_id.loc[ids[i]]
        after_rows.append([shorten(_clean_text(c, rec[c]), 30, placeholder="…") if c != "order_id"
                           else ids[i] for c in COLS])
        after_colors.append([GOOD if d else "white" for d in dirty_cells(i)])

    plt.rcParams["font.family"] = "DejaVu Sans"
    fig = plt.figure(figsize=(24, 0.42 * (size + 1) + 1.6))
    left = fig.add_axes([0.01, 0.1, 0.505, 0.76])
    right = fig.add_axes([0.535, 0.1, 0.455, 0.76])

    def draw(ax, rows, colors, titles, widths, title):
        ax.axis("off")
        share = (len(rows) + 1) / (size + 1)
        table = ax.table(cellText=rows, cellColours=colors, colLabels=titles, colWidths=widths,
                         cellLoc="left", bbox=[0, 1 - share, 1, share])
        table.auto_set_font_size(False)
        table.set_fontsize(9.5)
        for (r, c), cell in table.get_celld().items():
            cell.set_edgecolor("#C8CED6")
            cell.PAD = 0.04
            if r == 0:
                cell.set_facecolor(HEADER)
                cell.set_text_props(color="white", fontweight="bold", ha="center")
        ax.set_title(title, loc="left", fontsize=14, fontweight="bold", pad=10)
        return table

    t = draw(left, before_rows, before_colors, [""] + TITLES, [0.06, 0.1, 0.15, 0.27, 0.17, 0.13, 0.12],
             "ДО: raw_orders.xlsx")
    for r in range(1, len(before_rows) + 1):
        t[r, 0].set_text_props(color="#C0392B", fontweight="bold", ha="center")
    draw(right, after_rows, after_colors, TITLES, [0.11, 0.12, 0.33, 0.16, 0.16, 0.12],
         "ПОСЛЕ: clean_orders.xlsx")

    fig.text(0.5225, 0.55, "→", fontsize=34, ha="center", va="center", color=HEADER, fontweight="bold")
    fig.text(0.01, 0.97, "Очистка таблицы заказов: дубли, пустые строки, форматы дат, телефонов, сумм и городов",
             fontsize=17, fontweight="bold", va="top")
    fig.text(0.01, 0.03, "■ ", color="#E57373", fontsize=13, va="bottom")
    fig.text(0.022, 0.03, "ошибка в исходных данных", fontsize=11, va="bottom")
    fig.text(0.16, 0.03, "■ ", color="#5DBB7A", fontsize=13, va="bottom")
    fig.text(0.172, 0.03, "исправленное значение", fontsize=11, va="bottom")
    fig.text(0.29, 0.03, "·  лишний пробел", fontsize=11, va="bottom")
    fig.savefig(path, dpi=150, facecolor="white")
    plt.close(fig)
