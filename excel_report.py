"""Сборка clean_orders.xlsx: листы «Данные», «Справочник», «Сводка» с живыми формулами."""

import zipfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

from catalog import PRODUCTS

BIG_ORDER = 20000  # порог «крупного заказа», ₽
RUB = '#,##0 "₽"'
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
BIG_FILL = PatternFill("solid", fgColor="FFE699", bgColor="FFE699")  # в условном формате Excel берёт bgColor
THIN = Side(style="thin", color="BFBFBF")

DATA_COLUMNS = ["№ заказа", "Дата", "Клиент", "Телефон", "Город", "Товар", "Кол-во", "Цена за шт.", "Сумма"]


def _header(ws, row: int, first_col: int, titles: list[str]) -> None:
    for i, title in enumerate(titles):
        cell = ws.cell(row=row, column=first_col + i, value=title)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _widths(ws, widths: dict[str, float]) -> None:
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width


def _auto_widths(ws, min_width: float = 8, max_width: float = 45) -> None:
    for col in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 300)):
        letter = get_column_letter(col[0].column)
        texts = [c.value for c in col if c.value is not None and not str(c.value).startswith("=")]
        longest = max((len(str(v)) for v in texts), default=0)
        header = col[0].value
        header_need = len(str(header)) * 1.2 + 5 if header else 0  # жирный шрифт + кнопка фильтра
        ws.column_dimensions[letter].width = max(min_width, min(max(longest + 2, header_need), max_width))


def build_reference(wb: Workbook) -> None:
    ws = wb.create_sheet("Справочник")
    _header(ws, 1, 1, ["Товар", "Цена, ₽", "Категория"])
    for row, (product, (price, category)) in enumerate(sorted(PRODUCTS.items()), start=2):
        ws.cell(row=row, column=1, value=product)
        ws.cell(row=row, column=2, value=price).number_format = RUB
        ws.cell(row=row, column=3, value=category)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    _auto_widths(ws)


def build_data(wb: Workbook, df: pd.DataFrame) -> None:
    ws = wb.active
    ws.title = "Данные"
    _header(ws, 1, 1, DATA_COLUMNS)
    for i, rec in enumerate(df.itertuples(index=False), start=2):
        d = rec.date
        ws.cell(row=i, column=1, value=rec.order_id)
        ws.cell(row=i, column=2, value=datetime(d.year, d.month, d.day)).number_format = "DD.MM.YYYY"
        ws.cell(row=i, column=3, value=rec.name)
        ws.cell(row=i, column=4, value=rec.phone or None)
        ws.cell(row=i, column=5, value=rec.city)
        ws.cell(row=i, column=6, value=rec.product)
        ws.cell(row=i, column=7, value=int(rec.qty))
        # цена подтягивается из справочника формулой ВПР
        ws.cell(row=i, column=8, value=f"=VLOOKUP(F{i},'Справочник'!$A:$B,2,FALSE)").number_format = RUB
        ws.cell(row=i, column=9, value=float(rec.amount)).number_format = RUB
        for col in (1, 2, 7):
            ws.cell(row=i, column=col).alignment = Alignment(horizontal="center")

    last = len(df) + 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{last}"
    # подсветка крупных заказов: вся строка, если сумма >= порога
    ws.conditional_formatting.add(
        f"A2:I{last}", FormulaRule(formula=[f"$I2>={BIG_ORDER}"], fill=BIG_FILL, font=Font(bold=True)))
    _auto_widths(ws)
    _widths(ws, {"D": 16, "H": 15, "I": 14})


def build_summary(wb: Workbook, df: pd.DataFrame) -> None:
    ws = wb.create_sheet("Сводка")
    ws["A1"] = "Выручка по месяцам"
    ws["F1"] = "Выручка по городам"
    for ref in ("A1", "F1"):
        ws[ref].font = Font(bold=True, size=13)

    # --- по месяцам: СУММЕСЛИМН по диапазону дат [1-е число; 1-е число след. месяца)
    _header(ws, 3, 1, ["Месяц", "Заказов", "Выручка", "Средний чек"])
    months = sorted({date(d.year, d.month, 1) for d in df["date"]})
    for r, m in enumerate(months, start=4):
        ws.cell(row=r, column=1, value=datetime(m.year, m.month, 1)).number_format = "[$-419]MMMM YYYY"
        period = f"""'Данные'!$B:$B,">="&$A{r},'Данные'!$B:$B,"<"&EDATE($A{r},1)"""
        ws.cell(row=r, column=2, value=f"=COUNTIFS({period})")
        ws.cell(row=r, column=3, value=f"=SUMIFS('Данные'!$I:$I,{period})").number_format = RUB
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)").number_format = RUB
    m_total = 4 + len(months)
    ws.cell(row=m_total, column=1, value="Итого")
    ws.cell(row=m_total, column=2, value=f"=SUM(B4:B{m_total - 1})")
    ws.cell(row=m_total, column=3, value=f"=SUM(C4:C{m_total - 1})").number_format = RUB
    ws.cell(row=m_total, column=4, value=f"=IFERROR(C{m_total}/B{m_total},0)").number_format = RUB

    # --- по городам: СУММЕСЛИМН по названию города, порядок по убыванию выручки
    _header(ws, 3, 6, ["Город", "Заказов", "Выручка", "Доля"])
    cities = df.groupby("city")["amount"].sum().sort_values(ascending=False).index
    for r, city in enumerate(cities, start=4):
        ws.cell(row=r, column=6, value=city)
        ws.cell(row=r, column=7, value=f"=COUNTIFS('Данные'!$E:$E,$F{r})")
        ws.cell(row=r, column=8, value=f"=SUMIFS('Данные'!$I:$I,'Данные'!$E:$E,$F{r})").number_format = RUB
    c_total = 4 + len(cities)
    for r in range(4, c_total):
        ws.cell(row=r, column=9, value=f"=H{r}/$H${c_total}").number_format = "0.0%"
    ws.cell(row=c_total, column=6, value="Итого")
    ws.cell(row=c_total, column=7, value=f"=SUM(G4:G{c_total - 1})")
    ws.cell(row=c_total, column=8, value=f"=SUM(H4:H{c_total - 1})").number_format = RUB
    ws.cell(row=c_total, column=9, value=f"=SUM(I4:I{c_total - 1})").number_format = "0.0%"

    for row, cols in ((m_total, range(1, 5)), (c_total, range(6, 10))):
        for col in cols:
            cell = ws.cell(row=row, column=col)
            cell.font, cell.fill = Font(bold=True), TOTAL_FILL
    for r in range(4, max(m_total, c_total) + 1):
        for col in (*range(1, 5), *range(6, 10)):
            if (col < 5 and r <= m_total) or (col > 5 and r <= c_total):
                ws.cell(row=r, column=col).border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        for col in (2, 7, 9):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")

    # --- крупные заказы
    info = m_total + 2
    ws.cell(row=info, column=1, value="Крупные заказы").font = Font(bold=True)
    ws.cell(row=info, column=2, value=f"""=COUNTIFS('Данные'!$I:$I,">={BIG_ORDER}")""").alignment = Alignment(horizontal="center")
    ws.cell(row=info, column=3, value=f"""=SUMIFS('Данные'!$I:$I,'Данные'!$I:$I,">={BIG_ORDER}")""").number_format = RUB
    ws.cell(row=info, column=4, value=f"от {BIG_ORDER:,} ₽".replace(",", " ")).font = Font(color="808080")
    ws.cell(row=info + 1, column=1,
            value="Все числа на листе считаются формулами и пересчитываются при изменении листа «Данные».").font = \
        Font(italic=True, color="808080")

    _widths(ws, {"A": 16, "B": 10, "C": 15, "D": 14, "E": 3, "F": 19, "G": 10, "H": 15, "I": 9})

    chart = BarChart()
    chart.title = "Выручка по месяцам, ₽"
    chart.style = 10
    chart.legend = None
    chart.varyColors = False
    chart.height, chart.width = 8.5, 14.5
    chart.add_data(Reference(ws, min_col=3, min_row=3, max_row=m_total - 1), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=4, max_row=m_total - 1))
    chart.series[0].graphicalProperties.solidFill = "1F4E78"
    chart.series[0].graphicalProperties.line.solidFill = "1F4E78"
    chart.x_axis.number_format = "[$-419]MMM"
    chart.y_axis.majorGridlines = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    ws.add_chart(chart, "K3")


def _ignore_number_as_text(path: Path, sheet_xml: str, sqref: str) -> None:
    """Убирает зелёные треугольники «число сохранено как текст» у телефонов.

    openpyxl не умеет записывать <ignoredErrors>, поэтому дописываем тег в XML листа сами.
    По схеме xlsx он стоит перед <drawing>/<tableParts>/<extLst> или в конце листа.
    """
    tag = f'<ignoredErrors><ignoredError sqref="{sqref}" numberStoredAsText="1"/></ignoredErrors>'
    with zipfile.ZipFile(path) as src:
        files = {name: src.read(name) for name in src.namelist()}
    xml = files[sheet_xml].decode("utf-8")
    positions = [i for i in (xml.find(t) for t in ("<drawing", "<legacyDrawing", "<tableParts", "<extLst"))
                 if i != -1] or [xml.rfind("</worksheet>")]
    pos = min(positions)
    files[sheet_xml] = (xml[:pos] + tag + xml[pos:]).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, content in files.items():
            dst.writestr(name, content)


def save_report(df: pd.DataFrame, path: Path) -> None:
    wb = Workbook()
    build_data(wb, df)
    build_reference(wb)
    build_summary(wb, df)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)  # Excel пересчитает формулы при открытии
    wb.save(path)
    _ignore_number_as_text(path, "xl/worksheets/sheet1.xml", f"D2:D{len(df) + 1}")
