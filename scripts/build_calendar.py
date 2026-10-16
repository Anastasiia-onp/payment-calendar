# -*- coding: utf-8 -*-
"""
Платёжный календарь: расчёт по неделям, поиск кассовых разрывов, Excel-отчёт.

Что делает:
  1. Собирает обязательства: разовые из листа «Реестр», регулярные из правил,
     при наличии — импортированные из выгрузок.
  2. Раскладывает их по неделям горизонта.
  3. Строит Excel: сводный реестр, календарь, разрывы — всё живыми формулами.
  4. Контрольно пересчитывает всё на Python и сверяет с формулами Excel.

Исходный файл только читается. Результат пишется в 02_calc новым файлом с датой.

Запуск:
    python scripts/build_calendar.py 01_input/Реестр_обязательств_ДЕМО.xlsx
    python scripts/build_calendar.py <файл> --import 02_calc/Импорт_2026-10-16.xlsx
"""

import os
import sys
from datetime import date, datetime, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from expand_recurring import expand, read_horizon, as_date  # noqa: E402

# ---------------------------------------------------------------- оформление

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="D6DCE4")
IN_FILL = PatternFill("solid", fgColor="E2EFDA")
OUT_FILL = PatternFill("solid", fgColor="FCE4D6")
GAP_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0.00\\ "₽"'
MONEY_SHORT = '#,##0\\ "₽"'
DATE_FMT = "DD.MM.YYYY"

SHEET_REG = "Реестр"
SHEET_PARAMS = "Параметры"
SHEET_ACC = "Счета и лимиты"

REG_FIRST_ROW = 2
ACC_TOTAL_ROW = 21
CRD_TOTAL_ROW = 36

# порядок статей в календаре: сначала приход, потом расход по приоритету
INCOME_ARTICLES = ["Поступления от заказчиков", "Прочие поступления",
                   "Продажа активов", "Получение кредита"]


# ---------------------------------------------------------------- чтение данных

def read_params(path):
    """Параметры расчёта: дата начала, горизонт, порог, остаток, лимит."""
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET_PARAMS]
    params = {"start": None, "weeks": 13, "threshold": 0.0,
              "company": "", "balance": 0.0, "credit_free": 0.0}

    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=2).value
        value = ws.cell(row=r, column=3).value
        if not name:
            continue
        name = str(name)
        if "Дата начала" in name:
            params["start"] = as_date(value)
        elif "Горизонт" in name and isinstance(value, (int, float)):
            params["weeks"] = int(value)
        elif "Порог" in name and isinstance(value, (int, float)):
            params["threshold"] = float(value)
        elif "Название компании" in name:
            params["company"] = str(value or "")

    # остаток и лимит берём с листа счетов: в «Параметрах» это формулы
    if SHEET_ACC in wb.sheetnames:
        acc = wb[SHEET_ACC]
        total = acc.cell(row=ACC_TOTAL_ROW, column=4).value
        free = acc.cell(row=CRD_TOTAL_ROW, column=6).value
        if isinstance(total, (int, float)):
            params["balance"] = float(total)
        else:  # формула не вычислена — складываем сами
            params["balance"] = sum(
                v for v in (acc.cell(row=r, column=4).value for r in range(6, 21))
                if isinstance(v, (int, float)))
        if isinstance(free, (int, float)):
            params["credit_free"] = float(free)
        else:
            params["credit_free"] = sum(
                (acc.cell(row=r, column=4).value or 0) - (acc.cell(row=r, column=5).value or 0)
                for r in range(26, 36)
                if isinstance(acc.cell(row=r, column=4).value, (int, float)))

    return params


def read_registry(path):
    """Разовые обязательства с листа «Реестр»."""
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET_REG]
    rows = []
    for r in range(REG_FIRST_ROW, ws.max_row + 1):
        typ = ws.cell(row=r, column=2).value
        project = ws.cell(row=r, column=3).value
        contractor = ws.cell(row=r, column=4).value
        basis = ws.cell(row=r, column=5).value
        article = ws.cell(row=r, column=6).value
        amount = ws.cell(row=r, column=8).value
        due = as_date(ws.cell(row=r, column=9).value)
        status = ws.cell(row=r, column=11).value
        comment = ws.cell(row=r, column=12).value

        if not contractor and amount is None and due is None:
            continue
        rows.append((typ, project, contractor, basis, article,
                     amount, due, status, comment))
    return rows


def read_imported(path):
    """Строки из файла импорта (лист «Импорт», заголовки в строке 4)."""
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in range(5, ws.max_row + 1):
        typ = ws.cell(row=r, column=1).value
        if not typ:
            continue
        rows.append((
            typ,
            ws.cell(row=r, column=2).value,
            ws.cell(row=r, column=3).value,
            ws.cell(row=r, column=4).value,
            ws.cell(row=r, column=5).value,
            ws.cell(row=r, column=6).value,
            as_date(ws.cell(row=r, column=7).value),
            ws.cell(row=r, column=8).value,
            ws.cell(row=r, column=9).value,
        ))
    return rows


def read_articles(path):
    """Справочник статей: статья -> (приоритет, вид деятельности)."""
    wb = load_workbook(path, data_only=True)
    if "Справочники" not in wb.sheetnames:
        return {}
    ws = wb["Справочники"]
    out = {}
    for r in range(2, 62):
        article = ws.cell(row=r, column=1).value
        if not article:
            continue
        out[str(article)] = (ws.cell(row=r, column=3).value,
                             ws.cell(row=r, column=4).value)
    return out


# ---------------------------------------------------------------- проверки данных

def check_rows(rows, start, end):
    """Замечания к данным. Ничего не исправляем — только сообщаем."""
    notes = []
    for i, (typ, project, contractor, basis, article,
            amount, due, status, comment) in enumerate(rows, start=1):
        who = contractor or "без контрагента"
        if not typ:
            notes.append("{w}: не указан тип (поступление или платёж)".format(w=who))
        if amount is None:
            notes.append("{w}: нет суммы — строка в расчёт не попадёт".format(w=who))
        elif not isinstance(amount, (int, float)):
            notes.append("{w}: сумма не число".format(w=who))
        elif amount <= 0:
            notes.append("{w}: сумма {a} — должна быть положительной".format(w=who, a=amount))
        if due is None:
            notes.append("{w}: нет срока платежа — строка в расчёт не попадёт".format(w=who))
        elif not (start <= due <= end):
            notes.append("{w}: срок {d} вне горизонта прогноза".format(
                w=who, d=due.strftime("%d.%m.%Y")))
        if not article:
            notes.append("{w}: не указана статья ДДС".format(w=who))
    return notes


def valid_rows(rows, start, end):
    """Строки, пригодные для расчёта."""
    out = []
    for row in rows:
        typ, project, contractor, basis, article, amount, due, status, comment = row
        if not isinstance(amount, (int, float)) or amount <= 0:
            continue
        if due is None or not (start <= due <= end):
            continue
        if str(status or "").strip().lower() == "оплачено":
            continue
        out.append(row)
    return out


# ---------------------------------------------------------------- расчёт

def week_bounds(start, weeks):
    return [(start + timedelta(weeks=w), start + timedelta(weeks=w, days=6))
            for w in range(weeks)]


def calculate(rows, params):
    """Контрольный расчёт на Python — эталон для сверки с формулами Excel."""
    start, weeks = params["start"], params["weeks"]
    bounds = week_bounds(start, weeks)
    balance = params["balance"]

    result = []
    for w, (ws_, we_) in enumerate(bounds, start=1):
        income = sum(r[5] for r in rows
                     if r[0] == "Поступление" and ws_ <= r[6] <= we_)
        outcome = sum(r[5] for r in rows
                      if r[0] == "Платёж" and ws_ <= r[6] <= we_)
        opening = balance
        balance = opening + income - outcome
        result.append({
            "week": w, "from": ws_, "to": we_,
            "income": income, "outcome": outcome,
            "net": income - outcome,
            "opening": opening, "closing": balance,
        })
    return result


def find_gaps(weeks_data, rows, threshold, credit_free):
    """Кассовые разрывы и варианты их закрытия."""
    gaps = []
    for w in weeks_data:
        if w["closing"] >= threshold:
            continue
        deficit = -w["closing"] if w["closing"] < 0 else 0.0
        movable = [r for r in rows
                   if r[0] == "Платёж" and w["from"] <= r[6] <= w["to"]
                   and _priority_of(r) == 3]
        movable.sort(key=lambda r: -r[5])
        gaps.append({
            "week": w["week"], "from": w["from"], "to": w["to"],
            "closing": w["closing"],
            "deficit": deficit,
            "kind": "разрыв" if w["closing"] < 0 else "ниже порога",
            "movable": movable,
            "movable_sum": sum(r[5] for r in movable),
            "covered_by_credit": deficit <= credit_free,
        })
    return gaps


_ARTICLES_CACHE = {}


def _priority_of(row):
    article = row[4]
    info = _ARTICLES_CACHE.get(str(article))
    if info and isinstance(info[0], (int, float)):
        return int(info[0])
    return 2


# ---------------------------------------------------------------- Excel

def _head(ws, row, col, text, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def build_source_sheet(ws, rows, articles, bounds):
    """Сводный реестр — основа всех формул календаря."""
    ws.title = "Сводный реестр"
    ws.freeze_panes = "A2"

    headers = [("Тип", 13), ("Проект", 24), ("Контрагент", 22), ("Основание", 30),
               ("Статья ДДС", 24), ("Сумма", 15), ("Дата", 13), ("Статус", 13),
               ("Приоритет", 10), ("Вид деятельности", 16), ("Источник", 16),
               ("Неделя", 9), ("Период недели", 18), ("Комментарий", 30)]
    for i, (name, width) in enumerate(headers, start=1):
        _head(ws, 1, i, name, width)

    for r, (row, source) in enumerate(rows, start=2):
        typ, project, contractor, basis, article, amount, due, status, comment = row
        prio, activity = articles.get(str(article), (None, None))
        week = next((n for n, (a, b) in enumerate(bounds, start=1) if a <= due <= b), None)
        period = "{a} - {b}".format(
            a=bounds[week - 1][0].strftime("%d.%m"),
            b=bounds[week - 1][1].strftime("%d.%m")) if week else ""
        values = [typ, project, contractor, basis, article, amount, due, status,
                  prio, activity, source, week, period, comment]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = BORDER
            if c == 6:
                cell.number_format = MONEY
            elif c == 7:
                cell.number_format = DATE_FMT
    ws.auto_filter.ref = "A1:N{n}".format(n=len(rows) + 1)
    return len(rows) + 1


def build_calendar_sheet(ws, params, weeks_data, articles, used_articles, last_row):
    """Календарь по неделям. Все суммы — формулы СУММЕСЛИМН по сводному реестру."""
    ws.title = "Календарь"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.freeze_panes = "B5"

    weeks = params["weeks"]
    src = "'Сводный реестр'"
    rng = lambda col: "{s}!${c}$2:${c}${n}".format(s=src, c=col, n=last_row)

    ws.cell(row=1, column=1, value="Платёжный календарь").font = TITLE_FONT
    ws.cell(row=2, column=1, value=params["company"]).font = Font(size=10, color="808080")

    # шапка недель
    ws.cell(row=4, column=1, value="Неделя").fill = HEADER_FILL
    ws.cell(row=4, column=1).font = HEADER_FONT
    ws.cell(row=4, column=1).border = BORDER
    for i, w in enumerate(weeks_data):
        col = 2 + i
        # в шапке — весь период недели, иначе непонятно, какие дни попали в столбец
        _head(ws, 4, col, "Неделя {n}\n{a} — {b}".format(
            n=w["week"], a=w["from"].strftime("%d.%m"), b=w["to"].strftime("%d.%m")), 14)
        ws.cell(row=3, column=col, value=w["from"]).number_format = DATE_FMT
        ws.cell(row=3, column=col).font = Font(size=8, color="808080")
        ws.cell(row=3, column=col).alignment = Alignment(horizontal="center")
    _head(ws, 4, 2 + weeks, "Итого", 16)
    ws.row_dimensions[4].height = 30

    def money_row(row, label, fill=None, bold=False, indent=False):
        c = ws.cell(row=row, column=1, value=label)
        c.border = BORDER
        c.font = Font(bold=bold, size=10)
        if indent:
            c.alignment = Alignment(indent=1)
        if fill:
            c.fill = fill
        for i in range(weeks + 1):
            cell = ws.cell(row=row, column=2 + i)
            cell.border = BORDER
            cell.number_format = MONEY_SHORT
            if fill:
                cell.fill = fill
            if bold:
                cell.font = Font(bold=True, size=10)

    def sumifs(col_letter, typ, article, week_index):
        """Формула суммы по типу, статье и неделе."""
        w = weeks_data[week_index]
        parts = [
            rng("F"),
            rng("A"), '"{t}"'.format(t=typ),
            rng("G"), '">="&DATE({y},{m},{d})'.format(
                y=w["from"].year, m=w["from"].month, d=w["from"].day),
            rng("G"), '"<="&DATE({y},{m},{d})'.format(
                y=w["to"].year, m=w["to"].month, d=w["to"].day),
        ]
        if article is not None:
            parts += [rng("E"), '"{a}"'.format(a=article)]
        return "=SUMIFS({p})".format(p=",".join(parts))

    row = 5
    layout = {}

    # --- поступления
    money_row(row, "ПОСТУПЛЕНИЯ", SECTION_FILL, bold=True)
    layout["income_header"] = row
    row += 1

    income_articles = [a for a in used_articles if a in INCOME_ARTICLES]
    income_rows = []
    for article in income_articles:
        money_row(row, article, IN_FILL, indent=True)
        for i in range(weeks):
            ws.cell(row=row, column=2 + i, value=sumifs("F", "Поступление", article, i))
        ws.cell(row=row, column=2 + weeks,
                value="=SUM(B{r}:{c}{r})".format(r=row, c=get_column_letter(1 + weeks)))
        income_rows.append(row)
        row += 1

    money_row(row, "Итого поступления", None, bold=True)
    for i in range(weeks + 1):
        col = get_column_letter(2 + i)
        if income_rows:
            ws.cell(row=row, column=2 + i,
                    value="={s}".format(s="+".join("{c}{r}".format(c=col, r=x)
                                                   for x in income_rows)))
        else:
            ws.cell(row=row, column=2 + i, value=0)
    layout["income_total"] = row
    row += 2

    # --- платежи
    money_row(row, "ПЛАТЕЖИ", SECTION_FILL, bold=True)
    row += 1

    outcome_articles = [a for a in used_articles if a not in INCOME_ARTICLES]
    outcome_articles.sort(key=lambda a: (articles.get(a, (9, ""))[0] or 9, a))
    outcome_rows = []
    for article in outcome_articles:
        prio = articles.get(article, (None, None))[0]
        label = "{a}  (приоритет {p})".format(a=article, p=prio) if prio else article
        money_row(row, label, OUT_FILL, indent=True)
        for i in range(weeks):
            ws.cell(row=row, column=2 + i, value=sumifs("F", "Платёж", article, i))
        ws.cell(row=row, column=2 + weeks,
                value="=SUM(B{r}:{c}{r})".format(r=row, c=get_column_letter(1 + weeks)))
        outcome_rows.append(row)
        row += 1

    money_row(row, "Итого платежи", None, bold=True)
    for i in range(weeks + 1):
        col = get_column_letter(2 + i)
        ws.cell(row=row, column=2 + i,
                value="={s}".format(s="+".join("{c}{r}".format(c=col, r=x)
                                               for x in outcome_rows)))
    layout["outcome_total"] = row
    row += 2

    # --- итоги по неделям
    money_row(row, "Чистый денежный поток", None, bold=True)
    for i in range(weeks + 1):
        col = get_column_letter(2 + i)
        ws.cell(row=row, column=2 + i, value="={c}{a}-{c}{b}".format(
            c=col, a=layout["income_total"], b=layout["outcome_total"]))
    layout["net"] = row
    row += 1

    money_row(row, "Остаток на начало недели")
    for i in range(weeks):
        col = get_column_letter(2 + i)
        if i == 0:
            ws.cell(row=row, column=2, value="='Параметры и проверки'!$C$5")
        else:
            prev = get_column_letter(1 + i)
            ws.cell(row=row, column=2 + i,
                    value="={p}{r}".format(p=prev, r=row + 1))
    layout["opening"] = row
    row += 1

    money_row(row, "Остаток на конец недели", None, bold=True)
    for i in range(weeks):
        col = get_column_letter(2 + i)
        ws.cell(row=row, column=2 + i,
                value="={c}{o}+{c}{n}".format(c=col, o=layout["opening"], n=layout["net"]))
    layout["closing"] = row
    row += 1

    money_row(row, "Доступно с учётом кредитных лимитов")
    for i in range(weeks):
        col = get_column_letter(2 + i)
        ws.cell(row=row, column=2 + i,
                value="={c}{cl}+'Параметры и проверки'!$C$6".format(c=col, cl=layout["closing"]))
    layout["with_credit"] = row
    row += 1

    money_row(row, "Состояние недели", None, bold=True)
    for i in range(weeks):
        col = get_column_letter(2 + i)
        cell = ws.cell(row=row, column=2 + i)
        cell.value = ('=IF({c}{cl}<0,"РАЗРЫВ",IF({c}{cl}<\'Параметры и проверки\'!$C$7,'
                      '"Ниже порога","Норма"))'.format(c=col, cl=layout["closing"]))
        cell.number_format = "General"
        cell.alignment = Alignment(horizontal="center")
        cell.font = Font(bold=True, size=9)
    layout["state"] = row

    # подсветка факта (значения уже известны из контрольного расчёта)
    for i, w in enumerate(weeks_data):
        col = 2 + i
        if w["closing"] < 0:
            fill = GAP_FILL
        elif w["closing"] < params["threshold"]:
            fill = WARN_FILL
        else:
            fill = OK_FILL
        for r in (layout["closing"], layout["state"]):
            ws.cell(row=r, column=col).fill = fill

    # --- график остатка
    chart = LineChart()
    chart.title = "Остаток денежных средств по неделям"
    chart.height, chart.width = 8, 26
    data = Reference(ws, min_col=2, max_col=1 + weeks,
                     min_row=layout["closing"], max_row=layout["closing"])
    cats = Reference(ws, min_col=2, max_col=1 + weeks, min_row=3, max_row=3)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.y_axis.numFmt = '#,##0'
    ws.add_chart(chart, "A{r}".format(r=layout["state"] + 3))

    return layout


def build_gaps_sheet(ws, gaps, params):
    """Разрывы и варианты их закрытия."""
    ws.title = "Разрывы"
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEFG", [8, 14, 18, 18, 20, 36, 22]):
        ws.column_dimensions[col].width = width

    ws.cell(row=1, column=1, value="Кассовые разрывы и варианты закрытия").font = TITLE_FONT

    if not gaps:
        ws.cell(row=3, column=1,
                value="Разрывов в горизонте прогноза нет.").font = Font(bold=True)
        return

    headers = ["Неделя", "Дата", "Состояние", "Нехватка", "Можно перенести",
               "Что переносить", "Покрытие лимитом"]
    for i, name in enumerate(headers, start=1):
        _head(ws, 3, i, name)

    row = 4
    for g in gaps:
        ws.cell(row=row, column=1, value=g["week"]).border = BORDER
        c = ws.cell(row=row, column=2, value=g["from"])
        c.number_format = DATE_FMT
        c.border = BORDER
        ws.cell(row=row, column=3, value=g["kind"]).border = BORDER
        c = ws.cell(row=row, column=4, value=g["deficit"] or g["closing"])
        c.number_format = MONEY
        c.border = BORDER
        c = ws.cell(row=row, column=5, value=g["movable_sum"])
        c.number_format = MONEY
        c.border = BORDER

        if g["movable"]:
            what = "; ".join("{c} — {a}".format(
                c=m[2], a=format(m[5], ",.0f").replace(",", " ")) for m in g["movable"][:3])
        else:
            what = "платежей приоритета 3 на этой неделе нет"
        ws.cell(row=row, column=6, value=what).border = BORDER
        ws.cell(row=row, column=6).alignment = Alignment(wrap_text=True)

        if g["deficit"] <= 0:
            verdict = "не требуется"
        elif g["covered_by_credit"]:
            verdict = "да, лимита хватает"
        else:
            verdict = "НЕТ, лимита мало"
        c = ws.cell(row=row, column=7, value=verdict)
        c.border = BORDER
        c.fill = OK_FILL if "да" in verdict or "не требуется" in verdict else GAP_FILL

        for col in range(1, 8):
            if g["kind"] == "разрыв" and col in (1, 3, 4):
                ws.cell(row=row, column=col).fill = GAP_FILL
        row += 1

    row += 1
    ws.cell(row=row, column=1,
            value="Перенос предлагается, но не применяется автоматически: "
                  "решение принимает финансист.").font = Font(size=9, color="808080")


def build_params_sheet(ws, params, notes, control):
    """Параметры расчёта и результаты проверок."""
    ws.title = "Параметры и проверки"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 60

    ws.cell(row=2, column=2, value="Параметры расчёта").font = TITLE_FONT

    items = [
        ("Дата начала прогноза", params["start"], DATE_FMT),
        ("Остаток денежных средств на старте", params["balance"], MONEY),
        ("Свободные кредитные лимиты", params["credit_free"], MONEY),
        ("Порог тревоги", params["threshold"], MONEY),
        ("Горизонт, недель", params["weeks"], "0"),
        ("Компания", params["company"], "@"),
        ("Дата расчёта", date.today(), DATE_FMT),
    ]
    row = 4
    for name, value, fmt in items:
        ws.cell(row=row, column=2, value=name).font = Font(bold=True, size=10)
        c = ws.cell(row=row, column=3, value=value)
        c.number_format = fmt
        c.border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=2, value="Контрольные суммы").font = TITLE_FONT
    row += 1
    for name, value, fmt in [
        ("Обязательств в расчёте, шт.", control["count"], "0"),
        ("Поступления за горизонт", control["income"], MONEY),
        ("Платежи за горизонт", control["outcome"], MONEY),
        ("Чистый поток", control["net"], MONEY),
        ("Остаток на конец горизонта", control["closing"], MONEY),
        ("Минимальный остаток", control["min_balance"], MONEY),
    ]:
        ws.cell(row=row, column=2, value=name).font = Font(bold=True, size=10)
        c = ws.cell(row=row, column=3, value=value)
        c.number_format = fmt
        c.border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=2, value="Замечания к данным").font = TITLE_FONT
    row += 1
    if not notes:
        ws.cell(row=row, column=2, value="Замечаний нет: все строки прошли проверку.").fill = OK_FILL
    else:
        for note in notes:
            c = ws.cell(row=row, column=2, value=note)
            c.fill = WARN_FILL
            c.alignment = Alignment(wrap_text=True)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
            row += 1


# ---------------------------------------------------------------- сборка

def build(path, import_paths=None):
    params = read_params(path)
    if params["start"] is None:
        raise SystemExit("На листе «Параметры» не заполнена дата начала — данных нет.")

    start = params["start"]
    end = start + timedelta(weeks=params["weeks"]) - timedelta(days=1)

    articles = read_articles(path)
    _ARTICLES_CACHE.clear()
    _ARTICLES_CACHE.update(articles)

    sourced = [(r, "Реестр") for r in read_registry(path)]
    recurring, recur_problems = expand(path, start, end)
    sourced += [(r, "Регулярные") for r in recurring]
    for extra in (import_paths or []):
        sourced += [(r, "Импорт") for r in read_imported(extra)]

    all_rows = [r for r, _ in sourced]
    notes = check_rows(all_rows, start, end)
    notes += ["Правила повтора — " + str(p) for p in recur_problems]

    good = valid_rows(all_rows, start, end)
    good_sourced = [(r, s) for r, s in sourced if r in good]

    weeks_data = calculate(good, params)
    gaps = find_gaps(weeks_data, good, params["threshold"], params["credit_free"])

    control = {
        "count": len(good),
        "income": sum(w["income"] for w in weeks_data),
        "outcome": sum(w["outcome"] for w in weeks_data),
        "net": sum(w["net"] for w in weeks_data),
        "closing": weeks_data[-1]["closing"] if weeks_data else params["balance"],
        "min_balance": min((w["closing"] for w in weeks_data), default=params["balance"]),
    }

    used = []
    for r in good:
        if r[4] and r[4] not in used:
            used.append(str(r[4]))

    wb = Workbook()
    build_params_sheet(wb.active, params, notes, control)
    last_row = build_source_sheet(wb.create_sheet(), good_sourced, articles,
                                  week_bounds(start, params["weeks"]))
    build_calendar_sheet(wb.create_sheet(), params, weeks_data, articles, used, last_row)
    build_gaps_sheet(wb.create_sheet(), gaps, params)

    # календарь первым листом — его открывают чаще всего
    wb.move_sheet("Календарь", offset=-2)

    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(base, "02_calc")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Платежный_календарь_{d}.xlsx".format(
        d=date.today().strftime("%Y-%m-%d")))
    wb.save(out_path)

    return out_path, params, weeks_data, gaps, control, notes


def main():
    args = sys.argv[1:]
    if not args:
        print("Укажите файл реестра:")
        print("  python scripts/build_calendar.py 01_input/Реестр_обязательств_ДЕМО.xlsx")
        return 1

    path = args[0]
    imports = []
    i = 1
    while i < len(args):
        if args[i] in ("--import", "-i") and i + 1 < len(args):
            imports.append(args[i + 1])
            i += 2
        else:
            i += 1

    out, params, weeks_data, gaps, control, notes = build(path, imports)
    f = lambda x: format(x, ",.0f").replace(",", " ")

    print("Платёжный календарь: {c}".format(c=params["company"] or "компания не указана"))
    print("Горизонт: {a} - {b}, {w} недель".format(
        a=params["start"].strftime("%d.%m.%Y"),
        b=(params["start"] + timedelta(weeks=params["weeks"], days=-1)).strftime("%d.%m.%Y"),
        w=params["weeks"]))
    print("Обязательств в расчёте: {n}".format(n=control["count"]))
    print()
    print("{:<4} {:<12} {:>14} {:>14} {:>15}  {}".format(
        "Нед", "Начало", "Поступления", "Платежи", "Остаток", "Состояние"))
    for w in weeks_data:
        state = "РАЗРЫВ" if w["closing"] < 0 else (
            "ниже порога" if w["closing"] < params["threshold"] else "")
        print("{:<4} {:<12} {:>14} {:>14} {:>15}  {}".format(
            w["week"], w["from"].strftime("%d.%m.%Y"),
            f(w["income"]), f(w["outcome"]), f(w["closing"]), state))

    print()
    print("Поступления за горизонт: {s} ₽".format(s=f(control["income"])))
    print("Платежи за горизонт:     {s} ₽".format(s=f(control["outcome"])))
    print("Чистый поток:            {s} ₽".format(s=f(control["net"])))
    print("Минимальный остаток:     {s} ₽".format(s=f(control["min_balance"])))
    print("Остаток на конец:        {s} ₽".format(s=f(control["closing"])))

    if gaps:
        print()
        print("Проблемных недель: {n}".format(n=len(gaps)))
        for g in gaps:
            print("  неделя {w} ({d}): {k}, остаток {c}{cov}".format(
                w=g["week"], d=g["from"].strftime("%d.%m.%Y"), k=g["kind"],
                c=f(g["closing"]),
                cov=", покрывается лимитом" if g["covered_by_credit"] and g["deficit"]
                else (", ЛИМИТА НЕ ХВАТАЕТ" if g["deficit"] else "")))

    if notes:
        print()
        print("Замечания к данным ({n}):".format(n=len(notes)))
        for note in notes[:15]:
            print("  " + note)
        if len(notes) > 15:
            print("  ... ещё {n}".format(n=len(notes) - 15))

    print()
    print("Файл: 02_calc/" + os.path.basename(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
