# -*- coding: utf-8 -*-
"""
Генератор шаблона реестра обязательств для платёжного календаря.

Создаёт два файла в папке 01_input:
  1. Реестр_обязательств_Шаблон.xlsx — пустая форма для заполнения
  2. Реестр_обязательств_ДЕМО.xlsx   — та же форма с условными данными

Листы формы: Инструкция, Параметры, Счета и лимиты, Реестр, Справочники.

Запуск:  python scripts/make_template.py
"""

import os
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- оформление

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SECTION_FONT = Font(bold=True, size=11, color="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # жёлтый — поля для ввода
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")    # зелёный — формулы
TOTAL_FILL = PatternFill("solid", fgColor="D6DCE4")   # серый — итоги
DEMO_FILL = PatternFill("solid", fgColor="FCE4D6")    # оранжевый — предупреждение

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0.00\\ "₽"'
PERCENT = '0.00"%"'
DATE_FMT = "DD.MM.YYYY"

# ---------------------------------------------------------------- справочники

# статья ДДС, тип по умолчанию, приоритет, вид деятельности
ARTICLES = [
    ("Поступления от заказчиков", "Поступление", 0, "Операционная"),
    ("Прочие поступления", "Поступление", 0, "Операционная"),
    ("Налоги", "Платёж", 1, "Операционная"),
    ("ФОТ", "Платёж", 1, "Операционная"),
    ("Страховые взносы", "Платёж", 1, "Операционная"),
    ("Подрядчики", "Платёж", 2, "Операционная"),
    ("Материалы", "Платёж", 2, "Операционная"),
    ("Аренда", "Платёж", 2, "Операционная"),
    ("Командировки", "Платёж", 3, "Операционная"),
    ("Прочие расходы", "Платёж", 3, "Операционная"),
    ("Покупка оборудования", "Платёж", 3, "Инвестиционная"),
    ("Продажа активов", "Поступление", 0, "Инвестиционная"),
    ("Получение кредита", "Поступление", 0, "Финансовая"),
    ("Погашение кредита", "Платёж", 1, "Финансовая"),
    ("Проценты по кредитам", "Платёж", 1, "Финансовая"),
    ("Лизинг", "Платёж", 2, "Финансовая"),
    ("Дивиденды", "Платёж", 3, "Финансовая"),
]
ART_ROWS = 60              # сколько строк отведено под статьи (справочник открытый)
ART_LAST = ART_ROWS + 1    # последняя строка блока статей

ACTIVITIES = ["Операционная", "Инвестиционная", "Финансовая"]
ACTIVITY_MEANING = [
    ("Операционная", "Основная деятельность: выручка, зарплата, налоги, материалы, подрядчики"),
    ("Инвестиционная", "Покупка и продажа основных средств"),
    ("Финансовая", "Кредиты, лизинг, проценты, дивиденды"),
]

PRIORITY_MEANING = [
    (1, "Критично — просрочка даёт штрафы и остановку работы"),
    (2, "Важно — просрочка портит отношения с контрагентом"),
    (3, "Можно перенести без последствий"),
]

STATUSES = ["План", "Подтверждено", "Оплачено"]
TYPES = ["Поступление", "Платёж"]
CREDIT_KINDS = ["Овердрафт", "Кредитная линия", "Кредит", "Факторинг"]

PROJ_ROWS = 30             # строк под список проектов
PROJ_LAST = PROJ_ROWS + 1

FREQUENCIES = ["Ежемесячно", "Ежеквартально", "Раз в 2 недели", "Еженедельно"]
SHIFT_RULES = ["Раньше", "Позже", "Не переносить"]
YES_NO = ["Да", "Нет"]

# нерабочие дни РФ (праздники). Переносы выходных на 2027 год официально
# не утверждены — список правится пользователем на листе «Справочники».
HOLIDAYS = [
    date(2026, 11, 4),
    date(2027, 1, 1), date(2027, 1, 2), date(2027, 1, 3), date(2027, 1, 4),
    date(2027, 1, 5), date(2027, 1, 6), date(2027, 1, 7), date(2027, 1, 8),
    date(2027, 2, 23), date(2027, 3, 8),
    date(2027, 5, 1), date(2027, 5, 9), date(2027, 6, 12), date(2027, 11, 4),
]
HOL_ROWS = 40
HOL_LAST = HOL_ROWS + 1

DEMO_PROJECTS = [
    "Объект 1 — ТЦ, остекление",
    "Объект 2 — БЦ, фасад",
    "Объект 3 — склад, окна",
    "Общие",
]
EMPTY_PROJECTS = ["Общие"]

# ---------------------------------------------------------------- лист «Счета и лимиты»

ACC_HEAD_ROW = 5                       # шапка таблицы счетов
ACC_FIRST = ACC_HEAD_ROW + 1
ACC_ROWS = 15
ACC_LAST = ACC_FIRST + ACC_ROWS - 1
ACC_TOTAL = ACC_LAST + 1               # строка итога по счетам

CRD_TITLE_ROW = ACC_TOTAL + 3
CRD_HEAD_ROW = CRD_TITLE_ROW + 1
CRD_FIRST = CRD_HEAD_ROW + 1
CRD_ROWS = 10
CRD_LAST = CRD_FIRST + CRD_ROWS - 1
CRD_TOTAL = CRD_LAST + 1               # строка итога по лимитам

DEMO_ACCOUNTS = [
    ("Банк 1", "Расчётный счёт, основной", 1850000, "Основные расчёты с заказчиками"),
    ("Банк 2", "Расчётный счёт", 720000, "Обслуживание кредитной линии"),
    ("Банк 1", "Спецсчёт по госконтракту", 280000, "Целевые средства, снимать нельзя"),
]

DEMO_CREDITS = [
    ("Банк 1", "Овердрафт", 3000000, 0, 22.0, date(2027, 6, 30),
     "Лимит не выбран, доступен сразу"),
    ("Банк 2", "Кредитная линия", 5000000, 2000000, 19.5, date(2026, 12, 31),
     "Выбрано 2 млн, продление в декабре"),
]

# ---------------------------------------------------------------- лист «Регулярные платежи»

RECUR_COLS = [
    ("Включить", 11),
    ("Проект", 24),
    ("Контрагент", 22),
    ("Основание", 32),
    ("Статья ДДС", 24),
    ("Сумма, ₽", 16),
    ("Периодичность", 16),
    ("День месяца", 13),
    ("Действует с", 14),
    ("Действует по", 14),
    ("Перенос с выходного", 16),
    ("Комментарий", 30),
]
RECUR_HEAD = 4
RECUR_FIRST = RECUR_HEAD + 1
RECUR_ROWS = 40
RECUR_LAST = RECUR_FIRST + RECUR_ROWS - 1

# включить, проект, контрагент, основание, статья, сумма, периодичность,
# день, с, по, перенос, комментарий
DEMO_RECURRING = [
    ("Да", "Общие", "Персонал", "Аванс за {месяц}", "ФОТ", 1450000,
     "Ежемесячно", 20, date(2026, 10, 1), date(2027, 12, 31), "Раньше",
     "Зарплата идёт отдельно: сумма меняется"),
    ("Да", "Общие", "Арендодатель", "Аренда офиса и склада, {месяц}", "Аренда", 320000,
     "Ежемесячно", 5, date(2026, 10, 1), date(2027, 12, 31), "Раньше", ""),
    ("Да", "Общие", "Лизинговая компания", "Лизинг автовышки", "Лизинг", 185000,
     "Ежемесячно", 15, date(2026, 10, 1), date(2027, 12, 31), "Раньше",
     "График лизинга, 36 платежей"),
    ("Да", "Общие", "Прочие", "Связь, ПО, канцелярия", "Прочие расходы", 95000,
     "Ежемесячно", 30, date(2026, 10, 1), date(2027, 12, 31), "Раньше", ""),
    ("Да", "Общие", "Банк 2", "Проценты по кредитной линии, {месяц}",
     "Проценты по кредитам", 32500,
     "Ежемесячно", 30, date(2026, 10, 1), date(2026, 12, 31), "Позже",
     "2 000 000 × 19,5% / 12"),
]

# ---------------------------------------------------------------- лист «Реестр»

REGISTRY_COLS = [
    ("№", 6),
    ("Тип", 14),
    ("Проект", 26),
    ("Контрагент", 24),
    ("Основание", 30),
    ("Статья ДДС", 26),
    ("Вид деятельности", 17),
    ("Сумма, ₽", 16),
    ("Срок платежа", 15),
    ("Приоритет", 11),
    ("Статус", 15),
    ("Комментарий", 32),
]
COL_NUM, COL_ACTIVITY, COL_PRIORITY = 1, 7, 10
CALC_COLS = (COL_NUM, COL_ACTIVITY, COL_PRIORITY)

MAX_ROWS = 300  # до какой строки растянуты форматы и проверки


# ---------------------------------------------------------------- вспомогательное

def put_header(ws, row, col, text, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def style_input(cell, fmt=None):
    cell.fill = INPUT_FILL
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    return cell


def style_calc(cell, fmt=None):
    cell.fill = CALC_FILL
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    return cell


# ---------------------------------------------------------------- листы

def build_instruction(ws, demo):
    ws.title = "Инструкция"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 112

    lines = [
        ("Реестр обязательств — исходные данные платёжного календаря", "title"),
        ("", ""),
        ("Порядок заполнения", "h"),
        ("1. «Счета и лимиты» — остатки по всем расчётным счетам и доступные кредитные лимиты.", ""),
        ("2. «Параметры» — дата начала прогноза и порог тревоги. Остаток подтянется со счетов сам.", ""),
        ("3. «Справочники» — свои объекты и, при необходимости, свои статьи ДДС.", ""),
        ("4. «Реестр» — все ожидаемые поступления и платежи на 13 недель вперёд.", ""),
        ("5. Сохраните файл в папку 01_input и запустите расчёт.", ""),
        ("", ""),
        ("Как заполнять «Реестр»", "h"),
        ("Одна строка — одно обязательство с одной датой. Платёж частями = несколько строк.", ""),
        ("Жёлтые ячейки заполняете вы. Зелёные считаются формулой — их трогать не нужно.", ""),
        ("Сумма всегда положительная. Направление задаётся колонкой «Тип».", ""),
        ("Тип, Проект, Статья ДДС и Статус выбираются из выпадающего списка.", ""),
        ("Вид деятельности и Приоритет подставляются автоматически по статье ДДС.", ""),
        ("Приоритет можно переписать вручную, если по конкретному платежу он другой.", ""),
        ("Даты — только в формате ДД.ММ.ГГГГ.", ""),
        ("", ""),
        ("Счета и кредитные лимиты", "h"),
        ("Счетов может быть сколько угодно — остаток на старте считается как сумма по всем.", ""),
        ("Целевые и заблокированные средства вносите отдельной строкой с пометкой в комментарии:", ""),
        ("в расчёт они попадут, но вы будете видеть, что свободно распоряжаться ими нельзя.", ""),
        ("Кредитные линии и овердрафты — это подушка на случай разрыва.", ""),
        ("Свободный лимит считается формулой: лимит минус выбрано.", ""),
        ("При разрыве расчёт покажет, покрывается ли он лимитом и во сколько обойдутся проценты.", ""),
        ("", ""),
        ("Свои статьи ДДС", "h"),
        ("Справочник статей открытый — дописывайте свои строки на листе «Справочники».", ""),
        ("Для каждой новой статьи заполните все четыре колонки:", ""),
        ("статья, тип по умолчанию, приоритет (1–3) и вид деятельности.", ""),
        ("Без них по этой статье не подтянется ни приоритет, ни вид деятельности.", ""),
        ("", ""),
        ("Проект", "h"),
        ("Привязка к объекту показывает, какой проект приносит деньги, а какой их потребляет.", ""),
        ("Что к объекту не относится — налоги, аренда офиса, административный ФОТ — «Общие».", ""),
        ("", ""),
        ("Вид деятельности", "h"),
        ("Операционная — основная деятельность: выручка, ФОТ, налоги, материалы, подрядчики.", ""),
        ("Инвестиционная — покупка и продажа основных средств.", ""),
        ("Финансовая — кредиты, лизинг, проценты, дивиденды.", ""),
        ("Разрез нужен, чтобы платёжный календарь сходился со структурой БДДС.", ""),
        ("", ""),
        ("Что означает приоритет", "h"),
        ("1 — критично: налоги, зарплата, взносы, погашение кредита. Просрочка = штрафы и пени.", ""),
        ("2 — важно: подрядчики, материалы, аренда, лизинг. Просрочка портит отношения.", ""),
        ("3 — можно перенести: командировки, прочие расходы. Переносятся первыми при нехватке денег.", ""),
        ("", ""),
        ("Статусы", "h"),
        ("План — сумма и дата предполагаются.", ""),
        ("Подтверждено — есть счёт, договор или заявка с точной датой.", ""),
        ("Оплачено — платёж уже прошёл, в прогноз не попадает.", ""),
        ("", ""),
        ("Правила, которые соблюдает расчёт", "h"),
        ("Исходный файл не изменяется — результат пишется в новые файлы с датой в имени.", ""),
        ("Ни одна цифра не додумывается: пустое поле остаётся пустым и попадает в список замечаний.", ""),
        ("Все расчёты в Excel — живыми формулами, значения не вбиваются руками.", ""),
    ]

    if demo:
        lines.insert(1, ("", ""))
        lines.insert(2, ("ВНИМАНИЕ: это демонстрационный файл. Все данные условные и придуманы "
                         "для показа работы инструмента. Для расчёта по реальной компании "
                         "используйте «Реестр_обязательств_Шаблон.xlsx».", "warn"))

    row = 2
    for text, kind in lines:
        c = ws.cell(row=row, column=2, value=text)
        if kind == "title":
            c.font = TITLE_FONT
        elif kind == "h":
            c.font = SECTION_FONT
        elif kind == "warn":
            c.font = Font(bold=True, color="C00000", size=10)
            c.fill = DEMO_FILL
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[row].height = 45
        row += 1


def build_params(ws, demo):
    ws.title = "Параметры"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 60

    ws.cell(row=2, column=2, value="Параметры расчёта").font = TITLE_FONT

    # name, value, format, note, is_formula
    params = [
        ("Дата начала прогноза", date(2026, 10, 19) if demo else None, DATE_FMT,
         "Понедельник первой недели горизонта", False),
        ("Остаток денежных средств на дату начала",
         "='Счета и лимиты'!$D${r}".format(r=ACC_TOTAL), MONEY,
         "Считается как сумма по всем счетам с листа «Счета и лимиты»", True),
        ("Свободные кредитные лимиты",
         "='Счета и лимиты'!$F${r}".format(r=CRD_TOTAL), MONEY,
         "Подушка на случай кассового разрыва, в остаток не входит", True),
        ("Горизонт прогноза, недель", 13, "0",
         "13 недель — стандарт казначейского прогноза (квартал)", False),
        ("Порог тревоги, ₽", 500000 if demo else None, MONEY,
         "Минимальный остаток, ниже которого включается предупреждение", False),
        ("Название компании", "Демо-компания (условные данные)" if demo else None, "@",
         "Выводится в заголовке отчёта и дашборда", False),
    ]

    row = 4
    for name, value, fmt, note, is_formula in params:
        ws.cell(row=row, column=2, value=name).font = Font(bold=True, size=10)
        c = ws.cell(row=row, column=3, value=value)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right" if fmt != "@" else "left")
        style_calc(c) if is_formula else style_input(c)
        ws.cell(row=row, column=4, value=note).font = Font(size=9, color="808080")
        row += 1

    ws.cell(row=row + 1, column=2,
            value="Жёлтые ячейки — для заполнения, зелёные считаются формулой."
            ).font = Font(size=9, color="808080")


def build_accounts(ws, accounts, credits):
    ws.title = "Счета и лимиты"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    for col, width in [(2, 24), (3, 32), (4, 18), (5, 18), (6, 18), (7, 12), (8, 15), (9, 34)]:
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.cell(row=2, column=2,
            value="Денежные средства и источники покрытия").font = TITLE_FONT

    # ---- остатки на счетах
    ws.cell(row=ACC_HEAD_ROW - 1, column=2,
            value="Остатки на счетах на дату начала прогноза").font = SECTION_FONT
    for col, name in enumerate(["Банк", "Счёт / назначение", "Остаток, ₽", "Комментарий"],
                               start=2):
        put_header(ws, ACC_HEAD_ROW, col, name)

    for i in range(ACC_FIRST, ACC_LAST + 1):
        idx = i - ACC_FIRST
        data = accounts[idx] if idx < len(accounts) else (None, None, None, None)
        for col, value in enumerate(data, start=2):
            cell = ws.cell(row=i, column=col, value=value)
            style_input(cell, MONEY if col == 4 else None)

    ws.cell(row=ACC_TOTAL, column=2, value="Итого денежных средств").font = Font(bold=True)
    total = ws.cell(row=ACC_TOTAL, column=4,
                    value="=SUM(D{a}:D{b})".format(a=ACC_FIRST, b=ACC_LAST))
    total.number_format = MONEY
    total.font = Font(bold=True)
    for col in range(2, 6):
        ws.cell(row=ACC_TOTAL, column=col).fill = TOTAL_FILL
        ws.cell(row=ACC_TOTAL, column=col).border = BORDER

    # ---- кредитные линии
    ws.cell(row=CRD_TITLE_ROW, column=2,
            value="Кредитные линии и овердрафты").font = SECTION_FONT
    for col, name in enumerate(["Банк", "Вид", "Лимит, ₽", "Выбрано, ₽",
                                "Свободно, ₽", "Ставка, %", "Действует до", "Комментарий"],
                               start=2):
        put_header(ws, CRD_HEAD_ROW, col, name)
    ws.row_dimensions[CRD_HEAD_ROW].height = 28

    for i in range(CRD_FIRST, CRD_LAST + 1):
        idx = i - CRD_FIRST
        if idx < len(credits):
            bank, kind, limit, used, rate, until, note = credits[idx]
        else:
            bank = kind = limit = used = rate = until = note = None

        style_input(ws.cell(row=i, column=2, value=bank))
        style_input(ws.cell(row=i, column=3, value=kind))
        style_input(ws.cell(row=i, column=4, value=limit), MONEY)
        style_input(ws.cell(row=i, column=5, value=used), MONEY)
        # свободный лимит — формула
        free = ws.cell(row=i, column=6,
                       value='=IF(D{r}="","",D{r}-N(E{r}))'.format(r=i))
        style_calc(free, MONEY)
        style_input(ws.cell(row=i, column=7, value=rate), PERCENT)
        style_input(ws.cell(row=i, column=8, value=until), DATE_FMT)
        style_input(ws.cell(row=i, column=9, value=note))

    ws.cell(row=CRD_TOTAL, column=2, value="Итого свободный лимит").font = Font(bold=True)
    tot_free = ws.cell(row=CRD_TOTAL, column=6,
                       value="=SUM(F{a}:F{b})".format(a=CRD_FIRST, b=CRD_LAST))
    tot_free.number_format = MONEY
    tot_free.font = Font(bold=True)
    for col in range(2, 10):
        ws.cell(row=CRD_TOTAL, column=col).fill = TOTAL_FILL
        ws.cell(row=CRD_TOTAL, column=col).border = BORDER

    note_row = CRD_TOTAL + 2
    ws.cell(row=note_row, column=2,
            value="Свободный лимит не входит в остаток денежных средств. "
                  "Это источник покрытия разрыва, а не деньги компании."
            ).font = Font(size=9, color="808080")

    dv_kind = DataValidation(type="list",
                             formula1='"{v}"'.format(v=",".join(CREDIT_KINDS)),
                             allow_blank=True)
    ws.add_data_validation(dv_kind)
    dv_kind.add("C{a}:C{b}".format(a=CRD_FIRST, b=CRD_LAST))

    dv_money = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                              allow_blank=True, showErrorMessage=True,
                              errorTitle="Некорректная сумма",
                              error="Сумма не может быть отрицательной.")
    ws.add_data_validation(dv_money)
    dv_money.add("D{a}:D{b}".format(a=ACC_FIRST, b=ACC_LAST))
    dv_money.add("D{a}:E{b}".format(a=CRD_FIRST, b=CRD_LAST))


def build_recurring(ws, rules):
    """Правила повторяющихся платежей: одна строка разворачивается в серию дат."""
    ws.title = "Регулярные платежи"
    ws.freeze_panes = "A{r}".format(r=RECUR_FIRST)
    ws.sheet_view.showGridLines = False

    ws.cell(row=2, column=1,
            value="Повторяющиеся платежи — описываются правилом, а не строками").font = TITLE_FONT

    for i, (name, width) in enumerate(RECUR_COLS, start=1):
        put_header(ws, RECUR_HEAD, i, name, width)
    ws.row_dimensions[RECUR_HEAD].height = 30

    for r in range(RECUR_FIRST, RECUR_LAST + 1):
        idx = r - RECUR_FIRST
        data = rules[idx] if idx < len(rules) else (None,) * len(RECUR_COLS)
        for col, value in enumerate(data, start=1):
            cell = ws.cell(row=r, column=col, value=value)
            fmt = None
            if col == 6:
                fmt = MONEY
            elif col in (9, 10):
                fmt = DATE_FMT
            elif col == 8:
                fmt = "0"
            style_input(cell, fmt)

    note = RECUR_LAST + 2
    hints = [
        "Одна строка = правило. Генератор разворачивает его в платежи на весь горизонт.",
        "«День месяца» — число от 1 до 31. Для периодичности «Еженедельно» — день недели (1 = понедельник).",
        "«Перенос с выходного»: Раньше — предыдущий рабочий день (зарплата, аренда); "
        "Позже — следующий рабочий день (налоги, проценты).",
        "В «Основании» можно поставить {месяц} — подставится название месяца платежа.",
        "Платежи с меняющейся суммой (налоги, зарплата с премией) вносите в «Реестр» строками.",
        "Нерабочие дни правятся на листе «Справочники».",
    ]
    for i, text in enumerate(hints):
        ws.cell(row=note + i, column=1, value=text).font = Font(size=9, color="808080")

    dv_yes = DataValidation(type="list", formula1='"{v}"'.format(v=",".join(YES_NO)),
                            allow_blank=True)
    dv_freq = DataValidation(type="list", formula1='"{v}"'.format(v=",".join(FREQUENCIES)),
                             allow_blank=True)
    dv_shift = DataValidation(type="list", formula1='"{v}"'.format(v=",".join(SHIFT_RULES)),
                              allow_blank=True)
    dv_proj = DataValidation(type="list",
                             formula1="=Справочники!$P$2:$P${n}".format(n=PROJ_LAST),
                             allow_blank=True)
    dv_art = DataValidation(type="list",
                            formula1="=Справочники!$A$2:$A${n}".format(n=ART_LAST),
                            allow_blank=True)
    dv_day = DataValidation(type="whole", operator="between", formula1="1", formula2="31",
                            allow_blank=True, showErrorMessage=True,
                            errorTitle="День месяца",
                            error="Введите число от 1 до 31.")
    for dv, col in [(dv_yes, "A"), (dv_proj, "B"), (dv_art, "E"),
                    (dv_freq, "G"), (dv_day, "H"), (dv_shift, "K")]:
        ws.add_data_validation(dv)
        dv.add("{c}{a}:{c}{b}".format(c=col, a=RECUR_FIRST, b=RECUR_LAST))


def build_reference(ws, projects):
    ws.title = "Справочники"
    ws.sheet_view.showGridLines = False

    # ---- открытый справочник статей: A..D
    for col, (name, width) in enumerate(
            [("Статья ДДС", 28), ("Тип по умолчанию", 18),
             ("Приоритет", 11), ("Вид деятельности", 18)], start=1):
        put_header(ws, 1, col, name, width)

    for i in range(2, ART_LAST + 1):
        idx = i - 2
        row_data = ARTICLES[idx] if idx < len(ARTICLES) else (None, None, None, None)
        article, typ, prio, activity = row_data
        style_input(ws.cell(row=i, column=1, value=article))
        style_input(ws.cell(row=i, column=2, value=typ))
        style_input(ws.cell(row=i, column=3, value=prio if prio else None), "0")
        style_input(ws.cell(row=i, column=4, value=activity))

    # справочные блоки только для чтения
    for col, (name, values, width) in enumerate_blocks():
        put_header(ws, 1, col, name, width)
        for i, v in enumerate(values, start=2):
            c = ws.cell(row=i, column=col, value=v)
            c.border = BORDER

    # ---- открытый список проектов: P
    put_header(ws, 1, 16, "Проект", 30)
    for i in range(2, PROJ_LAST + 1):
        idx = i - 2
        value = projects[idx] if idx < len(projects) else None
        style_input(ws.cell(row=i, column=16, value=value))

    ws.cell(row=PROJ_LAST + 2, column=16,
            value="Впишите свои объекты").font = Font(size=9, color="808080")

    # ---- нерабочие дни: R
    put_header(ws, 1, 18, "Нерабочие дни", 16)
    for i in range(2, HOL_LAST + 1):
        idx = i - 2
        value = HOLIDAYS[idx] if idx < len(HOLIDAYS) else None
        style_input(ws.cell(row=i, column=18, value=value), DATE_FMT)
    ws.cell(row=HOL_LAST + 2, column=18,
            value="Праздники и перенесённые выходные.").font = Font(size=9, color="808080")
    ws.cell(row=HOL_LAST + 3, column=18,
            value="Субботы и воскресенья учитываются автоматически."
            ).font = Font(size=9, color="808080")
    ws.cell(row=HOL_LAST + 4, column=18,
            value="Переносы на 2027 год — проверьте по производственному календарю."
            ).font = Font(size=9, color="C00000")
    ws.cell(row=ART_LAST + 2, column=1,
            value="Справочник открытый: дописывайте свои статьи, "
                  "заполняя все четыре колонки").font = Font(size=9, color="808080")

    # проверки внутри самого справочника — чтобы новые статьи заполнялись корректно
    dv_t = DataValidation(type="list", formula1='"{v}"'.format(v=",".join(TYPES)),
                          allow_blank=True)
    dv_a = DataValidation(type="list", formula1='"{v}"'.format(v=",".join(ACTIVITIES)),
                          allow_blank=True)
    dv_p = DataValidation(type="whole", operator="between", formula1="1", formula2="3",
                          allow_blank=True, showErrorMessage=True,
                          errorTitle="Приоритет",
                          error="Приоритет — целое число от 1 до 3. "
                                "Для поступлений оставьте пусто.")
    for dv, col in [(dv_t, "B"), (dv_p, "C"), (dv_a, "D")]:
        ws.add_data_validation(dv)
        dv.add("{c}2:{c}{n}".format(c=col, n=ART_LAST))


def enumerate_blocks():
    """Справочные блоки-подсказки: (колонка, (заголовок, значения, ширина))."""
    return [
        (6, ("Тип", TYPES, 16)),
        (8, ("Статус", STATUSES, 16)),
        (10, ("Приоритет", [p[0] for p in PRIORITY_MEANING], 11)),
        (11, ("Что означает", [p[1] for p in PRIORITY_MEANING], 56)),
        (13, ("Вид деятельности", [a[0] for a in ACTIVITY_MEANING], 18)),
        (14, ("Что входит", [a[1] for a in ACTIVITY_MEANING], 62)),
    ]


def build_registry(ws, rows):
    ws.title = "Реестр"
    ws.freeze_panes = "A2"

    for i, (name, width) in enumerate(REGISTRY_COLS, start=1):
        put_header(ws, 1, i, name, width)
    ws.row_dimensions[1].height = 30

    art_range = "Справочники!$A$2:$D${n}".format(n=ART_LAST)

    for r in range(2, MAX_ROWS + 2):
        num = ws.cell(row=r, column=COL_NUM)
        num.value = '=IF(D{r}="","",ROW()-1)'.format(r=r)
        num.number_format = "0"

        act = ws.cell(row=r, column=COL_ACTIVITY)
        act.value = ('=IF(F{r}="","",IFERROR(VLOOKUP(F{r},{rng},4,FALSE),"?"))'
                     .format(r=r, rng=art_range))

        prio = ws.cell(row=r, column=COL_PRIORITY)
        prio.value = ('=IF(F{r}="","",IFERROR(IF(N(VLOOKUP(F{r},{rng},3,FALSE))=0,"",'
                      'VLOOKUP(F{r},{rng},3,FALSE)),"?"))'.format(r=r, rng=art_range))
        prio.number_format = "0"

        ws.cell(row=r, column=8).number_format = MONEY
        ws.cell(row=r, column=9).number_format = DATE_FMT

        for col in range(1, len(REGISTRY_COLS) + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDER
            if col in CALC_COLS:
                cell.fill = CALC_FILL
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.fill = INPUT_FILL

    dv_type = DataValidation(type="list", formula1="=Справочники!$F$2:$F$3", allow_blank=True)
    dv_proj = DataValidation(type="list",
                             formula1="=Справочники!$P$2:$P${n}".format(n=PROJ_LAST),
                             allow_blank=True)
    dv_art = DataValidation(type="list",
                            formula1="=Справочники!$A$2:$A${n}".format(n=ART_LAST),
                            allow_blank=True)
    dv_status = DataValidation(type="list", formula1="=Справочники!$H$2:$H$4", allow_blank=True)
    dv_sum = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                            allow_blank=True, showErrorMessage=True,
                            errorTitle="Некорректная сумма",
                            error="Сумма должна быть положительной. "
                                  "Направление задаётся колонкой «Тип».")
    for dv, col in [(dv_type, "B"), (dv_proj, "C"), (dv_art, "F"),
                    (dv_sum, "H"), (dv_status, "K")]:
        ws.add_data_validation(dv)
        dv.add("{c}2:{c}{n}".format(c=col, n=MAX_ROWS + 1))

    for r, item in enumerate(rows, start=2):
        typ, project, contractor, basis, article, amount, due, status, comment = item
        ws.cell(row=r, column=2, value=typ)
        ws.cell(row=r, column=3, value=project)
        ws.cell(row=r, column=4, value=contractor)
        ws.cell(row=r, column=5, value=basis)
        ws.cell(row=r, column=6, value=article)
        ws.cell(row=r, column=8, value=amount).number_format = MONEY
        ws.cell(row=r, column=9, value=due).number_format = DATE_FMT
        ws.cell(row=r, column=11, value=status)
        ws.cell(row=r, column=12, value=comment)

    ws.auto_filter.ref = "A1:L{n}".format(n=MAX_ROWS + 1)


def make_file(path, rows, demo, projects, accounts, credits, rules):
    wb = Workbook()
    build_instruction(wb.active, demo)
    build_params(wb.create_sheet(), demo)
    build_accounts(wb.create_sheet(), accounts, credits)
    build_recurring(wb.create_sheet(), rules)
    build_registry(wb.create_sheet(), rows)
    build_reference(wb.create_sheet(), projects)

    # именованные диапазоны — чтобы на них можно было ссылаться в расчётах
    wb.defined_names.add(DefinedName(
        "ОстатокНаСтарте",
        attr_text="'Счета и лимиты'!$D${r}".format(r=ACC_TOTAL)))
    wb.defined_names.add(DefinedName(
        "СвободныйЛимит",
        attr_text="'Счета и лимиты'!$F${r}".format(r=CRD_TOTAL)))

    wb.save(path)
    return path


# ---------------------------------------------------------------- демо-данные

def demo_rows():
    """Условные данные подрядной компании. Все контрагенты обезличены."""
    def d(day, month, year=2026):
        return date(year, month, day)

    P, F = "Поступление", "Платёж"
    OB1, OB2, OB3, GEN = DEMO_PROJECTS
    return [
        # поступления
        (P, OB1, "Заказчик А", "Договор 14/2026, окончательный расчёт",
         "Поступления от заказчиков", 1800000, d(23, 10), "Подтверждено", "Акт подписан"),
        (P, OB2, "Заказчик Б", "Договор 21/2026, аванс 30%",
         "Поступления от заказчиков", 4200000, d(6, 11), "Подтверждено",
         "Аванс по подписанному договору"),
        (P, OB3, "Заказчик В", "Договор 19/2026, этап 1",
         "Поступления от заказчиков", 3200000, d(20, 11), "План",
         "Срок зависит от подписания КС-2"),
        (P, OB3, "Заказчик Д", "Договор 24/2026, аванс 40%",
         "Поступления от заказчиков", 3600000, d(2, 12), "План",
         "Договор подписан, ждём оплату"),
        (P, OB2, "Заказчик Б", "Договор 21/2026, этап 2",
         "Поступления от заказчиков", 5200000, d(11, 12), "План", ""),
        (P, OB3, "Заказчик В", "Договор 19/2026, этап 2",
         "Поступления от заказчиков", 4100000, d(16, 12), "План", ""),
        (P, OB1, "Заказчик Г", "Договор 23/2026, аванс",
         "Поступления от заказчиков", 2800000, d(25, 12), "План", "Договор на согласовании"),
        (P, OB1, "Заказчик Г", "Договор 23/2026, этап 1",
         "Поступления от заказчиков", 2300000, d(5, 1, 2027), "План", ""),
        (P, OB1, "Заказчик А", "Доп. соглашение 2",
         "Поступления от заказчиков", 650000, d(16, 1, 2027), "План", "Дополнительные работы"),

        # ФОТ: авансы идут правилом на листе «Регулярные платежи»,
        # здесь только зарплата — её сумма меняется от месяца к месяцу
        (F, GEN, "Персонал", "Зарплата за октябрь", "ФОТ", 1500000, d(5, 11), "Подтверждено", ""),
        (F, GEN, "Персонал", "Зарплата за ноябрь", "ФОТ", 1500000, d(4, 12), "План", ""),
        (F, GEN, "Персонал", "Зарплата за декабрь", "ФОТ", 1620000, d(12, 1, 2027), "План",
         "Выше обычного: годовая премия"),

        # налоги и взносы
        (F, GEN, "ФНС", "НДС за III квартал, 1/3", "Налоги", 1150000, d(28, 10),
         "Подтверждено", ""),
        (F, GEN, "ФНС", "Налог на прибыль за 9 месяцев", "Налоги", 480000, d(28, 10),
         "Подтверждено", ""),
        (F, GEN, "ФНС", "НДС за III квартал, 2/3", "Налоги", 1150000, d(28, 11), "План", ""),
        (F, GEN, "ФНС", "НДС за III квартал, 3/3", "Налоги", 1150000, d(28, 12), "План", ""),
        (F, GEN, "ФНС", "Страховые взносы за октябрь", "Страховые взносы", 870000, d(28, 10),
         "Подтверждено", ""),
        (F, GEN, "ФНС", "Страховые взносы за ноябрь", "Страховые взносы", 870000, d(28, 11),
         "План", ""),
        (F, GEN, "ФНС", "Страховые взносы за декабрь", "Страховые взносы", 890000, d(28, 12),
         "План", ""),

        # подрядчики
        (F, OB2, "Подрядчик 1", "Монтажные работы, этап 1", "Подрядчики", 1850000, d(12, 11),
         "Подтверждено", "Счёт выставлен"),
        (F, OB2, "Подрядчик 1", "Монтажные работы, этап 2", "Подрядчики", 1600000, d(26, 11),
         "План", ""),
        (F, OB3, "Подрядчик 2", "Демонтаж и подготовка", "Подрядчики", 740000, d(3, 12),
         "План", ""),
        (F, OB2, "Подрядчик 1", "Монтажные работы, этап 3", "Подрядчики", 1900000, d(22, 12),
         "План", ""),

        # материалы
        (F, OB1, "Поставщик 1", "Стеклопакеты, партия 1", "Материалы", 900000, d(29, 10),
         "Подтверждено", ""),
        (F, OB2, "Поставщик 2", "Алюминиевый профиль", "Материалы", 900000, d(19, 11), "План", ""),
        (F, OB2, "Поставщик 1", "Стеклопакеты, партия 2", "Материалы", 1340000, d(9, 12),
         "План", ""),
        (F, OB3, "Поставщик 3", "Фурнитура и крепёж", "Материалы", 380000, d(23, 12), "План", ""),

        # аренда, лизинг, связь и проценты по кредиту — правилами
        # на листе «Регулярные платежи»

        (F, GEN, "Банк 2", "Погашение части кредитной линии", "Погашение кредита",
         500000, d(15, 1, 2027), "План", "По графику продления"),

        # прочее — кандидаты на перенос
        (F, OB3, "Прочие", "Командировка на объект", "Командировки", 140000, d(13, 11), "План",
         "Можно перенести"),
        (F, GEN, "Прочие", "Обучение сотрудников", "Прочие расходы", 210000, d(4, 12), "План",
         "Можно перенести"),
    ]


def main():
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(base, "01_input")
    os.makedirs(out, exist_ok=True)

    tpl = make_file(os.path.join(out, "Реестр_обязательств_Шаблон.xlsx"), [],
                    demo=False, projects=EMPTY_PROJECTS, accounts=[], credits=[], rules=[])
    dem = make_file(os.path.join(out, "Реестр_обязательств_ДЕМО.xlsx"), demo_rows(),
                    demo=True, projects=DEMO_PROJECTS,
                    accounts=DEMO_ACCOUNTS, credits=DEMO_CREDITS, rules=DEMO_RECURRING)

    print("Создано:")
    print("  " + os.path.basename(tpl) + " - пустая форма")
    print("  " + os.path.basename(dem) + " - {n} разовых обязательств + {k} правил повтора"
          .format(n=len(demo_rows()), k=len(DEMO_RECURRING)))


if __name__ == "__main__":
    main()
