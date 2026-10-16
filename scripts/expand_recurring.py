# -*- coding: utf-8 -*-
"""
Разворачивает правила повторяющихся платежей в отдельные обязательства.

Читает лист «Регулярные платежи» и возвращает список строк того же формата,
что и лист «Реестр». Исходный файл не изменяется — только чтение.

Правила переноса с нерабочего дня:
    Раньше        — предыдущий рабочий день (зарплата, аренда, лизинг)
    Позже         — следующий рабочий день (налоги, взносы, проценты)
    Не переносить — дата остаётся как есть

Нерабочие дни: суббота, воскресенье и список праздников
с листа «Справочники» (колонка «Нерабочие дни»).

Запуск напрямую — показывает, во что развернутся правила:
    python scripts/expand_recurring.py 01_input/Реестр_обязательств_ДЕМО.xlsx
"""

import calendar
import sys
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

MONTHS_GEN = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]

SHEET_RECUR = "Регулярные платежи"
SHEET_REF = "Справочники"
SHEET_PARAMS = "Параметры"

RECUR_FIRST_ROW = 5      # первая строка данных на листе правил
RECUR_LAST_ROW = 44      # последняя строка данных; ниже идут подсказки
HOLIDAY_COL = 18         # колонка «Нерабочие дни» на листе «Справочники»


class Problem(object):
    """Замечание к данным — не исключение, а то, что нужно показать пользователю."""

    def __init__(self, row, text):
        self.row = row
        self.text = text

    def __str__(self):
        return "строка {r}: {t}".format(r=self.row, t=self.text)


# ---------------------------------------------------------------- календарь

def read_holidays(wb):
    """Список нерабочих дней с листа «Справочники»."""
    if SHEET_REF not in wb.sheetnames:
        return set()
    ws = wb[SHEET_REF]
    out = set()
    for row in ws.iter_rows(min_row=2, min_col=HOLIDAY_COL, max_col=HOLIDAY_COL):
        value = as_date(row[0].value)
        if value is not None:
            out.add(value)
    return out


def is_working_day(day, holidays):
    return day.weekday() < 5 and day not in holidays


def shift_to_working(day, holidays, rule):
    """Перенос даты с нерабочего дня по правилу."""
    if rule == "Не переносить" or is_working_day(day, holidays):
        return day, False
    step = -1 if rule == "Раньше" else 1
    moved = day
    for _ in range(30):
        moved += timedelta(days=step)
        if is_working_day(moved, holidays):
            return moved, True
    return day, False


# ---------------------------------------------------------------- даты серии

def month_dates(start, end, day_of_month, step_months):
    """Даты по числу месяца от start до end с шагом step_months."""
    out = []
    year, month = start.year, start.month
    while True:
        last_day = calendar.monthrange(year, month)[1]
        current = date(year, month, min(day_of_month, last_day))
        if current > end:
            break
        if current >= start:
            out.append(current)
        month += step_months
        while month > 12:
            month -= 12
            year += 1
        if year > end.year + 1:
            break
    return out


def week_dates(start, end, weekday, step_days):
    """Даты по дню недели от start до end с шагом step_days."""
    out = []
    current = start
    while current.weekday() != (weekday - 1) % 7:
        current += timedelta(days=1)
        if current > end:
            return out
    while current <= end:
        out.append(current)
        current += timedelta(days=step_days)
    return out


def series_dates(freq, day, period_from, period_to):
    if freq == "Ежемесячно":
        return month_dates(period_from, period_to, day, 1)
    if freq == "Ежеквартально":
        return month_dates(period_from, period_to, day, 3)
    if freq == "Еженедельно":
        return week_dates(period_from, period_to, day, 7)
    if freq == "Раз в 2 недели":
        return week_dates(period_from, period_to, day, 14)
    return []


# ---------------------------------------------------------------- разворот правил

def as_date(value):
    """openpyxl отдаёт даты как datetime — приводим всё к date."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    return value if isinstance(value, date) else None


def expand(path, horizon_from=None, horizon_to=None):
    """
    Возвращает (rows, problems).

    rows — обязательства в формате листа «Реестр»:
        (тип, проект, контрагент, основание, статья, сумма, дата, статус, комментарий)
    problems — замечания к правилам, которые нельзя развернуть.
    """
    wb = load_workbook(path, data_only=True)
    if SHEET_RECUR not in wb.sheetnames:
        return [], [Problem(0, "лист «{s}» не найден".format(s=SHEET_RECUR))]

    holidays = read_holidays(wb)
    ws = wb[SHEET_RECUR]

    rows, problems = [], []

    for r in range(RECUR_FIRST_ROW, min(ws.max_row, RECUR_LAST_ROW) + 1):
        (enabled, project, contractor, basis, article, amount,
         freq, day, valid_from, valid_to, shift, comment) = [
            ws.cell(row=r, column=c).value for c in range(1, 13)
        ]

        # пустая строка — не правило
        if not any([enabled, contractor, article, amount]):
            continue
        if str(enabled).strip().lower() in ("нет", "no", "0", "false"):
            continue

        missing = [name for name, value in [
            ("контрагент", contractor), ("статья ДДС", article),
            ("сумма", amount), ("периодичность", freq), ("день", day),
        ] if value in (None, "")]
        if missing:
            problems.append(Problem(r, "не заполнено: " + ", ".join(missing)))
            continue

        if not isinstance(amount, (int, float)) or amount <= 0:
            problems.append(Problem(r, "сумма должна быть положительным числом"))
            continue

        valid_from = as_date(valid_from)
        valid_to = as_date(valid_to)
        period_from = max(x for x in [valid_from, horizon_from] if x is not None) \
            if (valid_from or horizon_from) else None
        period_to = min(x for x in [valid_to, horizon_to] if x is not None) \
            if (valid_to or horizon_to) else None

        if period_from is None or period_to is None:
            problems.append(Problem(r, "не заданы границы периода действия"))
            continue
        if period_from > period_to:
            continue  # правило не действует внутри горизонта — это не ошибка

        try:
            day = int(day)
        except (TypeError, ValueError):
            problems.append(Problem(r, "день должен быть числом"))
            continue

        dates = series_dates(freq, day, period_from, period_to)
        if not dates:
            problems.append(Problem(r, "периодичность «{f}» не распознана".format(f=freq)))
            continue

        shift = shift or "Не переносить"
        for due in dates:
            moved_date, was_moved = shift_to_working(due, holidays, shift)
            text = (basis or "").replace("{месяц}", MONTHS_GEN[due.month - 1])
            note = comment or ""
            if was_moved:
                mark = "перенос с {d} ({r})".format(
                    d=due.strftime("%d.%m.%Y"), r=shift.lower())
                note = (note + "; " + mark) if note else mark
            rows.append((
                "Платёж", project, contractor, text, article,
                float(amount), moved_date, "План", note,
            ))

    rows.sort(key=lambda x: (x[6], x[2]))
    return rows, problems


def read_horizon(path):
    """Дата начала и конца горизонта с листа «Параметры»."""
    wb = load_workbook(path, data_only=True)
    if SHEET_PARAMS not in wb.sheetnames:
        return None, None
    ws = wb[SHEET_PARAMS]
    start = weeks = None
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=2).value
        value = ws.cell(row=r, column=3).value
        if not name:
            continue
        if "Дата начала" in str(name):
            start = as_date(value)
        elif "Горизонт" in str(name):
            weeks = value
    if start is None:
        return None, None
    weeks = int(weeks) if isinstance(weeks, (int, float)) else 13
    return start, start + timedelta(weeks=weeks) - timedelta(days=1)


def main():
    if len(sys.argv) < 2:
        print("Укажите файл реестра:")
        print("  python scripts/expand_recurring.py 01_input/Реестр_обязательств_ДЕМО.xlsx")
        return 1

    path = sys.argv[1]
    start, end = read_horizon(path)
    if start is None:
        print("На листе «Параметры» не заполнена дата начала прогноза — данных нет.")
        return 1

    rows, problems = expand(path, start, end)

    print("Горизонт: {a} - {b}".format(
        a=start.strftime("%d.%m.%Y"), b=end.strftime("%d.%m.%Y")))
    print("Развёрнуто обязательств: {n}".format(n=len(rows)))
    print()

    if rows:
        print("{:<12} {:<22} {:<34} {:>14}  {}".format(
            "Дата", "Контрагент", "Основание", "Сумма", "Примечание"))
        for typ, project, contractor, basis, article, amount, due, status, note in rows:
            print("{:<12} {:<22} {:<34} {:>14}  {}".format(
                due.strftime("%d.%m.%Y"), contractor[:22], basis[:34],
                format(amount, ",.2f").replace(",", " "), note))
        print()
        print("Итого: {s} ₽".format(
            s=format(sum(r[5] for r in rows), ",.2f").replace(",", " ")))

    if problems:
        print()
        print("Замечания по правилам:")
        for p in problems:
            print("  " + str(p))

    return 0


if __name__ == "__main__":
    sys.exit(main())
