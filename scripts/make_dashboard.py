# -*- coding: utf-8 -*-
"""
Собирает интерактивный дашборд платёжного календаря.

Берёт тот же расчёт, что и Excel-версия, и встраивает данные в HTML.
Дашборд работает без интернета и без сервера: график рисуется на SVG,
пересчёт при изменении условий идёт в браузере.

Запуск:
    python scripts/make_dashboard.py 01_input/Реестр_обязательств_ДЕМО.xlsx
    python scripts/make_dashboard.py <файл> --anon      # обезличить контрагентов
"""

import json
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_calendar import build  # noqa: E402

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "dashboard_template.html")


def anonymize(rows):
    """Заменяет названия контрагентов на обезличенные — для публикации."""
    mapping, counters = {}, {}
    keep = ("ФНС", "Персонал", "Прочие")
    for row in rows:
        name = row["contractor"]
        if not name or name in keep or name in mapping:
            continue
        kind = "Контрагент"
        for word in ("Заказчик", "Подрядчик", "Поставщик", "Банк", "Арендодатель",
                     "Лизинговая", "Сервис"):
            if name.startswith(word):
                kind = word if word != "Лизинговая" else "Лизингодатель"
                break
        counters[kind] = counters.get(kind, 0) + 1
        mapping[name] = "{k} {n}".format(k=kind, n=counters[kind])
    for row in rows:
        if row["contractor"] in mapping:
            row["contractor"] = mapping[row["contractor"]]
    return rows


def collect(path, anon=False):
    out_path, params, weeks_data, gaps, control, notes = build(path)

    start = params["start"]
    rows = []
    # пересобираем строки заново: нужны номера недель и приоритеты
    from build_calendar import (read_articles, read_registry, valid_rows,
                                week_bounds)
    from expand_recurring import expand

    articles = read_articles(path)
    end = weeks_data[-1]["to"]
    all_rows = read_registry(path) + expand(path, start, end)[0]
    good = valid_rows(all_rows, start, end)
    bounds = week_bounds(start, params["weeks"])

    for i, row in enumerate(good, start=1):
        typ, project, contractor, basis, article, amount, due, status, comment = row
        week = next((n for n, (a, b) in enumerate(bounds, start=1) if a <= due <= b), None)
        if week is None:
            continue
        prio = articles.get(str(article), (None, None))[0]
        rows.append({
            "id": i,
            "type": typ,
            "project": project or "Общие",
            "contractor": str(contractor or ""),
            "basis": str(basis or ""),
            "article": str(article or ""),
            "amount": float(amount),
            "date": due.isoformat(),
            "week": week,
            "priority": int(prio) if isinstance(prio, (int, float)) and prio else 0,
        })

    if anon:
        rows = anonymize(rows)

    data = {
        "company": params["company"] if not anon else "Демо-компания (условные данные)",
        "calcDate": date.today().isoformat(),
        "balance": params["balance"],
        "creditFree": params["credit_free"],
        "threshold": params["threshold"],
        "source": os.path.basename(path) if not anon else "обезличенные демо-данные",
        "weeks": [{"n": w["week"], "from": w["from"].isoformat(),
                   "to": w["to"].isoformat()} for w in weeks_data],
        "rows": rows,
    }
    return data, out_path, control, gaps


def main():
    args = sys.argv[1:]
    if not args:
        print("Укажите файл реестра:")
        print("  python scripts/make_dashboard.py 01_input/Реестр_обязательств_ДЕМО.xlsx")
        return 1

    path = args[0]
    anon = "--anon" in args

    data, excel_path, control, gaps = collect(path, anon)

    with open(TEMPLATE, "r", encoding="utf-8") as fh:
        html = fh.read()

    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # защита от преждевременного закрытия тега script внутри JSON
    payload = payload.replace("</", "<\\/")
    html = html.replace("__DATA__", payload)

    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(base, "04_dashboard")
    os.makedirs(out_dir, exist_ok=True)
    name = "Платежный_календарь_Дашборд{a}.html".format(a="_demo" if anon else "")
    out_path = os.path.join(out_dir, name)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(html)

    f = lambda x: format(x, ",.0f").replace(",", " ")
    print("Дашборд собран.")
    print("  Обязательств:        {n}".format(n=len(data["rows"])))
    print("  Недель:              {n}".format(n=len(data["weeks"])))
    print("  Остаток на старте:   {s} ₽".format(s=f(data["balance"])))
    print("  Свободный лимит:     {s} ₽".format(s=f(data["creditFree"])))
    print("  Минимальный остаток: {s} ₽".format(s=f(control["min_balance"])))
    print("  Проблемных недель:   {n}".format(n=len(gaps)))
    if anon:
        print("  Контрагенты обезличены.")
    print()
    print("  04_dashboard/" + name)
    print("  02_calc/" + os.path.basename(excel_path))

    # проверка, что плейсхолдер заменён и данные внутри корректны
    if "__DATA__" in html:
        print("ОШИБКА: данные не подставлены в шаблон.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
