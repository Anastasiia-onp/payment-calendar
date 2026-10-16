# -*- coding: utf-8 -*-
"""
Формирует выводы по платёжному календарю: факт → причина → решение.

Все цифры берутся из расчёта. Ничего не додумывается: если данных нет,
так и пишется. Формулировки рассчитаны на руководителя, а не на финансиста.

Запуск:
    python scripts/make_conclusions.py 01_input/Реестр_обязательств_ДЕМО.xlsx
"""

import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_calendar import build, read_articles, read_registry, valid_rows, week_bounds  # noqa: E402
from expand_recurring import expand  # noqa: E402


def money(v):
    return format(round(v), ",.0f").replace(",", " ") + " ₽"


def dmy(d):
    return d.strftime("%d.%m.%Y")


def week_drivers(rows, w_from, w_to, limit=3):
    """Самые крупные платежи недели — объясняют, откуда дефицит."""
    pays = [r for r in rows if r[0] == "Платёж" and w_from <= r[6] <= w_to]
    pays.sort(key=lambda r: -r[5])
    return pays[:limit]


def build_text(path):
    out_path, params, weeks, gaps, control, notes = build(path)

    start = params["start"]
    end = weeks[-1]["to"]
    articles = read_articles(path)
    rows = valid_rows(read_registry(path) + expand(path, start, end)[0], start, end)
    bounds = week_bounds(start, params["weeks"])

    real_gaps = [g for g in gaps if g["kind"] == "разрыв"]
    low = [g for g in gaps if g["kind"] != "разрыв"]
    min_week = min(weeks, key=lambda w: w["closing"])
    deficit = abs(min_week["closing"]) if min_week["closing"] < 0 else 0

    movable_total = 0.0
    for g in real_gaps:
        movable_total += g["movable_sum"]

    L = []
    add = L.append

    add("# Платёжный календарь: выводы")
    add("")
    add("{c}  ".format(c=params["company"] or "Компания не указана"))
    add("Горизонт: {a} — {b} ({w} недель)  ".format(a=dmy(start), b=dmy(end), w=params["weeks"]))
    add("Расчёт от {d}".format(d=dmy(date.today())))
    add("")
    add("---")
    add("")

    # ---------------- главный вывод
    add("## Главное")
    add("")
    if not real_gaps:
        add("**Кассовых разрывов в горизонте нет.** Минимальный остаток — {m} "
            "на неделе {n} ({d}).".format(
                m=money(min_week["closing"]), n=min_week["week"], d=dmy(min_week["from"])))
        if low:
            add("")
            add("Но в {n} нед. остаток опускается ниже порога тревоги {t}. "
                "Запас на непредвиденные платежи минимальный.".format(
                    n=len(low), t=money(params["threshold"])))
    else:
        add("**Денег не хватит в {n} нед. из {total}.** Первый разрыв — {d}, "
            "неделя {w}.".format(n=len(real_gaps), total=params["weeks"],
                                 d=dmy(real_gaps[0]["from"]), w=real_gaps[0]["week"]))
        add("")
        add("Максимальная нехватка — {m} на неделе {n} ({d}).".format(
            m=money(deficit), n=min_week["week"], d=dmy(min_week["from"])))
        add("")
        if params["credit_free"] >= deficit:
            add("Свободный кредитный лимит {c} разрыв покрывает. "
                "Это рабочий вариант, но платный.".format(c=money(params["credit_free"])))
        elif params["credit_free"] > 0:
            add("Свободного лимита {c} на покрытие не хватает: нужно ещё {n}.".format(
                c=money(params["credit_free"]), n=money(deficit - params["credit_free"])))
        else:
            add("Свободных кредитных лимитов нет — закрывать разрыв нечем, "
                "кроме переноса платежей и ускорения поступлений.")
    add("")

    # ---------------- картина за горизонт
    add("## Что происходит с деньгами")
    add("")
    add("| Показатель | Сумма |")
    add("|---|---|")
    add("| Остаток на старте | {s} |".format(s=money(params["balance"])))
    add("| Поступления за горизонт | {s} |".format(s=money(control["income"])))
    add("| Платежи за горизонт | {s} |".format(s=money(control["outcome"])))
    add("| Чистый денежный поток | {s} |".format(s=money(control["net"])))
    add("| Остаток на конец горизонта | {s} |".format(s=money(control["closing"])))
    add("| Минимальный остаток | {s} |".format(s=money(control["min_balance"])))
    add("| Свободные кредитные лимиты | {s} |".format(s=money(params["credit_free"])))
    add("")

    # отклонение меньше 2% оборота — это баланс в ноль, а не дефицит
    scale = max(control["income"], 1)
    if control["net"] >= 0:
        add("За горизонт компания зарабатывает деньги: поток положительный. "
            "Проблема не в убыточности, а в сроках — платежи приходятся на недели, "
            "когда поступления ещё не пришли.")
    elif abs(control["net"]) / scale < 0.02:
        add("Поступления и платежи за горизонт сходятся практически в ноль: "
            "разница {s}, это {p}% оборота. Компания работает «с колёс» — "
            "весь входящий поток уходит на обязательства, запас не растёт. "
            "Разрывы возникают из-за сроков, а не из-за убытка.".format(
                s=money(abs(control["net"])),
                p=round(abs(control["net"]) / scale * 100, 1)))
    else:
        add("За горизонт денежный поток отрицательный: платежи превышают поступления "
            "на {s}. Остаток держится только за счёт начального запаса — "
            "это не разрыв по срокам, а системный дефицит.".format(
                s=money(abs(control["net"]))))
    add("")

    # ---------------- разбор каждого разрыва
    if real_gaps:
        add("## Откуда берётся нехватка")
        add("")
        for g in real_gaps:
            add("### Неделя {n}: {a} — {b}".format(
                n=g["week"], a=dmy(g["from"]), b=dmy(g["to"])))
            add("")
            w = next(x for x in weeks if x["week"] == g["week"])
            add("Остаток на конец недели — **{c}**. Поступлений {i}, платежей {o}.".format(
                c=money(w["closing"]), i=money(w["income"]), o=money(w["outcome"])))
            add("")
            drivers = week_drivers(rows, g["from"], g["to"])
            if drivers:
                add("Основные платежи недели:")
                add("")
                add("| Контрагент | Основание | Сумма | Приоритет |")
                add("|---|---|---|---|")
                for d in drivers:
                    prio = articles.get(str(d[4]), (None, None))[0]
                    add("| {c} | {b} | {s} | {p} |".format(
                        c=d[2], b=d[3] or d[4], s=money(d[5]),
                        p=prio if prio else "—"))
                add("")
            if g["movable"]:
                add("Можно перенести без последствий: {s} "
                    "(платежи приоритета 3).".format(s=money(g["movable_sum"])))
            else:
                add("Платежей приоритета 3 на этой неделе нет — "
                    "переносить без последствий нечего.")
            add("")

    # ---------------- что делать
    add("## Что делать")
    add("")
    if not real_gaps:
        add("1. Держать текущий график платежей — он исполним.")
        if low:
            add("2. На неделях с остатком ниже порога не планировать новых расходов.")
        add("")
    else:
        step = 1
        if movable_total > 0:
            add("{n}. **Перенести платежи приоритета 3** — всего {s} "
                "в проблемных неделях. Это командировки, обучение, закупки для офиса: "
                "сдвиг на 2–3 недели не несёт штрафов.".format(
                    n=step, s=money(movable_total)))
            step += 1
        if movable_total < deficit:
            add("{n}. **Договориться об отсрочке с подрядчиками.** Переноса "
                "мелких расходов не хватает: нужно закрыть ещё {s}. "
                "Платежи приоритета 2 (подрядчики, материалы) переносятся "
                "по договорённости, без штрафов по договору.".format(
                    n=step, s=money(max(deficit - movable_total, 0))))
            step += 1
        if params["credit_free"] > 0:
            add("{n}. **Держать кредитный лимит наготове.** Свободно {c}. "
                "Выборка закрывает разрыв сразу, но добавляет процентные расходы — "
                "это последний вариант, а не первый.".format(
                    n=step, c=money(params["credit_free"])))
            step += 1
        add("{n}. **Ускорить поступления.** Проверить, какие акты можно подписать "
            "раньше, и запросить авансы по договорам, где они предусмотрены.".format(n=step))
        step += 1
        add("{n}. **Не брать новых обязательств** со сроком оплаты в недели "
            "{w} до закрытия разрыва.".format(
                n=step, w=", ".join(str(g["week"]) for g in real_gaps)))
        add("")

    add("## Чего этот расчёт не показывает")
    add("")
    add("- Просроченную задолженность прошлых периодов — в реестре её нет.")
    add("- Вероятность, что плановые поступления придут в срок: все суммы "
        "со статусом «План» взяты как подтверждённые.")
    add("- Налоговые доначисления, штрафы и прочие незапланированные платежи.")
    add("")

    if notes:
        add("## Замечания к данным")
        add("")
        add("Расчёт выполнен, но на эти строки стоит посмотреть:")
        add("")
        for note in notes[:20]:
            add("- " + note)
        if len(notes) > 20:
            add("- … ещё {n} замечаний, полный список — на листе "
                "«Параметры и проверки».".format(n=len(notes) - 20))
        add("")

    add("---")
    add("")
    add("Источник: `{s}`  ".format(s=os.path.basename(path)))
    add("Расчёт: `{s}`".format(s=os.path.basename(out_path)))

    return "\n".join(L), out_path


def main():
    if len(sys.argv) < 2:
        print("Укажите файл реестра:")
        print("  python scripts/make_conclusions.py 01_input/Реестр_обязательств_ДЕМО.xlsx")
        return 1

    path = sys.argv[1]
    text, excel = build_text(path)

    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(base, "03_output")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(
        out_dir, "{d}_Платежный_календарь_Выводы.md".format(
            d=date.today().strftime("%Y-%m-%d")))
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(text + "\n")

    print("Выводы готовы: 03_output/" + os.path.basename(out_path))
    print()
    print(text[:1200])
    print("...")
    return 0


if __name__ == "__main__":
    sys.exit(main())
