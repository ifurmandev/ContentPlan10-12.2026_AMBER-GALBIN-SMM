# -*- coding: utf-8 -*-
"""Збирає TikTok-план на жовтень 2026 як HTML для імпорту в Google Документи.

Календар зверху: кожне відео — посилання на картку сценарію нижче,
у кожній картці — посилання «↑ До календаря».
"""
import datetime as dt
import html
import sys

import content_october as C

RCLASS = {C.R1: "r1", C.R2: "r2", C.R3: "r3", C.R4: "r4"}

SHORT_HOLIDAYS = C.SHORT_HOLIDAYS

CSS = (
    "body{font-family:Arial;color:#333333}"
    "table{border-collapse:collapse;width:100%}"
    "td{border:1px solid #D9D9D9;vertical-align:top;padding:5px;font-size:10pt}"
    ".wd{color:#6B6B6B;font-size:7pt;text-align:center}"
    ".num{font-size:16pt;color:#5A5A5A;margin:0}"
    ".hol{font-size:7pt;color:#B5452B;font-weight:bold;margin:0}"
    ".v{font-size:8pt;font-weight:bold;margin:0}"
    ".r1{background-color:#FFF1CC}.r1 a,.t1{color:#7A5200}"
    ".r2{background-color:#E3EFE0}.r2 a,.t2{color:#2F5A2A}"
    ".r3{background-color:#DDEBF7}.r3 a,.t3{color:#1F4E79}"
    ".r4{background-color:#FBE0CF}.r4 a,.t4{color:#8A3B12}"
    ".k{width:120pt;color:#8C8C8C;font-size:9pt;font-weight:bold}"
    ".muted{color:#8C8C8C;font-size:9pt}"
    "h2{font-size:18pt;color:#444444;page-break-before:always}"
    "h3{font-size:13pt;color:#555555}"
)


def e(s):
    return html.escape(str(s), quote=False)


def nl(s):
    return "<br>".join(e(x) for x in str(s).split("\n"))


def calendar(videos):
    by_date = {v["date"]: (i + 1, v) for i, v in enumerate(videos)}
    out = ['<table><tr>']
    for d in ["ПОНЕДІЛОК", "ВІВТОРОК", "СЕРЕДА", "ЧЕТВЕР", "П’ЯТНИЦЯ", "СУБОТА", "НЕДІЛЯ"]:
        out.append(f'<td class="wd">{d}</td>')
    out.append("</tr>")
    start = dt.date(2026, 9, 28)
    for w in range(5):
        out.append("<tr>")
        for i in range(7):
            day = start + dt.timedelta(days=w * 7 + i)
            if day.month != 10:
                out.append('<td><p class="num"> </p></td>')
                continue
            cls = ""
            body = f'<p class="num">{day.day}</p>'
            if day in SHORT_HOLIDAYS:
                body += f'<p class="hol">{e(SHORT_HOLIDAYS[day])}</p>'
            if day in by_date:
                n, v = by_date[day]
                cls = f' class="{RCLASS[v["rubric"]]}"'
                emoji = v["rubric"].split(" ")[0]
                body += (f'<p class="v"><a href="#v{n}">{emoji} {e(v["time"])}<br>'
                         f'{e(v["title"])}</a></p>')
            else:
                body += '<p class="v"> </p><p class="v"> </p>'
            out.append(f"<td{cls}>{body}</td>")
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


def legend():
    chips = " ".join(
        f'<span class="{RCLASS[r["name"]]} t{RCLASS[r["name"]][1]}"><b> {e(r["name"])} · {e(r["slot"])} </b></span>'
        for r in C.RUBRICS)
    return (f'<p style="font-size:8pt">{chips} <span style="color:#B5452B"><b>червоним — церковні, державні і професійні свята</b></span></p>'
            '<p class="muted">Натисніть на відео в календарі — відкриється його сценарій. '
            'У кожному сценарії є посилання «↑ До календаря».</p>')


def card(n, v):
    holiday = {h[0]: h[1] for h in C.HOLIDAYS}.get(v["date"])
    rc = RCLASS[v["rubric"]]
    rows = [
        ("Суть відео", v["essence"]),
        ("Хук: перші 1–2 с", v["hook"]),
        ("Розкадровка", v["board"]),
        ("Тривалість", v["dur"]),
        ("Звук", v["sound"]),
        ("Опис і хештеги", v["caption"]),
        ("Закріплений коментар", v["pinned"]),
        ("Мета: що дасть", v["goal"]),
        ("KPI-орієнтир", v["kpi"]),
        ("Референс", v["ref"]),
        ("Що потрібно для зйомки", v["need"]),
        ("Статус", "Ідея"),
        ("Коментар клієнта", " "),
    ]
    day = v["date"].strftime("%d.%m")
    meta = f'{e(v["day"])} {day} · {e(v["time"])}'
    if holiday:
        meta += f' · <span style="color:#B5452B"><b>{e(holiday)}</b></span>'
    t = "".join(f'<tr><td class="k">{e(k)}</td><td>{nl(val)}</td></tr>' for k, val in rows)
    return (f'<h2 id="v{n}">№{n} · {e(v["title"])}</h2>'
            f'<p><span class="{rc} t{rc[1]}"><b> {e(v["rubric"])} </b></span>  {meta}</p>'
            f'<table>{t}</table>'
            '<p><a href="#top">↑ До календаря</a></p>')


def simple_table(headers, rows, widths=None):
    h = "".join(f'<td class="k">{e(x)}</td>' for x in headers)
    b = "".join("<tr>" + "".join(f'<td>{c}</td>' for c in r) + "</tr>" for r in rows)
    return f'<table><tr>{h}</tr>{b}</table>'


def appendix():
    parts = ['<h2 id="rubrics">Рубрики</h2>',
             f'<p class="muted">{e(C.BASIS)}</p>']
    parts.append(simple_table(
        ["Рубрика", "Слот", "Суть", "Референс і цифри", "Мета", "Ціль переглядів"],
        [[f'<span class="{RCLASS[r["name"]]} t{RCLASS[r["name"]][1]}"><b>{e(r["name"])}</b></span>',
          e(r["slot"]), e(r["essence"]), e(r["ref"]), e(r["goal"] + " " + r["kpi"]), f'{r["target"]:,}'.replace(",", " ")]
         for r in C.RUBRICS]))
    parts.append('<h3 id="rules">Правила для всіх відео</h3><ol>')
    parts += [f"<li><p>{e(x)}</p></li>" for x in C.RULES]
    parts.append("</ol>")

    parts.append(f'<h2 id="holidays">Свята жовтень–грудень 2026</h2><p class="muted">{e(C.HOLIDAYS_RULE)}</p>')
    parts.append(simple_table(
        ["Дата", "Свято", "Тип", "Як використати", "Відео"],
        [[h[0].strftime("%d.%m.%Y"), f'<span style="color:#B5452B"><b>{e(h[1])}</b></span>', e(h[2]), e(h[4]), e(h[5])]
         for h in C.HOLIDAYS]))

    parts.append('<h2 id="refs">Референси: що спрацювало і чому</h2>'
                 '<p class="muted">Цифри зняті з екранних записів TikTok у папці referens. '
                 'ER = (вподобання + коментарі + збереження + поширення) / перегляди.</p>')
    rows = []
    for r in C.REFERENCES:
        stats = "—"
        if r["views"]:
            stats = f'{r["views"]:,}'.replace(",", " ") + " переглядів"
            if r["likes"] is not None:
                er = (r["likes"] + r["comments"] + r["saves"] + r["shares"]) / r["views"] * 100
                stats += (f'<br>{r["likes"]:,} вподобань · {r["comments"]} коментарів · {r["saves"]} збережень · '
                          f'{r["shares"]} поширень<br><b>ER {er:.1f}%</b>').replace(",", " ")
        date = r["date"].strftime("%d.%m.%Y") if r["date"] else "—"
        rows.append([f"<b>{e(r['name'])}</b><br>{date}", stats, e(r["frame"]), e(r["why"]), e(r["take"])])
    parts.append(simple_table(["Відео", "Цифри", "Що в кадрі", "Чому спрацювало / ні", "Що беремо в план"], rows))

    parts.append('<h2 id="metrics">Заміри після публікації</h2>'
                 '<p class="muted">Заповнюємо через 24–48 год з TikTok Studio. Орієнтири: «Орел» — ER 6,9%, '
                 '«Сирець» — 2,7%, «Доміно» — 3,8%. Щоп’ятниці вирішуємо, що повторюємо.</p>')
    mrows = [[str(i + 1), v["date"].strftime("%d.%m"), e(v["title"])] + [" "] * 6 for i, v in enumerate(C.S)]
    parts.append(simple_table(["№", "Дата", "Відео", "Перегляди", "Сер. час, %", "Поширення",
                               "Збереження", "Нові підписники", "Висновок"], mrows))
    return "".join(parts)


def build():
    toc = ('<p class="muted">Також у документі: <a href="#rubrics">Рубрики</a> · <a href="#rules">Правила</a> · '
           '<a href="#holidays">Свята</a> · <a href="#refs">Референси</a> · <a href="#metrics">Заміри</a></p>')
    head = (f'<p id="top" style="font-size:40pt;color:#555555;margin:0"><b>Жовтень</b> '
            f'<span style="color:#BDBDBD"><b>2026</b></span></p>'
            '<p class="muted">TikTok · Amber Galbin · 4 відео на тиждень · Пн 09:00 · Ср 21:00 · Пт 21:00 · Нд 11:00</p>')
    cards = "".join(card(i + 1, v) for i, v in enumerate(C.S))
    return (f'<html><head><meta charset="utf-8"><style>{CSS}</style></head><body>'
            f"{head}{calendar(C.S)}{legend()}{toc}{cards}{appendix()}</body></html>")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Amber_Galbin_TikTok_October_2026.html"
    open(out, "w", encoding="utf-8").write(build())
