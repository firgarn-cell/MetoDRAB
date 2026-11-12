#!/usr/bin/env python3
"""Generate working programs for programmer PMs from curriculum JSON."""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "pm_curriculum_is2511.json"
OUT_DIR = ROOT / "РабочиеПрограммы"


def set_run_font(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def add_p(doc, text="", *, center=False, bold=False, size=12, space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold)
    return p


def add_heading_local(doc, text, level=1):
    sizes = {1: 14, 2: 13, 3: 12}
    p = add_p(doc, text, bold=True, size=sizes.get(level, 12), space_before=10, space_after=8)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def set_cell_text(cell, text, *, bold=False, center=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(str(text))
    set_run_font(run, size=size, bold=bold)


def fill_table(table, rows, header=True):
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            set_cell_text(
                table.rows[r_idx].cells[c_idx],
                value,
                bold=header and r_idx == 0,
                center=True,
                size=9 if len(row) > 5 else 10,
            )


def h(v):
    return "" if v is None else str(v)


def build_document(meta: dict, module: dict) -> Document:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)

    # Title page
    for line in [
        "РОССИЙСКАЯ ФЕДЕРАЦИЯ",
        "ХАНТЫ-МАНСИЙСКИЙ АВТОНОМНЫЙ ОКРУГ — ЮГРА",
        "",
        "АВТОНОМНАЯ НЕКОММЕРЧЕСКАЯ ПРОФЕССИОНАЛЬНАЯ ОБРАЗОВАТЕЛЬНАЯ ОРГАНИЗАЦИЯ",
        "«СУРГУТСКИЙ ИНСТИТУТ ЭКОНОМИКИ, УПРАВЛЕНИЯ И ПРАВА»",
        "(АНПОО «СИЭУиП»)",
    ]:
        add_p(doc, line, center=True, bold=True, size=12 if line else 12, space_after=2)

    add_p(doc, "", space_after=18)
    add_p(doc, "РАБОЧАЯ ПРОГРАММА", center=True, bold=True, size=16)
    add_p(doc, f"профессионального модуля {module['code']}", center=True, bold=True, size=14)
    add_p(doc, module["name"], center=True, bold=True, size=13, space_after=12)

    info = [
        ("Специальность", meta["specialty"]),
        ("Квалификация", meta["qualification"]),
        ("Форма обучения / срок", f"{meta['form']}, {meta['duration']}"),
        ("Учебная группа", meta["group"]),
        ("Основание по часам", "учебный план ИС 25-11 (программист)"),
    ]
    table = doc.add_table(rows=len(info), cols=2)
    table.style = "Table Grid"
    for i, (k, v) in enumerate(info):
        set_cell_text(table.rows[i].cells[0], k, bold=True, size=11)
        set_cell_text(table.rows[i].cells[1], v, size=11)

    add_p(doc, "", space_after=18)
    add_p(doc, "Сургут, 2025", center=True, size=12)
    doc.add_page_break()

    add_p(
        doc,
        (
            f"Рабочая программа {module['code']} «{module['name']}» составлена в соответствии "
            "с ФГОС СПО по специальности 09.02.07 «Информационные системы и программирование» "
            "(приказ Минобрнауки России от 09.12.2016 № 1547) и учебным планом АНПОО «СИЭУиП» "
            f"для квалификации «{meta['qualification']}», группа {meta['group']}."
        ),
        size=11,
    )
    add_p(
        doc,
        "В программе учтена только квалификация «Программист». Разбивка аудиторных часов "
        "по семестрам (лекции / практические занятия) приведена строго по учебному плану.",
        size=11,
    )
    add_p(doc, "Кафедра информационных систем. Корпус СП7.", size=11)

    # 1. Passport
    add_heading_local(doc, "1. ОБЩАЯ ХАРАКТЕРИСТИКА РАБОЧЕЙ ПРОГРАММЫ ПРОФЕССИОНАЛЬНОГО МОДУЛЯ", 1)
    add_heading_local(doc, "1.1. Цель и планируемые результаты освоения", 2)
    add_p(
        doc,
        (
            f"В результате изучения профессионального модуля обучающийся должен освоить вид деятельности "
            f"«{module['vd']}» и соответствующие профессиональные и общие компетенции."
        ),
    )

    add_heading_local(doc, "1.1.1. Перечень общих компетенций", 3)
    ok_table = doc.add_table(rows=1 + len(meta["ok"]), cols=2)
    ok_table.style = "Table Grid"
    fill_table(ok_table, [["Код", "Наименование общих компетенций"]] + meta["ok"])

    add_heading_local(doc, "1.1.2. Перечень профессиональных компетенций", 3)
    pk_table = doc.add_table(rows=2 + len(module["pk"]), cols=2)
    pk_table.style = "Table Grid"
    set_cell_text(pk_table.rows[0].cells[0], "Код", bold=True, center=True)
    set_cell_text(pk_table.rows[0].cells[1], "Вид деятельности и профессиональные компетенции", bold=True, center=True)
    set_cell_text(pk_table.rows[1].cells[0], "ВД", bold=True, center=True)
    set_cell_text(pk_table.rows[1].cells[1], module["vd"], size=10)
    for i, (code, name) in enumerate(module["pk"], start=2):
        set_cell_text(pk_table.rows[i].cells[0], code, center=True, size=10)
        set_cell_text(pk_table.rows[i].cells[1], name, size=10)

    add_heading_local(doc, "1.1.3. Практический опыт, умения и знания", 3)
    zu = doc.add_table(rows=3, cols=2)
    zu.style = "Table Grid"
    set_cell_text(zu.rows[0].cells[0], "Иметь практический опыт", bold=True, size=10)
    set_cell_text(zu.rows[0].cells[1], module["experience"], size=10)
    set_cell_text(zu.rows[1].cells[0], "Уметь", bold=True, size=10)
    set_cell_text(zu.rows[1].cells[1], module["able"], size=10)
    set_cell_text(zu.rows[2].cells[0], "Знать", bold=True, size=10)
    set_cell_text(zu.rows[2].cells[1], module["know"], size=10)

    # Hours overview
    add_heading_local(doc, "1.2. Количество часов на освоение профессионального модуля", 2)
    t = module["totals"]
    hours_rows = [
        ["Показатель", "Часы"],
        ["Максимальная учебная нагрузка", t["max"]],
        ["Самостоятельная работа", t["self"]],
        ["Обязательная аудиторная нагрузка по МДК (всего)", t["ob"]],
        ["  в том числе лекции, уроки", t["lec"]],
        ["  в том числе практические занятия", t["pr"]],
        ["Курсовая работа / проектирование (входит в обязательную)", t.get("kurs", 0) or 0],
        ["Учебная практика", t["up"]],
        ["Производственная практика", t["pp"]],
        ["Промежуточная аттестация (в т.ч. экзамен по модулю)", t["att"]],
    ]
    hours_table = doc.add_table(rows=len(hours_rows), cols=2)
    hours_table.style = "Table Grid"
    fill_table(hours_table, hours_rows)

    add_p(doc, "Распределение обязательной аудиторной нагрузки модуля по семестрам:", size=11, space_before=8)
    sem_header = ["Семестр", "Макс.", "Самост.", "Обязат.", "Лекции", "Практика", "Курс.р.", "ПА"]
    sem_rows = [sem_header]
    for sem in sorted(module["semesters"], key=int):
        s = module["semesters"][sem]
        sem_rows.append(
            [
                sem,
                h(s.get("max")),
                h(s.get("self")),
                h(s.get("ob")),
                h(s.get("lec")),
                h(s.get("pr")),
                h(s.get("kurs")),
                h(s.get("att")),
            ]
        )
    sem_table = doc.add_table(rows=len(sem_rows), cols=len(sem_header))
    sem_table.style = "Table Grid"
    fill_table(sem_table, sem_rows)

    # 2. Structure
    add_heading_local(doc, "2. СТРУКТУРА И СОДЕРЖАНИЕ ПРОФЕССИОНАЛЬНОГО МОДУЛЯ", 1)
    add_heading_local(doc, "2.1. Структура профессионального модуля (квалификация — программист)", 2)

    struct_header = [
        "Элемент ПМ",
        "Всего обяз.",
        "Лекции",
        "Практика",
        "Курс.р.",
        "Самост.",
        "ПА",
        "Семестры (лек/практ)",
    ]
    struct_rows = [struct_header]
    for mdk in module["mdk"]:
        sem_parts = []
        for sem in sorted(mdk["semesters"], key=int):
            s = mdk["semesters"][sem]
            sem_parts.append(f"{sem}: {s.get('lec', 0)}/{s.get('pr', 0)}")
        tt = mdk["totals"]
        struct_rows.append(
            [
                f"{mdk['code']} {mdk['name']}",
                tt["ob"],
                tt["lec"],
                tt["pr"],
                tt.get("kurs", "") or "",
                tt.get("self", "") or "",
                tt.get("att", "") or "",
                "; ".join(sem_parts),
            ]
        )
    for prac in module["practices"]:
        sem_parts = [f"{sem}: {v['hours']} ч ({v['weeks']} нед.)" for sem, v in sorted(prac["semesters"].items(), key=lambda x: int(x[0]))]
        struct_rows.append(
            [
                f"{prac['code']} {prac['name']}",
                prac["hours"],
                "—",
                "—",
                "",
                "",
                "",
                "; ".join(sem_parts),
            ]
        )
    em = module["exam_module"]
    struct_rows.append([f"{em['code']} {em['name']}", "", "", "", "", "", em["hours"], f"сем. {em['sem']}"])

    struct_table = doc.add_table(rows=len(struct_rows), cols=len(struct_header))
    struct_table.style = "Table Grid"
    fill_table(struct_table, struct_rows)

    add_p(
        doc,
        "Примечание: для МДК в колонке «Семестры» указаны часы лекций/практических занятий "
        "по каждому семестру в соответствии с учебным планом.",
        size=10,
        space_before=6,
    )

    # Detailed semester tables per MDK
    add_heading_local(doc, "2.2. Разбивка часов МДК по семестрам", 2)
    for mdk in module["mdk"]:
        add_p(doc, f"{mdk['code']} «{mdk['name']}». Аттестация: {mdk['attestation']}", bold=True, size=11)
        rows = [["Семестр", "Макс.", "Самост.", "Обязат.", "Лекции", "Практ.", "Курс.р.", "ПА"]]
        for sem in sorted(mdk["semesters"], key=int):
            s = mdk["semesters"][sem]
            rows.append(
                [
                    sem,
                    h(s.get("max")),
                    h(s.get("self")),
                    h(s.get("ob")),
                    h(s.get("lec")),
                    h(s.get("pr")),
                    h(s.get("kurs")),
                    h(s.get("att")),
                ]
            )
        tot = mdk["totals"]
        rows.append(
            [
                "Итого",
                tot["max"],
                tot.get("self", ""),
                tot["ob"],
                tot["lec"],
                tot["pr"],
                tot.get("kurs", "") or "",
                tot.get("att", "") or "",
            ]
        )
        tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
        tbl.style = "Table Grid"
        fill_table(tbl, rows)
        add_p(doc, "", space_after=4)

    # Thematic plan
    add_heading_local(doc, "2.3. Тематический план и содержание (с разбивкой по семестрам)", 2)
    for mdk in module["mdk"]:
        add_p(doc, f"{mdk['code']} {mdk['name']}", bold=True, size=12, space_before=8)
        theme_header = ["Наименование тем", "Содержание / практические задания", "Сем.", "Лек.", "Практ."]
        theme_rows = [theme_header]
        for topic in mdk["topics"]:
            content_lines = ["Содержание:"] + [f"— {c}" for c in topic["content"]]
            if topic.get("practices"):
                content_lines += ["Практические занятия:"] + [f"— {c}" for c in topic["practices"]]
            theme_rows.append(
                [
                    topic["title"],
                    "\n".join(content_lines),
                    topic["sem"],
                    topic["lec"],
                    topic["pr"],
                ]
            )
        # semester subtotals
        for sem in sorted({str(t["sem"]) for t in mdk["topics"]}, key=int):
            lec = sum(t["lec"] for t in mdk["topics"] if str(t["sem"]) == sem)
            pr = sum(t["pr"] for t in mdk["topics"] if str(t["sem"]) == sem)
            theme_rows.append([f"Итого за семестр {sem}", "", sem, lec, pr])
        theme_rows.append(
            [
                "Всего по МДК (ауд.)",
                "",
                "",
                mdk["totals"]["lec"],
                mdk["totals"]["pr"],
            ]
        )
        theme_table = doc.add_table(rows=len(theme_rows), cols=5)
        theme_table.style = "Table Grid"
        for r_idx, row in enumerate(theme_rows):
            for c_idx, value in enumerate(row):
                set_cell_text(
                    theme_table.rows[r_idx].cells[c_idx],
                    value,
                    bold=r_idx == 0 or str(row[0]).startswith("Итого") or str(row[0]).startswith("Всего"),
                    center=c_idx >= 2,
                    size=9,
                )

    # Practices detail
    add_heading_local(doc, "2.4. Учебная и производственная практики", 2)
    for prac in module["practices"]:
        add_p(doc, f"{prac['code']} {prac['name']} — {prac['hours']} ч ({prac['weeks']} нед.). {prac['attestation']}", bold=True, size=11)
        rows = [["Семестр", "Часы", "Недели"]]
        for sem, v in sorted(prac["semesters"].items(), key=lambda x: int(x[0])):
            rows.append([sem, v["hours"], v["weeks"]])
        pt = doc.add_table(rows=len(rows), cols=3)
        pt.style = "Table Grid"
        fill_table(pt, rows)

    add_p(
        doc,
        f"{module['exam_module']['code']} «{module['exam_module']['name']}» — "
        f"{module['exam_module']['hours']} ч, семестр {module['exam_module']['sem']}.",
        size=11,
        space_before=8,
    )

    # 3. Conditions
    add_heading_local(doc, "3. УСЛОВИЯ РЕАЛИЗАЦИИ ПРОФЕССИОНАЛЬНОГО МОДУЛЯ", 1)
    add_heading_local(doc, "3.1. Материально-техническое обеспечение", 2)
    add_p(
        doc,
        "Реализация модуля предполагает наличие лабораторий программирования и баз данных, "
        "программного обеспечения и сопровождения компьютерных систем; персональных компьютеров "
        "с выходом в сеть Интернет; проекционного оборудования. Для проектов обучающихся "
        "используется имеющееся оборудование кафедры, включая VR-ресурсы института.",
    )
    add_heading_local(doc, "3.2. Информационное обеспечение обучения", 2)
    add_p(doc, "Основные источники: учебники и учебные пособия по специальности 09.02.07; документация используемых языков и сред разработки; ФГОС СПО; учебный план ИС 25-11.")
    add_p(doc, "Дополнительные источники: официальная документация СУБД, IDE, систем контроля версий; материалы Demo Day и проектной деятельности обучающихся.")
    add_heading_local(doc, "3.3. Организация образовательного процесса", 2)
    add_p(
        doc,
        "Освоение модуля осуществляется в соответствии с календарным учебным графиком и учебным планом. "
        "Практические занятия выполняются индивидуально или в малых группах. Курсовая работа "
        "(при наличии) выполняется в семестре, указанном в учебном плане.",
    )

    # 4. Assessment
    add_heading_local(doc, "4. КОНТРОЛЬ И ОЦЕНКА РЕЗУЛЬТАТОВ ОСВОЕНИЯ ПРОФЕССИОНАЛЬНОГО МОДУЛЯ", 1)
    assess_rows = [["Результаты обучения", "Формы и методы контроля"]]
    for code, name in module["pk"]:
        assess_rows.append(
            [
                f"{code}: {name}",
                "Практические работы, курсовая работа (при наличии), отчёты по практике, экзамен/дифф. зачёт по МДК, экзамен по модулю",
            ]
        )
    assess_rows.append(
        [
            "ОК 01–ОК 11",
            "Наблюдение на занятиях и практике, оценка выполнения заданий, защита результатов",
        ]
    )
    at = doc.add_table(rows=len(assess_rows), cols=2)
    at.style = "Table Grid"
    fill_table(at, assess_rows)

    add_p(
        doc,
        "Критерии оценки: «отлично» — задание выполнено полностью в соответствии с требованиями; "
        "«хорошо» — выполнено с незначительными замечаниями; «удовлетворительно» — выполнен минимальный "
        "обязательный объём; «неудовлетворительно» — требования не выполнены.",
        size=11,
        space_before=8,
    )
    return doc


def filename_for(module: dict) -> str:
    mapping = {
        "ПМ.01": "ПМ_01_Разработка_модулей_программного_обеспечения_для_компьютерных.docx",
        "ПМ.02": "ПМ_02_Осуществление_интеграции_программных_модулей.docx",
        "ПМ.04": "ПМ_04_Сопровождение_и_обслуживание_программного_обеспечения_компьютерных.docx",
        "ПМ.11": "ПМ_11_Разработка_администрирование_и_защита_баз_данных.docx",
    }
    return mapping[module["code"]]


def main():
    meta = json.loads(DATA.read_text(encoding="utf-8"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # Remove obsolete multi-qualification docs by overwriting programmer set.
    for module in meta["modules"]:
        doc = build_document(meta, module)
        out = OUT_DIR / filename_for(module)
        doc.save(out)
        print(f"Wrote {out.relative_to(ROOT)}")

    # Summary markdown for quick review
    summary = ROOT / "РабочиеПрограммы" / "Сводка_часов_ПМ_программист_ИС2511.md"
    lines = [
        "# Сводка часов ПМ (квалификация «Программист», группа ИС 25/11)",
        "",
        "Источник: учебный план `УчебныйПлан/ИС 25-11 (первый курс) (программист).xls`.",
        "В рабочих программах оставлены только ПМ для программиста: ПМ.01, ПМ.02, ПМ.04, ПМ.11.",
        "",
    ]
    for module in meta["modules"]:
        lines.append(f"## {module['code']} {module['name']}")
        t = module["totals"]
        lines.append(
            f"- Макс. {t['max']} ч; ауд. МДК {t['ob']} ч (лекции {t['lec']}, практика {t['pr']}); "
            f"УП {t['up']}; ПП {t['pp']}; ПА {t['att']}; СРС {t['self']}."
        )
        lines.append("")
        lines.append("| Элемент | Сем. | Лек. | Практ. | Обязат. | Примечание |")
        lines.append("|---|---:|---:|---:|---:|---|")
        for mdk in module["mdk"]:
            for sem, s in sorted(mdk["semesters"].items(), key=lambda x: int(x[0])):
                note = []
                if s.get("kurs"):
                    note.append(f"курс.р. {s['kurs']}")
                if s.get("att"):
                    note.append(f"ПА {s['att']}")
                lines.append(
                    f"| {mdk['code']} | {sem} | {s.get('lec', 0)} | {s.get('pr', 0)} | {s.get('ob', 0)} | {'; '.join(note)} |"
                )
        for prac in module["practices"]:
            for sem, s in sorted(prac["semesters"].items(), key=lambda x: int(x[0])):
                lines.append(
                    f"| {prac['code']} | {sem} | — | — | {s['hours']} | {s['weeks']} нед. |"
                )
        em = module["exam_module"]
        lines.append(f"| {em['code']} | {em['sem']} | — | — | — | {em['hours']} ч |")
        lines.append("")
    summary.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {summary.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
