#!/usr/bin/env python3
"""Rebuild PM working programs with curriculum-style semester matrix tables."""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


ROOT = Path(__file__).resolve().parents[1]
MATRIX = json.loads((ROOT / "data" / "plan_semester_matrix.json").read_text(encoding="utf-8"))
CURRIC = json.loads((ROOT / "data" / "pm_curriculum_is2511.json").read_text(encoding="utf-8"))
OUT_DIR = ROOT / "РабочиеПрограммы" / "Готовые_ИС2511_программист"
ORIG_PM11 = ROOT / "РабочиеПрограммы" / "оригиналы" / "ПМ.11_оригинал_веб.docx"

SEMESTERS = ["1", "2", "3", "4", "5", "6"]


def set_run_font(run, size=11, bold=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold


def add_p(doc, text="", *, center=False, bold=False, size=12, after=6, before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.1
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold)
    return p


def set_cell(cell, text, *, bold=False, center=True, size=8):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    set_run_font(p.add_run(str(text)), size=size, bold=bold)


def items_for_module(code: str):
    """Return matrix items belonging to a PM (module + MDK + practices + exam)."""
    items = MATRIX["items"]
    # find start
    start = None
    for i, it in enumerate(items):
        if it["index"] == code:
            start = i
            break
    if start is None:
        raise KeyError(code)
    out = [items[start]]
    prefix = code.replace("ПМ", "").replace("ПM", "")  # ".01" etc — better use code number
    num = code.split(".")[-1]  # 01, 02, 04, 11
    for it in items[start + 1 :]:
        idx = it["index"]
        if idx.startswith("ПМ.") or idx.startswith("ПM.") and "ЭК" not in idx and idx != code:
            # next module
            if idx.split(".")[-1] != num and "ЭК" not in idx:
                break
        if idx.startswith(("ПМ.", "ПM.")) and idx != code and "ЭК" not in idx:
            break
        if not idx:
            continue
        # include MDK/UP/PP/EK of same module number
        if (
            idx.startswith(f"МДК.{num}.")
            or idx.startswith(f"УП.{num}.")
            or idx.startswith(f"ПП.{num}.")
            or idx in (f"ПМ.{num}.ЭК", f"ПM.{num}.ЭК")
            or idx == code
        ):
            if it not in out:
                out.append(it)
        elif idx.startswith(("ПМ.", "ПM.", "МДК.", "УП.", "ПП.")):
            # different module
            if not (idx.endswith(f".{num}") or f".{num}." in idx or idx.endswith(f"{num}.ЭК")):
                break
    # cleaner filter
    filtered = []
    for it in items[start:]:
        idx = it["index"]
        if not idx:
            continue
        if it == items[start]:
            filtered.append(it)
            continue
        if idx.startswith(("ПМ.", "ПM.")) and "ЭК" not in idx:
            break
        if (
            idx.startswith(f"МДК.{num}.")
            or idx.startswith(f"УП.{num}.")
            or idx.startswith(f"ПП.{num}.")
            or idx in (f"ПМ.{num}.ЭК", f"ПM.{num}.ЭК")
        ):
            filtered.append(it)
    return filtered


def cell_text(sem_data):
    if not sem_data:
        return "", ""
    if sem_data.get("kind") == "practice":
        h = sem_data.get("hours")
        w = sem_data.get("weeks")
        return (f"{h} ч" if h is not None else ""), (f"{w} нед." if w is not None else "")
    lec = sem_data.get("lec")
    pr = sem_data.get("pr")
    lec_s = "" if lec in (None, 0) and not sem_data.get("att") and not sem_data.get("kurs") else str(lec or "")
    extras = []
    if sem_data.get("kurs"):
        extras.append(f"кп {sem_data['kurs']}")
    if sem_data.get("att"):
        extras.append(f"па {sem_data['att']}")
    pr_s = "" if pr in (None, 0) and not extras else str(pr or "")
    if extras:
        pr_s = (pr_s + " " if pr_s else "") + "(" + ", ".join(extras) + ")"
    # for exam-only rows lec/pr empty but att present
    if sem_data.get("kind") == "mdk" and lec in (None, 0) and pr in (None, 0) and sem_data.get("att"):
        return "", f"па {sem_data['att']}"
    return lec_s, pr_s


def add_matrix_table(doc, module_items):
    """Curriculum-like matrix: left = disciplines, top = semesters, cells = lec | pr."""
    add_p(
        doc,
        "Таблица соответствует учебному плану: семестры — в шапке, элементы ПМ — слева; "
        "на пересечении для МДК указаны часы лекций и практических занятий, для практик — часы и недели.",
        size=10,
        before=4,
    )
    # header rows: 2 (semester names + lec/pr), then data
    cols = 2 + len(SEMESTERS) * 2
    rows = 2 + len(module_items)
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    set_cell(table.rows[0].cells[0], "Индекс", bold=True, size=8)
    set_cell(table.rows[0].cells[1], "Наименование", bold=True, size=8)
    set_cell(table.rows[1].cells[0], "", size=8)
    set_cell(table.rows[1].cells[1], "", size=8)

    for i, sem in enumerate(SEMESTERS):
        c0 = 2 + i * 2
        weeks = MATRIX["header"]["semesters"][sem]
        set_cell(table.rows[0].cells[c0], f"Семестр {sem}", bold=True, size=7)
        set_cell(table.rows[0].cells[c0 + 1], weeks, bold=True, size=7)
        # merge visually by repeating label; subheaders:
        set_cell(table.rows[1].cells[c0], "Лек.", bold=True, size=7)
        set_cell(table.rows[1].cells[c0 + 1], "Практ.", bold=True, size=7)

    for ri, it in enumerate(module_items):
        r = table.rows[ri + 2]
        set_cell(r.cells[0], it["index"], bold=True, center=False, size=7)
        set_cell(r.cells[1], it["name"], center=False, size=7)
        for i, sem in enumerate(SEMESTERS):
            lec_s, pr_s = cell_text(it["by_semester"].get(sem))
            set_cell(r.cells[2 + i * 2], lec_s, size=7)
            set_cell(r.cells[2 + i * 2 + 1], pr_s, size=7)
    return table


def module_meta(code: str):
    for m in CURRIC["modules"]:
        if m["code"] == code:
            return m
    raise KeyError(code)


def build_pm_doc(code: str) -> Document:
    meta = module_meta(code)
    items = items_for_module(code)
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(1.2)

    for line in [
        "РОССИЙСКАЯ ФЕДЕРАЦИЯ",
        "ХАНТЫ-МАНСИЙСКИЙ АВТОНОМНЫЙ ОКРУГ — ЮГРА",
        "АНПОО «СУРГУТСКИЙ ИНСТИТУТ ЭКОНОМИКИ, УПРАВЛЕНИЯ И ПРАВА»",
    ]:
        add_p(doc, line, center=True, bold=True, size=11, after=1)

    add_p(doc, "РАБОЧАЯ ПРОГРАММА", center=True, bold=True, size=14, before=10)
    add_p(doc, f"профессионального модуля {code}", center=True, bold=True, size=12)
    add_p(doc, meta["name"], center=True, bold=True, size=12, after=8)

    info = [
        ("Специальность", CURRIC["specialty"]),
        ("Квалификация", "Программист"),
        ("Форма / срок", f"{CURRIC['form']}, {CURRIC['duration']}"),
        ("Группа", CURRIC["group"]),
        ("Источник часов", "учебный план ИС 25-11 (программист), лист «План»"),
    ]
    t = doc.add_table(rows=len(info), cols=2)
    t.style = "Table Grid"
    for i, (k, v) in enumerate(info):
        set_cell(t.rows[i].cells[0], k, bold=True, center=False, size=10)
        set_cell(t.rows[i].cells[1], v, center=False, size=10)

    add_p(doc, "Сургут, 2025", center=True, size=11, before=10)
    doc.add_page_break()

    add_p(
        doc,
        f"Рабочая программа {code} «{meta['name']}» составлена по ФГОС СПО 09.02.07 и учебному плану "
        f"АНПОО «СИЭУиП» для квалификации «Программист», группа {CURRIC['group']}. "
        "Разбивка часов по семестрам взята с перекрёстка строк элементов ПМ и столбцов семестров учебного плана "
        "(шапка: Семестр 1–6; в том числе — «Лекции, уроки» и «Пр. занятия»).",
        size=11,
    )

    # Results short
    add_p(doc, "1. РЕЗУЛЬТАТЫ ОСВОЕНИЯ", bold=True, size=13, before=8)
    add_p(doc, f"Вид деятельности: {meta['vd']}", size=11)
    add_p(doc, "Профессиональные компетенции:", bold=True, size=11)
    for c, n in meta["pk"]:
        add_p(doc, f"{c} — {n}", size=10, after=2)
    add_p(doc, "Иметь практический опыт: " + meta["experience"], size=10)
    add_p(doc, "Уметь: " + meta["able"], size=10)
    add_p(doc, "Знать: " + meta["know"], size=10)

    tot = meta["totals"]
    add_p(doc, "2. ОБЪЁМ МОДУЛЯ (итог по учебному плану)", bold=True, size=13, before=8)
    hrs = [
        ["Показатель", "Часы"],
        ["Максимальная учебная нагрузка", tot["max"]],
        ["Самостоятельная работа", tot["self"]],
        ["Обязательная аудиторная по МДК", tot["ob"]],
        ["  лекции, уроки", tot["lec"]],
        ["  практические занятия", tot["pr"]],
        ["Курсовое проектирование (входит в обязательную)", tot.get("kurs") or 0],
        ["Учебная практика", tot["up"]],
        ["Производственная практика", tot["pp"]],
        ["Промежуточная аттестация", tot["att"]],
    ]
    ht = doc.add_table(rows=len(hrs), cols=2)
    ht.style = "Table Grid"
    for i, row in enumerate(hrs):
        set_cell(ht.rows[i].cells[0], row[0], bold=i == 0, center=False, size=10)
        set_cell(ht.rows[i].cells[1], row[1], bold=i == 0, size=10)

    add_p(doc, "3. РАСПРЕДЕЛЕНИЕ ПО СЕМЕСТРАМ (как в учебном плане)", bold=True, size=13, before=10)
    add_matrix_table(doc, items)

    # Thematic content by MDK — WITHOUT inventing per-topic semester labels that conflict with plan.
    # Instead: show plan hours per semester for MDK, then thematic list with total hours.
    add_p(doc, "4. ТЕМАТИЧЕСКОЕ СОДЕРЖАНИЕ МДК", bold=True, size=13, before=12)
    add_p(
        doc,
        "Ниже для каждого МДК сначала повторяется строка учебного плана (часы лекций/практики по семестрам), "
        "затем тематическое содержание. Сумма аудиторных часов тем равна итогу МДК по плану; "
        "проведение тем распределяется внутри семестров, где по плану стоят часы данного МДК.",
        size=10,
    )

    for mdk in meta["mdk"]:
        add_p(doc, f"{mdk['code']} «{mdk['name']}»", bold=True, size=12, before=8)
        add_p(doc, f"Аттестация: {mdk['attestation']}", size=10)
        # mini matrix one row
        mdk_item = next(it for it in items if it["index"] == mdk["code"])
        rows_data = [["Семестр"] + [f"Сем. {s}" for s in SEMESTERS], ["Лекции"], ["Практ. занятия"], ["Курс. проект / ПА"]]
        for s in SEMESTERS:
            cell = mdk_item["by_semester"].get(s)
            if not cell:
                rows_data[1].append("")
                rows_data[2].append("")
                rows_data[3].append("")
            else:
                rows_data[1].append(cell.get("lec") or "")
                rows_data[2].append(cell.get("pr") or "")
                extras = []
                if cell.get("kurs"):
                    extras.append(f"кп {cell['kurs']}")
                if cell.get("att"):
                    extras.append(f"па {cell['att']}")
                rows_data[3].append(", ".join(extras))
        mt = doc.add_table(rows=4, cols=7)
        mt.style = "Table Grid"
        for r_i, row in enumerate(rows_data):
            for c_i, val in enumerate(row):
                set_cell(mt.rows[r_i].cells[c_i], val, bold=r_i == 0 or c_i == 0, size=8)

        # topics — show content + hours; mark allowed semesters from plan
        allowed = sorted(mdk_item["by_semester"].keys(), key=int)
        add_p(doc, f"Семестры проведения по учебному плану: {', '.join(allowed)}.", size=10, before=4)

        th = doc.add_table(rows=1 + len(mdk["topics"]), cols=4)
        th.style = "Table Grid"
        for c_i, h in enumerate(["Тема", "Содержание / практические задания", "Лек.", "Практ."]):
            set_cell(th.rows[0].cells[c_i], h, bold=True, size=8)
        for r_i, topic in enumerate(mdk["topics"], start=1):
            body = ["Содержание:"] + [f"— {x}" for x in topic["content"]]
            if topic.get("practices"):
                body += ["Практические занятия:"] + [f"— {x}" for x in topic["practices"]]
            set_cell(th.rows[r_i].cells[0], topic["title"], center=False, size=8)
            set_cell(th.rows[r_i].cells[1], "\n".join(body), center=False, size=7)
            set_cell(th.rows[r_i].cells[2], topic["lec"], size=8)
            set_cell(th.rows[r_i].cells[3], topic["pr"], size=8)

        add_p(
            doc,
            f"Итого по МДК: лекции {mdk['totals']['lec']} ч, практические {mdk['totals']['pr']} ч"
            + (f", курсовой проект {mdk['totals'].get('kurs')} ч" if mdk["totals"].get("kurs") else "")
            + ".",
            size=10,
            before=4,
        )

    # Practices
    add_p(doc, "5. УЧЕБНАЯ И ПРОИЗВОДСТВЕННАЯ ПРАКТИКИ", bold=True, size=13, before=10)
    for prac in meta["practices"]:
        add_p(doc, f"{prac['code']} {prac['name']} — {prac['hours']} ч ({prac['weeks']} нед.). {prac['attestation']}", bold=True, size=11)
        add_p(doc, "По учебному плану (перекрёсток строки практики и столбцов семестров):", size=10)
        pt = doc.add_table(rows=1 + len(prac["semesters"]), cols=3)
        pt.style = "Table Grid"
        set_cell(pt.rows[0].cells[0], "Семестр", bold=True, size=9)
        set_cell(pt.rows[0].cells[1], "Часы", bold=True, size=9)
        set_cell(pt.rows[0].cells[2], "Недели", bold=True, size=9)
        for i, (sem, v) in enumerate(sorted(prac["semesters"].items(), key=lambda x: int(x[0])), start=1):
            set_cell(pt.rows[i].cells[0], sem, size=9)
            set_cell(pt.rows[i].cells[1], v["hours"], size=9)
            set_cell(pt.rows[i].cells[2], v["weeks"], size=9)

    em = meta["exam_module"]
    add_p(doc, f"{em['code']} «{em['name']}» — {em['hours']} ч, семестр {em['sem']} (по учебному плану).", size=11, before=6)

    add_p(doc, "6. УСЛОВИЯ РЕАЛИЗАЦИИ И КОНТРОЛЬ", bold=True, size=13, before=10)
    add_p(
        doc,
        "Реализация модуля предполагает лаборатории программирования и баз данных / ПО и сопровождения КС, "
        "АРМ обучающихся, проекционное оборудование. Контроль: практические работы, курсовой проект (при наличии), "
        "отчёты по практикам, экзамены/дифф. зачёты по МДК, экзамен по модулю.",
        size=10,
    )
    return doc


def write_summary_matrix():
    lines = [
        "# Сводка часов ПМ по учебному плану ИС 25/11 (программист)",
        "",
        "Источник: лист «План» файла `УчебныйПлан/ИС 25-11 (первый курс) (программист).xls`.",
        "",
        "Формат как в учебном плане: **семестры в шапке**, **элементы ПМ слева**, "
        "на пересечении — **лекции / практические занятия** (для УП/ПП — **часы / недели**).",
        "",
    ]
    for code in ["ПМ.01", "ПМ.02", "ПМ.04", "ПМ.11"]:
        meta = module_meta(code)
        items = items_for_module(code)
        lines.append(f"## {code} {meta['name']}")
        lines.append("")
        header = "| Индекс | Наименование |"
        sep = "|---|---|"
        for s in SEMESTERS:
            header += f" Сем.{s} лек | Сем.{s} пр |"
            sep += "---:|---:|"
        lines.append(header)
        lines.append(sep)
        for it in items:
            row = f"| {it['index']} | {it['name']} |"
            for s in SEMESTERS:
                lec_s, pr_s = cell_text(it["by_semester"].get(s))
                row += f" {lec_s} | {pr_s} |"
            lines.append(row)
        lines.append("")
    path = OUT_DIR / "Сводка_часов.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {path.relative_to(ROOT)}")


FILENAMES = {
    "ПМ.01": "ПМ.01_Разработка_модулей_ПО_программист.docx",
    "ПМ.02": "ПМ.02_Интеграция_программных_модулей_программист.docx",
    "ПМ.04": "ПМ.04_Сопровождение_и_обслуживание_ПО_программист.docx",
    "ПМ.11": "ПМ.11_Разработка_администрирование_защита_БД_программист.docx",
}


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    write_summary_matrix()
    for code, fn in FILENAMES.items():
        doc = build_pm_doc(code)
        out = OUT_DIR / fn
        doc.save(out)
        print(f"Wrote {out.relative_to(ROOT)}")

    # also refresh underscore copies used earlier
    mirror = {
        "ПМ.01": ROOT / "РабочиеПрограммы" / "ПМ_01_Разработка_модулей_программного_обеспечения_для_компьютерных.docx",
        "ПМ.02": ROOT / "РабочиеПрограммы" / "ПМ_02_Осуществление_интеграции_программных_модулей.docx",
        "ПМ.04": ROOT / "РабочиеПрограммы" / "ПМ_04_Сопровождение_и_обслуживание_программного_обеспечения_компьютерных.docx",
        "ПМ.11": ROOT / "РабочиеПрограммы" / "ПМ_11_Разработка_администрирование_и_защита_баз_данных.docx",
    }
    for code, path in mirror.items():
        build_pm_doc(code).save(path)
        print(f"Wrote {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
