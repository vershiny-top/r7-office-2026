#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a .docbuilder script from the CSVs and execute it via R7 Document Builder."""
import os
import shutil
import subprocess
import sys
import tempfile

INPUTS = {
    "products":  "file_1_Товары каталог.txt",
    "suppliers": "file_2_Поставщики.txt",
    "sales":     "file_3_Продажи.txt",
    "clients":   "file_4_Клиенты.txt",
}


def find_input(name):
    for base in (".", "data"):
        p = os.path.join(base, name)
        if os.path.exists(p):
            return p
    return None


def read_rows(path):
    if not path:
        return []
    with open(path, encoding="utf-8") as f:
        return [ln.rstrip("\r\n").split(";") for ln in f if ln.strip()]


def js(v):
    """Render a Python value as a JS literal (string or number)."""
    if v is None:
        return "null"
    if isinstance(v, (int, float)):
        return repr(v)
    s = str(v)
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n") + '"'


def call(sheet_var, cell, payload, is_formula=False):
    fn = "SetFormula" if is_formula else "SetValue"
    return f'{sheet_var}.GetRange("{cell}").{fn}({payload});'


def gen_docbuilder():
    prod = read_rows(find_input(INPUTS["products"]))
    sup = read_rows(find_input(INPUTS["suppliers"]))
    sal = read_rows(find_input(INPUTS["sales"]))
    cli = read_rows(find_input(INPUTS["clients"]))

    L = []
    L.append('builder.CreateFile("xlsx");')
    L.append('var s0 = Api.GetActiveSheet();')
    L.append('s0.SetName("Товары_исходные");')

    def dump_rows(target_var, rows):
        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row, start=1):
                col = chr(ord("A") + c - 1) if c <= 26 else "A"
                L.append(call(target_var, f"{col}{r}", js(val)))

    def add_sheet(name):
        v = f"s{len(L)}"
        L.append(f'var {v} = Api.AddSheet({js(name)});')
        return v

    # --- Справочники first (formulas reference it) ---
    ref = add_sheet("Справочники")
    ref_data = [
        (1, [("Исходная", "Канон"),
             ("Смартфоны", "Смартфоны"), ("Телефоны", "Смартфоны"),
             ("Мобильные телефоны", "Смартфоны"),
             ("Ноутбуки", "Ноутбуки"), ("Лэптопы", "Ноутбуки"),
             ("Планшеты", "Планшеты"), ("Телевизоры", "Телевизоры"),
             ("Наушники", "Наушники"), ("Мониторы", "Мониторы"),
             ("Аксессуары", "Аксессуары"), ("Часы", "Часы")]),
        (4, [("Исходная", "Канон"), ("шт", "шт."), ("штука", "шт."),
             ("шт.", "шт."), ("штуки", "шт.")]),
        (7, [("Исходная", "Канон"),
             ("Завершен", "Завершен"), ("завершен", "Завершен"),
             ("completed", "Завершен"), ("выполнен", "Завершен"),
             ("отменен", "Отменен"), ("Отменен", "Отменен"),
             ("canceled", "Отменен"), ("возврат", "Возврат"), ("Возврат", "Возврат")]),
        (10, [("Исходная", "Канон"),
              ("Москва", "Москва"), ("г.Москва", "Москва"),
              ("МСК", "Москва"), ("г Москва", "Москва"),
              ("Санкт-Петербург", "Санкт-Петербург"),
              ("СПб", "Санкт-Петербург"), ("СПБ", "Санкт-Петербург")]),
    ]
    for start_col, rows in ref_data:
        for r, row in enumerate(rows, start=1):
            c1 = chr(ord("A") + start_col - 1)
            c2 = chr(ord("A") + start_col)
            L.append(call(ref, f"{c1}{r}", js(row[0])))
            L.append(call(ref, f"{c2}{r}", js(row[1])))

    # --- Товары_исходные (raw) ---
    dump_rows("s0", prod)

    # --- Товары_очищенные (formulas) ---
    pr = add_sheet("Товары_очищенные")
    if len(prod) > 1:
        for c, v in enumerate(prod[0], start=1):
            col = chr(ord("A") + c - 1)
            L.append(call(pr, f"{col}1", js(v)))
        L.append(call(pr, "H1", js("Проверка_цены")))
        L.append(call(pr, "I1", js("Проверка_уникальности")))
        n = len(prod)
        for i in range(2, n + 1):
            L.append(call(pr, f"A{i}",
                '"=\"АРТ-\"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Товары_исходные!A' + str(i) + ',\\"№\\",\\"\\"),\\"АРТ-\\",\\"\\"),\\"АРТ_\\",\\"\\"),\\"/\\",\\"-\\"),3)"',
                is_formula=True))
            L.append(call(pr, f"B{i}",
                '"=PROPER(TRIM(IFERROR(LEFT(Товары_исходные!B' + str(i) + ',FIND(\\"(\\",Товары_исходные!B' + str(i) + ')-1),Товары_исходные!B' + str(i) + ')))"',
                is_formula=True))
            L.append(call(pr, f"C{i}",
                '"=IFERROR(VLOOKUP(Товары_исходные!C' + str(i) + ',\'Справочники\'!$A$2:$B$50,2,FALSE),Товары_исходные!C' + str(i) + ')"',
                is_formula=True))
            L.append(call(pr, f"D{i}",
                '"=PROPER(TRIM(Товары_исходные!D' + str(i) + '))"',
                is_formula=True))
            for col in ("E", "F"):
                L.append(call(pr, f"{col}{i}",
                    '"=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Товары_исходные!' + col + str(i) + ',\\" \\",\\"\\"),\\"руб.\\",\\"\\"),\\"руб\\",\\"\\"),\\",\\",\\".\\")),0)"',
                    is_formula=True))
            L.append(call(pr, f"G{i}",
                '"=IFERROR(VLOOKUP(Товары_исходные!G' + str(i) + ',\'Справочники\'!$D$2:$E$50,2,FALSE),\\"шт.\\")"',
                is_formula=True))
            L.append(call(pr, f"H{i}",
                '"=IF(F' + str(i) + '>=E' + str(i) + ',\\"OK\\",\\"Розница ниже закупки\\")"',
                is_formula=True))
            L.append(call(pr, f"I{i}",
                '"=IF(COUNTIF($A$2:$A$' + str(n) + ',A' + str(i) + ')=1,\\"OK\\",\\"Дубликат\\")"',
                is_formula=True))

    # --- Поставщики ---
    psrc = add_sheet("Поставщики_исходные")
    dump_rows(psrc, sup)
    psup = add_sheet("Поставщики_очищенные")
    if len(sup) > 1:
        for c, v in enumerate(sup[0], start=1):
            col = chr(ord("A") + c - 1)
            L.append(call(psup, f"{col}1", js(v)))
        for col, h in [("H", "Индекс"), ("I", "Адрес_остаток"),
                       ("J", "Проверка_ИНН"), ("K", "Проверка_Email")]:
            L.append(call(psup, f"{col}1", js(h)))
        n = len(sup)
        for i in range(2, n + 1):
            L.append(call(psup, f"A{i}", f'"=Поставщики_исходные!A{i}"', is_formula=True))
            L.append(call(psup, f"B{i}", f'"=PROPER(TRIM(Поставщики_исходные!B{i}))"', is_formula=True))
            L.append(call(psup, f"C{i}", f'"=SUMPRODUCT(MID(0&Поставщики_исходные!C{i},LARGE(INDEX(ISNUMBER(--MID(Поставщики_исходные!C{i},ROW(INDIRECT(\\"1:\\"&LEN(Поставщики_исходные!C{i}))),1))*ROW(INDIRECT(\\"1:\\"&LEN(Поставщики_исходные!C{i}))),0),ROW(INDIRECT(\\"1:\\"&LEN(Поставщики_исходные!C{i}))))+1,1)*10^ROW(INDIRECT(\\"1:\\"&LEN(Поставщики_исходные!C{i})))/10)"', is_formula=True))
            L.append(call(psup, f"D{i}", f'"=Поставщики_исходные!D{i}"', is_formula=True))
            L.append(call(psup, f"E{i}", f'"=Поставщики_исходные!E{i}"', is_formula=True))
            L.append(call(psup, f"F{i}", f'"=TRIM(Поставщики_исходные!F{i})"', is_formula=True))
            L.append(call(psup, f"G{i}", f'"=TRIM(Поставщики_исходные!G{i})"', is_formula=True))
            L.append(call(psup, f"H{i}", f'"=IF(ISNUMBER(VALUE(LEFT(TRIM(Поставщики_исходные!G{i}),6))),LEFT(TRIM(Поставщики_исходные!G{i}),6),\\"\\")"', is_formula=True))
            L.append(call(psup, f"I{i}", f'"=TRIM(MID(TRIM(Поставщики_исходные!G{i}),7,255))"', is_formula=True))
            L.append(call(psup, f"J{i}", f'"=IF(OR(LEN(C{i})=10,LEN(C{i})=12),\\"OK\\",\\"Ошибка ИНН\\")"', is_formula=True))
            L.append(call(psup, f"K{i}", f'"=IF(AND(ISNUMBER(FIND(\\"@\\",F{i})),ISNUMBER(FIND(\\".\\",MID(F{i},FIND(\\"@\\",F{i})+1,255)))),\\"OK\\",\\"Ошибка Email\\")"', is_formula=True))

    # --- Продажи ---
    ssrc = add_sheet("Продажи_исходные")
    dump_rows(ssrc, sal)

    # Dedup indices (keep first occurrence)
    seen = set()
    kept = []
    for orig_i, row in enumerate(sal[1:], start=2):
        if len(row) < 5:
            continue
        k = (row[0], row[1], row[4])
        if k in seen:
            continue
        seen.add(k)
        kept.append((orig_i, row))

    ssal = add_sheet("Продажи_очищенные")
    if kept:
        for c, v in enumerate(sal[0], start=1):
            col = chr(ord("A") + c - 1)
            L.append(call(ssal, f"{col}1", js(v)))
        for col, h in [("H", "Розничная_цена"), ("I", "Категория_товара"), ("J", "Проверка_цены")]:
            L.append(call(ssal, f"{col}1", js(h)))
        for new_i, (orig_i, _) in enumerate(kept, start=2):
            L.append(call(ssal, f"A{new_i}",
                f'"=DATE(RIGHT(Продажи_исходные!A{orig_i},4),LEFT(Продажи_исходные!A{orig_i},2),MID(Продажи_исходные!A{orig_i},4,2))"',
                is_formula=True))
            L.append(call(ssal, f"B{new_i}",
                f'"=\\"АРТ-\\"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!B{orig_i},\\"№\\",\\"\\"),\\"АРТ-\\",\\"\\"),\\"АРТ_\\",\\"\\"),\\"/\\",\\"-\\"),3)"',
                is_formula=True))
            L.append(call(ssal, f"C{new_i}",
                f'"=IFERROR(VALUE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!C{orig_i},\\" шт.\\",\\"\\"),\\" шт\\",\\"\\"),\\"шт.\\",\\"\\"),\\"шт\\",\\"\\"))),0)"',
                is_formula=True))
            L.append(call(ssal, f"D{new_i}",
                f'"=IF(LEFT(Продажи_исходные!D{orig_i},1)=\\"$\\",VALUE(MID(Продажи_исходные!D{orig_i},2,255))*95,IF(LEFT(Продажи_исходные!D{orig_i},1)=\\"€\\",VALUE(MID(Продажи_исходные!D{orig_i},2,255))*100,VALUE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!D{orig_i},\\"₽\\",\\"\\"),\\" \\",\\"\\"))))"',
                is_formula=True))
            L.append(call(ssal, f"E{new_i}", f'"=Продажи_исходные!E{orig_i}"', is_formula=True))
            L.append(call(ssal, f"F{new_i}", f'"=Продажи_исходные!F{orig_i}"', is_formula=True))
            L.append(call(ssal, f"G{new_i}",
                f'"=IFERROR(VLOOKUP(Продажи_исходные!G{orig_i},\'Справочники\'!$G$2:$H$50,2,FALSE),Продажи_исходные!G{orig_i})"',
                is_formula=True))
            L.append(call(ssal, f"H{new_i}",
                f'"=IFERROR(VLOOKUP(B{new_i},\'Товары_очищенные\'!$A$1:$G$1000,6,FALSE),\\"нет в каталоге\\")"',
                is_formula=True))
            L.append(call(ssal, f"I{new_i}",
                f'"=IFERROR(VLOOKUP(B{new_i},\'Товары_очищенные\'!$A$1:$G$1000,3,FALSE),\\"нет в каталоге\\")"',
                is_formula=True))
            L.append(call(ssal, f"J{new_i}",
                f'"=IF(D{new_i}>0,\\"OK\\",\\"Некорректная цена\\")"',
                is_formula=True))

    # --- Клиенты ---
    csrc = add_sheet("Клиенты_исходные")
    dump_rows(csrc, cli)
    ccli = add_sheet("Клиенты_очищенные")
    if len(cli) > 1:
        for c, v in enumerate(cli[0], start=1):
            col = chr(ord("A") + c - 1)
            L.append(call(ccli, f"{col}1", js(v)))
        for col, h in [("H", "Возраст"), ("I", "Проверка_возраста"), ("J", "Проверка_Email")]:
            L.append(call(ccli, f"{col}1", js(h)))
        n = len(cli)
        for i in range(2, n + 1):
            L.append(call(ccli, f"A{i}", f'"=Клиенты_исходные!A{i}"', is_formula=True))
            L.append(call(ccli, f"B{i}", f'"=PROPER(TRIM(Клиенты_исходные!B{i}))"', is_formula=True))
            L.append(call(ccli, f"C{i}",
                f'"=IFERROR(IF(ISNUMBER(FIND(\\".\\",Клиенты_исходные!C{i})),DATE(RIGHT(Клиенты_исходные!C{i},4),MID(Клиенты_исходные!C{i},4,2),LEFT(Клиенты_исходные!C{i},2)),DATE(LEFT(Клиенты_исходные!C{i},4),MID(Клиенты_исходные!C{i},6,2),RIGHT(Клиенты_исходные!C{i},2))),Клиенты_исходные!C{i})"',
                is_formula=True))
            L.append(call(ccli, f"D{i}", f'"=Клиенты_исходные!D{i}"', is_formula=True))
            L.append(call(ccli, f"E{i}", f'"=TRIM(Клиенты_исходные!E{i})"', is_formula=True))
            L.append(call(ccli, f"F{i}",
                f'"=IFERROR(VLOOKUP(Клиенты_исходные!F{i},\'Справочники\'!$J$2:$K$50,2,FALSE),Клиенты_исходные!F{i})"',
                is_formula=True))
            L.append(call(ccli, f"G{i}",
                f'"=IFERROR(IF(ISNUMBER(FIND(\\".\\",Клиенты_исходные!G{i})),DATE(RIGHT(Клиенты_исходные!G{i},4),MID(Клиенты_исходные!G{i},4,2),LEFT(Клиенты_исходные!G{i},2)),DATE(LEFT(Клиенты_исходные!G{i},4),MID(Клиенты_исходные!G{i},6,2),RIGHT(Клиенты_исходные!G{i},2))),Клиенты_исходные!G{i})"',
                is_formula=True))
            L.append(call(ccli, f"H{i}", f'"=DATEDIF(C{i},TODAY(),\\"Y\\")"', is_formula=True))
            L.append(call(ccli, f"I{i}",
                f'"=IF(OR(H{i}<18,H{i}>80),\\"Ошибка даты рождения\\",\\"OK\\")"',
                is_formula=True))
            L.append(call(ccli, f"J{i}",
                f'"=IF(AND(ISNUMBER(FIND(\\"@\\",E{i})),ISNUMBER(FIND(\\".\\",MID(E{i},FIND(\\"@\\",E{i})+1,255)))),\\"OK\\",\\"Ошибка Email\\")"',
                is_formula=True))

    # --- Итоговый_отчет ---
    rep = add_sheet("Итоговый_отчет")
    L.append(call(rep, "A1", js("Файл")))
    L.append(call(rep, "B1", js("Строк до")))
    L.append(call(rep, "C1", js("Строк после")))
    L.append(call(rep, "D1", js("Удалено дублей")))
    for r, (label, s1, s2) in enumerate([
        ("Товары", "Товары_исходные", "Товары_очищенные"),
        ("Поставщики", "Поставщики_исходные", "Поставщики_очищенные"),
        ("Продажи", "Продажи_исходные", "Продажи_очищенные"),
        ("Клиенты", "Клиенты_исходные", "Клиенты_очищенные"),
    ], start=2):
        L.append(call(rep, f"A{r}", js(label)))
        L.append(call(rep, f"B{r}", f'"=COUNTA(\'{s1}\'!A:A)-1"', is_formula=True))
        L.append(call(rep, f"C{r}", f'"=COUNTA(\'{s2}\'!A:A)-1"', is_formula=True))
        L.append(call(rep, f"D{r}", f'"=B{r}-C{r}"', is_formula=True))

    L.append(call(rep, "A7", js("Ошибки по типам")))
    L.append(call(rep, "B7", js("Кол-во")))
    checks = [
        ("Ошибка ИНН", '=COUNTIF(\'Поставщики_очищенные\'!J:J,"Ошибка ИНН")'),
        ("Ошибка Email (поставщики)", '=COUNTIF(\'Поставщики_очищенные\'!K:K,"Ошибка Email")'),
        ("Ошибка Email (клиенты)", '=COUNTIF(\'Клиенты_очищенные\'!J:J,"Ошибка Email")'),
        ("Ошибка даты рождения", '=COUNTIF(\'Клиенты_очищенные\'!I:I,"Ошибка даты рождения")'),
        ("Розница ниже закупки", '=COUNTIF(\'Товары_очищенные\'!H:H,"Розница ниже закупки")'),
        ("Дубли артикулов", '=COUNTIF(\'Товары_очищенные\'!I:I,"Дубликат")'),
    ]
    for r, (label, formula) in enumerate(checks, start=8):
        L.append(call(rep, f"A{r}", js(label)))
        L.append(call(rep, f"B{r}", js(formula), is_formula=True))

    # --- Инструкция ---
    ins = add_sheet("Инструкция")
    for r, txt in enumerate([
        "Инструкция к файлу",
        "",
        "Сборка выполнена R7 Document Builder (documentconstructor) — родной инструмент Р7.",
        "Формулы записаны через SetFormula, данные — через SetValue.",
        "",
        "Ключевые функции: SUBSTITUTE, PROPER, TRIM, VLOOKUP, FIND, LEFT,",
        "SUMPRODUCT+MID+ROW+INDIRECT (извлечение цифр), DATE, DATEDIF, IFERROR,",
        "COUNTIF, COUNTA.",
        "",
        "Проверка: откройте Результат.xlsx, дайте пересчитать (F9), сверьте",
        "соответствие *_исходные / *_очищенные и Итоговый_отчет.",
    ], start=1):
        L.append(call(ins, f"A{r}", js(txt)))

    L.append('builder.SaveFile("xlsx", "Результат.xlsx");')
    L.append("builder.CloseFile();")
    return "\n".join(L) + "\n"


def find_binary():
    for name in ("documentconstructor", "docbuilder"):
        p = shutil.which(name)
        if p:
            return p
    return None


def main():
    binary = find_binary()
    if not binary:
        print("ERROR: neither 'documentconstructor' nor 'docbuilder' found in PATH.", file=sys.stderr)
        print("       Check the R7 Document Builder install steps.", file=sys.stderr)
        sys.exit(2)

    script = gen_docbuilder()
    with tempfile.NamedTemporaryFile("w", suffix=".docbuilder",
                                     delete=False, encoding="utf-8") as f:
        f.write(script)
        path = f.name

    print(f"Generated {path} ({len(script)} bytes)")
    print(f"Running: {binary} {path}")
    try:
        subprocess.run([binary, path], check=True, cwd=os.getcwd())
    finally:
        os.unlink(path)

    out = os.path.abspath("Результат.xlsx")
    if os.path.exists(out):
        print(f"OK: {out} ({os.path.getsize(out)} bytes)")
    else:
        print("ERROR: Результат.xlsx not produced.", file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
