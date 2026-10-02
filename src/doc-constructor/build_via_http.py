#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build Результат.xlsx via R7 dc-httpd (multipart /run)."""
import os
import sys
import time

import requests

try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:
    pass

INPUTS = {
    "products":  "file_1_Товары каталог.txt",
    "suppliers": "file_2_Поставщики.txt",
    "sales":     "file_3_Продажи.txt",
    "clients":   "file_4_Клиенты.txt",
}

BASE = "http://docbuilder:8080"
SCRIPT_LOCAL = "/workdir/script.docbuilder"
# dc-httpd writes here (bind-mounted to ./workdir on host and to /workdir in client)
RESULT_IN_SERVER = "/app/files/Result.xlsx"
RESULT_CLIENT_TMP = "/workdir/Result.xlsx"
RESULT_FINAL = "/workdir/Результат.xlsx"


def find_input(name):
    for base in ("/project", "/project/data", "."):
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


def F(formula):
    """Wrap an Excel formula (a Python string) as a JS string literal."""
    return js(formula)


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
                L.append(f'{target_var}.GetRange("{col}{r}").SetValue({js(val)});')

    def add_sheet(name):
        v = f"s{len(L)}"
        L.append(f'var {v} = Api.AddSheet({js(name)});')
        return v

    def put(sh, cell, value):
        L.append(f'{sh}.GetRange("{cell}").SetValue({js(value)});')

    def formula(sh, cell, formula):
        L.append(f'{sh}.GetRange("{cell}").SetFormula({F(formula)});')

    # --- Справочники ---
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
            put(ref, f"{c1}{r}", row[0])
            put(ref, f"{c2}{r}", row[1])

    # --- Товары_исходные ---
    dump_rows("s0", prod)

    # --- Товары_очищенные ---
    pr = add_sheet("Товары_очищенные")
    if len(prod) > 1:
        for c, v in enumerate(prod[0], start=1):
            col = chr(ord("A") + c - 1)
            put(pr, f"{col}1", v)
        put(pr, "H1", "Проверка_цены")
        put(pr, "I1", "Проверка_уникальности")
        n = len(prod)
        for i in range(2, n + 1):
            formula(pr, f"A{i}",
                f'="АРТ-"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Товары_исходные!A{i},"№",""),"АРТ-",""),"АРТ_",""),"/","-"),3)')
            formula(pr, f"B{i}",
                f'=PROPER(TRIM(IFERROR(LEFT(Товары_исходные!B{i},FIND("(",Товары_исходные!B{i})-1),Товары_исходные!B{i})))')
            formula(pr, f"C{i}",
                f"=IFERROR(VLOOKUP(Товары_исходные!C{i},'Справочники'!$A$2:$B$50,2,FALSE),Товары_исходные!C{i})")
            formula(pr, f"D{i}",
                f'=PROPER(TRIM(Товары_исходные!D{i}))')
            for col in ("E", "F"):
                formula(pr, f"{col}{i}",
                    f'=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Товары_исходные!{col}{i}," ",""),"руб.",""),"руб",""),",",".")),0)')
            formula(pr, f"G{i}",
                f"=IFERROR(VLOOKUP(Товары_исходные!G{i},'Справочники'!$D$2:$E$50,2,FALSE),\"шт.\")")
            formula(pr, f"H{i}",
                f'=IF(F{i}>=E{i},"OK","Розница ниже закупки")')
            formula(pr, f"I{i}",
                f'=IF(COUNTIF($A$2:$A${n},A{i})=1,"OK","Дубликат")')

    # --- Поставщики ---
    psrc = add_sheet("Поставщики_исходные")
    dump_rows(psrc, sup)
    psup = add_sheet("Поставщики_очищенные")
    if len(sup) > 1:
        for c, v in enumerate(sup[0], start=1):
            col = chr(ord("A") + c - 1)
            put(psup, f"{col}1", v)
        for col, h in [("H", "Индекс"), ("I", "Адрес_остаток"),
                       ("J", "Проверка_ИНН"), ("K", "Проверка_Email")]:
            put(psup, f"{col}1", h)
        n = len(sup)
        for i in range(2, n + 1):
            formula(psup, f"A{i}", f"=Поставщики_исходные!A{i}")
            formula(psup, f"B{i}", f"=PROPER(TRIM(Поставщики_исходные!B{i}))")
            formula(psup, f"C{i}",
                f'=SUMPRODUCT(MID(0&Поставщики_исходные!C{i},LARGE(INDEX(ISNUMBER(--MID(Поставщики_исходные!C{i},ROW(INDIRECT("1:"&LEN(Поставщики_исходные!C{i}))),1))*ROW(INDIRECT("1:"&LEN(Поставщики_исходные!C{i}))),0),ROW(INDIRECT("1:"&LEN(Поставщики_исходные!C{i}))))+1,1)*10^ROW(INDIRECT("1:"&LEN(Поставщики_исходные!C{i})))/10)')
            formula(psup, f"D{i}", f"=Поставщики_исходные!D{i}")
            formula(psup, f"E{i}", f"=Поставщики_исходные!E{i}")
            formula(psup, f"F{i}", f"=TRIM(Поставщики_исходные!F{i})")
            formula(psup, f"G{i}", f"=TRIM(Поставщики_исходные!G{i})")
            formula(psup, f"H{i}",
                f'=IF(ISNUMBER(VALUE(LEFT(TRIM(Поставщики_исходные!G{i}),6))),LEFT(TRIM(Поставщики_исходные!G{i}),6),"")')
            formula(psup, f"I{i}",
                f"=TRIM(MID(TRIM(Поставщики_исходные!G{i}),7,255))")
            formula(psup, f"J{i}",
                f'=IF(OR(LEN(C{i})=10,LEN(C{i})=12),"OK","Ошибка ИНН")')
            formula(psup, f"K{i}",
                f'=IF(AND(ISNUMBER(FIND("@",F{i})),ISNUMBER(FIND(".",MID(F{i},FIND("@",F{i})+1,255)))),"OK","Ошибка Email")')

    # --- Продажи ---
    ssrc = add_sheet("Продажи_исходные")
    dump_rows(ssrc, sal)

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
            put(ssal, f"{col}1", v)
        for col, h in [("H", "Розничная_цена"), ("I", "Категория_товара"), ("J", "Проверка_цены")]:
            put(ssal, f"{col}1", h)
        for new_i, (orig_i, _) in enumerate(kept, start=2):
            formula(ssal, f"A{new_i}",
                f"=DATE(RIGHT(Продажи_исходные!A{orig_i},4),LEFT(Продажи_исходные!A{orig_i},2),MID(Продажи_исходные!A{orig_i},4,2))")
            formula(ssal, f"B{new_i}",
                f'="АРТ-"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!B{orig_i},"№",""),"АРТ-",""),"АРТ_",""),"/","-"),3)')
            formula(ssal, f"C{new_i}",
                f'=IFERROR(VALUE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!C{orig_i}," шт.","")," шт",""),"шт.",""),"шт",""))),0)')
            formula(ssal, f"D{new_i}",
                f'=IF(LEFT(Продажи_исходные!D{orig_i},1)="$",VALUE(MID(Продажи_исходные!D{orig_i},2,255))*95,IF(LEFT(Продажи_исходные!D{orig_i},1)="€",VALUE(MID(Продажи_исходные!D{orig_i},2,255))*100,VALUE(SUBSTITUTE(SUBSTITUTE(Продажи_исходные!D{orig_i},"₽","")," ",""))))')
            formula(ssal, f"E{new_i}", f"=Продажи_исходные!E{orig_i}")
            formula(ssal, f"F{new_i}", f"=Продажи_исходные!F{orig_i}")
            formula(ssal, f"G{new_i}",
                f"=IFERROR(VLOOKUP(Продажи_исходные!G{orig_i},'Справочники'!$G$2:$H$50,2,FALSE),Продажи_исходные!G{orig_i})")
            formula(ssal, f"H{new_i}",
                f"=IFERROR(VLOOKUP(B{new_i},'Товары_очищенные'!$A$1:$G$1000,6,FALSE),\"нет в каталоге\")")
            formula(ssal, f"I{new_i}",
                f"=IFERROR(VLOOKUP(B{new_i},'Товары_очищенные'!$A$1:$G$1000,3,FALSE),\"нет в каталоге\")")
            formula(ssal, f"J{new_i}",
                f'=IF(D{new_i}>0,"OK","Некорректная цена")')

    # --- Клиенты ---
    csrc = add_sheet("Клиенты_исходные")
    dump_rows(csrc, cli)
    ccli = add_sheet("Клиенты_очищенные")
    if len(cli) > 1:
        for c, v in enumerate(cli[0], start=1):
            col = chr(ord("A") + c - 1)
            put(ccli, f"{col}1", v)
        for col, h in [("H", "Возраст"), ("I", "Проверка_возраста"), ("J", "Проверка_Email")]:
            put(ccli, f"{col}1", h)
        n = len(cli)
        for i in range(2, n + 1):
            formula(ccli, f"A{i}", f"=Клиенты_исходные!A{i}")
            formula(ccli, f"B{i}", f"=PROPER(TRIM(Клиенты_исходные!B{i}))")
            formula(ccli, f"C{i}",
                f'=IFERROR(IF(ISNUMBER(FIND(".",Клиенты_исходные!C{i})),DATE(RIGHT(Клиенты_исходные!C{i},4),MID(Клиенты_исходные!C{i},4,2),LEFT(Клиенты_исходные!C{i},2)),DATE(LEFT(Клиенты_исходные!C{i},4),MID(Клиенты_исходные!C{i},6,2),RIGHT(Клиенты_исходные!C{i},2))),Клиенты_исходные!C{i})')
            formula(ccli, f"D{i}", f"=Клиенты_исходные!D{i}")
            formula(ccli, f"E{i}", f"=TRIM(Клиенты_исходные!E{i})")
            formula(ccli, f"F{i}",
                f"=IFERROR(VLOOKUP(Клиенты_исходные!F{i},'Справочники'!$J$2:$K$50,2,FALSE),Клиенты_исходные!F{i})")
            formula(ccli, f"G{i}",
                f'=IFERROR(IF(ISNUMBER(FIND(".",Клиенты_исходные!G{i})),DATE(RIGHT(Клиенты_исходные!G{i},4),MID(Клиенты_исходные!G{i},4,2),LEFT(Клиенты_исходные!G{i},2)),DATE(LEFT(Клиенты_исходные!G{i},4),MID(Клиенты_исходные!G{i},6,2),RIGHT(Клиенты_исходные!G{i},2))),Клиенты_исходные!G{i})')
            formula(ccli, f"H{i}", f'=DATEDIF(C{i},TODAY(),"Y")')
            formula(ccli, f"I{i}",
                f'=IF(OR(H{i}<18,H{i}>80),"Ошибка даты рождения","OK")')
            formula(ccli, f"J{i}",
                f'=IF(AND(ISNUMBER(FIND("@",E{i})),ISNUMBER(FIND(".",MID(E{i},FIND("@",E{i})+1,255)))),"OK","Ошибка Email")')

    # --- Итоговый_отчет ---
    rep = add_sheet("Итоговый_отчет")
    put(rep, "A1", "Файл")
    put(rep, "B1", "Строк до")
    put(rep, "C1", "Строк после")
    put(rep, "D1", "Удалено дублей")
    for r, (label, s1, s2) in enumerate([
        ("Товары", "Товары_исходные", "Товары_очищенные"),
        ("Поставщики", "Поставщики_исходные", "Поставщики_очищенные"),
        ("Продажи", "Продажи_исходные", "Продажи_очищенные"),
        ("Клиенты", "Клиенты_исходные", "Клиенты_очищенные"),
    ], start=2):
        put(rep, f"A{r}", label)
        formula(rep, f"B{r}", f"=COUNTA('{s1}'!A:A)-1")
        formula(rep, f"C{r}", f"=COUNTA('{s2}'!A:A)-1")
        formula(rep, f"D{r}", f"=B{r}-C{r}")

    put(rep, "A7", "Ошибки по типам")
    put(rep, "B7", "Кол-во")
    checks = [
        ("Ошибка ИНН", "=COUNTIF('Поставщики_очищенные'!J:J,\"Ошибка ИНН\")"),
        ("Ошибка Email (поставщики)", "=COUNTIF('Поставщики_очищенные'!K:K,\"Ошибка Email\")"),
        ("Ошибка Email (клиенты)", "=COUNTIF('Клиенты_очищенные'!J:J,\"Ошибка Email\")"),
        ("Ошибка даты рождения", "=COUNTIF('Клиенты_очищенные'!I:I,\"Ошибка даты рождения\")"),
        ("Розница ниже закупки", "=COUNTIF('Товары_очищенные'!H:H,\"Розница ниже закупки\")"),
        ("Дубли артикулов", "=COUNTIF('Товары_очищенные'!I:I,\"Дубликат\")"),
    ]
    for r, (label, fml) in enumerate(checks, start=8):
        put(rep, f"A{r}", label)
        formula(rep, f"B{r}", fml)

    # --- Инструкция ---
    ins = add_sheet("Инструкция")
    for r, txt in enumerate([
        "Инструкция к файлу",
        "",
        "Сборка выполнена R7 Document Builder (dc-httpd).",
        "Формулы записаны через SetFormula, данные — через SetValue.",
        "",
        "Ключевые функции: SUBSTITUTE, PROPER, TRIM, VLOOKUP, FIND, LEFT,",
        "SUMPRODUCT+MID+ROW+INDIRECT, DATE, DATEDIF, IFERROR, COUNTIF, COUNTA.",
    ], start=1):
        put(ins, f"A{r}", txt)

    # Save as Latin name into the shared volume. Renamed to «Результат.xlsx» by client.
    L.append(f'builder.SaveFile("xlsx", {js(RESULT_IN_SERVER)});')
    L.append("builder.CloseFile();")
    return "\n".join(L) + "\n"


def main():
    script = gen_docbuilder()
    with open(SCRIPT_LOCAL, "w", encoding="utf-8") as f:
        f.write(script)
    print(f"OK: wrote {SCRIPT_LOCAL} ({len(script)} bytes)", flush=True)

    # Wait for server.
    for attempt in range(30):
        try:
            r = requests.get(BASE + "/", timeout=5)
            if r.status_code < 500:
                print(f"server up (HTTP {r.status_code})", flush=True)
                break
        except Exception as e:
            print(f"waiting: {e!r}", flush=True)
        time.sleep(2)
    else:
        print("ERROR: server not reachable", file=sys.stderr, flush=True)
        sys.exit(2)

    # Clean stale.
    for p in (RESULT_CLIENT_TMP, RESULT_FINAL):
        if os.path.exists(p):
            os.unlink(p)

    print(f"POST {BASE}/run (multipart field 'filename')", flush=True)
    with open(SCRIPT_LOCAL, "rb") as fh:
        files = {"filename": ("script.docbuilder", fh, "application/octet-stream")}
        r = requests.post(BASE + "/run", files=files, timeout=900)

    print(f"HTTP {r.status_code} | body {len(r.content)} bytes | "
          f"CT={r.headers.get('Content-Type')}", flush=True)
    print(f"RAW (first 500): {r.content[:500]!r}", flush=True)

    # Case A: body IS the xlsx.
    if r.content[:2] == b"PK":
        with open(RESULT_FINAL, "wb") as out:
            out.write(r.content)
        print(f"OK (case A): {RESULT_FINAL} ({os.path.getsize(RESULT_FINAL)} bytes)", flush=True)
        return

    # Case B: dc-httpd wrote into its own /app/files (= our shared volume).
    for _ in range(60):
        if os.path.exists(RESULT_CLIENT_TMP) and os.path.getsize(RESULT_CLIENT_TMP) > 0:
            os.replace(RESULT_CLIENT_TMP, RESULT_FINAL)
            print(f"OK (case B): {RESULT_FINAL} ({os.path.getsize(RESULT_FINAL)} bytes)", flush=True)
            return
        time.sleep(2)

    print(f"ERROR: {RESULT_CLIENT_TMP} not produced.", file=sys.stderr, flush=True)
    print("Check: docker logs r7-docbuilder-server", file=sys.stderr, flush=True)
    print("And:   docker exec r7-docbuilder-server ls -la /app/files", file=sys.stderr, flush=True)
    sys.exit(3)


if __name__ == "__main__":
    main()
