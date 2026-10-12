#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собирает со справки Microsoft все статистические функции Excel:
название, описание, синтаксис, аргументы, замечания, ссылку и пример-таблицу.
Результат: один файл excel_statistical_functions.xlsx
  * лист «Все функции» - сводная таблица (+ ссылки на страницы и на листы с примерами)
  * по одному листу на каждую функцию - пример-таблица в том же виде, что у Microsoft,
    плюс колонка «Результат (Excel)» с НАСТОЯЩЕЙ формулой, которую Excel считает сам.

Запуск:   python build_excel.py
Пробный:  python build_excel.py --limit 5
"""
import argparse
import datetime
import pathlib
import re
import sys
import time
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

INDEX_URL = "https://support.microsoft.com/ru-ru/excel/statistical-functions-reference"
BASE_FUNC_URL = "https://support.microsoft.com/ru-ru/excel/functions/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9",
}
CACHE = pathlib.Path("cache")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E2F3")
LIVE_FILL = PatternFill("solid", fgColor="E2EFDA")


# ----------------------------------------------------------------- загрузка
def fetch(url):
    """Скачивает страницу (с кэшем в папке cache/, чтобы не качать дважды)."""
    url = url.split("#")[0]
    CACHE.mkdir(exist_ok=True)
    p = CACHE / (re.sub(r"[^A-Za-z0-9]+", "_", url)[-150:] + ".html")
    if p.exists():
        return p.read_text(encoding="utf-8")
    err = None
    for _ in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
            r.raise_for_status()
            r.encoding = "utf-8"
            p.write_text(r.text, encoding="utf-8")
            time.sleep(0.4)  # вежливая пауза
            return r.text
        except Exception as e:  # noqa
            err = e
            time.sleep(2)
    raise err


def clean(s):
    s = (s or "").replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", s).strip()


# ------------------------------------------------------------ список функций
def en_name(href):
    u = urlparse(href)
    frag = u.fragment
    slug = u.path.rstrip("/").split("/")[-1]
    special = {
        "forecast-and-forecast-linear-functions": "FORECAST",
        "use-the-countif-function-in-microsoft-excel": "COUNTIF",
    }
    if frag:
        return frag.lstrip("_").upper()
    if slug in special:
        return special[slug]
    if slug.endswith("-function"):
        return slug[:-9].replace("-", ".").upper()
    return ""


def page_url(href):
    u = urlparse(href)
    if u.fragment:  # напр. ...forecasting-functions-reference#_forecast.ets
        slug = u.fragment.lstrip("_").replace(".", "-")
        return BASE_FUNC_URL + slug + "-function"
    return href.split("#")[0]


def parse_index(html):
    soup = BeautifulSoup(html, "html.parser")
    funcs, seen = [], set()
    for tr in soup.find_all("tr"):
        tds = tr.find_all(["td", "th"])
        if len(tds) < 2:
            continue
        a = tds[0].find("a", href=True)
        if not a:
            continue
        name = re.sub(r"^Функция\s+", "", clean(a.get_text()))
        href = urljoin(INDEX_URL, a["href"])
        if (name, href) in seen:
            continue
        seen.add((name, href))
        funcs.append({
            "name": name,
            "en": en_name(href),
            "desc": clean(tds[1].get_text(" ")),
            "index_href": href,
            "url": page_url(href),
        })
    return funcs


def parse_nav_names(html):
    """Берёт со страницы функции боковой список «Связанные темы»:
    русское название -> английское (по адресу ссылки). Нужен для перевода формул."""
    soup = BeautifulSoup(html, "html.parser")
    out = {}
    for a in soup.find_all("a", href=True):
        m = re.search(r"/excel/functions/([a-z0-9\-]+)-function$", a["href"])
        if not m:
            continue
        ru = re.sub(r"^Функция\s+", "", clean(a.get_text()))
        ru = re.sub(r"\.\s+", ".", ru).upper()
        if re.fullmatch(r"[\w.]+", ru):
            out.setdefault(ru, m.group(1).replace("-", ".").upper())
    return out


# ------------------------------------------------------------ страница функции
def table_to_grid(table):
    grid = []
    for tr in table.find_all("tr"):
        row = []
        for c in tr.find_all(["td", "th"]):
            txt = clean(c.get_text(" "))
            bold = c.name == "th" or bool(c.find(["b", "strong"]))
            try:
                span = int(c.get("colspan", 1) or 1)
            except ValueError:
                span = 1
            row.append((txt, bold))
            row.extend([("", False)] * (span - 1))
        grid.append(row)
    w = max((len(r) for r in grid), default=0)
    for r in grid:
        r.extend([("", False)] * (w - len(r)))
    # выкидываем полностью пустые строки
    return [r for r in grid if any(t for t, _ in r)]


def parse_function_page(html):
    soup = BeautifulSoup(html, "html.parser")
    h1 = soup.find("h1")
    res = {"title": clean(h1.get_text()) if h1 else "", "desc": "", "syntax": "",
           "args": [], "remarks": [], "tables": []}
    if not h1:
        return res
    sections, cur = {"_intro": []}, "_intro"
    for el in h1.find_all_next(["h2", "h3", "p", "li", "table"]):
        if el.name in ("h2", "h3"):
            t = clean(el.get_text()).lower()
            if t.startswith("полезны ли"):
                break
            cur = t
            sections.setdefault(cur, [])
            continue
        if el.name != "table" and el.find_parent("table"):
            continue
        if el.name == "p" and el.find_parent("li"):
            continue
        sections[cur].append(el)

    def keys(*prefixes):
        return [k for k in sections if any(k.startswith(p) for p in prefixes)]

    def texts(ks, names=("p", "li")):
        out = []
        for k in ks:
            for el in sections[k]:
                if el.name in names:
                    t = clean(el.get_text(" "))
                    if t:
                        out.append(t)
        return out

    res["desc"] = " ".join(texts(keys("описание")))
    syn_keys = keys("синтаксис")
    ps = texts(syn_keys, ("p",))
    cand = [p for p in ps if "(" in p and not p.lower().startswith("аргумент")]
    res["syntax"] = (cand or ps or [""])[0]
    res["args"] = texts(syn_keys + keys("аргумент"), ("li",))
    if not res["args"]:  # аргументы иногда лежат в таблице
        for k in syn_keys + keys("аргумент"):
            for el in sections[k]:
                if el.name == "table":
                    for r in table_to_grid(el):
                        res["args"].append(" — ".join(t for t, _ in r if t))
    res["remarks"] = texts(keys("замечани", "примечани"))
    ex_keys = keys("пример")
    tabs = [el for k in ex_keys for el in sections[k] if el.name == "table"]
    if not tabs:
        tabs = [el for k, v in sections.items() if k not in syn_keys
                for el in v if el.name == "table"]
    res["tables"] = [g for g in (table_to_grid(t) for t in tabs) if g]
    return res


# ------------------------------------------------- перевод формул RU -> EN
# Функции, появившиеся после Excel 2007 и записываемые в xlsx с префиксом _xlfn.
# (все функции с точкой в названии - STDEV.S, NORM.DIST и т.д. - тоже новые)
NEW_NO_DOT = {"GAMMA", "GAUSS", "PHI", "PERMUTATIONA", "MAXIFS", "MINIFS", "IFS", "SWITCH",
              "CONCAT", "TEXTJOIN", "IFNA", "XOR", "DAYS", "ISOWEEKNUM", "XLOOKUP", "XMATCH",
              "NUMBERVALUE", "UNICHAR", "UNICODE", "FORMULATEXT", "ISFORMULA", "SHEET", "SHEETS",
              "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT", "COMBINA", "PDURATION",
              "RRI", "DECIMAL", "ARABIC", "BASE", "ACOT", "ACOTH", "COT", "COTH", "CSC", "CSCH",
              "SEC", "SECH", "MUNIT", "RRI", "IMSEC", "IMSECH", "IMCSC", "IMCSCH", "IMCOT",
              "IMSIN", "IMCOS"}


def xl_name(en):
    return "_xlfn." + en if ("." in en or en in NEW_NO_DOT) else en


class Translator:
    def __init__(self, name_map):
        self.map = {k.upper(): xl_name(v) for k, v in name_map.items()}
        names = sorted(self.map, key=len, reverse=True)
        self.re = re.compile(r"(?<![\w.])(" + "|".join(re.escape(n) for n in names) + r")\(",
                             re.I) if names else None

    def __call__(self, f):
        """Возвращает формулу на английском для записи в xlsx или None, если не получилось."""
        f = f.strip()
        if not f.startswith("=") or "{" in f:
            return None
        for a, b in (("“", '"'), ("”", '"'), ("«", '"'), ("»", '"'), ("−", "-"), ("–", "-")):
            f = f.replace(a, b)
        parts = re.split(r'("[^"]*")', f)
        out = []
        for i, p in enumerate(parts):
            if i % 2:
                out.append(p)
                continue
            p = re.sub(r"(?<=\d),(?=\d)", ".", p)  # десятичная запятая
            p = p.replace(";", ",")
            if self.re:
                p = self.re.sub(lambda m: self.map[m.group(1).upper()] + "(", p)
            p = re.sub(r"\bИСТИНА\b", "TRUE", p, flags=re.I)
            p = re.sub(r"\bЛОЖЬ\b", "FALSE", p, flags=re.I)
            out.append(p)
        res = "".join(out)
        outside = "".join(out[0::2])
        if re.search(r"[А-Яа-яЁё]", outside) or res.count("(") != res.count(")"):
            return None
        return res


def to_value(t):
    t = t.strip()
    if re.fullmatch(r"-?\d+", t):
        return int(t)
    if re.fullmatch(r"-?\d+[.,]\d+", t):
        return float(t.replace(",", "."))
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", t)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return t


def set_text(cell, text):
    cell.value = text
    cell.data_type = "s"  # чтобы «=СРЗНАЧ(...)» осталось текстом, а не формулой


# --------------------------------------------------------- построение книги
def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:28] or "func"
    n, i = base, 2
    while n.lower() in used:
        n = f"{base}_{i}"
        i += 1
    used.add(n.lower())
    return n


def write_function_sheet(wb, sheet_name, rec, translate):
    """Рисует пример-таблицу, начиная с A1 (чтобы ссылки A2:A6 из примера работали)."""
    ws = wb.create_sheet(sheet_name)
    status = []
    if not rec["tables"]:
        ws["A1"] = "На странице Microsoft нет примера-таблицы"
        return ws, "", "нет примера"
    grid = rec["tables"][0]
    ncols = len(grid[0])
    hdr = next((i for i, r in enumerate(grid)
                if any(t.lower() == "формула" for t, _ in r)), None)
    live_col = ncols + 1
    translated_any = failed = 0
    for ri, row in enumerate(grid, start=1):
        row_has_formula_text = False
        for ci, (txt, bold) in enumerate(row, start=1):
            c = ws.cell(ri, ci)
            if txt.startswith("="):
                if hdr is not None and ri - 1 < hdr:
                    f = translate(txt)  # формула среди исходных данных - считаем на месте
                    if f:
                        c.value = f
                        translated_any += 1
                    else:
                        set_text(c, txt)
                        failed += 1
                else:
                    set_text(c, txt)
                    row_has_formula_text = True
            elif txt:
                c.value = to_value(txt)
            c.font = Font(bold=bold or (hdr is not None and ri - 1 == hdr))
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BORDER
            if hdr is not None and ri - 1 == hdr:
                c.fill = HEAD_FILL
        if hdr is not None and ri - 1 == hdr:
            lc = ws.cell(ri, live_col, "Результат (Excel)")
            lc.font = Font(bold=True)
            lc.fill = LIVE_FILL
            lc.border = BORDER
        if row_has_formula_text:
            ftxt = next(t for t, _ in row if t.startswith("="))
            f = translate(ftxt)
            lc = ws.cell(ri, live_col)
            lc.fill = LIVE_FILL
            lc.border = BORDER
            if f:
                lc.value = f
                translated_any += 1
            else:
                lc.value = "не пересчитано (массивная/нестандартная формула)"
                failed += 1
    if hdr is None and not translated_any:
        # таблица без заголовка «Формула» - просто показываем её как есть
        pass
    for ci in range(1, live_col + 1):
        L = get_column_letter(ci)
        width = max((len(str(ws.cell(r, ci).value or "")) for r in range(1, len(grid) + 1)),
                    default=10)
        ws.column_dimensions[L].width = min(max(12, width + 2), 55)
    # дополнительные таблицы - ниже, как обычные значения
    r0 = len(grid) + 3
    for extra in rec["tables"][1:]:
        ws.cell(r0 - 1, 1, "Ещё одна таблица со страницы (без пересчёта):").font = Font(italic=True)
        for ri, row in enumerate(extra):
            for ci, (txt, bold) in enumerate(row, start=1):
                c = ws.cell(r0 + ri, ci)
                set_text(c, txt) if txt.startswith("=") else setattr(c, "value", to_value(txt) if txt else None)
                c.font = Font(bold=bold)
                c.border = BORDER
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r0 += len(extra) + 2
    rng = f"A1:{get_column_letter(live_col)}{len(grid)}"
    if failed:
        status.append("часть формул не пересчитана")
    return ws, rng, "; ".join(status) or "OK"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="excel_statistical_functions.xlsx")
    ap.add_argument("--limit", type=int, default=0, help="для проверки: только первые N функций")
    args = ap.parse_args()

    print("Загружаю список функций...")
    funcs = parse_index(fetch(INDEX_URL))
    if args.limit:
        funcs = funcs[: args.limit]
    print(f"Найдено функций: {len(funcs)}")

    name_map, records = {}, []
    for i, fn in enumerate(funcs, 1):
        print(f"[{i}/{len(funcs)}] {fn['name']}", flush=True)
        rec = {"error": ""}
        try:
            html = fetch(fn["url"])
            if not name_map:
                name_map.update(parse_nav_names(html))
            rec.update(parse_function_page(html))
        except Exception as e:  # noqa
            rec.update({"title": "", "desc": "", "syntax": "", "args": [], "remarks": [],
                        "tables": [], "error": f"ошибка загрузки: {e}"})
        records.append(rec)
    for fn in funcs:  # названия из списка надёжнее, чем из бокового меню
        if fn["en"]:
            name_map[fn["name"].upper()] = fn["en"]
    translate = Translator(name_map)

    wb = Workbook()
    ms = wb.active
    ms.title = "Все функции"
    heads = ["№", "Функция", "Англ. название", "Описание", "Синтаксис", "Аргументы",
             "Замечания", "Ссылка на справку", "Лист с примером", "Диапазон таблицы", "Статус"]
    ms.append(heads)
    for c in ms[1]:
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")

    used = {"все функции"}
    for i, (fn, rec) in enumerate(zip(funcs, records), 1):
        sheet = safe_sheet_name(fn["name"], used)
        ws, rng, status = write_function_sheet(wb, sheet, rec, translate)
        if rec["error"]:
            status = rec["error"]
        # справка справа от таблицы на листе функции
        col = (ws.max_column or 1) + 2
        info = [("Функция", fn["name"]), ("Англ. название", fn["en"]),
                ("Описание", rec["desc"] or fn["desc"]), ("Синтаксис", rec["syntax"]),
                ("Аргументы", "\n".join(rec["args"])), ("Замечания", "\n".join(rec["remarks"])),
                ("Ссылка", fn["url"])]
        for k, (lab, val) in enumerate(info, start=1):
            a = ws.cell(k, col, lab)
            a.font = Font(bold=True)
            a.alignment = Alignment(vertical="top")
            b = ws.cell(k, col + 1, val)
            b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(col)].width = 16
        ws.column_dimensions[get_column_letter(col + 1)].width = 80

        ms.append([i, fn["name"], fn["en"], rec["desc"] or fn["desc"], rec["syntax"],
                   "\n".join(rec["args"]), "\n".join(rec["remarks"]), fn["url"], sheet, rng,
                   status])
        r = ms.max_row
        ms.cell(r, 8).hyperlink = fn["url"]
        ms.cell(r, 8).font = Font(color="0563C1", underline="single")
        ms.cell(r, 9).hyperlink = f"#'{sheet}'!A1"
        ms.cell(r, 9).font = Font(color="0563C1", underline="single")
        for c in ms[r]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BORDER

    for L, w in zip("ABCDEFGHIJK", (5, 22, 20, 50, 38, 60, 60, 45, 20, 14, 28)):
        ms.column_dimensions[L].width = w
    ms.freeze_panes = "C2"
    ms.auto_filter.ref = ms.dimensions
    wb.save(args.out)
    bad = [fn["name"] for fn, rec in zip(funcs, records) if rec["error"]]
    print(f"\nГотово: {args.out}")
    if bad:
        print("Не загрузились (запустите скрипт ещё раз - уже скачанные берутся из кэша):", ", ".join(bad))


if __name__ == "__main__":
    sys.exit(main())
