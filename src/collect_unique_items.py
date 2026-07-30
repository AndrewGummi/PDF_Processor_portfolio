#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
collect_unique_items.py — збирає УНІКАЛЬНІ найменування майна з готових
відомостей і додає відсутні в аркуш "Словник" макета.

ЗВІДКИ БЕРЕ
Проходить усі .xlsx у папці готових відомостей (кожен файл — кілька
аркушів, кожен аркуш — окремий епізод). У рядку позиції:
    A = № з/п,  B = найменування,  C = одиниця виміру,  D = кількість
Трапляється зсув на одну колонку (найменування в C, одиниця в D) —
обидва варіанти розпізнаються за формою даних, а не за жорсткою
позицією: № з/п має бути числом, одиниця — коротким словом зі списку.

ЩО ВІДКИДАЄ
- хвости на кшталт "2025р.", "2024 р.", "(2025р.)" в кінці назви —
  це рік переоцінки, а не частина найменування;
- службові рядки відомості ("Всього:", "Примітка:", "Голова комісії");
- позиції, які вже є у "Словнику" (точний збіг або дуже близький).

ЯК ПОЗНАЧАЄ ДОДАНЕ
Нові рядки дописуються В КІНЕЦЬ аркуша — щоб колонки B (одиниця) і
C (термін служби) наявних рядків лишились на своїх місцях і не
роз'їхались відносно назв. У новому рядку заповнюються:
    A — найменування
    B — одиниця виміру (з тієї ж відомості, звідки взято назву)
    C — НЕ заповнюється: термін служби береться з наказу, вигадувати
        його не можна
    D — коментар: скільки разів трапилось і в скількох файлах
Рядок заливається кольором за категорією:
    зелений  — трапляється часто (>= 3 файлів), назва стабільна
    жовтий   — трапилось 1-2 рази, потребує перевірки людиною

ВИКОРИСТАННЯ
    python collect_unique_items.py <папка_відомостей> <макет.xlsm> [--apply]

Без --apply лише друкує звіт і пише CSV поруч з макетом (нічого не
змінює). З --apply дописує рядки у копію макета з суфіксом
"_зі_словником" — оригінал не чіпається ніколи.
"""

import csv
import json
import shutil
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
except ImportError:
    print("ПОМИЛКА: потрібен openpyxl (pip install openpyxl)")
    raise

try:
    from rapidfuzz import fuzz
except ImportError:
    print("ПОМИЛКА: потрібен rapidfuzz (pip install rapidfuzz)")
    raise

PROJECT_ROOT = Path(__file__).resolve().parent

# Порогова кількість файлів, з якої вважаємо назву "частою".
FREQUENT_MIN_FILES = 3

# Наскільки близькою має бути назва до наявної у Словнику, щоб вважати
# її тією самою і НЕ додавати. Високий поріг свідомо: краще показати
# людині зайвий дубль, ніж мовчки проковтнути справді нову позицію.
DUPLICATE_SCORE = 93

# Відомі одиниці виміру — за ними впізнаємо, де саме колонка назви.
KNOWN_UNITS = {
    "шт", "шт.", "штук", "к-т", "к-т.", "кт", "комплект", "комплекти",
    "пара", "пари", "пар", "од", "од.", "одиниць", "компл", "компл.",
    "м", "м.", "кг", "л", "уп", "уп.", "набір",
}

# Рядки відомості, які не є позиціями майна.
SERVICE_ROW_RE = re.compile(
    r"^(всього|разом|примітка|голова|член|звання|посада|підпис|начальник|"
    r"командир|відомість|найменування|№|витрати|усього)", re.IGNORECASE)

# Хвіст "2025р." / "2024 р." / "(2025р.)" наприкінці назви.
YEAR_TAIL_RE = re.compile(
    r"[\s(,]*\b(19|20)\d{2}\s*(р|рік|року)?\.?\s*\)?\s*$", re.IGNORECASE)


# ================================================================
# ДОВІДНИК З НАКАЗУ 232 (офіційне джерело найменувань і одиниць)
# ================================================================
# PDF наказу лежить у Zvirka_slovnuka/norms_cache/. Текст із нього
# витягується pdftotext -layout, де рядок номенклатури має вигляд:
#     32  Чоботи гумові        пара      1   3   3   3   3   2   8
# тобто: номер, назва, одиниця, далі норми по роках. Беремо лише
# назву й одиницю — терміни служби залежать від категорії
# військовослужбовця, і зводити їх в одне число не можна.
NAKAZ_UNITS = ["штука", "штук", "пара", "пар", "комплект", "комплектів",
               "кілограм", "метр", "пачка", "набір"]
NAKAZ_ROW_RE = re.compile(
    r"^\s*\d{1,3}\s+([А-ЯІЇЄҐ][^\d]{3,70}?)\s{2,}(" + "|".join(NAKAZ_UNITS) + r")\b",
    re.M)

# Одиниці в наказі писані повними словами, у відомостях — скорочено.
NAKAZ_UNIT_SHORT = {
    "штука": "шт.", "штук": "шт.", "пара": "пара", "пар": "пара",
    "комплект": "к-т", "комплектів": "к-т", "кілограм": "кг",
    "метр": "м", "пачка": "уп.", "набір": "к-т",
}


def load_nakaz_reference(txt_path):
    """Повертає {назва_lower: (назва, одиниця)} з тексту наказу 232."""
    if not txt_path or not Path(txt_path).is_file():
        return {}
    try:
        text = Path(txt_path).read_text(encoding="utf-8")
    except Exception:
        return {}
    ref = {}
    for m in NAKAZ_ROW_RE.finditer(text):
        name = re.sub(r"\s+", " ", m.group(1)).strip(" .,;*")
        unit = NAKAZ_UNIT_SHORT.get(m.group(2), m.group(2))
        if len(name) >= 4:
            ref.setdefault(name.lower(), (name, unit))
    return ref


def match_nakaz(name, nakaz_ref, min_score=88):
    """Шукає позицію в наказі 232. Повертає (назва, одиниця, оцінка) або None.
    Поріг високий: позначка "підтверджено наказом" має щось означати,
    інакше вона знецінюється — краще не підтвердити, ніж підтвердити хибно."""
    if not nakaz_ref:
        return None
    low = name.lower()
    if low in nakaz_ref:
        n, u = nakaz_ref[low]
        return n, u, 100.0
    best, best_score = None, -1.0
    for key, (n, u) in nakaz_ref.items():
        score = max(fuzz.ratio(low, key), fuzz.token_sort_ratio(low, key))
        if score > best_score:
            best, best_score = (n, u), score
    if best is not None and best_score >= min_score:
        return best[0], best[1], best_score
    return None


def _flag_value(argv, flag, default=None):
    """Значення іменованого прапорця виду '--flag ЗНАЧЕННЯ'."""
    if flag not in argv:
        return default
    i = argv.index(flag)
    if i + 1 >= len(argv):
        print(f"ПОМИЛКА: після {flag} очікується значення.")
        sys.exit(1)
    return argv[i + 1]


def log(msg=""):
    print(msg, flush=True)


def clean_name(raw):
    """Прибирає рік переоцінки в кінці й зайві пробіли/розділові."""
    name = re.sub(r"\s+", " ", str(raw)).strip()
    prev = None
    # Хвіст може бути подвійним: "Мішок спальний 2025р. 2025р."
    while prev != name:
        prev = name
        name = YEAR_TAIL_RE.sub("", name).strip()
    return name.strip(" .,;-–—")


def looks_like_unit(value):
    if value is None:
        return False
    return str(value).strip().lower().strip(".") in {u.strip(".") for u in KNOWN_UNITS}


def is_service_row(name):
    if not name or len(name) < 3:
        return True
    return bool(SERVICE_ROW_RE.match(name))


def extract_from_sheet(ws):
    """Повертає список (назва, одиниця) з одного аркуша відомості."""
    found = []
    for row in ws.iter_rows(min_row=1, max_col=6, values_only=True):
        if not row or row[0] is None:
            continue
        # № з/п має бути числом (або числом-рядком)
        idx = str(row[0]).strip()
        if not idx.isdigit():
            continue

        # Варіант 1: B=назва, C=одиниця. Варіант 2 (зсув): C=назва, D=одиниця.
        for name_i, unit_i in ((1, 2), (2, 3)):
            if len(row) <= unit_i:
                continue
            name_raw, unit_raw = row[name_i], row[unit_i]
            if name_raw is None:
                continue
            name = clean_name(name_raw)
            if is_service_row(name):
                continue
            # Назва має бути текстом, а не числом (інакше це рядок
            # заголовка "1|2|3|4", який теж починається з цифри).
            if name.isdigit() or len(name) < 4:
                continue
            if looks_like_unit(unit_raw) or unit_raw is None:
                unit = str(unit_raw).strip() if unit_raw is not None else ""
                found.append((name, unit))
                break
    return found


def load_dictionary(template_path):
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    if "Словник" not in wb.sheetnames:
        wb.close()
        raise SystemExit(f"ПОМИЛКА: у макеті немає аркуша 'Словник': {template_path}")
    ws = wb["Словник"]
    entries = []
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] and str(row[0]).strip():
            entries.append(str(row[0]).strip())
    wb.close()
    return entries


def already_in_dictionary(name, dict_lower, dict_entries):
    low = name.lower()
    if low in dict_lower:
        return True, dict_lower[low], 100.0
    best, best_score = None, -1.0
    for entry in dict_entries:
        score = max(fuzz.ratio(low, entry.lower()),
                    fuzz.token_sort_ratio(low, entry.lower()))
        if score > best_score:
            best_score, best = score, entry
    return (best_score >= DUPLICATE_SCORE), best, best_score


def main():
    argv = sys.argv[1:]
    positional = [a for a in argv if not a.startswith("--")]
    apply_changes = "--apply" in argv

    if len(positional) < 2:
        print(__doc__)
        sys.exit(1)

    src_dir = Path(positional[0])
    template = Path(positional[1])

    if not src_dir.is_dir():
        print(f"ПОМИЛКА: папка не знайдена: {src_dir}")
        sys.exit(1)
    if not template.is_file():
        print(f"ПОМИЛКА: макет не знайдено: {template}")
        sys.exit(1)

    # Скан 1700+ файлів триває ~8 хв, тому результат кешується. Повторний
    # запуск (напр. щоб змінити пороги чи кольори) стартує миттєво;
    # --rescan примусово перечитує папку.
    scan_cache = PROJECT_ROOT / "Data" / "unique_items_scan.json"
    stats = None
    if scan_cache.is_file() and "--rescan" not in argv:
        try:
            with open(scan_cache, encoding="utf-8") as f:
                raw = json.load(f)
            stats = {k: {"files": set(v["files"]), "count": v["count"],
                         "units": defaultdict(int, v["units"])}
                     for k, v in raw.items()}
            log(f"Взято з кешу сканування: {len(stats)} унікальних назв "
                f"({scan_cache.name}; --rescan щоб перечитати папку)")
        except Exception:
            stats = None

    if stats is None:
        files = sorted(src_dir.glob("*.xlsx")) + sorted(src_dir.glob("*.xlsm"))
        log(f"Файлів у папці: {len(files)}")
        stats = defaultdict(lambda: {"files": set(), "count": 0,
                                      "units": defaultdict(int)})
        broken = 0
        t0 = time.time()

        for i, path in enumerate(files, 1):
            if i % 200 == 0:
                log(f"  ...оброблено {i}/{len(files)} файлів, "
                    f"унікальних назв поки {len(stats)}")
            try:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            except Exception:
                broken += 1
                continue
            try:
                for sheet in wb.worksheets:
                    for name, unit in extract_from_sheet(sheet):
                        rec = stats[name]
                        rec["files"].add(path.name)
                        rec["count"] += 1
                        if unit:
                            rec["units"][unit] += 1
            except Exception:
                broken += 1
            finally:
                wb.close()

        log(f"Прочитано за {time.time() - t0:.0f} сек. Унікальних назв: {len(stats)}"
            + (f", не вдалося відкрити: {broken}" if broken else ""))

        scan_cache.parent.mkdir(parents=True, exist_ok=True)
        with open(scan_cache, "w", encoding="utf-8") as f:
            json.dump({k: {"files": sorted(v["files"]), "count": v["count"],
                           "units": dict(v["units"])}
                       for k, v in stats.items()}, f, ensure_ascii=False)

    dict_entries = load_dictionary(template)
    dict_lower = {e.lower(): e for e in dict_entries}
    log(f"У Словнику зараз: {len(dict_entries)} позицій")

    nakaz_txt = _flag_value(argv, "--nakaz")
    nakaz_ref = load_nakaz_reference(nakaz_txt)
    if nakaz_txt and not nakaz_ref:
        log(f"УВАГА: не вдалося прочитати номенклатуру з {nakaz_txt} — "
            f"позначки 'підтверджено наказом' не буде")
    log(f"Наказ 232: {len(nakaz_ref)} позицій номенклатури"
        if nakaz_ref else "Наказ 232: не задано (--nakaz ФАЙЛ.txt)")

    # Зливаємо назви, що відрізняються лише регістром чи пробілами:
    # "Чоботи гумові" і "чоботи гумові" — це одна позиція, а не дві.
    # Без цього кроку в Словник потрапляли дублі (виміряно: 30 штук),
    # а дубль у словнику шкідливий — нечіткий пошук починає обирати між
    # ними довільно. Лишаємо найчастіше написання як канонічне.
    merged = {}
    for name, rec in stats.items():
        key = " ".join(name.lower().split())
        if key not in merged:
            merged[key] = {"variants": defaultdict(int), "files": set(),
                           "count": 0, "units": defaultdict(int)}
        m = merged[key]
        m["variants"][name] += rec["count"]
        m["files"] |= rec["files"]
        m["count"] += rec["count"]
        for u, c in rec["units"].items():
            m["units"][u] += c

    collapsed = len(stats) - len(merged)
    if collapsed:
        log(f"Злито варіантів написання (регістр/пробіли): {collapsed}")
    stats = {max(m["variants"].items(), key=lambda kv: kv[1])[0]: m
             for m in merged.values()}

    new_items = []
    known = 0
    for name, rec in stats.items():
        is_dup, closest, score = already_in_dictionary(name, dict_lower, dict_entries)
        if is_dup:
            known += 1
            continue
        unit = max(rec["units"].items(), key=lambda kv: kv[1])[0] if rec["units"] else ""
        nakaz = match_nakaz(name, nakaz_ref)
        new_items.append({
            "name": name,
            "unit": unit,
            "files": len(rec["files"]),
            "count": rec["count"],
            "closest": closest or "",
            "closest_score": round(score, 1),
            "nakaz_name": nakaz[0] if nakaz else "",
            "nakaz_unit": nakaz[1] if nakaz else "",
            "nakaz_score": round(nakaz[2], 1) if nakaz else 0.0,
        })

    new_items.sort(key=lambda d: (-d["files"], -d["count"], d["name"]))
    confirmed = [d for d in new_items if d["nakaz_name"]]
    frequent = [d for d in new_items if not d["nakaz_name"] and d["files"] >= FREQUENT_MIN_FILES]
    rare = [d for d in new_items if not d["nakaz_name"] and d["files"] < FREQUENT_MIN_FILES]

    def category(d):
        if d["nakaz_name"]:
            return "підтверджено наказом 232"
        return "часта" if d["files"] >= FREQUENT_MIN_FILES else "рідкісна"

    log("")
    log("=" * 62)
    log(f"Уже є у Словнику:              {known}")
    log(f"НОВИХ позицій:                 {len(new_items)}")
    log(f"  підтверджені наказом 232:    {len(confirmed)}")
    log(f"  часті (>= {FREQUENT_MIN_FILES} файлів):         {len(frequent)}")
    log(f"  рідкісні (1-2 файли):        {len(rare)}")
    log("=" * 62)

    csv_path = template.parent / "нові_позиції_для_словника.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Найменування", "Одиниця", "Файлів", "Згадок", "Категорія",
                    "Наказ 232: назва", "Наказ 232: одиниця", "Збіг з наказом %",
                    "Найближче у Словнику", "Схожість %"])
        for d in new_items:
            w.writerow([d["name"], d["unit"], d["files"], d["count"], category(d),
                        d["nakaz_name"], d["nakaz_unit"], d["nakaz_score"],
                        d["closest"], d["closest_score"]])
    log(f"Звіт CSV: {csv_path}")

    log("")
    log("--- ТОП-25 частих нових позицій ---")
    for d in frequent[:25]:
        log(f"  {d['files']:>4} файлів | {d['unit']:<6} | {d['name'][:64]}")

    if not apply_changes:
        log("")
        log("Це був ПЕРЕГЛЯД. Щоб дописати їх у копію макета — додай --apply")
        return

    # ---------- запис У САМ МАКЕТ ----------
    # Пишемо в оригінальний файл: решта скриптів (excel_report_generator,
    # ocr_text_corrector, doc_processor) читають Словник саме звідти, і
    # копія з суфіксом їм не видима. Перед записом — бекап з міткою часу
    # поруч, як це вже робить universal_dict_extractor.py.
    keep_vba = template.suffix.lower() in {".xlsm", ".xltm"}
    backup = template.with_name(
        f"{template.stem}_backup_{time.strftime('%Y-%m-%d_%H-%M-%S')}{template.suffix}")
    shutil.copy2(template, backup)
    log(f"Бекап оригіналу: {backup.name}")

    out_path = template
    wb = openpyxl.load_workbook(template, keep_vba=keep_vba)
    ws = wb["Словник"]

    blue = PatternFill("solid", fgColor="BDD7EE")    # підтверджено наказом 232
    green = PatternFill("solid", fgColor="C6EFCE")   # часті
    yellow = PatternFill("solid", fgColor="FFEB9C")  # рідкісні

    row = ws.max_row + 1
    if not ws.cell(1, 4).value:
        ws.cell(1, 4).value = "Джерело / коментар"
        ws.cell(1, 4).font = Font(bold=True)

    unit_conflicts = 0
    for d in new_items:
        ws.cell(row, 1).value = d["name"]
        # Одиниця: якщо позиція підтверджена наказом — беремо одиницю
        # ЗВІДТИ (офіційна норма важливіша за те, як написали у відомості).
        unit = d["nakaz_unit"] or d["unit"]
        ws.cell(row, 2).value = unit
        # Колонку C (термін служби) НЕ заповнюємо: у наказі він залежить
        # від категорії військовослужбовця й пори року, зводити це в
        # одне число не можна — нехай лишається порожньою для людини.

        parts = [f"З готових відомостей: {d['files']} файл(ів), {d['count']} згадок"]
        if d["nakaz_name"]:
            parts.append(f"ПІДТВЕРДЖЕНО наказом 232: «{d['nakaz_name']}», "
                         f"од. {d['nakaz_unit']} (збіг {d['nakaz_score']:.0f}%)")
            if d["unit"] and d["nakaz_unit"] and \
               d["unit"].strip(".").lower() != d["nakaz_unit"].strip(".").lower():
                parts.append(f"УВАГА: у відомостях одиниця «{d['unit']}», "
                             f"у наказі «{d['nakaz_unit']}» — перевір")
                unit_conflicts += 1
        if d["closest"]:
            parts.append(f"найближче у Словнику: «{d['closest']}» "
                         f"({d['closest_score']:.0f}%)")
        ws.cell(row, 4).value = "; ".join(parts)

        fill = blue if d["nakaz_name"] else (
            green if d["files"] >= FREQUENT_MIN_FILES else yellow)
        for col in (1, 2, 3, 4):
            ws.cell(row, col).fill = fill
        row += 1

    wb.save(out_path)
    log("")
    log(f"ДОДАНО {len(new_items)} рядків у: {out_path}")
    log(f"  СИНІЙ   — підтверджено наказом 232:      {len(confirmed)}")
    log(f"  ЗЕЛЕНИЙ — часті (>= {FREQUENT_MIN_FILES} файлів), без наказу: {len(frequent)}")
    log(f"  ЖОВТИЙ  — рідкісні (1-2 файли):          {len(rare)}")
    if unit_conflicts:
        log(f"  УВАГА: розбіжність одиниці з наказом у {unit_conflicts} позиціях "
            f"(деталі в колонці D)")
    log("Колонки B і C наявних рядків не змінювались: нові позиції")
    log("дописані В КІНЕЦЬ, тому нічого не зсунулось.")
    log("Оригінал макета не змінювався.")


if __name__ == "__main__":
    main()
