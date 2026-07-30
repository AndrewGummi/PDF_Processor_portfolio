#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cleanup_dictionary.py — чистить аркуш "Словник" від того, що не є
номенклатурою, і зводить базові назви до канонічних.

ПРАВИЛО ПРЕДМЕТНОЇ ОБЛАСТІ (від замовника)
Наказ 232 — еталон. Позиція в обліку не може бути "голою" базовою
назвою: вона або збігається з наказом, або має припис модифікації, як
"Бронежилет модульний PLASTOON-UAFE (тип 2, рівень П, вид 5)". Зокрема:
    "окуляри захисні"  -> завжди "балістичні" (припис або саме слово)
    "окуляри маска"    -> модифікація або "захисні балістичні"

ЩО РОБИТЬ
1. Strips rank/surname annotations in parentheses, e.g. "(сол. ПРІЗВИЩЕ І.П)"
   or "(Сол. ПРІЗВИЩЕ А.О., сол ПРІЗВИЩЕ А. М.)". These are per-person notes
   carried over from the statement, not part of the item name - because of
   them a single item multiplied into dozens of "unique" entries.
2. Зводить окуляри до канонічних форм за правилом вище.
3. Після заміни зливає рядки, що стали однаковими: підсумовує
   статистику в коментарі й лишає один рядок.

ЧОГО НЕ РОБИТЬ СВІДОМО
Не оголошує "коротка назва = неповна". Перевірено на реальних даних:
"Чоботи гумові", "Рушник вафельний", "Матрац ватяний", "Простирадло
бавовняне" — короткі, але це точні назви з наказу 232. Правило про
обов'язковий припис стосується конкретних груп (бронежилети, окуляри,
шоломи), а не довжини рядка, тому інші сумнівні позиції лише
ПОКАЗУЮТЬСЯ у звіті — рішення за людиною.

ВИКОРИСТАННЯ
    python cleanup_dictionary.py <макет.xlsm> [--apply]
Без --apply лише показує, що буде змінено.
"""

import re
import shutil
import sys
import time
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ПОМИЛКА: потрібен openpyxl (pip install openpyxl)")
    raise

# A parenthesised block containing a rank and a surname, e.g.
# "(сол. ПРІЗВИЩЕ І.П)", "(Сол. ПРІЗВИЩЕ А.О., сол ПРІЗВИЩЕ А. М.)",
# "(серж.ПРІЗВИЩЕ.І.Г)".
RANK_WORDS = r"(?:сол|солдат|серж|сержант|ст\.?\s*сол|мол\.?\s*серж|м/с|ряд|" \
             r"к-н|л-т|ст\.?\s*л-т|м-р|прап|курс)"
PERSON_PAREN_RE = re.compile(
    r"\s*\([^)]*\b" + RANK_WORDS + r"\b[^)]*\)", re.IGNORECASE)
# Дужки з чистим ПІБ без звання: "(Шевченко Т.Г)"
NAME_PAREN_RE = re.compile(
    r"\s*\([А-ЯІЇЄҐ][а-яіїєґ']+\s+[А-ЯІЇЄҐ]\.\s*[А-ЯІЇЄҐ]\.?\s*\)")

# Канонізація окулярів за правилом замовника.
# Ключ — назва в нижньому регістрі після нормалізації пробілів/дефісів.
EYEWEAR_CANON = {
    "окуляри захисні": "Окуляри захисні балістичні",
    "окуляри захисні захисні балістичні": "Окуляри захисні балістичні",
    "окуляри маска": "Окуляри-маска захисні балістичні",
    "окуляри маска балістичні": "Окуляри-маска захисні балістичні",
    "окуляри маска захисні": "Окуляри-маска захисні балістичні",
    "окуляри маска захистні балістичні": "Окуляри-маска захисні балістичні",
    "окуляри маска захисні балістичні": "Окуляри-маска захисні балістичні",
}


def log(msg=""):
    print(msg, flush=True)


def norm_key(name):
    """Ключ для порівняння: без регістру, зайвих пробілів і різниці
    'окуляри-маска' / 'окуляри - маска' / 'окуляри маска'."""
    s = " ".join(str(name).lower().split())
    s = re.sub(r"\s*-\s*", " ", s)
    return " ".join(s.split())


def strip_person(name):
    """Прибирає дужки з прізвищами/званнями. Повертає (назва, чи_змінено)."""
    original = name
    prev = None
    while prev != name:
        prev = name
        name = PERSON_PAREN_RE.sub("", name)
        name = NAME_PAREN_RE.sub("", name)
    name = re.sub(r"\s+", " ", name).strip(" .,;-–—")
    return name, name != original


def canon_eyewear(name):
    """Зводить окуляри до канонічної форми. Повертає (назва, чи_змінено)."""
    key = norm_key(name)
    if key in EYEWEAR_CANON:
        new = EYEWEAR_CANON[key]
        return new, norm_key(new) != key or new != name
    return name, False


def main():
    argv = sys.argv[1:]
    positional = [a for a in argv if not a.startswith("--")]
    apply_changes = "--apply" in argv

    if not positional:
        print(__doc__)
        sys.exit(1)

    template = Path(positional[0])
    if not template.is_file():
        print(f"ПОМИЛКА: макет не знайдено: {template}")
        sys.exit(1)

    keep_vba = template.suffix.lower() in {".xlsm", ".xltm"}
    wb = openpyxl.load_workbook(template, keep_vba=keep_vba)
    if "Словник" not in wb.sheetnames:
        print("ПОМИЛКА: немає аркуша 'Словник'")
        sys.exit(1)
    ws = wb["Словник"]

    # Межа дописаного блоку — лише для звітності (щоб було видно, де
    # чиї рядки). Чистка застосовується до ВСЬОГО словника: оригінальні
    # рядки теж заповнювались вручну, тому там трапляються подвійні
    # пробіли, друкарські помилки ("захистні") і дублі з різним
    # написанням.
    first_added = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 4).value:
            first_added = r
            break

    log(f"Словник: {ws.max_row - 1} позицій")
    if first_added:
        log(f"  рядки 2-{first_added - 1} — оригінальні (ручне заповнення)")
        log(f"  рядки {first_added}-{ws.max_row} — дописані з відомостей")
    log("Чистка застосовується до ВСІХ рядків.")

    changed_person, changed_eyewear = [], []
    plan = {}  # row -> нова назва
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        name = str(name).strip()
        new = name

        stripped, had_person = strip_person(new)
        if had_person and stripped != new:
            changed_person.append((new, stripped))
            new = stripped

        canon, is_eyewear = canon_eyewear(new)
        if is_eyewear and canon != new:
            changed_eyewear.append((new, canon))
            new = canon

        if new and new != name:
            plan[r] = new

    log("")
    log(f"Прибрано прізвищ/звань з назв: {len(changed_person)}")
    for old, new in changed_person[:8]:
        log(f"   «{old[:60]}»")
        log(f"      -> «{new[:60]}»")
    if len(changed_person) > 8:
        log(f"   ... ще {len(changed_person) - 8}")

    log("")
    log(f"Зведено окулярів до канонічних форм: {len(changed_eyewear)}")
    for old, new in changed_eyewear:
        log(f"   «{old[:46]}» -> «{new}»")

    # Групуємо ВСІ рядки за нормалізованою назвою. У групі лишається
    # один рядок, і це не обов'язково перший: зберігаємо той, у якого
    # найповніші дані. Термін служби (колонка C) заповнений лише в ~117
    # рядках і взятий з наказу — втратити його при злитті не можна,
    # тому відсутні значення переносяться в рядок, що лишається.
    groups = defaultdict(list)
    for r in range(2, ws.max_row + 1):
        name = plan.get(r) or ws.cell(r, 1).value
        if name:
            groups[norm_key(name)].append(r)

    def completeness(r):
        return (1 if ws.cell(r, 3).value else 0) + (1 if ws.cell(r, 2).value else 0)

    to_delete = []
    carry = {}  # рядок-переможець -> {"unit":..., "term":...}
    for key, rws in groups.items():
        if len(rws) < 2:
            continue
        keeper = max(rws, key=lambda r: (completeness(r), -r))
        unit = ws.cell(keeper, 2).value or next(
            (ws.cell(r, 2).value for r in rws if ws.cell(r, 2).value), None)
        term = ws.cell(keeper, 3).value or next(
            (ws.cell(r, 3).value for r in rws if ws.cell(r, 3).value), None)
        if unit != ws.cell(keeper, 2).value or term != ws.cell(keeper, 3).value:
            carry[keeper] = {"unit": unit, "term": term}
        for r in rws:
            if r != keeper:
                name = plan.get(r) or ws.cell(r, 1).value
                to_delete.append((r, name, f"дубль -> лишається рядок {keeper}"))

    log("")
    log(f"Дублів буде злито: {len(to_delete)}")
    for r, name, why in to_delete[:10]:
        log(f"   r{r}: «{str(name)[:46]}» — {why}")
    if len(to_delete) > 10:
        log(f"   ... ще {len(to_delete) - 10}")
    if carry:
        log(f"Перенесено одиницю/термін у рядок-переможець: {len(carry)} випадків")

    if not apply_changes:
        log("")
        log("Це був ПЕРЕГЛЯД. Щоб застосувати — додай --apply")
        return

    backup = template.with_name(
        f"{template.stem}_backup_{time.strftime('%Y-%m-%d_%H-%M-%S')}{template.suffix}")
    shutil.copy2(template, backup)
    log("")
    log(f"Бекап: {backup.name}")

    # 1) Перейменування на місці — лише колонка A.
    for r, new in plan.items():
        ws.cell(r, 1).value = new

    # 2) Перенесення одиниці/терміну в рядок, що лишається (ДО видалення,
    #    поки номери рядків ще не зсунулись).
    for r, vals in carry.items():
        if vals["unit"] is not None:
            ws.cell(r, 2).value = vals["unit"]
        if vals["term"] is not None:
            ws.cell(r, 3).value = vals["term"]

    # 3) Видалення дублів — знизу вгору, щоб не поїхали номери рядків.
    for r, _, _ in sorted(to_delete, key=lambda t: -t[0]):
        ws.delete_rows(r, 1)

    wb.save(template)
    log(f"ЗАПИСАНО: {template.name}")
    log(f"  перейменовано:      {len(plan)}")
    log(f"  злито дублів:       {len(to_delete)}")
    log(f"  перенесено даних:   {len(carry)}")
    log(f"  позицій у Словнику: {ws.max_row - 1}")


if __name__ == "__main__":
    main()
