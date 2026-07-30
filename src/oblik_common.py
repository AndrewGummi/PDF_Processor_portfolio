#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
oblik_common.py -- спільні функції для скриптів обліку майна бригади:
    cards_photo_to_excel.py     (Журнал руху -- з фото карток, vision-LLM)
    import_pidrozdily_xlsm.py   (Підрозділи -- з файлу "Залишки_На_ХХ_20ХХ.xlsm")
    build_master_workbook.py    (початкова збірка всього файлу разом)

Тут же -- rebuild_rollup(), яка перебудовує "Залишки — Загалом" і "Звірка"
на основі того, що ЗАРАЗ реально лежить в "Журнал руху" / "Склади" /
"Підрозділи" (об'єднання унікальних найменувань майна з усіх трьох).
Викликати її треба щоразу після зміни будь-якого з трьох джерел --
інакше зведення й звірка лишаться зі старим списком найменувань.
"""

import re
from datetime import date
from pathlib import Path

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
EXAMPLE_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

ITEM_NAME_HEADER = "Найменування майна"


def load_env_file(path):
    """Читає локальний .env-файл (KEY=VALUE, по рядку, # -- коментар) і
    прописує знайдені пари в os.environ ЦЬОГО процесу (не зберігається
    ніде більше, не потрапляє в жоден лог). Повертає СПИСОК ІМЕН
    завантажених змінних -- НІКОЛИ самих значень, щоб ключі не
    з'являлись у виводі/логах навіть випадково."""
    import os
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Файл не знайдено: {path}")
    loaded = []
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and value:
            os.environ[key] = value
            loaded.append(key)
    return loaded


def style_header(ws, headers, row=1):
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = f"A{row + 1}"


def add_or_resize_table(ws, name, n_cols, n_rows, style="TableStyleMedium9"):
    last_col = get_column_letter(n_cols)
    ref = f"A1:{last_col}{max(n_rows, 2)}"
    if name in getattr(ws, "tables", {}):
        ws.tables[name].ref = ref
    else:
        tbl = Table(displayName=name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True,
                                             showFirstColumn=False, showLastColumn=False,
                                             showColumnStripes=False)
        ws.add_table(tbl)


def _table_column_values(ws, table_name, header_name):
    """Повертає список значень стовпця header_name таблиці table_name на
    аркуші ws (без рядка заголовка). Порожній список, якщо таблиці чи
    аркуша нема -- це нормальний стан, поки якесь із джерел ще не заповнене."""
    if ws is None or table_name not in getattr(ws, "tables", {}):
        return []
    tbl = ws.tables[table_name]
    min_col, min_row, max_col, max_row = _ref_bounds(tbl.ref)
    headers = [ws.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)]
    if header_name not in headers:
        return []
    col_idx = min_col + headers.index(header_name)
    values = []
    for r in range(min_row + 1, max_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v not in (None, ""):
            values.append(str(v).strip())
    return values


def _ref_bounds(ref: str):
    from openpyxl.utils import range_boundaries
    return range_boundaries(ref)


def collect_unique_items(wb, sources):
    """sources -- список (sheet_name, table_name). Повертає відсортований
    список унікальних найменувань майна, зібраних з усіх джерел, які
    реально існують у книзі зараз."""
    items = set()
    for sheet_name, table_name in sources:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else None
        for v in _table_column_values(ws, table_name, ITEM_NAME_HEADER):
            items.add(v)
    return sorted(items, key=lambda s: s.lower())


def rebuild_rollup(wb, journal_sheet="Журнал руху", journal_table="ZhurnalRukhu",
                    sklady_sheet="Склади", sklady_table="Sklady",
                    pidrozdily_sheet="Підрозділи", pidrozdily_table="Pidrozdily",
                    zalyshky_sheet="Залишки — Загалом", zvirka_sheet="Звірка"):
    """Перебудовує 'Залишки — Загалом' і 'Звірка' з нуля на основі
    ОБ'ЄДНАННЯ унікальних найменувань майна з Журналу/Складів/Підрозділів.
    Виклик безпечний повторно -- старі версії цих двох аркушів видаляються
    і будуються заново, решта аркушів не чіпається."""
    items = collect_unique_items(wb, [
        (journal_sheet, journal_table),
        (sklady_sheet, sklady_table),
        (pidrozdily_sheet, pidrozdily_table),
    ])

    for name in (zalyshky_sheet, zvirka_sheet):
        if name in wb.sheetnames:
            del wb[name]

    ws_z = wb.create_sheet(zalyshky_sheet)
    headers = ["Найменування майна", "Надійшло (журнал)", "Вибуло (журнал)",
               "Залишок за журналом (бригада)", "Дата останнього запису"]
    style_header(ws_z, headers)
    for i, item in enumerate(items, start=2):
        safe_item = item.replace('"', '""')
        ws_z.cell(row=i, column=1, value=item).font = Font(name="Arial")
        c2 = ws_z.cell(row=i, column=2,
                        value=f'=SUMIFS({journal_table}[Надійшло],{journal_table}[Найменування майна],"{safe_item}")')
        c3 = ws_z.cell(row=i, column=3,
                        value=f'=SUMIFS({journal_table}[Вибуло],{journal_table}[Найменування майна],"{safe_item}")')
        c4 = ws_z.cell(row=i, column=4, value=f'=B{i}-C{i}')
        c5 = ws_z.cell(row=i, column=5,
                        value=f'=IFERROR(_xlfn.MAXIFS({journal_table}[Дата запису],{journal_table}[Найменування майна],"{safe_item}"),"")')
        for c in (c2, c3, c4):
            c.font = Font(name="Arial")
            c.number_format = "#,##0"
        c5.font = Font(name="Arial")
        c5.number_format = "DD.MM.YYYY"
    add_or_resize_table(ws_z, "ZalyshkyZagalom", 5, len(items) + 1)
    for col_idx, w in zip(range(1, 6), [34, 18, 16, 22, 20]):
        ws_z.column_dimensions[get_column_letter(col_idx)].width = w

    ws_v = wb.create_sheet(zvirka_sheet)
    headers = ["Найменування майна", "Залишок за журналом (бригада)",
               "Разом по складах", "Разом по підрозділах",
               "Разом (склади + підрозділи)", "Різниця", "Статус"]
    style_header(ws_v, headers)
    for i, item in enumerate(items, start=2):
        safe_item = item.replace('"', '""')
        z_row = i  # той самий порядок рядків, що й у "Залишки — Загалом"
        ws_v.cell(row=i, column=1, value=item).font = Font(name="Arial")
        c2 = ws_v.cell(row=i, column=2, value=f"='{zalyshky_sheet}'!D{z_row}")
        c3 = ws_v.cell(row=i, column=3,
                        value=f'=SUMIFS({sklady_table}[Кількість],{sklady_table}[Найменування майна],"{safe_item}")')
        c4 = ws_v.cell(row=i, column=4,
                        value=f'=SUMIFS({pidrozdily_table}[Кількість],{pidrozdily_table}[Найменування майна],"{safe_item}")')
        c5 = ws_v.cell(row=i, column=5, value=f'=C{i}+D{i}')
        c6 = ws_v.cell(row=i, column=6, value=f'=B{i}-E{i}')
        c7 = ws_v.cell(row=i, column=7, value=f'=IF(F{i}=0,"OK","ПЕРЕВІРИТИ")')
        for c in (c2, c3, c4, c5, c6):
            c.font = Font(name="Arial")
            c.number_format = "#,##0"
        c7.font = Font(name="Arial", bold=True)
    add_or_resize_table(ws_v, "Zvirka", 7, len(items) + 1)
    for col_idx, w in zip(range(1, 8), [34, 22, 14, 16, 20, 10, 14]):
        ws_v.column_dimensions[get_column_letter(col_idx)].width = w

    note = ("'ПЕРЕВІРИТИ' тут не завжди означає помилку: якщо позиція є лише в "
            "Підрозділах, але ще жодного разу не проходила через Журнал руху "
            "бригади (чи навпаки) -- це очікувано покаже розбіжність, доки "
            "джерела не наздоженуть одне одного.")
    ws_v.cell(row=len(items) + 3, column=1, value=note).font = Font(name="Arial", italic=True, size=9)
    return len(items)


# ================================================================
# ПАРСЕР "Залишки_На_ХХ_20ХХ.xlsm" (по одному аркушу на підрозділ)
# ================================================================
_TRAILING_MONTH_RE = re.compile(r"[\s\-–—]*\d{1,2}[\s\-–—]*$")


def upsert_location_rows(ws, table_name, headers, location_col_header, location_name,
                          new_data_rows):
    """Замінює в таблиці table_name на аркуші ws УСІ рядки, де стовпець
    location_col_header == location_name, на new_data_rows (список
    кортежів значень у порядку headers, ПРОПУСКАЮЧИ '№', колонку
    локації і будь-які формула-колонки (Відповідальна особа/Телефон) --
    саме так, як формують їх скрипти-імпортери). Рядки інших локацій
    лишаються недоторканими. Використовується докс-імпортером, коли
    приходить файл лише по ОДНОМУ складу/підрозділу.

    Формули INDEX/MATCH (Відповідальна особа, Телефон) тут НЕ
    заповнюються -- після виклику треба окремо викликати
    stamp_responsible_lookup()."""
    if table_name not in getattr(ws, "tables", {}):
        raise ValueError(f"Таблиця {table_name} відсутня на аркуші {ws.title}")

    tbl = ws.tables[table_name]
    min_col, min_row, max_col, max_row = _ref_bounds(tbl.ref)
    loc_col_idx = min_col + headers.index(location_col_header)
    formula_cols = {"Відповідальна особа", "Телефон"}
    fillable_idx = [i for i, h in enumerate(headers)
                    if h != "№" and h != location_col_header and h not in formula_cols]

    kept_rows = []
    for r in range(min_row + 1, max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(min_col, max_col + 1)]
        if row_vals[loc_col_idx - min_col] != location_name:
            kept_rows.append(row_vals)

    # Перебудовуємо аркуш: заголовок як був, далі kept_rows, далі нові
    # рядки для location_name. Простіше й надійніше, ніж вставляти/
    # видаляти рядки посеред існуючої таблиці.
    for r in range(min_row + 1, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).value = None

    next_row = min_row + 1
    for row_vals in kept_rows:
        for c_offset, v in enumerate(row_vals):
            ws.cell(row=next_row, column=min_col + c_offset, value=v)
        next_row += 1

    for data_row in new_data_rows:
        full_row = [None] * len(headers)
        full_row[headers.index("№")] = f"=ROW()-{min_row}"
        full_row[headers.index(location_col_header)] = location_name
        for di, ci in enumerate(fillable_idx):
            if di < len(data_row):
                full_row[ci] = data_row[di]
        for c_offset, (v, h) in enumerate(zip(full_row, headers)):
            if h in formula_cols:
                continue  # заповнюється окремо через stamp_responsible_lookup()
            cell = ws.cell(row=next_row, column=min_col + c_offset, value=v)
            cell.font = Font(name="Arial")
            if "Кількість" in h:
                cell.number_format = "#,##0"
            elif "Дата" in h:
                cell.number_format = "DD.MM.YYYY"
        next_row += 1

    add_or_resize_table(ws, table_name, len(headers), next_row - min_row)
    return next_row - min_row - 1  # кількість рядків даних після операції


def stamp_responsible_lookup(ws, table_name, headers, location_col_header):
    """Проходить по ВСІХ поточних рядках даних таблиці й (пере)записує
    формули INDEX/MATCH у колонках 'Відповідальна особа' / 'Телефон',
    підтягуючи їх з довідника Vidpovidalni за значенням колонки
    location_col_header. Викликати після будь-якого запису/оновлення
    даних у Sklady чи Pidrozdily."""
    if table_name not in getattr(ws, "tables", {}):
        return
    if "Відповідальна особа" not in headers or "Телефон" not in headers:
        return
    tbl = ws.tables[table_name]
    min_col, min_row, max_col, max_row = _ref_bounds(tbl.ref)
    loc_col = min_col + headers.index(location_col_header)
    name_col = min_col + headers.index("Відповідальна особа")
    phone_col = min_col + headers.index("Телефон")
    loc_letter = get_column_letter(loc_col)
    for r in range(min_row + 1, max_row + 1):
        if ws.cell(row=r, column=loc_col).value in (None, ""):
            continue
        name_f, phone_f = responsible_lookup_formulas(loc_letter, r)
        c1 = ws.cell(row=r, column=name_col, value=name_f)
        c2 = ws.cell(row=r, column=phone_col, value=phone_f)
        c1.font = Font(name="Arial")
        c2.font = Font(name="Arial")


def clean_subdivision_name(sheet_name: str) -> str:
    """'батрАР - 06' -> 'батрАР'. Якщо після чистки нічого не лишається
    (назва аркуша й без того без суфіксу місяця) -- повертає як є."""
    cleaned = _TRAILING_MONTH_RE.sub("", sheet_name).strip()
    return cleaned if cleaned else sheet_name


def parse_subdivisions_xlsm(wb_source, year: int):
    """wb_source -- вже завантажена (data_only=True) книга
    'Залишки_На_ХХ_20ХХ.xlsm'. Кожен аркуш = один підрозділ, з рядком
    заголовка (Найменування матеріальних засобів, Одиниця виміру, 1..12)
    і рядками-позиціями майна. Бере ОСТАННІЙ місяць, де є хоч якесь
    значення (по кожному аркушу окремо -- на випадок якщо не всі
    підрозділи синхронно заповнені), як ПОТОЧНИЙ залишок.

    Повертає список кортежів:
        (назва_підрозділу, найменування_майна, одиниця, кількість, дата_обліку)
    """
    records = []
    for sheet_name in wb_source.sheetnames:
        ws = wb_source[sheet_name]
        subdivision = clean_subdivision_name(sheet_name)

        rows = list(ws.iter_rows(min_row=2, values_only=True))
        last_month = 0
        for row in rows:
            if not row or not row[0]:
                continue
            months = row[2:14]
            for mi, v in enumerate(months, start=1):
                if v is not None:
                    last_month = max(last_month, mi)
        if last_month == 0:
            continue  # аркуш без жодного заповненого місяця -- пропускаємо

        record_date = date(year, last_month, 25)  # "станом на 25 число", як у ВІДОМОСТІ

        for row in rows:
            if not row or not row[0]:
                continue
            item_name = str(row[0]).strip()
            unit = row[1]
            qty_row = row[2:14]
            qty = qty_row[last_month - 1] if last_month - 1 < len(qty_row) else None
            if qty is None:
                continue
            records.append((subdivision, item_name, unit, qty, record_date))
    return records


# ================================================================
# ПАРСЕР ОКРЕМОГО DOCX-ФАЙЛУ "ВІДОМІСТЬ наявності речового майна <Х>"
# (той самий формат, що й один аркуш xlsm, тільки один підрозділ/склад
# за раз -- напр. коли підрозділ надсилає лише свій файл, без зведеного
# xlsm по всій бригаді)
# ================================================================
def parse_vidomist_docx(path: str, year: int, location_name: str = None):
    """Повертає (location_name, records), records -- список кортежів
    (найменування_майна, одиниця, кількість, дата_обліку) для ОДНОГО
    підрозділу/складу з docx-файлу формату "ВІДОМІСТЬ наявності
    речового майна <назва>". Якщо location_name не задано -- береться
    з тексту параграфа "наявності речового майна <назва>" у файлі."""
    import docx

    d = docx.Document(path)

    if location_name is None:
        for p in d.paragraphs:
            m = re.search(r"наявності\s+речового\s+майна\s+(.+)", p.text.strip(), re.IGNORECASE)
            if m:
                location_name = m.group(1).strip()
                break
    if location_name is None:
        location_name = Path(path).stem

    if not d.tables:
        return location_name, []

    table = d.tables[0]
    rows = [[c.text.strip() for c in row.cells] for row in table.rows]

    # Рядок-заголовок місяців ("01".."12") -- шукаємо його, бо позиція
    # трохи гуляє між файлами (16 чи 17 колонок залежно від файлу).
    month_header_row_idx = None
    for i, row in enumerate(rows):
        digits = [c for c in row if re.fullmatch(r"0?[1-9]|1[0-2]", c)]
        if len(digits) >= 6:  # принаймні пів року присутнє
            month_header_row_idx = i
            break
    if month_header_row_idx is None:
        return location_name, []

    header_row = rows[month_header_row_idx]
    # Індекси колонок, що відповідають місяцям 1..12 (можуть повторюватись
    # через об'єднані комірки Word -- беремо ПЕРШЕ входження кожного місяця)
    month_col_idx = {}
    for col_idx, val in enumerate(header_row):
        if re.fullmatch(r"0?[1-9]|1[0-2]", val):
            mi = int(val)
            if mi not in month_col_idx:
                month_col_idx[mi] = col_idx

    # Колонки "Найменування" / "Одиниця виміру" -- типово одразу після
    # колонки "№ з/п" (перший стовпець), тобто col 1 і col 2. Перевіряємо
    # це по рядку одразу під заголовком-назвами (не по числовому рядку).
    name_col, unit_col = 1, 2

    records = []
    for row in rows[month_header_row_idx + 1:]:
        if not row or not row[name_col].strip():
            continue
        if row[name_col].strip().upper().startswith("ВСЬОГО"):
            break
        item_name = row[name_col].strip()
        unit = row[unit_col].strip() if unit_col < len(row) else ""

        last_month, last_val = 0, None
        for mi in range(1, 13):
            ci = month_col_idx.get(mi)
            if ci is None or ci >= len(row):
                continue
            val = row[ci].strip()
            if val:
                last_month, last_val = mi, val
        if last_val is None:
            continue
        try:
            qty = float(last_val.replace(",", "."))
            qty = int(qty) if qty.is_integer() else qty
        except ValueError:
            qty = last_val  # нечислове значення -- лишаємо як є, не вигадуємо число

        record_date = date(year, last_month, 25)
        records.append((item_name, unit, qty, record_date))

    return location_name, records


def find_images_recursive(root_folder, extensions=(".jpg", ".jpeg", ".png", ".webp")):
    """Рекурсивний пошук фото в папці й усіх підпапках (для ~300+ фото,
    розкладених по підпапках, а не в одній купі)."""
    root = Path(root_folder)
    files = [p for p in root.rglob("*") if p.suffix.lower() in extensions]
    return sorted(str(p) for p in files)


# ================================================================
# ДОВІДНИК ВІДПОВІДАЛЬНИХ ОСІБ (МВО) ПО СКЛАДАХ/ПІДРОЗДІЛАХ
# ================================================================
RESPONSIBLE_SHEET = "Відповідальні особи"
RESPONSIBLE_TABLE = "Vidpovidalni"
RESPONSIBLE_HEADERS = ["Назва (склад/підрозділ)", "Тип", "Відповідальна особа",
                        "Телефон", "Примітка"]


def ensure_responsible_sheet(wb):
    """Створює аркуш 'Відповідальні особи' якщо його ще нема (порожній,
    для ручного заповнення -- жодне із джерел даних це не приносить
    автоматично, це те, що користувач вписує сам)."""
    if RESPONSIBLE_SHEET in wb.sheetnames:
        return wb[RESPONSIBLE_SHEET]
    ws = wb.create_sheet(RESPONSIBLE_SHEET)
    style_header(ws, RESPONSIBLE_HEADERS)
    example = ["Приклад: Склад №1 (речовий)", "Склад", "Прізвище Ім'я По батькові",
               "+380XXXXXXXXX", "ПРИКЛАД -- замініть/додайте реальні рядки"]
    for col, value in enumerate(example, start=1):
        c = ws.cell(row=2, column=col, value=value)
        c.font = Font(name="Arial")
        c.fill = EXAMPLE_FILL
    add_or_resize_table(ws, RESPONSIBLE_TABLE, len(RESPONSIBLE_HEADERS), 2)
    for col_idx, w in zip(range(1, 6), [26, 12, 26, 16, 40]):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    return ws


def responsible_lookup_formulas(location_col_letter: str, row: int):
    """Формули INDEX/MATCH для колонок 'Відповідальна особа' і 'Телефон',
    що підтягують дані з довідника Vidpovidalni за назвою складу/підрозділу
    -- замість дублювання імені й телефону в кожному з тисяч рядків майна."""
    name_f = (f'=IFERROR(INDEX({RESPONSIBLE_TABLE}[Відповідальна особа],'
              f'MATCH({location_col_letter}{row},{RESPONSIBLE_TABLE}[Назва (склад/підрозділ)],0)),"")')
    phone_f = (f'=IFERROR(INDEX({RESPONSIBLE_TABLE}[Телефон],'
               f'MATCH({location_col_letter}{row},{RESPONSIBLE_TABLE}[Назва (склад/підрозділ)],0)),"")')
    return name_f, phone_f
