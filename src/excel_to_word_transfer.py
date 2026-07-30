# -*- coding: utf-8 -*-
"""
excel_to_word_transfer.py

Python-заміна VBA-макросу TransferByPageBreaks.

Логіка:
  1. Відкриваємо Excel-файл з даними (openpyxl).
  2. Для кожного аркуша визначаємо межі "сторінок" за розривами сторінок
     (row page breaks), як у VBA (HPageBreaks).
  3. На кожній сторінці шукаємо рядок-шапку, де колонка A == "№ п/п".
  4. Збираємо рядки даних (колонка A - число, колонка B - не число і не
     порожня) аж до рядка "Всього" (об'єднані комірки A+B з текстом
     "Всього").
  5. Копіюємо Word-шаблон під назвою аркуша, дописуємо потрібну кількість
     рядків у першу таблицю документа і записуємо туди дані.

Немає залежності від встановленого MS Word / COM - все через python-docx,
тому працює навіть без Office.

Встановлення залежностей (Git Bash):
    pip install openpyxl python-docx
"""

import os
import re
import shutil
import zipfile
from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple

import openpyxl
from docx import Document

LogFunc = Callable[[str], None]

# Колонки, які читаємо з Excel (1-indexed, як в openpyxl)
COL_A = 1   # № п/п
COL_B = 2
COL_C = 3
COL_D = 4
COL_E = 5
COL_M = 13

INVALID_FS_CHARS = r'\/:*?"<>|'


class TemplateError(Exception):
    """Шаблон .docx непридатний для використання (не .docx / пошкоджений)."""


def validate_docx_template(template_path: str) -> None:
    """Перевіряє шаблон ОДИН РАЗ, до обробки аркушів — щоб дати одну чітку
    помилку одразу, а не намагатись відкрити копію шаблону для КОЖНОГО
    аркуша і плутати користувача криптичним повідомленням python-docx
    (напр. "is not a Word file, content type is '...themeManager+xml'"),
    яке насправді означає лише одне: файл — не справжній .docx-документ.
    """
    if not os.path.exists(template_path):
        raise TemplateError(f"Шаблон не знайдено: '{template_path}'")

    ext = os.path.splitext(template_path)[1].lower()
    if ext != ".docx":
        raise TemplateError(
            f"Шаблон '{template_path}' має розширення '{ext}', а очікується "
            f"'.docx'. Стара двійкова '.doc' (Word 97-2003) чи шаблон "
            f"'.dotx'/'.dotm' тут НЕ підійде — python-docx працює лише з "
            f"форматом .docx. Відкрий файл у Word і зроби "
            f"'Файл -> Зберегти як...' -> тип файлу 'Документ Word (*.docx)', "
            f"і передай скрипту саме цей новий файл."
        )

    if not zipfile.is_zipfile(template_path):
        raise TemplateError(
            f"Шаблон '{template_path}' не є коректним .docx-файлом (це не "
            f"ZIP-архів, хоча має розширення .docx) — ймовірно, це стара "
            f".doc-структура з підміненим розширенням, файл пошкоджений, "
            f"або скачаний не повністю. Перезбережи його у Word як "
            f"'Документ Word (*.docx)' і спробуй ще раз."
        )

    try:
        Document(template_path)
    except Exception as exc:
        raise TemplateError(
            f"Шаблон '{template_path}' не вдалось відкрити як Word-документ "
            f"({exc}). Найімовірніше, файл насправді не .docx (напр. це "
            f"Excel/PowerPoint файл з переплутаним розширенням, або "
            f"пошкоджений .docx). Перезбережи файл у Word: "
            f"'Файл -> Зберегти як...' -> 'Документ Word (*.docx)'."
        ) from exc


def _default_log(msg: str) -> None:
    print(msg)


def sanitize_filename(name: str) -> str:
    """Аналог заміни недопустимих символів у назві файлу, як у VBA."""
    result = name
    for ch in INVALID_FS_CHARS:
        result = result.replace(ch, "_")
    return result.strip()


def _cell(ws, row: int, col: int):
    return ws.cell(row=row, column=col)


def _is_numeric(value) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip()
    if text == "":
        return False
    try:
        float(text.replace(",", "."))
        return True
    except ValueError:
        return False


def get_merge_range(ws, row: int, col: int):
    """Повертає MultiCellRange-об'єкт, якщо (row, col) входить у злиття."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng
    return None


def is_vsogo_row(ws, row: int) -> bool:
    """Аналог IsVsogo: A і B об'єднані разом і містять слово 'Всього'."""
    range_a = get_merge_range(ws, row, COL_A)
    range_b = get_merge_range(ws, row, COL_B)
    if range_a is None or range_b is None:
        return False
    if range_a.coord != range_b.coord:
        return False
    master = ws.cell(row=range_a.min_row, column=range_a.min_col)
    value = str(master.value) if master.value is not None else ""
    return "всього" in value.lower()


def is_data_row(ws, row: int) -> bool:
    """Аналог IsDataRow: A - число і не порожнє, B - НЕ число і не порожнє."""
    a_val = _cell(ws, row, COL_A).value
    b_val = _cell(ws, row, COL_B).value
    a_ok = _is_numeric(a_val) and str(a_val).strip() != ""
    b_ok = (not _is_numeric(b_val)) and b_val is not None and str(b_val).strip() != ""
    return a_ok and b_ok


def find_header_row(ws, start_row: int, end_row: int) -> int:
    """Шукає рядок, де колонка A == '№ п/п'. Повертає 0, якщо не знайдено."""
    for r in range(start_row, end_row + 1):
        val = _cell(ws, r, COL_A).value
        if val is not None and str(val).strip() == "№ п/п":
            return r
    return 0


def get_last_row(ws) -> int:
    last_row = 1
    for r in range(ws.max_row, 0, -1):
        if _cell(ws, r, COL_A).value not in (None, ""):
            last_row = r
            break
    return last_row


def get_page_bounds(ws) -> List[Tuple[int, int]]:
    """
    Визначає межі сторінок за row page breaks (аналог HPageBreaks у VBA).
    Якщо розривів немає - весь аркуш вважається однією сторінкою.

    openpyxl: Break.id - номер рядка, ПІСЛЯ якого стоїть розрив
    (тобто останній рядок попередньої сторінки).
    """
    last_row = get_last_row(ws)
    break_ids = sorted({brk.id for brk in ws.row_breaks.brk}) if ws.row_breaks else []
    break_ids = [b for b in break_ids if 0 < b < last_row]

    bounds = []
    start = 1
    for bid in break_ids:
        bounds.append((start, bid))
        start = bid + 1
    bounds.append((start, last_row))
    return bounds


@dataclass
class SheetResult:
    sheet_name: str
    rows_collected: int
    output_path: Optional[str]
    skipped_reason: Optional[str] = None


def collect_sheet_data(ws) -> List[List[str]]:
    """Проходить усі сторінки аркуша і збирає рядки даних у список списків
    [B, C, D, E, M] як текст (як dataArr у VBA)."""
    data: List[List[str]] = []
    for page_start, page_end in get_page_bounds(ws):
        header_row = find_header_row(ws, page_start, page_end)
        if header_row == 0:
            continue
        data_start = header_row + 2
        r = data_start
        while r <= page_end:
            if is_vsogo_row(ws, r):
                break
            if is_data_row(ws, r):
                row_vals = [
                    _cell(ws, r, COL_B).value,
                    _cell(ws, r, COL_C).value,
                    _cell(ws, r, COL_D).value,
                    _cell(ws, r, COL_E).value,
                    _cell(ws, r, COL_M).value,
                ]
                data.append(["" if v is None else str(v) for v in row_vals])
            r += 1
    return data


def fill_word_table(template_path: str, save_path: str, data: List[List[str]]) -> None:
    """Копіює шаблон, дописує потрібну кількість рядків у першу таблицю
    і записує туди дані (аналог блоку роботи з Word у VBA)."""
    shutil.copy(template_path, save_path)
    doc = Document(save_path)

    if not doc.tables:
        os.remove(save_path)
        raise ValueError("У шаблоні немає жодної таблиці")

    table = doc.tables[0]

    needed_rows = 7 + len(data) - 1  # аналог neededRows у VBA
    rows_to_add = needed_rows - len(table.rows)
    for _ in range(max(rows_to_add, 0)):
        table.add_row()

    word_row = 7  # 1-indexed, як у VBA
    for row_data in data:
        b, c, d, e, m = row_data
        r_idx = word_row - 1  # python-docx 0-indexed
        table.cell(r_idx, 2).text = b    # Cell(wordRow, 3)
        table.cell(r_idx, 4).text = c    # Cell(wordRow, 5)
        table.cell(r_idx, 5).text = d    # Cell(wordRow, 6)
        table.cell(r_idx, 6).text = e    # Cell(wordRow, 7)
        table.cell(r_idx, 7).text = m    # Cell(wordRow, 8)
        table.cell(r_idx, 17).text = m   # Cell(wordRow, 18)
        word_row += 1

    doc.save(save_path)


def process_workbook(
    template_path: str,
    excel_path: str,
    log: Optional[LogFunc] = None,
) -> List[SheetResult]:
    """Головна функція - аналог TransferByPageBreaks."""
    log = log or _default_log

    validate_docx_template(template_path)

    template_ext = os.path.splitext(template_path)[1]
    output_folder = os.path.dirname(os.path.abspath(template_path))

    log(f"Відкриваємо Excel: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    results: List[SheetResult] = []
    total_sheets = len(wb.sheetnames)

    for idx, sheet_name in enumerate(wb.sheetnames, start=1):
        ws = wb[sheet_name]
        log(f"[{idx}/{total_sheets}] Аркуш: {sheet_name}")

        data = collect_sheet_data(ws)
        if not data:
            log(f"    -> пропущено (немає рядків даних)")
            results.append(SheetResult(sheet_name, 0, None, "немає даних"))
            continue

        safe_name = sanitize_filename(sheet_name)
        save_path = os.path.join(output_folder, f"{safe_name}{template_ext}")

        try:
            fill_word_table(template_path, save_path, data)
        except Exception as exc:
            log(f"    -> ПОМИЛКА: {exc}")
            results.append(SheetResult(sheet_name, len(data), None, str(exc)))
            continue

        log(f"    -> записано {len(data)} рядків -> {save_path}")
        results.append(SheetResult(sheet_name, len(data), save_path))

    wb.close()
    log("Готово.")
    return results


if __name__ == "__main__":
    # Простий запуск з командного рядка (без GUI):
    #   python excel_to_word_transfer.py шаблон.docx дані.xlsx
    import sys

    if len(sys.argv) != 3:
        print("Використання: python excel_to_word_transfer.py <шаблон.docx> <дані.xlsx>")
        sys.exit(1)

    try:
        process_workbook(sys.argv[1], sys.argv[2])
    except TemplateError as exc:
        print(f"ПОМИЛКА ШАБЛОНУ: {exc}")
        sys.exit(1)
