# -*- coding: utf-8 -*-
"""Batch JSON bridge for OCR txt -> LLM extracted report data.

Usage:
    python batch_processor.py C:/path/to/output_txt
    python batch_processor.py C:/path/to/output_txt --no-llm
"""

import argparse
import os
import json
import re
import sys
import traceback
from pathlib import Path

# Додаємо шлях до директорії з llm_extractor.py до sys.path
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from llm_extractor import (  # type: ignore
    DictionaryMatcher,
    LlmExtractor,
    LlmExtractorError,
    extract_items_block,
    fallback_extract_items,
    normalize_qty,
    normalize_unit,
)

from excel_utils.excel_writer import ( # type: ignore
    TEMPLATE_SHEET_NAME,
    CELL_VIDOMIST_NUMBER,
    CELL_UNIT,
    CELL_DATE_PLACE,
    ROW_ITEMS_START,
    COLOR_YELLOW,
    _get_next_vidomist_number,
    clean_sheet_name,
    safe_clear_cell,
    safe_clear_color,
    fill_error,
    write_json_items_to_sheet,
    create_workbook_from_template,
    reset_vidomist_counter,
    START_VIDOMIST_NUMBER,
    SKIP_NUMBER, # Для логування
)

from typing import List, Dict, Any, Optional
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def read_text(path):
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")

def find_txt_files(root, recursive=True):
    pattern = "**/*.txt" if recursive else "*.txt"
    return sorted(p for p in root.glob(pattern) if p.is_file())

def fallback_fields(text):
    date = ""
    location = ""
    unit = ""

    date_match = re.search(r"\b(\d{2}\.\d{2}\.\d{4})\b", text)
    if date_match:
        date = date_match.group(1)

    place_match = re.search(
        r"(?:н\.?\s*п\.?|населен[а-яіїєґ]+\s+пункт[а-яіїєґ]*|районі)\s+([А-ЯІЇЄҐ][А-Яа-яІіЇїЄєҐґ'’\-]+)",
        text,
    )
    if place_match:
        location = place_match.group(1).strip(" .,;:")

    unit_match = re.search(r"([А-Яа-яІіЇїЄєҐґ'’\-\s]{6,80}(?:дивізіону|батальйону|бригади|роти|служби))", text)
    if unit_match:
        unit = re.sub(r"\s+", " ", unit_match.group(1)).strip()

    return {"unit": unit, "date": date, "place": location}

def normalise_item(item, matcher):
    raw_name = str(item.get("name", "")).strip()
    qty = normalize_qty(item.get("qty", item.get("quantity", "")))
    unit = normalize_unit(item.get("unit", ""))
    match = matcher.match(raw_name) if matcher is not None else None
    canonical = match.canonical if match and match.found else raw_name
    out = {
        "name": canonical,
        "raw_name": raw_name,
        "unit": unit,
        "quantity": qty,
    }
    if match is not None:
        out["match_score"] = round(match.score, 3)
        out["match_found"] = bool(match.found)
    return out

def process_txt_file_content(txt_path: Path, text_content: str, extractor: LlmExtractor, matcher: DictionaryMatcher, use_llm: bool):
    """
    Обробляє вміст одного TXT-файлу, витягуючи поля та позиції майна.
    Повертає словник з витягнутими даними або інформацію про помилку.
    """
    warnings = []

    if not text_content.strip():
        warnings.append("Файл порожній або нечитабельний")
        return {
            "status": "error",
            "error_message": "Файл порожній або нечитабельний",
            "fields": {},
            "items": [],
            "warnings": warnings,
        }

    items_block = extract_items_block(text_content)
    if not items_block:
        warnings.append("Блок речової служби не знайдено, для items використано весь текст")
        print(f"WARNING: Блок речової служби не знайдено у '{txt_path.name}', для items використано весь текст.")
        items_block = text_content
    else:
        print(f"DEBUG: Блок речової служби знайдено у '{txt_path.name}' ({len(items_block)} символів).")

    fields = None
    raw_items = None

    if use_llm and extractor is not None:
        print(f"INFO: Використовую LLM для витягнення полів та позицій майна з '{txt_path.name}'.")
        # Extract fields
        llm_fields = extractor.extract_report_fields(text_content)
        if llm_fields is None:
            warnings.append("LLM не витягнула поля документа, застосовано fallback")
            print(f"WARNING: LLM не витягнула поля документа з '{txt_path.name}', застосовую fallback.")
            fields = fallback_fields(text_content)
        else:
            fields = llm_fields
            print(f"DEBUG: Поля документа для '{txt_path.name}' успішно витягнуто LLM.")

        # Extract items
        llm_items = extractor.extract_items(items_block)
        if llm_items is None:
            warnings.append("LLM не витягнула items, застосовано fallback")
            print(f"WARNING: LLM не витягнула позиції майна з '{txt_path.name}', застосовую fallback.")
            raw_items = fallback_extract_items(items_block)
        else:
            raw_items = llm_items
            print(f"DEBUG: Позиції майна для '{txt_path.name}' успішно витягнуто LLM.")
    else:
        print(f"INFO: Режим без LLM (--no-llm) або екстрактор недоступний для '{txt_path.name}', використовую fallback.")
        fields = fallback_fields(text_content)
        raw_items = fallback_extract_items(items_block)
        print(f"DEBUG: Використано fallback для полів та позицій майна для '{txt_path.name}'.")

    if not raw_items:
        warnings.append("Не вдалося знайти жодної позиції майна")
        print(f"WARNING: Не вдалося знайти жодної позиції майна у файлі '{txt_path.name}'.")
        return {
            "status": "partial_error",
            "error_message": "Не вдалося знайти жодної позиції майна",
            "fields": fields,
            "items": [],
            "warnings": warnings,
        }

    items = [normalise_item(item, matcher) for item in raw_items]

    return {
        "status": "ok",
        "fields": fields,
        "items": items,
        "warnings": warnings,
    }

def main(argv=None):
    parser = argparse.ArgumentParser(description="Generate intermediate JSON from OCR txt reports.")
    parser.add_argument("folder", help="Folder with OCR .txt files")
    parser.add_argument("--output-dir", default="", help="Output directory path for Excel files (defaults to input folder)")
    parser.add_argument("--no-llm", action="store_true", help="Use only deterministic fallback extraction")
    parser.add_argument("--non-recursive", action="store_true", help="Do not search subfolders")
    parser.add_argument("--template", default="template.xlsm", help="Path to the Excel template file")
    # Додамо опцію для overwrite, якщо Excel-файл вже існує
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing Excel files")
    args = parser.parse_args(argv)

    root_folder = Path(args.folder).expanduser().resolve()
    if not root_folder.is_dir():
        print(f"ERROR: Вхідна папка не знайдена: {root_folder}", file=sys.stderr)
        return 2

    # Визначення шляху до шаблону
    template_path = Path(args.template)
    if not template_path.is_absolute():
        # Спершу спробуємо знайти шаблон відносно директорії самого скрипта (python_scripts/batch_processor.py)
        temp_check_path_relative_to_script = SCRIPT_DIR / template_path
        if temp_check_path_relative_to_script.is_file():
            template_path = temp_check_path_relative_to_script
        else:
            # Якщо не знайдено, спробуємо знайти в папці 'templates' на рівні кореня проекту
            # SCRIPT_DIR.parent -> project_root/
            template_path = SCRIPT_DIR.parent / "templates" / args.template

    if not template_path.is_file():
        print(f"ERROR: Файл-шаблон Excel не знайдено за шляхом: {template_path}", file=sys.stderr)
        return 2

    print(f"INFO: Використовується шаблон Excel: {template_path}")

    # Групуємо файли за батьківською папкою, як у VBA
    all_txt_files = find_txt_files(root_folder, recursive=not args.non_recursive)
    if not all_txt_files:
        print(f"Не знайдено .txt файлів у {root_folder}")
        return 0

    grouped_txt_files: Dict[Path, List[Path]] = {}
    for txt_path in all_txt_files:
        parent_folder = txt_path.parent
        if parent_folder not in grouped_txt_files:
            grouped_txt_files[parent_folder] = []
        grouped_txt_files[parent_folder].append(txt_path)

    if not grouped_txt_files:
        print(f"Не знайдено папок з .txt файлами у {root_folder}")
        return 0

    matcher = DictionaryMatcher()
    use_llm = not args.no_llm
    extractor = None
    total_failed_files = 0

    print(f"INFO: Знайдено папок з .txt файлами: {len(grouped_txt_files)}")

    try:
        if use_llm:
            print("INFO: Ініціалізація LLM екстрактора...")
            extractor = LlmExtractor()
            extractor.start()
        else:
            print("INFO: Режим без LLM (--no-llm), LLM екстратор не запускається.")

        # Скидаємо лічильник номерів відомостей для кожної серії обробки
        reset_vidomist_counter(START_VIDOMIST_NUMBER)

        for folder_idx, (folder_path, txt_files_in_folder) in enumerate(grouped_txt_files.items(), 1):
            print(f"\nINFO: Обробка папки [{folder_idx}/{len(grouped_txt_files)}]: {folder_path.name} ({len(txt_files_in_folder)} файлів)")

            # Визначення шляху для збереження результуючого Excel-файлу
            output_excel_dir = Path(args.output_dir) if args.output_dir else folder_path
            output_excel_dir.mkdir(parents=True, exist_ok=True)

            # Назва файлу як ім'я папки
            excel_filename = folder_path.name + ".xlsx"
            output_excel_path = output_excel_dir / excel_filename

            if output_excel_path.exists() and not args.overwrite:
                print(f"INFO: Пропускаю існуючий Excel-файл '{output_excel_path.name}' (використайте --overwrite для перезапису).")
                continue

            wb: Optional[Workbook] = None
            try:
                wb = create_workbook_from_template(str(template_path))
                ws_template = wb[TEMPLATE_SHEET_NAME]
                print(f"INFO: Створено робочу книгу з шаблону '{TEMPLATE_SHEET_NAME}'.")

                for file_idx, txt_path in enumerate(txt_files_in_folder, 1):
                    print(f"  INFO: Обробка файлу [{file_idx}/{len(txt_files_in_folder)}]: {txt_path.name}")
                    text_content = read_text(txt_path)

                    # Пропускаємо файли з json_output директорії, якщо вони випадково потрапили
                    if output_excel_dir in txt_path.parents:
                        print(f"  INFO: Пропускаю файл '{txt_path.name}' з директорії виводу.")
                        continue

                    processed_data = None
                    try:
                        processed_data = process_txt_file_content(txt_path, text_content, extractor, matcher, use_llm)
                    except Exception as exc:
                        error_msg = f"Помилка при обробці '{txt_path.name}': {exc}\n{traceback.format_exc(limit=5)}"
                        print(f"  ERROR: {error_msg}", file=sys.stderr)
                        processed_data = {
                            "status": "error",
                            "error_message": error_msg,
                            "fields": {},
                            "items": [],
                            "warnings": [],
                        }

                    # Створюємо новий аркуш
                    sheet_name = clean_sheet_name(txt_path.stem)
                    # openpyxl додає аркуші після існуючих, а не на початку
                    ws_new = wb.copy_worksheet(ws_template) # Копіюємо шаблон
                    ws_new.title = sheet_name
                    print(f"  INFO: Створено аркуш: '{sheet_name}'.")

                    # Заповнення загальних полів відомості
                    vid_number = _get_next_vidomist_number()
                    ws_new[CELL_VIDOMIST_NUMBER].value = f"ВІДОМІСТЬ № {vid_number}"
                    print(f"  DEBUG: Аркуш '{sheet_name}': номер відомості {vid_number}.")

                    if processed_data["status"] != "ok":
                        total_failed_files += 1
                        print(f"  WARNING: Обробка '{txt_path.name}' завершилася зі статусом '{processed_data['status']}'.")
                        print(f"  ERROR: {processed_data.get('error_message', 'Невідома помилка')}", file=sys.stderr)

                        # Заповнюємо поля помилки
                        fill_error(ws_new, ws_new[CELL_UNIT].row, ws_new[CELL_UNIT].column, "[ НЕ ЗНАЙДЕНО — перевірте вручну ]")
                        fill_error(ws_new, ws_new[CELL_DATE_PLACE].row, ws_new[CELL_DATE_PLACE].column, "[ НЕ ЗНАЙДЕНО — перевірте вручну ]")
                        fill_error(ws_new, ROW_ITEMS_START, 2, "[ МАЙНО не знайдено — перевірте вручну ]")

                        # Додаємо попередження до комірки для діагностики
                        warn_text = "; ".join(processed_data.get("warnings", []))
                        if warn_text:
                             ws_new.cell(row=ROW_ITEMS_START + 1, column=2).value = f"ПОПЕРЕДЖЕННЯ: {warn_text}"
                             ws_new.cell(row=ROW_ITEMS_START + 1, column=2).fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid") # Помаранчевий

                    else:
                        fields = processed_data["fields"]
                        items = processed_data["items"]

                        # Заповнення полів рапорту
                        unit = str(fields.get("unit", "")).strip()
                        date = str(fields.get("date", "")).strip()
                        place = str(fields.get("place", fields.get("location", ""))).strip() # 'place' або 'location'

                        ws_new[CELL_UNIT].value = unit
                        ws_new[CELL_DATE_PLACE].value = f"{date} {place}".strip()

                        if not unit or not date or not place:
                             # Якщо якісь поля відсутні, позначаємо жовтим
                             if not unit:
                                 fill_error(ws_new, ws_new[CELL_UNIT].row, ws_new[CELL_UNIT].column)
                             if not date or not place:
                                 fill_error(ws_new, ws_new[CELL_DATE_PLACE].row, ws_new[CELL_DATE_PLACE].column)
                             print(f"  WARNING: Аркуш '{sheet_name}': деякі поля рапорту неповні. unit='{unit}', date='{date}', place='{place}'.")
                        else:
                             print(f"  DEBUG: Аркуш '{sheet_name}': поля заповнено: unit='{unit}', date='{date}', place='{place}'.")

                        # Заповнення позицій майна
                        write_json_items_to_sheet(ws_new, items, base_row=ROW_ITEMS_START)
                        print(f"  DEBUG: Аркуш '{sheet_name}': заповнено {len(items)} позицій майна.")

                    # Очищаємо тимчасовий аркуш шаблону, якщо він був скопійований,
                    # але потрібно переконатися, що він завжди є як джерело для копіювання.
                    # openpyxl.Workbook.copy_worksheet створює копію, яку потім можна перейменувати.
                    # Оригінальний шаблон не видаляється до збереження.

                # Видаляємо оригінальний аркуш шаблону перед збереженням
                if TEMPLATE_SHEET_NAME in wb.sheetnames:
                    del wb[TEMPLATE_SHEET_NAME]

                # Зберігаємо готову робочу книгу
                wb.save(output_excel_path)
                print(f"INFO: Збережено Excel-файл до: {output_excel_path}")

            except Exception as exc:
                total_failed_files += 1
                error_msg = f"Критична помилка при обробці папки '{folder_path.name}': {exc}\n{traceback.format_exc(limit=5)}"
                print(f"CRITICAL ERROR: {error_msg}", file=sys.stderr)
                if wb:
                    try:
                        # Якщо сталася помилка, спробуємо зберегти з назвою помилки, якщо можливо
                        error_excel_path = output_excel_dir / f"ERROR_{folder_path.name}.xlsx"
                        wb.save(error_excel_path)
                        print(f"INFO: Збережено Excel-файл з помилкою до: {error_excel_path}")
                    except Exception as save_exc:
                        print(f"ERROR: Не вдалося зберегти Excel-файл з помилкою: {save_exc}", file=sys.stderr)
            finally:
                if wb:
                    # openpyxl не має явного .close() для Workbook, об'єкт збирається збирачем сміття
                    pass

    except LlmExtractorError as exc:
        total_failed_files += 1
        message = f"Критична помилка запуску LLM: {exc}"
        print(f"CRITICAL ERROR: {message}", file=sys.stderr)
        # У випадку критичної помилки LLM, ми не можемо обробити файли,
        # але повинні повернути ненульовий код.
    except Exception as exc:
        total_failed_files += 1
        print(f"CRITICAL ERROR: Загальна помилка в main: {exc}\n{traceback.format_exc(limit=5)}", file=sys.stderr)
    finally:
        if extractor is not None:
            print("INFO: Зупиняю LLM екстрактор...")
            extractor.stop()
        print("INFO: Процес batch_processor завершено.")

    return 1 if total_failed_files else 0

if __name__ == "__main__":
    raise SystemExit(main())
