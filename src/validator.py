# -*- coding: utf-8 -*-
"""Скрипт для валідації JSON-файлів звітів, згенерованих batch_processor.py."""

import json
import argparse
from pathlib import Path
import sys

# Імпортуємо _UNIT_ALIASES з llm_extractor для валідації одиниць виміру.
# Передбачається, що llm_extractor.py знаходиться в тому ж каталозі.
try:
    from llm_extractor import _UNIT_ALIASES
    KNOWN_UNITS = set(_UNIT_ALIASES.values())
except ImportError:
    # Запасний варіант, якщо llm_extractor недоступний або _UNIT_ALIASES не знайдено.
    # В цьому випадку перевірка одиниць виміру буде обмеженою.
    sys.stderr.write("WARNING: Не вдалося імпортувати _UNIT_ALIASES з llm_extractor.py. Перевірка одиниць виміру буде обмеженою.\n")
    KNOWN_UNITS = {"шт.", "пара", "к-т"} # Базові запасні одиниці


# Допоміжні функції для валідації
def _is_non_empty_string(value):
    """Перевіряє, чи є значення непорожнім рядком."""
    return isinstance(value, str) and len(value.strip()) > 0

def _is_float_between_0_and_1(value):
    """Перевіряє, чи є значення числом з плаваючою точкою (або цілим) між 0.0 і 1.0 включно."""
    if not isinstance(value, (int, float)):
        return False
    return 0.0 <= float(value) <= 1.0

def _is_valid_unit(value):
    """Перевіряє, чи є значення дійсною одиницею виміру з KNOWN_UNITS."""
    return isinstance(value, str) and value in KNOWN_UNITS

def write_json(path: Path, data):
    """
    Записує JSON-дані у файл.
    Створює батьківські директорії, якщо вони не існують.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def validate_report_json(json_path: Path):
    """
    Валідує структуру та вміст JSON-файлу звіту.

    Args:
        json_path (Path): Шлях до JSON-файлу.

    Returns:
        tuple[bool, list[str]]: Кортеж, що містить:
                                - bool: True, якщо JSON валідний, False інакше.
                                - list[str]: Список повідомлень про помилки, якщо вони є.
    """
    errors = []
    
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        errors.append(f"Файл не знайдено: {json_path.name}")
        return False, errors
    except json.JSONDecodeError as e:
        errors.append(f"Недійсний JSON: {e}")
        return False, errors
    except Exception as e:
        errors.append(f"Помилка читання файлу: {e}")
        return False, errors
    except Exception as e:
        errors.append(f"Помилка читання файлу: {e}")
        return False, errors

    # Перевірка, що кореневий елемент JSON є об'єктом
    if not isinstance(data, dict):
        errors.append("Кореневий елемент JSON повинен бути об'єктом, а не списком.")
        return False, errors

    # 1. Валідація поля 'status'
    status = data.get("status")
    if status != "ok":
        error_message = data.get("error", "відсутня")
        errors.append(f"Поле 'status' не 'ok'. Поточний статус: '{status}'. Помилка: '{error_message}'")

    # 2. Валідація поля 'document'
    document = data.get("document")
    if not isinstance(document, dict):
        errors.append("Поле 'document' відсутнє або не є об'єктом.")
    else:
        for field in ["date", "location", "unit"]:
            if not _is_non_empty_string(document.get(field)):
                errors.append(f"Поле 'document.{field}' відсутнє або порожнє.")

    # 3. Валідація поля 'items'
    items = data.get("items")
    if not isinstance(items, list) or not items:
        errors.append("Поле 'items' відсутнє, не є списком або порожнє.")
    else:
        for i, item in enumerate(items):
            if not isinstance(item, dict):
                errors.append(f"Елемент 'items[{i}]' не є об'єктом.")
                continue

            # Перевірка наявності обов'язкових полів
            mandatory_fields = ["name", "quantity", "unit", "match_found", "match_score"]
            for field in mandatory_fields:
                if field not in item:
                    errors.append(f"Елемент 'items[{i}]' не містить поля '{field}'.")

            # Валідація name, quantity, unit (непорожні рядки)
            if not _is_non_empty_string(item.get("name")):
                errors.append(f"Елемент 'items[{i}].name' відсутній або порожній.")
            if not _is_non_empty_string(item.get("quantity")):
                errors.append(f"Елемент 'items[{i}].quantity' відсутній або порожній.")
            if not _is_non_empty_string(item.get("unit")):
                errors.append(f"Елемент 'items[{i}].unit' відсутній або порожній.")

            # Валідація match_found (булеве значення)
            match_found = item.get("match_found")
            if not isinstance(match_found, bool):
                errors.append(f"Елемент 'items[{i}].match_found' відсутній або не є булевим значенням.")

            # Валідація match_score (число з плаваючою точкою між 0 і 1)
            match_score = item.get("match_score")
            if not _is_float_between_0_and_1(match_score):
                errors.append(f"Елемент 'items[{i}].match_score' відсутній або не є числом між 0.0 і 1.0.")

            # Валідація unit (одна з відомих одиниць)
            item_unit = item.get("unit")
            if not _is_valid_unit(item_unit):
                errors.append(f"Елемент 'items[{i}].unit' ('{item_unit}') не є відомою одиницею виміру (дозволені: {', '.join(sorted(list(KNOWN_UNITS))) if KNOWN_UNITS else 'не визначено'}).")

    return not bool(errors), errors

def find_json_files(root_dir: Path, recursive: bool = True):
    """
    Знаходить JSON-файли у вказаній директорії.

    Args:
        root_dir (Path): Коренева директорія для пошуку.
        recursive (bool): Чи шукати у підпапках.

    Returns:
        list[Path]: Відсортований список шляхів до знайдених JSON-файлів.
    """
    pattern = "**/*.json" if recursive else "*.json"
    return sorted(p for p in root_dir.glob(pattern) if p.is_file())

def main():
    """Головна функція для виконання валідації з командного рядка."""
    parser = argparse.ArgumentParser(description="Валідація JSON-файлів звітів.")
    parser.add_argument("folder", help="Шлях до папки з JSON-файлами для валідації.")
    parser.add_argument("--non-recursive", action="store_true", help="Не шукати у підпапках.")
    parser.add_argument("--summary", action="store_true", help="Виводити тільки підсумок результатів.")
    parser.add_argument("--report-file", type=Path, help="Шлях до файлу для збереження детального звіту JSON.")
    args = parser.parse_args()

    root_dir = Path(args.folder).expanduser().resolve()
    if not root_dir.is_dir():
        print(f"Помилка: Папку не знайдено: {root_dir}", file=sys.stderr)
        return 2

    json_files = find_json_files(root_dir, recursive=not args.non_recursive)

    if not json_files:
        print(f"Не знайдено жодного JSON-файлу у {root_dir}")
        return 0

    total_files = len(json_files)
    valid_files = 0
    invalid_files = 0
    
    validation_results = [] # Список для збору детальних результатів

    print(f"Починаю валідацію {total_files} JSON-файлів у {root_dir}...")

    for idx, json_file in enumerate(json_files):
        is_valid, errors = validate_report_json(json_file)
        
        result_entry = { # Додаємо запис до детального звіту
            "file": str(json_file.relative_to(root_dir)), # Відносний шлях для звіту
            "is_valid": is_valid,
            "errors": errors
        }
        validation_results.append(result_entry)

        if is_valid:
            valid_files += 1
            if not args.summary:
                print(f"[{idx+1}/{total_files}] OK: {json_file.name}")
        else:
            invalid_files += 1
            print(f"[{idx+1}/{total_files}] ПОМИЛКА: {json_file.name}")
            if not args.summary:
                for error in errors:
                    print(f"  - {error}")
    
    print("\n--- ПІДСУМОК ---")
    print(f"Всього файлів: {total_files}")
    print(f"Валідних файлів: {valid_files}")
    print(f"Невалідних файлів: {invalid_files}")

    # Збереження детального звіту, якщо вказано --report-file
    if args.report_file:
        try:
            write_json(args.report_file, validation_results)
            print(f"Детальний звіт збережено у: {args.report_file}")
        except Exception as e:
            print(f"Помилка збереження звіту у {args.report_file}: {e}", file=sys.stderr)

    return 1 if invalid_files > 0 else 0

if __name__ == "__main__":
    raise SystemExit(main())
