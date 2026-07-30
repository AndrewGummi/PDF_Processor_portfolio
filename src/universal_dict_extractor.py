# -*- coding: utf-8 -*-
"""
Універсальний словниковий екстрактор.

ЩО РОБИТЬ:
  1. Питає шлях до макету (xlsm з аркушем "Словник": A=найменування,
     B=одиниці виміру, C=сюди пишемо знайдені значення).
  2. Питає папку, де рекурсивно (усі підпапки й підпапки підпапок) шукати
     однотипні файли-джерела (.xlsx, .xlsm, .docx).
  3. У кожному файлі, у кожному аркуші/таблиці, знаходить рядки даних —
     БЕЗ прив'язки до фіксованої кількості рядків шапки: рядком даних
     вважається такий, де в колонці A стоїть ціле число (це і є
     "№ п/п", який завжди починається з 1 у кожній новій області друку
     і завершується об'єднаною коміркою A:B зі словом "Всього").
     Це працює однаково для 1 області друку чи для 20.
  4. Найменування (колонка B) звіряється зі словником нечітким пошуком
     (rapidfuzz, поріг задається через --threshold або за замовчуванням 85%).
     Значення з колонки J (за замовчуванням — константа TARGET_VALUE_COLUMN нижче)
     збирається для кожного впізнаного найменування з усіх файлів.
  5. Для кожного найменування зі словника пишеться НАЙЧАСТІШЕ значення
     серед усіх знайдених збігів (Counter.most_common).
  6. Перед записом робиться timestamped backup макету — оригінал
     ніколи не втрачається.
  7. У кінці друкується звіт: скільки файлів/таблиць/рядків оброблено,
     скільки записів словника заповнено, і окремий CSV з "незматченими"
     найменуваннями (для ручного розширення словника чи fuzzy-порогу).

ВСТАНОВЛЕННЯ:
  pip install openpyxl python-docx rapidfuzz

ЗАПУСК:
  python universal_dict_extractor.py [macket] [folder] [--threshold 85]
  (якщо аргументи не задано — інтерактивний режим)
"""

from __future__ import annotations

import csv
import io
import re
import shutil
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from rapidfuzz import fuzz, process

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None  # обробка docx стане недоступна, xlsx/xlsm працюватимуть

# Примусово UTF-8 для stdout/stderr — незалежно від кодової сторінки
# консолі Windows (cp866/cp1251) чи того, чи процес запущено з консоллю
# взагалі (як дочірній процес GUI-лаунчера, де pipe інакше падає на
# codepage за замовчуванням і кирилиця перетворюється на "      ").
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass  # старі Python без reconfigure — пропускаємо, не критично

# ---------------------------------------------------------------------------
# Налаштування — за потреби зміни ці константи
# ---------------------------------------------------------------------------

DICT_SHEET_NAME = "Словник"
NAME_COL = 2          # колонка B (найменування) — і в словнику, і в джерелах
VALUE_COL = 10         # колонка J (10-та) — те, що витягуємо. Було виявлено,
                       # що в наданому зразку це "tn" (нормативний термін
                       # служби, міс.), а не кількість (та в колонці D=4).
                       # Якщо треба інше значення — зміни цифру тут.
DEFAULT_THRESHOLD = 85
SOURCE_EXTENSIONS = {".xlsx", ".xlsm", ".docx"}

# ---------------------------------------------------------------------------
# Допоміжне
# ---------------------------------------------------------------------------

def _clean_path_input(raw: str) -> str:
    """Прибирає лапки/пробіли навколо шляху, вставленого з провідника."""
    return raw.strip().strip('"').strip("'").strip()


def _ask_path(prompt: str, must_exist: bool = True, is_dir: bool = False) -> Path:
    while True:
        raw = input(prompt).strip()
        if not raw:
            print("  Порожній шлях, спробуй ще раз.")
            continue
        p = Path(_clean_path_input(raw))
        if must_exist and not p.exists():
            print(f"  Не знайдено: {p}. Спробуй ще раз.")
            continue
        if is_dir and must_exist and not p.is_dir():
            print(f"  Це не папка: {p}. Спробуй ще раз.")
            continue
        return p


def _basic_normalize(text) -> str:
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = text.replace("’", "'").replace("ʼ", "'").replace("`", "'")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .,;")


def _is_int_like(value) -> bool:
    """№ п/п може прийти як int (xlsx) або як текст '1' (docx)."""
    if isinstance(value, int):
        return True
    if isinstance(value, float):
        return value.is_integer()
    if isinstance(value, str):
        return value.strip().isdigit()
    return False


# ---------------------------------------------------------------------------
# Читання словника
# ---------------------------------------------------------------------------

@dataclass
class DictEntry:
    row: int
    name: str
    name_norm: str
    unit: str


def load_dictionary(macket_path: Path) -> list[DictEntry]:
    wb = load_workbook(macket_path, data_only=True, keep_vba=True)
    if DICT_SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Аркуш '{DICT_SHEET_NAME}' не знайдено в {macket_path}. "
            f"Наявні аркуші: {wb.sheetnames}"
        )
    ws = wb[DICT_SHEET_NAME]
    entries = []
    for r in range(2, ws.max_row + 1):  # рядок 1 — заголовок
        name = ws.cell(row=r, column=1).value
        unit = ws.cell(row=r, column=2).value
        if not name:
            continue
        entries.append(DictEntry(
            row=r, name=str(name).strip(),
            name_norm=_basic_normalize(name),
            unit=str(unit).strip() if unit else "",
        ))
    wb.close()
    print(f"Словник завантажено: {len(entries)} записів.")
    return entries


# ---------------------------------------------------------------------------
# Пошук файлів-джерел
# ---------------------------------------------------------------------------

def find_source_files(folder: Path, macket_path: Path) -> list[Path]:
    files = []
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in SOURCE_EXTENSIONS:
            continue
        if p.name.startswith("~$"):  # тимчасові lock-файли Office
            continue
        if p.resolve() == macket_path.resolve():
            continue  # не читати сам макет як джерело даних
        files.append(p)
    return files


# ---------------------------------------------------------------------------
# Витяг рядків даних (name, value) з одного файлу
# ---------------------------------------------------------------------------

def extract_rows_xlsx(path: Path) -> list[tuple[str, object]]:
    results = []
    try:
        wb = load_workbook(path, data_only=True, keep_vba=path.suffix.lower() == ".xlsm")
    except Exception as e:
        print(f"  [!] Не вдалось відкрити {path.name}: {e}")
        return results

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for r in range(1, ws.max_row + 1):
            a_val = ws.cell(row=r, column=1).value
            if not _is_int_like(a_val):
                continue
            name = ws.cell(row=r, column=NAME_COL).value
            if not name or _is_int_like(name):
                continue  # рядок-"легенда" з номерами колонок, не дані
            value = ws.cell(row=r, column=VALUE_COL).value
            results.append((str(name).strip(), value))
    wb.close()
    return results


def extract_rows_docx(path: Path) -> list[tuple[str, object]]:
    results = []
    if DocxDocument is None:
        print("  [!] python-docx не встановлено — .docx файли пропущено "
              "(pip install python-docx)")
        return results
    try:
        # Читаємо байти файлу напряму, а не даємо python-docx тримати
        # відкритим ОС-хендл файлу до збирання сміття. python-docx не
        # надає метод .close() — при швидкому послідовному відкритті
        # десятків файлів у циклі це провокує "ValueError: I/O operation
        # on closed file" при відкладеному __del__ у zipfile (Python 3.13+).
        # BytesIO живе лише в пам'яті процесу і такої проблеми не має.
        with open(path, "rb") as f:
            file_bytes = io.BytesIO(f.read())
        doc = DocxDocument(file_bytes)
    except Exception as e:
        print(f"  [!] Не вдалось відкрити {path.name}: {e}")
        return results

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) <= max(NAME_COL - 1, VALUE_COL - 1):
                continue  # у цій таблиці замало колонок — пропускаємо рядок
            a_val = cells[0].text.strip()
            if not _is_int_like(a_val):
                continue
            name = cells[NAME_COL - 1].text.strip()
            if not name or _is_int_like(name):
                continue  # рядок-"легенда" з номерами колонок, не дані
            value = cells[VALUE_COL - 1].text.strip()
            results.append((name, value))
    return results


def extract_rows(path: Path) -> list[tuple[str, object]]:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return extract_rows_xlsx(path)
    if suffix == ".docx":
        return extract_rows_docx(path)
    return []


# ---------------------------------------------------------------------------
# Основний пайплайн
# ---------------------------------------------------------------------------

def _to_number_if_possible(value_str: str):
    """Записаний у комірку рядок має стати числом, якщо це число —
    інакше Excel зберігає його як текст (з зеленим трикутником-
    попередженням і без можливості рахувати формулами)."""
    s = value_str.strip().replace(",", ".")
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return value_str  # справді текст (напр. одиниця виміру) — лишаємо як є


@dataclass
class Stats:
    files_processed: int = 0
    files_failed: int = 0
    rows_seen: int = 0
    rows_matched: int = 0
    unmatched_names: Counter = field(default_factory=Counter)


def run(macket_path: Path, source_folder: Path, threshold: int = DEFAULT_THRESHOLD):
    dictionary = load_dictionary(macket_path)
    names_norm = [e.name_norm for e in dictionary]

    value_counters: dict[int, Counter] = defaultdict(Counter)  # dict.row -> Counter(значення)
    stats = Stats()

    source_files = find_source_files(source_folder, macket_path)
    print(f"Знайдено файлів-джерел: {len(source_files)}")
    print(f"Поріг нечіткого пошуку: {threshold}%\n")

    for i, path in enumerate(source_files, 1):
        print(f"[{i}/{len(source_files)}] {path.relative_to(source_folder)}")
        rows = extract_rows(path)
        if not rows:
            continue
        stats.files_processed += 1

        for name, value in rows:
            stats.rows_seen += 1
            query_norm = _basic_normalize(name)
            match = process.extractOne(query_norm, names_norm, scorer=fuzz.WRatio)
            if match and match[1] >= threshold:
                entry = dictionary[match[2]]
                if value is not None and str(value).strip() != "":
                    value_counters[entry.row][str(value).strip()] += 1
                    stats.rows_matched += 1
            else:
                stats.unmatched_names[name] += 1

    # --- Запис результатів у макет (з бекапом) ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = macket_path.with_name(
        f"{macket_path.stem}_backup_{timestamp}{macket_path.suffix}"
    )
    shutil.copy2(macket_path, backup_path)
    print(f"\nБекап макету створено: {backup_path.name}")

    wb = load_workbook(macket_path, keep_vba=True)
    ws = wb[DICT_SHEET_NAME]
    filled = 0
    for entry in dictionary:
        counter = value_counters.get(entry.row)
        if counter:
            best_value, freq = counter.most_common(1)[0]
            ws.cell(row=entry.row, column=3).value = _to_number_if_possible(best_value)
            filled += 1
            if len(counter) > 1:
                print(f"  [увага] '{entry.name}': кілька різних значень "
                      f"{dict(counter)} -> обрано найчастіше: {best_value}")
    wb.save(macket_path)
    wb.close()

    # --- CSV зі незматченими найменуваннями ---
    unmatched_path = macket_path.with_name(f"unmatched_{timestamp}.csv")
    with open(unmatched_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["найменування (не знайдено в словнику)", "к-сть зустрічей"])
        for name, count in stats.unmatched_names.most_common():
            writer.writerow([name, count])

    # --- Підсумок ---
    print("\n" + "=" * 60)
    print("ГОТОВО")
    print(f"Файлів оброблено:        {stats.files_processed}")
    print(f"Рядків даних переглянуто: {stats.rows_seen}")
    print(f"Рядків заматчено:        {stats.rows_matched}")
    print(f"Записів словника заповнено: {filled} з {len(dictionary)}")
    print(f"Унікальних незматчених найменувань: {len(stats.unmatched_names)}")
    print(f"  -> детальний список: {unmatched_path.name}")
    print(f"Результат записано у: {macket_path.name}")
    print(f"Бекап оригіналу:      {backup_path.name}")
    print("=" * 60)


# ---------------------------------------------------------------------------
# Точка входу
# ---------------------------------------------------------------------------

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Універсальний словниковий екстрактор — витягує значення "
                     "з однотипних xlsx/xlsm/docx файлів і записує їх у "
                     "аркуш 'Словник' макету."
    )
    parser.add_argument("macket", nargs="?", help="Шлях до макету (xlsm з аркушем 'Словник')")
    parser.add_argument("folder", nargs="?", help="Папка для рекурсивного пошуку файлів-джерел")
    parser.add_argument("--threshold", type=int, default=DEFAULT_THRESHOLD,
                        help=f"Поріг нечіткого пошуку у відсотках (за замовчуванням {DEFAULT_THRESHOLD})")
    args = parser.parse_args()

    print("=== Універсальний словниковий екстрактор ===\n")

    if args.macket and args.folder:
        # Запущено з аргументами (напр. з GUI-лаунчера) — валідуємо без input()
        macket = Path(_clean_path_input(args.macket))
        folder = Path(_clean_path_input(args.folder))
        if not macket.is_file():
            print(f"ПОМИЛКА: макет не знайдено: {macket}")
            sys.exit(1)
        if not folder.is_dir():
            print(f"ПОМИЛКА: папку не знайдено: {folder}")
            sys.exit(1)
    else:
        # Інтерактивний режим — питаємо шляхи в консолі
        macket = _ask_path(
            "Шлях до макету (xlsm з аркушем 'Словник'): ",
            must_exist=True, is_dir=False,
        )
        folder = _ask_path(
            "Папка для пошуку файлів-джерел (шукає рекурсивно, всі підпапки): ",
            must_exist=True, is_dir=True,
        )

    print()
    run(macket, folder, threshold=args.threshold)

    if not (args.macket and args.folder):
        input("\nНатисни Enter, щоб закрити...")


if __name__ == "__main__":
    main()