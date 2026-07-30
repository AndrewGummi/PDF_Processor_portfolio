# -*- coding: utf-8 -*-
"""
Запускатор для vidomist_report.py.

Рекурсивно шукає в обраній папці (і всіх підпапках) файли відомостей
(.xlsm/.xlsx) та для кожного намагається знайти відповідний файл-витяг
(<та сама назва>_витяг.doc/.docx у тій самій папці, або будь-де в дереві).
Для кожної пари друкує зведення (як у vidomist_report.py) і в кінці
зберігає один загальний звіт у обраній папці.

Запуск (Git Bash):
    python3 run_vidomist.py                       # відкриє діалог вибору папки
    python3 run_vidomist.py "шлях/до/папки"        # одразу обробка, без діалогу

Файли мають лежати поруч: run_vidomist.py та vidomist_report.py в одній папці.

(Сортування файлів за структурою відомості — окрема задача, окремий
скрипт: sort_vidomosti_gui.py. Запускається сам по собі:
    python3 sort_vidomosti_gui.py
Має лежати поруч з vidomist_report.py.)
"""

import os
import sys

from vidomist_report import (
    WordSession,
    build_report,
    extract_vidomosti,
    rows_from_vidomosti,
    upsert_xlsx_table,
)

SKIP_PREFIXES = ("~$",)  # тимчасові Excel-файли типу ~$файл.xlsm


def _is_own_output(name):
    low = name.lower()
    return low.endswith("_звіт.xlsx") or low == "звіт_загальний.xlsx"


def choose_root_folder():
    if len(sys.argv) > 1:
        return sys.argv[1]

    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        folder = filedialog.askdirectory(
            title="Оберіть папку з відомостями (шукатиму також у підпапках)"
        )
        root.destroy()
        if folder:
            return folder
    except Exception:
        pass

    return input("Введи шлях до папки з відомостями: ").strip()


def find_xlsm_files(root_folder):
    found = []
    for dirpath, _dirs, filenames in os.walk(root_folder):
        for name in filenames:
            if name.startswith(SKIP_PREFIXES) or _is_own_output(name):
                continue
            if name.lower().endswith((".xlsm", ".xlsx")):
                found.append(os.path.join(dirpath, name))
    return sorted(found)


def find_matching_doc(xlsm_path, root_folder):
    """Шукає файл-витяг для даної відомості.

    Порядок пошуку:
      1) точна назва "<стем>_витяг.doc/.docx" у тій самій папці;
      2) будь-який .doc/.docx у тій самій папці зі словом "витяг" у назві;
      3) точна назва "<стем>_витяг.*" будь-де в дереві папок.
    """
    stem = os.path.splitext(os.path.basename(xlsm_path))[0]
    same_dir = os.path.dirname(xlsm_path)
    exact_names = {f"{stem}_витяг.doc".lower(), f"{stem}_витяг.docx".lower()}

    try:
        siblings = os.listdir(same_dir)
    except OSError:
        siblings = []

    for name in siblings:
        if name.lower() in exact_names:
            return os.path.join(same_dir, name)

    for name in siblings:
        low = name.lower()
        if low.endswith((".doc", ".docx")) and "витяг" in low:
            return os.path.join(same_dir, name)

    for dirpath, _dirs, filenames in os.walk(root_folder):
        for name in filenames:
            if name.lower() in exact_names:
                return os.path.join(dirpath, name)

    return None


def dedupe_rows(rows):
    """Прибирає повні дублікати — одна й та сама ВІДОМІСТЬ (той самий
    № + дата + сума) знайдена в кількох файлах чи копіях папок.

    Лишає перший знайдений рядок на ключ.

    Повертає (унікальні_рядки, лог_дублікатів) — лог потрібен лише
    щоб надрукувати користувачу, що саме прибрали."""
    best = {}
    duplicates_log = []
    for r in rows:
        sum_val = r.get("sum")
        sum_key = round(sum_val, 2) if isinstance(sum_val, (int, float)) else sum_val
        key = (r.get("number"), r.get("date"), sum_key)

        existing = best.get(key)
        if existing is None:
            best[key] = r
            continue

        duplicates_log.append((key, existing.get("file"), r.get("file")))

    return list(best.values()), duplicates_log


def _run():
    root_folder = choose_root_folder()
    if not root_folder or not os.path.isdir(root_folder):
        print("Папку не знайдено або не обрано. Вихід.")
        return

    xlsm_files = find_xlsm_files(root_folder)
    if not xlsm_files:
        print(f"У папці {root_folder} (і підпапках) не знайдено жодного .xlsm/.xlsx файлу.")
        return

    print(f"Знайдено {len(xlsm_files)} файл(ів) відомостей. Обробка...\n")

    all_rows = []
    failed_files = []
    out_path = os.path.join(root_folder, "звіт_загальний.xlsx")

    try:
        with WordSession() as session:
            for xlsm_path in xlsm_files:
                rel = os.path.relpath(xlsm_path, root_folder)
                print(f"— {rel}")

                try:
                    vidomosti = extract_vidomosti(xlsm_path)
                except Exception as exc:
                    print(f"   ПОМИЛКА при читанні файлу: {exc} — пропускаю цей файл\n")
                    failed_files.append((rel, str(exc)))
                    continue

                if not vidomosti:
                    print("   (у файлі не знайдено жодної ВІДОМІСТЬ — пропускаю)\n")
                    continue

                doc_path = find_matching_doc(xlsm_path, root_folder)
                if doc_path:
                    nakaz_number, nakaz_date = session.extract_nakaz(doc_path)
                else:
                    print("   УВАГА: не знайдено відповідний файл-витяг (_витяг.doc)")
                    nakaz_number, nakaz_date = None, None

                report = build_report(vidomosti, nakaz_number, nakaz_date)
                print(report)
                print()
                all_rows.extend(rows_from_vidomosti(
                    vidomosti, nakaz_number, nakaz_date,
                    source_file=rel, source_path=os.path.abspath(xlsm_path),
                ))
    finally:
        # Зберігаємо все, що встигли зібрати, НАВІТЬ якщо стався збій —
        # один битий файл більше не має "з'їдати" роботу по решті файлів.
        if all_rows:
            unique_rows, dupes = dedupe_rows(all_rows)
            if dupes:
                print(f"\nПрибрано {len(dupes)} повтор(ів) (той самий № + дата + сума):")
                for (number, date, _sum), file_kept, file_dupe in dupes:
                    print(f"   №{number} від {date} — лишив '{file_kept}', прибрав дубль '{file_dupe}'")
            save_xlsx_table_added, save_xlsx_table_total, suspicious_drops, removed_dupes_log, healed_log = upsert_xlsx_table(
                unique_rows, out_path, include_file_column=True
            )
            print(f"Збережено: {out_path}")
            print(
                f"Додано нових рядків: {save_xlsx_table_added} "
                f"(усього в звіті: {save_xlsx_table_total})"
            )
            if healed_log:
                print(
                    f"\nАвтоматично підправлено {len(healed_log)} рядок(ів) за місцем у джерелі "
                    f"(стара дата -> нова, без ручного редагування — Наказ/Статус збережено):"
                )
                for number, old_date, new_date, fname in healed_log:
                    print(f"   №{number} у файлі '{fname}': {old_date} -> {new_date}")
            if removed_dupes_log:
                print(
                    f"\nПрибрано {len(removed_dupes_log)} накопичений(их) дублікат(ів), "
                    f"що лишались у самому звіті (зі старих запусків, до цього виправлення):"
                )
                for number, date, summ, fname in removed_dupes_log:
                    print(f"   №{number} від {date} у файлі '{fname}'")
                dupes_path = os.path.join(root_folder, "звіт_прибрані_дублі.txt")
                try:
                    with open(dupes_path, "w", encoding="utf-8") as f:
                        for number, date, summ, fname in removed_dupes_log:
                            f.write(f"№{number}\t{date}\t{summ}\t{fname}\n")
                    print(f"Список прибраних дублів: {dupes_path}")
                except Exception as exc:
                    print(f"Не вдалося зберегти лог прибраних дублів: {exc}")
            if suspicious_drops:
                print(f"\nУВАГА: {len(suspicious_drops)} запис(ів) мали підозріло занижену "
                      f"нову суму при повторному скануванні — стару суму залишено, "
                      f"нову ІГНОРОВАНО. Ось звідки саме (звірте ці файли вручну):")
                for number, old_sum, new_sum, fname in suspicious_drops:
                    print(f"   №{number} у файлі '{fname}': було {old_sum:.2f}, "
                          f"нове зчитування дало {new_sum:.2f}")
        if failed_files:
            fail_path = os.path.join(root_folder, "звіт_пропущені_файли.txt")
            with open(fail_path, "w", encoding="utf-8") as f:
                for rel, err in failed_files:
                    f.write(f"{rel}\t{err}\n")
            print(f"\nФайлів пропущено через помилку читання: {len(failed_files)}")
            print(f"Список і причини: {fail_path}")

    if not all_rows:
        print("Нічого не оброблено.")
        return root_folder

    print(f"\nГотово. Загальна таблиця: {out_path}")
    return root_folder


def main():
    root_folder = None
    try:
        root_folder = _run()
    except Exception:
        import traceback
        err_text = traceback.format_exc()
        print("\nПОМИЛКА:\n" + err_text, file=sys.stderr)
        fallback_dir = root_folder if root_folder and os.path.isdir(root_folder) else "."
        err_path = os.path.join(fallback_dir, "звіт_помилка.txt")
        try:
            with open(err_path, "w", encoding="utf-8") as f:
                f.write(err_text)
            print(f"Деталі помилки збережено: {err_path}")
        except Exception:
            pass
    finally:
        try:
            input("\nНатисни Enter, щоб закрити...")
        except (EOFError, OSError):
            pass


if __name__ == "__main__":
    main()
