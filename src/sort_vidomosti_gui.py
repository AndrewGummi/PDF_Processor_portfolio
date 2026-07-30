# -*- coding: utf-8 -*-
"""
Окреме віконце: сортує файли за структурою "відомості".

Береш папку-джерело (там може бути купа різного мотлоху — .xlsm, .docx,
.jpg, що завгодно) і папку-призначення. Скрипт рекурсивно проходить
джерело, і для кожного .xlsm/.xlsx файлу перевіряє ТИМ САМИМ кодом, що і
vidomist_report.py (функція extract_vidomosti) — чи є в файлі хоч один
блок "ВІДОМІСТЬ №...". Якщо є — файл відповідає структурі і копіюється
(оригінал НЕ видаляється і НЕ переміщується, лишається на місці) у
папку-призначення. Все інше (інші розширення, або .xlsm/.xlsx без жодної
ВІДОМІСТЬ) — пропускається, з причиною в лозі.

Запуск (Git Bash):
    python3 sort_vidomosti_gui.py

Або через меню в run_vidomist.py (варіант "2" при старті) — тоді
відкривається як окремий процес/вікно, а консольний запускатор лишається
вільним.

Файл має лежати поруч з vidomist_report.py (в тій самій папці).
"""

import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext

from vidomist_report import extract_vidomosti

SKIP_PREFIXES = ("~$",)  # тимчасові Excel-файли типу ~$файл.xlsm


def _is_own_output(name):
    """Власні звіти цього ж інструментарію — щоб не затягнути в сортування
    вже згенеровані звіти (вони теж .xlsx, і в них теж могли лишитись слова
    "Відомість" в заголовках, тож без цього фільтра могли б хибно
    "підійти під структуру")."""
    low = name.lower()
    return (
        low.endswith("_звіт.xlsx")
        or low == "звіт_загальний.xlsx"
        or low == "звіт_відомостей.xlsx"
        or low == "звіт_пропущені_файли.txt"
    )


def find_all_files(root_folder):
    """Всі файли в дереві (будь-яке розширення) — для повної картини в
    лозі, скільки взагалі 'сміття' було в папці."""
    found = []
    for dirpath, _dirs, filenames in os.walk(root_folder):
        for name in filenames:
            if name.startswith(SKIP_PREFIXES):
                continue
            found.append(os.path.join(dirpath, name))
    return sorted(found)


def unique_destination(dest_folder, filename):
    """Якщо файл з такою назвою вже є у призначенні — додає (2), (3)...
    щоб копіювання ніколи мовчки нічого не перезаписало."""
    base, ext = os.path.splitext(filename)
    candidate = filename
    i = 2
    while os.path.exists(os.path.join(dest_folder, candidate)):
        candidate = f"{base} ({i}){ext}"
        i += 1
    return os.path.join(dest_folder, candidate)


class App:
    def __init__(self, root):
        self.root = root
        root.title("Сортування файлів за структурою відомості")
        root.geometry("700x520")

        self.src_var = tk.StringVar()
        self.dst_var = tk.StringVar()

        tk.Label(root, text="Папка-джерело (де шукати, включно з підпапками):").pack(
            anchor="w", padx=10, pady=(10, 0)
        )
        row1 = tk.Frame(root)
        row1.pack(fill="x", padx=10)
        tk.Entry(row1, textvariable=self.src_var).pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="Обрати...", command=self.pick_src).pack(side="left", padx=(6, 0))

        tk.Label(root, text="Папка-призначення (куди копіювати підхожі файли):").pack(
            anchor="w", padx=10, pady=(10, 0)
        )
        row2 = tk.Frame(root)
        row2.pack(fill="x", padx=10)
        tk.Entry(row2, textvariable=self.dst_var).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="Обрати...", command=self.pick_dst).pack(side="left", padx=(6, 0))

        self.start_btn = tk.Button(
            root, text="Почати", command=self.start, bg="#4CAF50", fg="white", width=20
        )
        self.start_btn.pack(pady=12)

        tk.Label(root, text="(копіює, оригінали в джерелі не чіпає)", fg="#666666").pack()

        self.log = scrolledtext.ScrolledText(root, height=20)
        self.log.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    def pick_src(self):
        folder = filedialog.askdirectory(title="Оберіть папку-джерело")
        if folder:
            self.src_var.set(folder)

    def pick_dst(self):
        folder = filedialog.askdirectory(title="Оберіть папку-призначення")
        if folder:
            self.dst_var.set(folder)

    def write_log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.update_idletasks()

    def start(self):
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()
        if not src or not os.path.isdir(src):
            self.write_log("Помилка: вкажи коректну папку-джерело.")
            return
        if not dst:
            self.write_log("Помилка: вкажи папку-призначення.")
            return
        if os.path.abspath(src) == os.path.abspath(dst):
            self.write_log("Помилка: джерело і призначення — та сама папка.")
            return
        os.makedirs(dst, exist_ok=True)

        self.start_btn.config(state="disabled")
        self.log.delete("1.0", "end")
        threading.Thread(target=self.run_sort, args=(src, dst), daemon=True).start()

    def run_sort(self, src, dst):
        all_files = find_all_files(src)
        candidates = [
            p for p in all_files
            if p.lower().endswith((".xlsm", ".xlsx")) and not _is_own_output(os.path.basename(p))
        ]
        other_count = len(all_files) - len(candidates)

        self.write_log(
            f"У '{src}' знайдено {len(all_files)} файл(ів) усього, з них "
            f"{len(candidates)} .xlsm/.xlsx для перевірки структури "
            f"(решта {other_count} — інші розширення, пропущено без відкриття).\n"
        )

        copied, skipped, errors = [], [], []

        for path in candidates:
            rel = os.path.relpath(path, src)
            try:
                vidomosti = extract_vidomosti(path)
            except Exception as exc:
                self.write_log(f"[ПОМИЛКА ЧИТАННЯ] {rel}: {exc}")
                errors.append((rel, str(exc)))
                continue

            if vidomosti:
                dest_path = unique_destination(dst, os.path.basename(path))
                shutil.copy2(path, dest_path)
                self.write_log(
                    f"[ПІДХОДИТЬ] {rel}  ->  '{os.path.basename(dest_path)}'  "
                    f"({len(vidomosti)} відомост(і/ей) знайдено)"
                )
                copied.append(rel)
            else:
                self.write_log(f"[не та структура] {rel} — пропущено")
                skipped.append(rel)

        self.write_log(
            f"\nГотово. Скопійовано: {len(copied)}. "
            f"Пропущено (немає блоку ВІДОМІСТЬ): {len(skipped)}. "
            f"Помилок читання: {len(errors)}."
        )

        log_path = os.path.join(dst, "лог_сортування.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"Джерело: {src}\nПризначення: {dst}\n\n")
                f.write(f"Скопійовано, відповідають структурі ({len(copied)}):\n")
                for r in copied:
                    f.write(f"  {r}\n")
                f.write(f"\nПропущено, немає блоку ВІДОМІСТЬ ({len(skipped)}):\n")
                for r in skipped:
                    f.write(f"  {r}\n")
                if errors:
                    f.write(f"\nПомилки читання ({len(errors)}):\n")
                    for r, err in errors:
                        f.write(f"  {r}: {err}\n")
                f.write(f"\nІнші розширення (не перевірялись) — {other_count} шт.\n")
            self.write_log(f"Лог збережено: {log_path}")
        except Exception as exc:
            self.write_log(f"Не вдалося зберегти лог: {exc}")

        self.start_btn.config(state="normal")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
