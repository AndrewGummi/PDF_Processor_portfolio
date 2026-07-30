#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
launcher.py -- інтуітивний запускатор для всіх скриптів обліку майна.

Лежить разом з рештою файлів у корені проєкту.
Запуск (Git Bash):
    python launcher.py

Що вміє (усе в одному вікні, без командного рядка):
    1. "Імпортувати залишки" -- обрати файл Word (.docx, один
       склад/підрозділ) або Excel (.xlsm/.xlsx, усі підрозділи одразу,
       формат "Залишки_На_ХХ_20ХХ.xlsm") -> дані йдуть у майстер-файл.
    2. "Обробити фото карток" -- обрати папку (шукає фото РЕКУРСИВНО,
       у всіх підпапках -- для 300+ фото це і треба), кожне фото йде
       через vision-модель (OpenRouter) у "Журнал руху".
    3. "Оновити зведення і звірку" -- перебудувати "Залишки — Загалом"
       і "Звірка" вручну (якщо самі щось поправили в майстер-файлі).

Ніякої власної логіки тут немає -- launcher лише викликає функції з
cards_photo_to_excel.py / import_pidrozdily_xlsm.py /
import_balances_docx.py / oblik_common.py, які й так можна запускати
окремо з командного рядка.
"""

import os
import queue
import sys
import threading
import traceback
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    print("ERROR: tkinter відсутній у цьому Python. На Windows він зазвичай "
          "вбудований у стандартний інсталятор python.org -- перевстановіть "
          "Python звідти (галочка 'tcl/tk and IDLE' при встановленні).")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import oblik_common as oc  # noqa: E402
import cards_photo_to_excel as cpe  # noqa: E402
import import_pidrozdily_xlsm as imp_xlsm  # noqa: E402
import import_balances_docx as imp_docx  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402


class App:
    def __init__(self, root):
        self.root = root
        root.title("Облік майна бригади -- запускатор")
        root.geometry("760x560")

        self.log_queue = queue.Queue()
        self.master_path = tk.StringVar(value=str(SCRIPT_DIR / "Облік_майна_бригади.xlsx"))
        self.photo_folder = tk.StringVar(value=str(SCRIPT_DIR))
        self.provider = tk.StringVar(value=cpe.PROVIDER_ANTHROPIC)
        self.api_key = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.env_file = tk.StringVar(value="")

        self._build_ui()
        self.root.after(200, self._drain_log_queue)

    # ------------------------------------------------------------
    # UI
    # ------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        frm_master = ttk.LabelFrame(self.root, text="Майстер-файл обліку")
        frm_master.pack(fill="x", **pad)
        ttk.Entry(frm_master, textvariable=self.master_path, width=70).pack(
            side="left", fill="x", expand=True, padx=6, pady=6)
        ttk.Button(frm_master, text="Обрати...", command=self._pick_master).pack(
            side="left", padx=6)

        frm_import = ttk.LabelFrame(self.root, text="1. Імпортувати залишки (Word або Excel)")
        frm_import.pack(fill="x", **pad)
        ttk.Button(frm_import, text="Обрати файл залишків...",
                   command=self._pick_balances_file).pack(side="left", padx=6, pady=6)
        ttk.Label(frm_import,
                  text=".docx -- один склад/підрозділ  |  .xlsm/.xlsx -- усі підрозділи одразу"
                  ).pack(side="left", padx=6)

        frm_photos = ttk.LabelFrame(self.root, text="2. Обробити фото карток майна (з підпапками)")
        frm_photos.pack(fill="x", **pad)
        row1 = ttk.Frame(frm_photos)
        row1.pack(fill="x", padx=6, pady=4)
        ttk.Entry(row1, textvariable=self.photo_folder, width=55).pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="Обрати папку...", command=self._pick_photo_folder).pack(side="left", padx=6)

        row2 = ttk.Frame(frm_photos)
        row2.pack(fill="x", padx=6, pady=4)
        ttk.Label(row2, text="Постачальник:").pack(side="left")
        ttk.Combobox(row2, textvariable=self.provider, state="readonly", width=12,
                     values=[cpe.PROVIDER_ANTHROPIC, cpe.PROVIDER_OPENROUTER]).pack(side="left", padx=6)
        ttk.Button(row2, text="Завантажити ключі з .env файлу...",
                   command=self._pick_env_file).pack(side="left", padx=10)

        row3 = ttk.Frame(frm_photos)
        row3.pack(fill="x", padx=6, pady=4)
        ttk.Label(row3, text="API-ключ (або через .env вище):").pack(side="left")
        ttk.Entry(row3, textvariable=self.api_key, show="*", width=40).pack(side="left", padx=6)
        ttk.Button(row3, text="Обробити фото", command=self._run_photos).pack(side="left", padx=10)

        frm_rollup = ttk.LabelFrame(self.root, text="3. Обслуговування")
        frm_rollup.pack(fill="x", **pad)
        ttk.Button(frm_rollup, text="Оновити 'Залишки — Загалом' і 'Звірка'",
                   command=self._run_rollup).pack(side="left", padx=6, pady=6)

        frm_log = ttk.LabelFrame(self.root, text="Журнал виконання")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt_log = tk.Text(frm_log, height=16, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    # ------------------------------------------------------------
    # Логування в текстове поле (потокобезпечно через чергу)
    # ------------------------------------------------------------
    def log(self, msg):
        self.log_queue.put(str(msg))

    def _drain_log_queue(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.root.after(200, self._drain_log_queue)

    # ------------------------------------------------------------
    # Пікери файлів/папок
    # ------------------------------------------------------------
    def _pick_master(self):
        path = filedialog.asksaveasfilename(
            title="Майстер-файл обліку", initialdir=str(SCRIPT_DIR),
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.master_path.set(path)

    def _pick_photo_folder(self):
        path = filedialog.askdirectory(title="Папка з фото карток (шукає й у підпапках)",
                                        initialdir=self.photo_folder.get())
        if path:
            self.photo_folder.set(path)

    def _pick_env_file(self):
        path = filedialog.askopenfilename(
            title="Файл з ключами (.env)", initialdir=str(SCRIPT_DIR),
            filetypes=[(".env / текстовий", "*.env *.txt"), ("Усі файли", "*.*")])
        if not path:
            return
        try:
            loaded = oc.load_env_file(path)
            self.env_file.set(path)
            # Значення ключів НІКОЛИ не потрапляють у журнал -- лише назви змінних
            self.log(f"З .env завантажено: {', '.join(loaded) if loaded else '(нічого)'}")
            key_name = ("ANTHROPIC_API_KEY" if self.provider.get() == cpe.PROVIDER_ANTHROPIC
                        else "OPENROUTER_API_KEY")
            if key_name in loaded:
                self.api_key.set(os.environ[key_name])
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалось прочитати .env: {e}")

    def _pick_balances_file(self):
        path = filedialog.askopenfilename(
            title="Файл залишків",
            initialdir=str(SCRIPT_DIR),
            filetypes=[("Word / Excel", "*.docx *.xlsm *.xlsx"),
                       ("Word", "*.docx"), ("Excel", "*.xlsm *.xlsx")])
        if not path:
            return
        threading.Thread(target=self._import_balances, args=(path,), daemon=True).start()

    # ------------------------------------------------------------
    # Дії (кожна -- в окремому потоці, щоб вікно не зависало)
    # ------------------------------------------------------------
    def _import_balances(self, path):
        try:
            master = self.master_path.get()
            ext = Path(path).suffix.lower()
            self.log(f"Імпорт: {path}")

            if ext in (".xlsm", ".xlsx"):
                import warnings
                warnings.filterwarnings("ignore")
                wb_source = load_workbook(path, data_only=True)
                year = imp_xlsm.guess_year(path)
                records = oc.parse_subdivisions_xlsm(wb_source, year)
                self.log(f"Знайдено {len(records)} рядків майна, "
                         f"{len(wb_source.sheetnames)} підрозділів.")

                wb = load_workbook(master) if Path(master).exists() else self._new_wb()
                imp_xlsm.replace_pidrozdily_sheet(wb, records, Path(path).name)
                oc.ensure_responsible_sheet(wb)
                n = oc.rebuild_rollup(wb)
                wb.save(master)
                self.log(f"Готово. Унікальних найменувань у зведенні: {n}.")

            elif ext == ".docx":
                loc_type = self._ask_sklad_or_pidrozdil()
                if loc_type is None:
                    self.log("Скасовано.")
                    return
                import datetime
                location_name, records = oc.parse_vidomist_docx(
                    path, datetime.date.today().year)
                self.log(f"Локація: {location_name}. Знайдено позицій: {len(records)}.")
                cfg = imp_docx.LOCATION_CONFIG[loc_type]

                wb = load_workbook(master) if Path(master).exists() else self._new_wb()
                ws = imp_docx.ensure_sheet(wb, cfg)
                new_rows = [(item, unit, qty, rdate, f"з {Path(path).name}")
                            for item, unit, qty, rdate in records]
                n_rows = oc.upsert_location_rows(ws, cfg["table"], cfg["headers"],
                                                  cfg["location_col"], location_name, new_rows)
                oc.stamp_responsible_lookup(ws, cfg["table"], cfg["headers"], cfg["location_col"])
                oc.ensure_responsible_sheet(wb)
                n = oc.rebuild_rollup(wb)
                wb.save(master)
                self.log(f"Готово. Рядків у '{cfg['sheet']}': {n_rows}. "
                         f"Найменувань у зведенні: {n}.")
            else:
                self.log(f"Невідомий тип файлу: {ext}")
                return

            messagebox.showinfo("Готово", "Імпорт завершено.")
        except Exception:
            self.log("ПОМИЛКА:\n" + traceback.format_exc())
            messagebox.showerror("Помилка", "Щось пішло не так -- деталі в журналі виконання.")

    def _ask_sklad_or_pidrozdil(self):
        win = tk.Toplevel(self.root)
        win.title("Це склад чи підрозділ?")
        win.grab_set()
        result = {"value": None}

        def choose(v):
            result["value"] = v
            win.destroy()

        ttk.Label(win, text="Цей docx-файл -- склад чи підрозділ?").pack(padx=20, pady=10)
        row = ttk.Frame(win)
        row.pack(pady=10)
        ttk.Button(row, text="Склад", command=lambda: choose("sklad")).pack(side="left", padx=10)
        ttk.Button(row, text="Підрозділ", command=lambda: choose("pidrozdil")).pack(side="left", padx=10)
        self.root.wait_window(win)
        return result["value"]

    def _run_photos(self):
        folder = self.photo_folder.get()
        provider = self.provider.get()
        api_key = self.api_key.get().strip()
        if not api_key:
            messagebox.showwarning("Немає ключа", "Вкажіть API-ключ або завантажте його з .env файлу.")
            return
        if not Path(folder).exists():
            messagebox.showwarning("Немає папки", "Оберіть існуючу папку з фото.")
            return
        threading.Thread(target=self._process_photos, args=(folder, api_key, provider), daemon=True).start()

    def _process_photos(self, folder, api_key, provider):
        try:
            master = self.master_path.get()
            images = oc.find_images_recursive(folder)
            self.log(f"Знайдено фото (з підпапками): {len(images)}")
            if not images:
                self.log("Фото не знайдено.")
                return

            model = (cpe.ANTHROPIC_DEFAULT_MODEL if provider == cpe.PROVIDER_ANTHROPIC
                     else cpe.OPENROUTER_DEFAULT_MODEL)
            wb, ws = cpe.open_or_create_workbook(Path(master))
            total, failed = 0, []
            for i, img_path in enumerate(images, 1):
                name = Path(img_path).name
                self.log(f"[{i}/{len(images)}] {name}")
                card = cpe.extract_card(img_path, api_key, model, provider=provider)
                added, err = cpe.append_rows(ws, card, name)
                total += added
                if err:
                    failed.append((name, err))
                if i % 15 == 0:
                    cpe.register_journal_table(ws)
                    wb.save(master)

            cpe.autosize_columns(ws)
            cpe.register_journal_table(ws)
            oc.rebuild_rollup(wb)
            wb.save(master)

            self.log(f"Готово. Оброблено фото: {len(images)}, додано рядків: {total}.")
            if failed:
                self.log(f"НЕ вдалось обробити: {len(failed)} -- див. лог cards_photo_to_excel.log")
            messagebox.showinfo("Готово", f"Оброблено {len(images)} фото, додано {total} рядків.")
        except Exception:
            self.log("ПОМИЛКА:\n" + traceback.format_exc())
            messagebox.showerror("Помилка", "Щось пішло не так -- деталі в журналі виконання.")

    def _run_rollup(self):
        master = self.master_path.get()
        if not Path(master).exists():
            messagebox.showwarning("Нема файлу", "Спочатку створіть/оберіть майстер-файл.")
            return
        try:
            wb = load_workbook(master)
            n = oc.rebuild_rollup(wb)
            wb.save(master)
            self.log(f"Зведення й звірку перебудовано. Найменувань: {n}.")
            messagebox.showinfo("Готово", "Зведення й звірку оновлено.")
        except Exception:
            self.log("ПОМИЛКА:\n" + traceback.format_exc())
            messagebox.showerror("Помилка", "Щось пішло не так -- деталі в журналі виконання.")

    @staticmethod
    def _new_wb():
        wb = Workbook()
        wb.remove(wb.active)
        return wb


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
