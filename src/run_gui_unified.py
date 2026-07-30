#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_gui_unified.py — єдиний GUI-лаунчер для всіх інструментів обробки
військової адміністративної документації.

Об'єднує в одному вікні: меню інструментів зліва, кожен інструмент на власній сторінці.
  1. "Втрати майна"        -> doc_processor.py
  2. "Відомість (Excel)"   -> excel_report_generator.py (тепер включає й
                              підбір цін -- окремим фінальним кроком, ПІСЛЯ
                              усіх циклів розпізнавання найменувань і
                              зіставлення синонімів; колишню окрему
                              вкладку "Підбір цін" (price_matcher.py)
                              прибрано -- її логіку повністю перенесено
                              всередину excel_report_generator.py)
  3. "OCR PDF"              -> pdf_to_text_multithread.py
  4. "Excel -> Word"        -> excel_to_word_transfer.py (прямий імпорт)
  5. "Словн. екстрактор"    -> universal_dict_extractor.py

Принцип лишається той самий, що й у окремих run_*.py: ЖОДНОГО cmd.exe,
powershell.exe чи bash.exe. Дочірні скрипти запускаються НАПРЯМУ як
python.exe-процеси через subprocess.Popen(sys.executable, ...), без
shell=True. Єдиний виняток — excel_to_word_transfer.py: він, як і в
оригінальному run_gui_transfer.py, викликається прямим імпортом функції
process_workbook() у фоновому потоці (це не окремий процес, а функція
всередині цього ж python.exe).

Кожна вкладка живе у власному потоці (threading.Thread), тому кілька
інструментів технічно можна запускати "одночасно" (хоча звичайне
використання — по одному). Логи кожної вкладки незалежні.

Запуск:
    python.exe run_gui_unified.py
Поклади цей файл в ту саму папку, що й усі скрипти-оброблювачі
(doc_processor.py, excel_report_generator.py, pdf_to_text_multithread.py,
excel_to_word_transfer.py, universal_dict_extractor.py).
"""

import os
import sys
import gc
import json
import subprocess
import threading
import datetime
import tempfile
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from functools import partial

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ──────────────────────────────────────────────────────────────────────
# Спільні шляхи
# ──────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent

DOC_PROCESSOR = SCRIPT_DIR / "doc_processor.py"
EXCEL_GENERATOR_SCRIPT = SCRIPT_DIR / "excel_report_generator.py"
OCR_SCRIPT = SCRIPT_DIR / "pdf_to_text_multithread.py"
EXTRACTOR_SCRIPT = SCRIPT_DIR / "universal_dict_extractor.py"
POPPLER_BIN = SCRIPT_DIR / "poppler" / "poppler-24.08.0" / "Library" / "bin"

OUTPUT_DOCS_DIR = SCRIPT_DIR / "output_docs"
LOGS_DIR = SCRIPT_DIR / "Logs"

# Маркери протоколу "ручний вибір варіанта" — мають ЗБІГАТИСЯ дослівно з
# _MANUAL_MARK_BEGIN / _MANUAL_MARK_END у doc_processor.py. Між ними
# doc_processor.py друкує пронумерований список кандидатів і блокується на
# input(), чекаючи цифру зі stdin.
MANUAL_MARK_BEGIN = "##MANUAL_MATCH_BEGIN##"
MANUAL_MARK_END = "##MANUAL_MATCH_END##"

# Однорядковий маркер — doc_processor.py друкує його одразу після того, як
# оператор (або авто-таймаут спроб) вирішив "не змінювати" сиру назву.
# Побачивши цей рядок, GUI підсвічує відповідний запис у лозі блакитним,
# замість того щоб показувати сирий службовий рядок як є.
MANUAL_SKIP_PREFIX = "##MANUAL_SKIP##"


# ──────────────────────────────────────────────────────────────────────
# Спільні утиліти
# ──────────────────────────────────────────────────────────────────────
def utf8_env(extra=None):
    """PYTHONUTF8=1 — інакше на Windows дочірній python.exe пише stdout у
    кодовій сторінці консолі (cp866/cp1251), а ми читаємо як UTF-8 і
    кирилиця перетворюється на сміття."""
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    if extra:
        env.update(extra)
    return env


def stream_subprocess(cmd, cwd, env, log_callback, proc_holder=None, manual_prompt_handler=None):
    """Запускає cmd як прямий дочірній процес (без shell=True, без
    cmd/powershell/bash) і стрімить рядки stdout у log_callback.
    Повертає код завершення (-1 при помилці запуску, None якщо зупинено
    користувачем через kill_process_tree).

    proc_holder — опційний dict; якщо переданий, туди одразу після старту
    кладеться сам об'єкт Popen (proc_holder["proc"] = proc), щоб виклик
    з іншого потоку (натискання кнопки "Зупинити") міг дістати PID і
    примусово завершити процес разом з усіма його нащадками.

    manual_prompt_handler — опційна функція(block_lines: list[str]) -> str,
    що викликається щоразу, коли в stdout дочірнього процесу зустрічається
    блок між MANUAL_MARK_BEGIN/END (протокол "ручний вибір варіанта" з
    doc_processor.py). Має повернути обраний номер варіанта (рядок), який
    буде записано в stdin процесу. Якщо переданий — Popen відкриває stdin
    на запис (stdin=PIPE); якщо ні — stdin=DEVNULL (без змін поведінки)."""
    proc = None
    try:
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE if manual_prompt_handler else subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            cwd=str(cwd),
            env=env,
        )
        if proc_holder is not None:
            proc_holder["proc"] = proc

        collecting = False
        block_lines = []
        for line in proc.stdout:
            stripped = line.rstrip("\n").rstrip("\r")
            if manual_prompt_handler and stripped == MANUAL_MARK_BEGIN:
                collecting = True
                block_lines = []
                continue
            if collecting:
                if stripped == MANUAL_MARK_END:
                    collecting = False
                    try:
                        answer = manual_prompt_handler(block_lines)
                    except Exception as e:
                        log_callback(f"\nПОМИЛКА діалогу ручного вибору: {e}\n")
                        answer = "0"
                    try:
                        proc.stdin.write(f"{answer}\n")
                        proc.stdin.flush()
                    except Exception:
                        pass
                    continue
                block_lines.append(stripped)
                continue
            if stripped.startswith(MANUAL_SKIP_PREFIX):
                raw = stripped[len(MANUAL_SKIP_PREFIX):].strip()
                try:
                    log_callback(f"  ⏭ Залишено без змін (не змінювати): «{raw}»\n", "skip")
                except TypeError:
                    # log_callback без підтримки тегів (напр. просте логування
                    # в консоль без tkinter) — пишемо звичайним рядком.
                    log_callback(f"  ⏭ Залишено без змін (не змінювати): «{raw}»\n")
                continue
            log_callback(line)
        proc.wait()
        return proc.returncode
    except Exception as e:
        log_callback(f"\nПОМИЛКА ЗАПУСКУ: {e}\n")
        return -1
    finally:
        if proc_holder is not None:
            proc_holder["proc"] = None


def kill_process_tree(pid, log_callback=None):
    """Жорстко завершує процес PID і ВСІХ його нащадків одним викликом.

    Це критично саме через doc_processor.py: при LLM fallback він
    піднімає окремий дочірній процес llama-server.exe із завантаженою в
    оперативну пам'ять моделлю (кілька гігабайт). Якщо просто вбити сам
    python.exe (proc.terminate()), llama-server.exe лишається "сиротою",
    процес видно в диспетчері задач, і вся пам'ять моделі НЕ звільняється.

    taskkill.exe — системна утиліта Windows, викликається тут напряму
    списком аргументів (без shell=True, без cmd.exe як інтерпретатора) —
    так само, як у решті скрипта викликаються python.exe/magick.exe.
    /T — завершити разом з усім деревом нащадків, /F — примусово."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if log_callback:
            log_callback(f"  Дерево процесів (PID {pid}) примусово завершено — "
                         f"пам'ять, зайнята моделлю/дочірніми процесами, звільнена.\n")
    except Exception as e:
        if log_callback:
            log_callback(f"  Увага: не вдалося завершити процес через taskkill: {e}\n")


def open_in_explorer(path):
    try:
        os.startfile(str(path))  # тільки Провідник, без cmd/powershell
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────────
# Лог-панель (спільна для всіх вкладок)
# ──────────────────────────────────────────────────────────────────────
class LogPanel(ttk.Frame):
    def __init__(self, parent, height=18):
        super().__init__(parent)
        tk.Label(self, text="Лог виконання (наживо):").pack(anchor="w")
        self.text = scrolledtext.ScrolledText(self, height=height, font=("Consolas", 9))
        self.text.pack(fill="both", expand=True)
        # Тег "skip" — підсвічує рядки про рішення оператора "не змінювати"
        # блакитним, щоб їх було одразу видно в потоці логу.
        self.text.tag_configure("skip", foreground="#0d47a1", font=("Consolas", 9, "bold"))
        self.log_file = None
        self.log_file_path = None

    def log(self, text, tag=None):
        # Може викликатись з фонового потоку — плануємо у головний цикл tk.
        self.after(0, self._append, text, tag)

    def _append(self, text, tag=None):
        if tag:
            self.text.insert(tk.END, text, tag)
        else:
            self.text.insert(tk.END, text)
        self.text.see(tk.END)
        if self.log_file:
            try:
                self.log_file.write(text)
                self.log_file.flush()
                os.fsync(self.log_file.fileno())
            except Exception:
                pass

    def clear(self):
        self.text.delete("1.0", tk.END)

    def open_log_file(self, prefix="run"):
        try:
            LOGS_DIR.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            self.log_file_path = LOGS_DIR / f"{prefix}_{ts}.log"
            self.log_file = open(self.log_file_path, "w", encoding="utf-8", buffering=1)
        except Exception as e:
            self.log_file = None
            self.log_file_path = None
            self.log(f"УВАГА: не вдалося створити лог-файл: {e}\n")

    def close_log_file(self):
        if self.log_file:
            try:
                self.log_file.close()
            except Exception:
                pass
            self.log_file = None


# ──────────────────────────────────────────────────────────────────────
# Базовий клас вкладки: вибір шляхів + кнопка запуску + лог
# ──────────────────────────────────────────────────────────────────────
class BaseToolTab(ttk.Frame):
    def __init__(self, parent, script_path=None):
        super().__init__(parent)
        self.script_path = script_path
        self._row = 0
        self.is_running = False
        self.pad = {"padx": 10, "pady": 5}
        # Тримає посилання на активний Popen дочірнього процесу цієї вкладки
        # (заповнюється всередині stream_subprocess через proc_holder).
        # Саме через нього кнопка "Зупинити" дістає PID для kill_process_tree.
        self.current_proc = {"proc": None}
        self.stop_btn = None

        self.fields_frame = ttk.Frame(self)
        self.fields_frame.pack(fill="x")
        self.fields_frame.grid_columnconfigure(1, weight=1)

        if script_path is not None:
            self.status_label = tk.Label(self, text="", anchor="w", fg="#555555")
            self.status_label.pack(fill="x", padx=10)
            self._refresh_status()

    # ── статус наявності скрипта ──
    def _refresh_status(self):
        if self.script_path is None:
            return
        if self.script_path.is_file():
            self.status_label.config(text=f"✓ Скрипт знайдено: {self.script_path.name}", fg="#2e7d32")
        else:
            self.status_label.config(text=f"✗ Скрипт НЕ знайдено: {self.script_path}", fg="#c62828")

    # ── допоміжні конструктори полів (grid у fields_frame) ──
    def _next_row(self):
        r = self._row
        self._row += 1
        return r

    def add_open_file_row(self, label, var, filetypes, title=None):
        r = self._next_row()
        tk.Label(self.fields_frame, text=label).grid(row=r, column=0, sticky="w", **self.pad)
        tk.Entry(self.fields_frame, textvariable=var, width=68).grid(row=r, column=1, sticky="we", **self.pad)
        tk.Button(
            self.fields_frame, text="Вибрати...",
            command=lambda: self._pick_open_file(var, filetypes, title or label)
        ).grid(row=r, column=2, **self.pad)

    def add_save_file_row(self, label, var, filetypes, default_ext, title=None):
        r = self._next_row()
        tk.Label(self.fields_frame, text=label).grid(row=r, column=0, sticky="w", **self.pad)
        tk.Entry(self.fields_frame, textvariable=var, width=68).grid(row=r, column=1, sticky="we", **self.pad)
        tk.Button(
            self.fields_frame, text="Зберегти як...",
            command=lambda: self._pick_save_file(var, filetypes, default_ext, title or label)
        ).grid(row=r, column=2, **self.pad)

    def add_dir_row(self, label, var, title=None):
        r = self._next_row()
        tk.Label(self.fields_frame, text=label).grid(row=r, column=0, sticky="w", **self.pad)
        tk.Entry(self.fields_frame, textvariable=var, width=68).grid(row=r, column=1, sticky="we", **self.pad)
        tk.Button(
            self.fields_frame, text="Вибрати...",
            command=lambda: self._pick_dir(var, title or label)
        ).grid(row=r, column=2, **self.pad)

    def add_spinbox_row(self, label, var, from_, to_, hint=None):
        r = self._next_row()
        tk.Label(self.fields_frame, text=label).grid(row=r, column=0, sticky="w", **self.pad)
        tk.Spinbox(self.fields_frame, from_=from_, to=to_, increment=1, textvariable=var, width=10).grid(
            row=r, column=1, sticky="w", **self.pad
        )
        if hint:
            tk.Label(self.fields_frame, text=hint, fg="#555555").grid(row=r, column=2, sticky="w", **self.pad)

    def add_checkbox_row(self, label, var):
        r = self._next_row()
        tk.Checkbutton(self.fields_frame, text=label, variable=var).grid(
            row=r, column=0, columnspan=3, sticky="w", padx=10, pady=2
        )

    def add_hint(self, text):
        r = self._next_row()
        tk.Label(self.fields_frame, text=text, fg="#555555", justify="left").grid(
            row=r, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 4)
        )

    def add_run_button(self, text, command, color="#2e7d32"):
        btn = tk.Button(self, text=text, command=command, bg=color, fg="white",
                         font=("Segoe UI", 11, "bold"), height=2)
        btn.pack(fill="x", padx=10, pady=6)
        return btn

    def add_run_stop_row(self, run_text, run_command, color="#2e7d32", stop_command=None):
        """Створює пару кнопок в одному рядку: "Запустити" (зліва) і
        "Зупинити" (справа, спочатку неактивна — вмикається лише поки
        йде виконання). Повертає (run_btn, stop_btn)."""
        row = ttk.Frame(self)
        row.pack(fill="x", padx=10, pady=6)
        run_btn = tk.Button(row, text=run_text, command=run_command, bg=color, fg="white",
                             font=("Segoe UI", 11, "bold"), height=2)
        run_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))
        stop_btn = tk.Button(
            row, text="Зупинити", command=stop_command or self.stop_clicked,
            bg="#c62828", fg="white", font=("Segoe UI", 11, "bold"), height=2,
            state="disabled",
        )
        stop_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.stop_btn = stop_btn
        return run_btn, stop_btn

    # ── зупинка процесу (спільна для всіх вкладок, що йдуть через subprocess) ──
    def stop_clicked(self):
        """Примусово завершує дочірній процес цієї вкладки разом з усіма
        його нащадками (напр. осиротілий llama-server.exe з моделлю в
        пам'яті) і звільняє відповідну пам'ять. Якщо процес уже завершився
        сам — просто прибирає стан кнопок."""
        if not self.is_running:
            return
        proc = self.current_proc.get("proc")
        if proc is not None and proc.poll() is None:
            self.log_panel.log("\n>>> Зупинка за запитом користувача...\n")
            kill_process_tree(proc.pid, self.log_panel.log)
        else:
            self.log_panel.log("\n>>> Процес уже завершився, нема що зупиняти.\n")
        if self.stop_btn:
            self.stop_btn.config(state="disabled")
        # Реальне повернення run_btn у активний стан і фінальне логування
        # робить власний _worker/_finish вкладки — після kill читання
        # stdout процесу завершується саме собою і потік доходить до кінця.

    def _release_memory(self):
        """Викликати одразу після завершення/зупинки процесу вкладки.
        Прибирає посилання на завершений Popen і форсує збірку сміття в
        самому GUI-процесі (лог-буфери, тимчасові рядки виводу тощо)."""
        self.current_proc["proc"] = None
        gc.collect()

    # ── діалоги вибору (нативні Windows-діалоги через tkinter, без powershell) ──
    def _pick_open_file(self, var, filetypes, title):
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _pick_save_file(self, var, filetypes, default_ext, title):
        path = filedialog.asksaveasfilename(title=title, defaultextension=default_ext, filetypes=filetypes)
        if path:
            var.set(path)

    def _pick_dir(self, var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)


XLSX_TYPES = [("Excel файли", "*.xlsx *.xlsm"), ("Усі файли", "*.*")]


# ──────────────────────────────────────────────────────────────────────
# Вкладка 1: doc_processor.py
# ──────────────────────────────────────────────────────────────────────
class DocProcessorTab(BaseToolTab):
    def __init__(self, parent):
        super().__init__(parent, DOC_PROCESSOR)

        self.dict_path = tk.StringVar()
        self.remains_path = tk.StringVar()
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(OUTPUT_DOCS_DIR))
        self.check_remains = tk.BooleanVar(value=True)
        self.manual_match = tk.BooleanVar(value=False)

        self.add_open_file_row("Словник одиниць (xlsx/xlsm):", self.dict_path, XLSX_TYPES)
        self.add_open_file_row("Таблиця залишків (xlsx/xlsm):", self.remains_path, XLSX_TYPES)
        self.add_dir_row("Папка з документами:", self.input_dir)
        self.add_dir_row("Папка результату:", self.output_dir)
        self.add_checkbox_row(
            "Перевіряти залишок накопичувально між документами "
            "(вимкни, щоб фарбувати наіменування червоним при нестачі)",
            self.check_remains,
        )
        self.add_checkbox_row(
            "Ручний режим зіставлення назв (при неоднозначному збігу — "
            "запитає, який варіант правильний, замість автоматичного вибору)",
            self.manual_match,
        )

        self.run_btn, self.stop_btn = self.add_run_stop_row("Запустити обробку", self.run_clicked)
        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def run_clicked(self):
        if self.is_running:
            return
        if not self.dict_path.get() or not self.remains_path.get() or not self.input_dir.get():
            messagebox.showwarning("Увага", "Заповни всі три обов'язкові поля (словник, залишки, папка документів).")
            return
        self._refresh_status()
        if not DOC_PROCESSOR.is_file():
            messagebox.showerror("Помилка", f"doc_processor.py не знайдено:\n{DOC_PROCESSOR}")
            return

        out_dir = Path(self.output_dir.get() or OUTPUT_DOCS_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        threading.Thread(target=self._worker, args=(out_dir,), daemon=True).start()

    def _worker(self, out_dir):
        cmd = [sys.executable, "-u", str(DOC_PROCESSOR),
               self.dict_path.get(), self.remains_path.get(), self.input_dir.get(), str(out_dir)]
        if not self.check_remains.get():
            cmd.append("--no-check-remains")
        manual_on = self.manual_match.get()
        if manual_on:
            cmd.append("--manual-match")
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")
        exit_code = stream_subprocess(
            cmd, SCRIPT_DIR, utf8_env(), self.log_panel.log,
            proc_holder=self.current_proc,
            manual_prompt_handler=self._manual_prompt_handler if manual_on else None,
        )
        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log(f"Готово! Результат у: {out_dir}\n")
            open_in_explorer(out_dir)
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Завершено з кодом помилки {exit_code}.\n")
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Запустити обробку")
        self.stop_btn.config(state="disabled")
        self._release_memory()


# ──────────────────────────────────────────────────────────────────────
# Вкладка 6: «Інші інструменти» — автоматично підбирає решту скриптів
# ──────────────────────────────────────────────────────────────────────
class OtherToolsTab(BaseToolTab):
    """Вкладка "Інші інструменти" — ліворуч список скриптів (25%), праворуч
    дві панелі: зверху — налаштування для обраного скрипта, знизу — опис/живий лог.
    При запуску формуються стандартні прапори `--input-dir`, `--output-dir`,
    `--threshold` та ін., а також додаються довільні аргументи з поля "Додаткові аргументи".
    """

    def __init__(self, parent):
        super().__init__(parent, None)

        # Скрипти
        self.scripts = []  # list[Path]

        # Ліва колонка — список скриптів (прибл. 25% ширини)
        left = ttk.Frame(self)
        left.pack(side="left", fill="y", padx=(10, 6), pady=10)
        tk.Label(left, text="Доступні скрипти:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.listbox = tk.Listbox(left, width=32, height=30, activestyle="dotbox")
        self.listbox.pack(side="left", fill="y", expand=False)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        sb = ttk.Scrollbar(left, orient="vertical", command=self.listbox.yview)
        sb.pack(side="left", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        # Права область — розділ на дві рівні частини (top=settings, bottom=desc/log)
        right = ttk.Frame(self)
        right.pack(side="left", fill="both", expand=True, padx=(6, 10), pady=10)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        # Верхня половина — налаштування
        top = ttk.Frame(right, padding=(6, 6))
        top.grid(row=0, column=0, sticky="nsew")
        tk.Label(top, text="Налаштування скрипта", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        # container where per-script pages will be created
        self.settings_area = ttk.Frame(top)
        self.settings_area.pack(fill="both", expand=True, pady=(6, 0))
        self.settings_area.grid_columnconfigure(1, weight=1)

        # storage for per-script tkinter variables: {script_name: {varname: tk.Variable}}
        self.script_vars = {}
        self.watch_config = tk.BooleanVar(value=False)

        # mapping: script_name -> list of field definitions (type, key, label, extra)
        # types: dir, open_file, save_file, int, bool, text
        self.per_script_fields = {
            "pipeline_pdf_to_excel.py": [
                ("dir", "input_dir", "Папка з PDF (input):"),
                ("save_file", "output_file", "Вихідний Excel (save as):", XLSX_TYPES, ".xlsx"),
                ("open_file", "template", "Файл шаблону (xlsm/xlsx):", XLSX_TYPES),
                ("int", "threshold", "Поріг (%)", 50, 100, 90),
                ("bool", "llm", "Використовувати LLM (--llm):"),
                ("bool", "preprocess", "Застосувати препроцесинг (--preprocess):"),
            ],
            "pdf_to_text_multithread.py": [
                ("dir", "input_dir", "Папка з PDF (input):"),
                ("dir", "output_dir", "Папка з txt (output):"),
                ("bool", "no_magick", "Виключити ImageMagick (--no-magick):"),
                ("bool", "save_training", "Зберегти навчальні дані (--save-training):"),
                ("text", "search_phrase", "Фраза для пошуку (--phrase):"),
                ("int", "workers", "Кількість потоків:", 1, 32, 4),
            ],
            "ocr_text_corrector.py": [
                ("dir", "input_dir", "Папка з .txt (input):"),
                ("dir", "output_dir", "Папка результату (output):"),
                ("open_file", "template", "Макет (xlsm/xlsx):", XLSX_TYPES),
                ("bool", "llm", "Використовувати LLM (--llm):"),
            ],
            "excel_report_generator.py": [
                ("dir", "input_dir", "Папка з .txt (input):"),
                ("save_file", "output_file", "Вихідний Excel (save as):", XLSX_TYPES, ".xlsx"),
                ("open_file", "template", "Файл шаблону:", XLSX_TYPES),
                ("int", "threshold", "Поріг нечіткого пошуку (%):", 50, 100, 90),
                ("dir", "history_folder", "Папка історії (опційно):"),
            ],
            "doc_processor.py": [
                ("open_file", "dict_path", "Словник (xlsx/xlsm):", XLSX_TYPES),
                ("open_file", "remains_path", "Таблиця залишків (xlsx/xlsm):", XLSX_TYPES),
                ("dir", "input_dir", "Папка з документами:"),
                ("dir", "output_dir", "Папка результату:"),
                ("bool", "check_remains", "Перевіряти залишки (накопичувально):"),
                ("bool", "manual_match", "Ручний режим зіставлення (--manual-match):"),
            ],
            "batch_processor.py": [
                ("dir", "folder", "Папка з .txt файлами (input):"),
                ("dir", "output_dir", "Папка результату (--output-dir):"),
                ("open_file", "template", "Файл шаблону (template.xlsm):", XLSX_TYPES),
                ("bool", "no_llm", "Використовувати тільки детерміністичний режим (--no-llm):"),
                ("bool", "non_recursive", "Не шукати у підпапках (--non-recursive):"),
                ("bool", "overwrite", "Перезаписати наявні Excel файли (--overwrite):"),
            ],
            "cards_photo_to_excel.py": [
                ("dir", "input", "Папка з фото карток (--input):"),
                ("save_file", "output", "Вихідний Excel файл (--output):", XLSX_TYPES, ".xlsx"),
                ("text", "provider", "Провайдер (--provider):"),
                ("text", "model", "Модель (--model):"),
                ("text", "api_key", "API ключ (--api-key):"),
                ("open_file", "env_file", "Файл .env (--env-file):", [("ENV файли", "*.env"), ("Усі файли", "*.*")]),
            ],
            "image_preprocessor.py": [
                ("dir", "input", "Папка з фото (input):"),
                ("dir", "output", "Папка результату (output):"),
                ("int", "rotate", "Примусовий поворот (0/90/180/270):", 0, 270, 0),
                ("bool", "no_strips", "Не нарізати на смуги (--no-strips):"),
                ("bool", "no_gamma", "Без gamma варіанту (--no-gamma):"),
                ("bool", "no_osd", "Не використовувати OSD (--no-osd):"),
                ("bool", "no_curve_fix", "Без кривинного виправлення (--no-curve-fix):"),
                ("bool", "denoise", "Зробити шумозаглушення (--denoise):"),
                ("int", "workers", "Кількість потоків (--workers):", 1, 32, 3),
            ],
            "import_balances_docx.py": [
                ("open_file", "input", "Вхідний docx файл (--input):", [("Word files", "*.docx *.doc"), ("Усі файли", "*.*")]),
                ("save_file", "output", "Вихідний Excel файл (--output):", XLSX_TYPES, ".xlsx"),
                ("text", "type", "Тип документа (--type):"),
                ("text", "location_name", "Назва складу/підрозділу (--location-name):"),
                ("int", "year", "Рік (--year):", 2000, 2100, 2026),
            ],
            "import_pidrozdily_xlsm.py": [
                ("open_file", "input", "Вхідний xlsm файл (--input):", XLSX_TYPES),
                ("save_file", "output", "Вихідний xlsx файл (--output):", XLSX_TYPES, ".xlsx"),
                ("int", "year", "Рік (--year):", 2000, 2100, 2026),
            ],
            "validator.py": [
                ("dir", "folder", "Папка з JSON файлами:"),
                ("bool", "non_recursive", "Не шукати у підпапках (--non-recursive):"),
                ("bool", "summary", "Тільки підсумок (--summary):"),
                ("save_file", "report_file", "Звіт JSON (--report-file):", [("JSON файли", "*.json"), ("Усі файли", "*.*")], ".json"),
            ],
            "vidomist_report.py": [
                ("open_file", "xlsm", "Вхідна відомість (.xlsm/.xlsx):", XLSX_TYPES),
                ("open_file", "doc", "Вхідний наказ (.doc/.docx):", [("Word files", "*.docx *.doc"), ("Усі файли", "*.*")]),
                ("save_file", "report", "Звіт .xlsx (--report):", XLSX_TYPES, ".xlsx"),
            ],
            "universal_dict_extractor.py": [
                ("open_file", "macket", "Макет xlsm (macket):", XLSX_TYPES),
                ("dir", "folder", "Папка джерел (folder):"),
                ("int", "threshold", "Поріг нечіткого пошуку (%):", 50, 100, 85),
            ],
        }

        # default run/stop buttons
        btns = ttk.Frame(top)
        btns.pack(fill="x", pady=(6, 0))
        self.run_btn, self.stop_btn = self.add_run_stop_row("Запустити", self.run_clicked, color="#6a1b9a")

        # Нижня половина — опис / лог (показує опис при виборі, лог під час виконання)
        bottom = ttk.Frame(right, padding=(6, 6))
        bottom.grid(row=1, column=0, sticky="nsew")

        self.desc = scrolledtext.ScrolledText(bottom, height=10, font=("Consolas", 10))
        self.desc.pack(fill="both", expand=True)

        self.log_panel = LogPanel(bottom, height=10)
        # log_panel буде приховано до запуску процесу
        self.log_panel.pack_forget()

        self._populate_scripts()

        # auto-select first script so a page is visible on startup
        try:
            if self.scripts:
                self.listbox.selection_set(0)
                self.listbox.activate(0)
                self._on_select()
        except Exception:
            pass

        # config writer thread handle
        self._config_path = None
        self._config_writer_stop = None
        self.dynamic_vars = {}

    # ----------------- dynamic option discovery -----------------
    def _parse_help_options(self, script_path: Path, timeout: float = 1.0):
        """Try to run `<python> script --help` and parse possible options.
        Returns list of dicts: {name, has_value, help}.
        """
        opts = []
        try:
            proc = subprocess.Popen([sys.executable, str(script_path), "--help"],
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            try:
                out, _ = proc.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                proc.kill()
                out = ""
        except Exception:
            out = ""

        # simple parsing: lines that start with space then '-' contain options
        for ln in out.splitlines():
            s = ln.strip()
            if s.startswith("-"):
                parts = s.split()
                # find option token like --name
                name = None
                has_value = False
                for tok in parts:
                    if tok.startswith("--"):
                        name = tok
                        # if token contains '=' or next part is <...> assume value
                        if "=" in tok or "<" in tok or any('=' in p for p in parts):
                            has_value = True
                        break
                    if tok.startswith("-") and len(tok) == 2 and any(p.startswith("--") for p in parts):
                        # short form with long form present
                        continue
                help_text = " ".join(parts[1:]).strip()
                if name:
                    opts.append({"name": name, "has_value": has_value, "help": help_text})
        return opts

    def _build_dynamic_options(self, options):
        # remove previous dynamic widgets
        if hasattr(self, "dynamic_frame"):
            self.dynamic_frame.destroy()
        self.dynamic_frame = ttk.Frame(self)
        # insert dynamic frame just above bottom area (desc/log)
        # pack it into the right-top area by placing it after fixed settings
        self.dynamic_frame.pack_forget()
        # We'll pack it below the settings area by inserting before desc
        # For simplicity, append to settings_area in init, but here just create widgets
        r = 0
        self.dynamic_vars = {}
        for opt in options:
            key = opt["name"].lstrip("-").replace("-", "_")
            if opt["has_value"]:
                var = tk.StringVar()
                tk.Label(self.dynamic_frame, text=f"{opt['name']}: ").grid(row=r, column=0, sticky="w", padx=6, pady=2)
                ttk.Entry(self.dynamic_frame, textvariable=var).grid(row=r, column=1, sticky="we", padx=6, pady=2)
            else:
                var = tk.BooleanVar()
                ttk.Checkbutton(self.dynamic_frame, text=opt['name'], variable=var).grid(row=r, column=0, columnspan=2, sticky="w", padx=6, pady=2)
            self.dynamic_vars[key] = var
            r += 1
        if r:
            # place dynamic_frame into the UI: pack it above desc
            self.dynamic_frame.pack(fill="x", padx=10, pady=(6, 6))

    # ----------------- launcher-config writing -----------------
    def _start_config_writer(self, static_cfg: dict, interval: float = 2.0):
        # write initial config to a temp file
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        cfg_path = Path(tempfile.gettempdir()) / f"launcher_config_{ts}.json"
        self._config_path = str(cfg_path)
        def _write_once():
            try:
                cfg_path.write_text(json.dumps(static_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
            except Exception:
                pass
        _write_once()

        stop_event = threading.Event()

        def _writer_loop():
            while not stop_event.wait(interval):
                cfg = self._collect_current_config(static_cfg)
                try:
                    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    pass

        t = threading.Thread(target=_writer_loop, daemon=True)
        t.start()

        def stop():
            stop_event.set()
            try:
                t.join(timeout=0.2)
            except Exception:
                pass

        self._config_writer_stop = stop
        return self._config_path

    def _stop_config_writer(self):
        if self._config_writer_stop:
            try:
                self._config_writer_stop()
            except Exception:
                pass
            self._config_writer_stop = None

    def _collect_current_config(self, base: dict):
        cfg = dict(base)
        # try to pull standard fields from the currently built script vars (if any)
        sel = self.listbox.curselection()
        vars_map = {}
        if sel:
            try:
                script = self.scripts[sel[0]]
                vars_map = self.script_vars.get(script.name, {}) or {}
            except Exception:
                vars_map = {}

        def _get_var_value(name):
            v = vars_map.get(name)
            if v is None:
                # fallback to old attributes if present
                try:
                    attr = getattr(self, name)
                    return attr.get()
                except Exception:
                    return None
            try:
                return v.get()
            except Exception:
                return None

        cfg['input_dir'] = (_get_var_value('input_dir') or None)
        cfg['output_dir'] = (_get_var_value('output_dir') or None)
        cfg['template'] = (_get_var_value('template') or None)
        try:
            thr = _get_var_value('threshold')
            cfg['threshold'] = int(thr) if thr is not None and thr != '' else None
        except Exception:
            cfg['threshold'] = None
        cfg['llm'] = bool(_get_var_value('llm'))
        cfg['preprocess'] = bool(_get_var_value('preprocess'))
        cfg['no_magick'] = bool(_get_var_value('no_magick'))
        cfg['save_training'] = bool(_get_var_value('save_training'))
        cfg['keep_existing_txt'] = bool(_get_var_value('keep_existing_txt'))
        cfg['history_folder'] = (_get_var_value('history_folder') or None)
        cfg['search_phrase'] = (_get_var_value('search_phrase') or None)

        # dynamic vars
        dyn = {}
        for k, v in self.dynamic_vars.items():
            try:
                dyn[k] = v.get()
            except Exception:
                dyn[k] = None
        cfg['dynamic'] = dyn
        return cfg

    def _populate_scripts(self):
        exclude = {
            "run_gui_unified.py",
            "ocr_text_corrector.py",
            "image_preprocessor.py",
            "run_gui_image_preprocessor.py",
            "run_vidomist.py",
        }
        self.scripts = []
        for p in sorted(SCRIPT_DIR.glob("*.py")):
            if p.name in exclude:
                continue
            self.scripts.append(p)
        self.listbox.delete(0, tk.END)
        for p in self.scripts:
            self.listbox.insert(tk.END, p.name)

    def _clear_settings_area(self):
        for child in list(self.settings_area.winfo_children()):
            child.destroy()

    def _build_script_page(self, script_name: str):
        self._clear_settings_area()
        fields = self.per_script_fields.get(script_name)
        vars_map = self.script_vars.get(script_name, {})
        if not vars_map:
            vars_map = {}
        row = 0

        def add_var(name, var):
            vars_map[name] = var
            return var

        if fields is None:
            # fallback page: expose the generic options that the generic
            # arg collector can later convert into flags.
            input_dir = vars_map.get("input_dir") or add_var("input_dir", tk.StringVar())
            output_dir = vars_map.get("output_dir") or add_var("output_dir", tk.StringVar())
            template = vars_map.get("template") or add_var("template", tk.StringVar())
            threshold = vars_map.get("threshold") or add_var("threshold", tk.IntVar(value=90))
            llm = vars_map.get("llm") or add_var("llm", tk.BooleanVar(value=False))
            preprocess = vars_map.get("preprocess") or add_var("preprocess", tk.BooleanVar(value=False))
            no_magick = vars_map.get("no_magick") or add_var("no_magick", tk.BooleanVar(value=False))
            save_training = vars_map.get("save_training") or add_var("save_training", tk.BooleanVar(value=False))
            keep_existing_txt = vars_map.get("keep_existing_txt") or add_var("keep_existing_txt", tk.BooleanVar(value=False))
            history_folder = vars_map.get("history_folder") or add_var("history_folder", tk.StringVar())
            extra_args = vars_map.get("extra_args") or add_var("extra_args", tk.StringVar())

            tk.Label(self.settings_area, text="Папка-джерело (input):").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(self.settings_area, textvariable=input_dir).grid(row=row, column=1, sticky="we", pady=4)
            ttk.Button(self.settings_area, text="Вибрати", command=lambda v=input_dir: self._pick_dir(v, "Папка-джерело")).grid(row=row, column=2, padx=6)
            row += 1

            tk.Label(self.settings_area, text="Папка-результат (output):").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(self.settings_area, textvariable=output_dir).grid(row=row, column=1, sticky="we", pady=4)
            ttk.Button(self.settings_area, text="Вибрати", command=lambda v=output_dir: self._pick_dir(v, "Папка-результат")).grid(row=row, column=2, padx=6)
            row += 1

            tk.Label(self.settings_area, text="Файл шаблону (якщо є):").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(self.settings_area, textvariable=template).grid(row=row, column=1, sticky="we", pady=4)
            ttk.Button(self.settings_area, text="Вибрати", command=lambda v=template: self._pick_open_file(v, XLSX_TYPES, "Файл шаблону")).grid(row=row, column=2, padx=6)
            row += 1

            tk.Label(self.settings_area, text="Поріг нечіткого пошуку (%):").grid(row=row, column=0, sticky="w", pady=4)
            tk.Spinbox(self.settings_area, from_=0, to=100, increment=1, textvariable=threshold, width=8).grid(row=row, column=1, sticky="w", pady=4)
            row += 1

            ttk.Checkbutton(self.settings_area, text="Використовувати LLM (--llm)", variable=llm).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
            row += 1
            ttk.Checkbutton(self.settings_area, text="Застосувати препроцесинг (--preprocess)", variable=preprocess).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
            row += 1
            ttk.Checkbutton(self.settings_area, text="Вимкнути ImageMagick (--no-magick)", variable=no_magick).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
            row += 1
            ttk.Checkbutton(self.settings_area, text="Зберегти навчальні дані (--save-training)", variable=save_training).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
            row += 1
            ttk.Checkbutton(self.settings_area, text="Залишити існуючі txt-файли (--keep-existing-txt)", variable=keep_existing_txt).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
            row += 1

            tk.Label(self.settings_area, text="Папка історії (--history-folder):").grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(self.settings_area, textvariable=history_folder).grid(row=row, column=1, sticky="we", pady=4)
            ttk.Button(self.settings_area, text="Вибрати", command=lambda v=history_folder: self._pick_dir(v, "Папка історії")).grid(row=row, column=2, padx=6)
            row += 1

            tk.Label(self.settings_area, text="Додаткові аргументи:").grid(row=row, column=0, sticky="w", pady=8)
            ttk.Entry(self.settings_area, textvariable=extra_args).grid(row=row, column=1, columnspan=2, sticky="we", pady=8)
            row += 1
        else:
            for field in fields:
                ftype = field[0]
                key = field[1]
                label = field[2]
                existing = vars_map.get(key)
                if existing is None:
                    if ftype == "bool":
                        var = tk.BooleanVar(value=False)
                    elif ftype == "int":
                        default = field[5] if len(field) > 5 else 0
                        var = tk.IntVar(value=default)
                    else:
                        var = tk.StringVar()
                    vars_map[key] = var
                else:
                    var = existing

                if ftype == "dir":
                    tk.Label(self.settings_area, text=label).grid(row=row, column=0, sticky="w", pady=4)
                    ttk.Entry(self.settings_area, textvariable=var).grid(row=row, column=1, sticky="we", pady=4)
                    ttk.Button(self.settings_area, text="Вибрати", command=lambda v=var, t=label: self._pick_dir(v, t)).grid(row=row, column=2, padx=6)
                elif ftype == "open_file":
                    tk.Label(self.settings_area, text=label).grid(row=row, column=0, sticky="w", pady=4)
                    ttk.Entry(self.settings_area, textvariable=var).grid(row=row, column=1, sticky="we", pady=4)
                    ttk.Button(self.settings_area, text="Вибрати", command=lambda v=var, ft=field[3], t=label: self._pick_open_file(v, ft, t)).grid(row=row, column=2, padx=6)
                elif ftype == "save_file":
                    tk.Label(self.settings_area, text=label).grid(row=row, column=0, sticky="w", pady=4)
                    ttk.Entry(self.settings_area, textvariable=var).grid(row=row, column=1, sticky="we", pady=4)
                    ttk.Button(self.settings_area, text="Зберегти як...", command=lambda v=var, ft=field[3], de=field[4], t=label: self._pick_save_file(v, ft, de, t)).grid(row=row, column=2, padx=6)
                elif ftype == "int":
                    tk.Label(self.settings_area, text=label).grid(row=row, column=0, sticky="w", pady=4)
                    tk.Spinbox(self.settings_area, from_=field[3], to=field[4], increment=1, textvariable=var, width=8).grid(row=row, column=1, sticky="w", pady=4)
                elif ftype == "bool":
                    ttk.Checkbutton(self.settings_area, text=label, variable=var).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=2)
                else:
                    tk.Label(self.settings_area, text=label).grid(row=row, column=0, sticky="w", pady=4)
                    ttk.Entry(self.settings_area, textvariable=var).grid(row=row, column=1, sticky="we", pady=4)
                row += 1

        # always show watch config checkbox
        watch = vars_map.get("watch_config") or add_var("watch_config", self.watch_config)
        ttk.Checkbutton(self.settings_area, text="Слідкувати за змінами конфігурації (watch config)", variable=watch).grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(10, 4))
        row += 1

        self.script_vars[script_name] = vars_map

    def _collect_args_for_script(self, script: Path):
        vars_map = self.script_vars.get(script.name, {})
        args = []
        base_cfg = {}

        def add_flag(key, flag):
            var = vars_map.get(key)
            if isinstance(var, tk.BooleanVar) and var.get():
                args.append(flag)
                base_cfg[key] = True

        def add_value(key, flag=None, positional=False):
            var = vars_map.get(key)
            if var is None:
                return
            try:
                value = var.get()
            except Exception:
                value = None
            if value in (None, ""):
                return
            if isinstance(var, tk.BooleanVar):
                if value:
                    if flag:
                        args.append(flag)
                    base_cfg[key] = True
            else:
                if positional:
                    args.append(str(value))
                elif flag:
                    args.extend([flag, str(value)])
                else:
                    args.extend([f"--{key.replace('_', '-')}", str(value)])
                base_cfg[key] = value

        if script.name == "pipeline_pdf_to_excel.py":
            add_value("input_dir", positional=True)
            add_value("output_file", positional=True)
            add_value("template", positional=True)
            add_flag("llm", "--llm")
            add_flag("preprocess", "--preprocess")
            add_value("threshold", "--threshold")
            add_value("history_folder", "--history-folder")
            add_flag("skip_ocr", "--skip-ocr")
            add_flag("keep_existing_txt", "--keep-existing-txt")
        elif script.name == "pdf_to_text_multithread.py":
            add_value("input_dir", positional=True)
            add_flag("no_magick", "--no-magick")
            add_flag("save_training", "--save-training")
            add_flag("preprocess", "--preprocess")
            add_value("min_conf", "--min-conf")
            add_value("fuzzy_threshold", "--fuzzy-threshold")
            add_value("search_phrase", "--phrase")
            add_flag("preprocess_osd", "--preprocess-osd")
            add_flag("no_llm_recheck", "--no-llm-recheck")
        elif script.name == "ocr_text_corrector.py":
            add_value("input_dir", positional=True)
            add_value("output_dir", positional=True)
            add_value("template", positional=True)
            add_flag("llm", "--llm")
        elif script.name == "excel_report_generator.py":
            add_value("input_dir", positional=True)
            add_value("output_file", positional=True)
            add_value("template", positional=True)
            add_value("threshold", "--threshold")
            add_value("history_folder", "--history-folder")
        elif script.name == "doc_processor.py":
            add_value("dict_path", positional=True)
            add_value("remains_path", positional=True)
            add_value("input_dir", positional=True)
            add_value("output_dir", positional=True)
            add_flag("check_remains", "--no-check-remains")
            add_flag("manual_match", "--manual-match")
        elif script.name == "batch_processor.py":
            add_value("folder", positional=True)
            add_value("output_dir", "--output-dir")
            add_value("template", "--template")
            add_flag("no_llm", "--no-llm")
            add_flag("non_recursive", "--non-recursive")
            add_flag("overwrite", "--overwrite")
        elif script.name == "cards_photo_to_excel.py":
            add_value("input", "--input")
            add_value("output", "--output")
            add_value("provider", "--provider")
            add_value("model", "--model")
            add_value("api_key", "--api-key")
            add_value("env_file", "--env-file")
        elif script.name == "image_preprocessor.py":
            add_value("input", positional=True)
            add_value("output", positional=True)
            add_value("rotate", "--rotate")
            add_flag("no_strips", "--no-strips")
            add_flag("no_gamma", "--no-gamma")
            add_flag("no_osd", "--no-osd")
            add_flag("no_curve_fix", "--no-curve-fix")
            add_flag("denoise", "--denoise")
            add_value("workers", "--workers")
        elif script.name == "import_balances_docx.py":
            add_value("input", "--input")
            add_value("output", "--output")
            add_value("type", "--type")
            add_value("location_name", "--location-name")
            add_value("year", "--year")
        elif script.name == "import_pidrozdily_xlsm.py":
            add_value("input", "--input")
            add_value("output", "--output")
            add_value("year", "--year")
        elif script.name == "validator.py":
            add_value("folder", positional=True)
            add_flag("non_recursive", "--non-recursive")
            add_flag("summary", "--summary")
            add_value("report_file", "--report-file")
        elif script.name == "vidomist_report.py":
            add_value("xlsm", positional=True)
            add_value("doc", positional=True)
            add_value("report", "--report")
        elif script.name == "universal_dict_extractor.py":
            add_value("macket", positional=True)
            add_value("folder", positional=True)
            add_value("threshold", "--threshold")
        else:
            # generic fallback
            add_value("input_dir", "--input-dir")
            add_value("output_dir", "--output-dir")
            add_value("template", "--template")
            add_flag("llm", "--llm")
            add_flag("preprocess", "--preprocess")
            add_flag("no_magick", "--no-magick")
            add_flag("save_training", "--save-training")
            add_flag("keep_existing_txt", "--keep-existing-txt")
            add_value("threshold", "--threshold")
            add_value("history_folder", "--history-folder")
            extra = vars_map.get("extra_args")
            text = ""
            if extra is not None:
                try:
                    text = extra.get().strip()
                except Exception:
                    text = ""
                if text:
                    args.extend(text.split())
            base_cfg["extra_args"] = text if text else None

        return args, base_cfg

    def _read_docstring(self, path: Path) -> str:
        try:
            text = path.read_text(encoding="utf-8")
        except Exception:
            return "(Не вдалося прочитати файл.)"
        idx = text.find('"""')
        if idx == -1:
            idx = text.find("'''")
            quote = "'''"
        else:
            quote = '"""'
        if idx == -1:
            return "(Докстрінг відсутній)"
        end = text.find(quote, idx + len(quote))
        if end == -1:
            return text[idx + len(quote):].strip()[:4000]
        return text[idx + len(quote):end].strip()

    def _on_select(self, evt=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        path = self.scripts[sel[0]]
        doc = self._read_docstring(path)
        header = f"{path.name} — {path}\n\n"
        self.desc.delete("1.0", tk.END)
        self.desc.insert(tk.END, header + doc)
        # показати опис (якщо лог зараз не активний)
        if not self.is_running:
            self.log_panel.pack_forget()
            self.desc.pack(fill="both", expand=True)
        # Побудувати сторінку налаштувань для обраного скрипта
        try:
            self._build_script_page(path.name)
        except Exception:
            pass

    def _configure_for_script(self, name: str):
        # Прості евристики для яких полів зазвичай потрібні
        lower = name.lower()
        # загально: показуємо всі поля, але підказкою — можна очищати непотрібні
        # (нехай користувач не заповнює їх якщо не потрібно)
        # Щоб зробити UI менш заповненим, можна приховати непотрібні,
        # але тут ми залишаємо видимими всі поля для гнучкості.
        pass

    def run_clicked(self):
        if self.is_running:
            return
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("Увага", "Вибери скрипт зі списку.")
            return
        script = self.scripts[sel[0]]
        if not script.is_file():
            messagebox.showerror("Помилка", f"Файл не знайдено:\n{script}")
            return

        # Побудова аргументів/конфігурації на основі активної сторінки скрипта
        args, base_cfg = self._collect_args_for_script(script)
        # include dynamic vars if any
        base_cfg = self._collect_current_config(base_cfg)

        cmd = [sys.executable, "-u", str(script)] + args

        # start config writer if watch enabled (or always write a static config and pass path)
        cfg_path = None
        try:
            cfg_path = self._start_config_writer(base_cfg) if self.watch_config.get() else None
        except Exception:
            cfg_path = None
        if cfg_path:
            cmd += ["--launcher-config", cfg_path, "--watch-config"]
        else:
            # write a single static config file and pass it so script can read initial settings
            try:
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                tmp = Path(tempfile.gettempdir()) / f"launcher_config_{ts}.json"
                tmp.write_text(json.dumps(base_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
                cmd += ["--launcher-config", str(tmp)]
            except Exception:
                pass

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        # показати лог і ховати опис
        self.desc.pack_forget()
        self.log_panel.pack(fill="both", expand=True)
        self.log_panel.clear()
        if self.log_panel.log_file_path:
            self.log_panel.log(f"Лог-файл цього запуску: {self.log_panel.log_file_path}\n")
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")

        threading.Thread(target=self._worker, args=(cmd,), daemon=True).start()

    def _worker(self, cmd):
        extra_env = {}
        if POPPLER_BIN.is_dir():
            extra_env["PATH"] = str(POPPLER_BIN) + os.pathsep + os.environ.get("PATH", "")
        exit_code = stream_subprocess(cmd, SCRIPT_DIR, utf8_env(extra_env), self.log_panel.log,
                                       proc_holder=self.current_proc)
        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log("Готово!\n")
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Завершено з кодом помилки {exit_code}.\n")
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Запустити")
        self.stop_btn.config(state="disabled")
        self._release_memory()
        # stop config writer if running
        try:
            self._stop_config_writer()
        except Exception:
            pass

    def _manual_prompt_handler(self, block_lines):
        """Викликається stream_subprocess З ФОНОВОГО ПОТОКУ (worker), коли
        doc_processor.py в --manual-match режимі надіслав блок кандидатів
        між MANUAL_MARK_BEGIN/END. Tkinter-віджети можна створювати лише в
        головному потоці, тому саме вікно будується через self.after(0, ...),
        а цей метод блокується на threading.Event, доки користувач не
        натисне кнопку. Повертає рядок з номером обраного варіанта
        ('0' — "не змінювати"), який піде прямо в stdin процесу."""
        scope = ""
        raw_name = ""
        options = []  # [(номер: str, текст: str)]
        for ln in block_lines:
            if ln.startswith("SCOPE:"):
                scope = ln[len("SCOPE:"):]
            elif ln.startswith("RAW:"):
                raw_name = ln[len("RAW:"):]
            else:
                head, sep, tail = ln.partition(")")
                if sep and head.strip().isdigit():
                    options.append((head.strip(), tail.strip()))

        result_holder = {"answer": "0"}
        done_event = threading.Event()

        def build_dialog():
            win = tk.Toplevel(self)
            win.title("Ручний вибір варіанта — Втрати майна")
            win.transient(self.winfo_toplevel())
            win.grab_set()
            win.resizable(False, False)

            tk.Label(win, text=scope, font=("Segoe UI", 9, "bold"),
                     wraplength=440, justify="left").pack(anchor="w", padx=12, pady=(12, 2))
            tk.Label(win, text=f"Сира назва з документа:  «{raw_name}»",
                     wraplength=440, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

            def choose(answer):
                result_holder["answer"] = answer
                win.destroy()
                done_event.set()

            for num, text in options:
                if num == "0":
                    # "Не змінювати" — окремо підсвічуємо блакитним, щоб
                    # оператор одразу бачив, що це варіант "залишити як є"
                    # (а не ще один кандидат зі словника/залишків).
                    tk.Button(
                        win, text="0) Не змінювати / залишити як є", width=62,
                        bg="#e3f2fd", fg="#0d47a1", activebackground="#bbdefb",
                        command=lambda: choose("0"),
                    ).pack(fill="x", padx=12, pady=(2, 8))
                else:
                    ttk.Button(
                        win, text=f"{num}) {text}", width=62,
                        command=lambda n=num: choose(n),
                    ).pack(fill="x", padx=12, pady=2)

            # ── Власний варіант оператора ──
            # Якщо серед показаних кандидатів немає правильного, оператор
            # може ввести точну назву самостійно — вона звіряється з ПОВНОЮ
            # базою (не лише з тим, що показано вище) на боці doc_processor.py.
            ttk.Separator(win, orient="horizontal").pack(fill="x", padx=12, pady=(6, 6))
            tk.Label(win, text="Або введіть свій варіант (точну назву з бази):",
                     wraplength=440, justify="left").pack(anchor="w", padx=12, pady=(0, 2))
            entry_row = ttk.Frame(win)
            entry_row.pack(fill="x", padx=12, pady=(0, 12))
            manual_var = tk.StringVar()
            entry = ttk.Entry(entry_row, textvariable=manual_var, width=46)
            entry.pack(side="left", fill="x", expand=True)

            def submit_manual():
                text = manual_var.get().strip()
                if text:
                    choose(text)

            entry.bind("<Return>", lambda e: submit_manual())
            ttk.Button(entry_row, text="Застосувати", command=submit_manual).pack(side="left", padx=(6, 0))

            win.protocol("WM_DELETE_WINDOW", lambda: choose("0"))
            win.after(50, lambda: entry.focus_force())

        self.after(0, build_dialog)
        done_event.wait()
        return result_holder["answer"]


# ──────────────────────────────────────────────────────────────────────
# Вкладка 2: excel_report_generator.py (+ кнопка "Тест")
# ──────────────────────────────────────────────────────────────────────
class ExcelGeneratorTab(BaseToolTab):
    REQUIRED_SHEETS = ("Макет відомості на списання", "Словник_підрозділи")

    def __init__(self, parent):
        super().__init__(parent, EXCEL_GENERATOR_SCRIPT)

        self.input_txt_dir = tk.StringVar()
        self.output_excel_file = tk.StringVar(value=str(SCRIPT_DIR / "Відомість_на_списання.xlsx"))
        self.template_excel_path = tk.StringVar()
        self.price_threshold_var = tk.IntVar(value=90)  # поріг нечіткого пошуку ЦІН (перенесено з price_matcher.py)
        self.history_folder = tk.StringVar()  # опційна папка з уже готовими файлами (запасне джерело цін)

        self.add_dir_row("Папка з .txt файлами (рапорти):", self.input_txt_dir)
        self.add_save_file_row("Вихідний Excel-файл (зберегти як):", self.output_excel_file, XLSX_TYPES, ".xlsx")
        self.add_open_file_row(
            "Шаблон Excel (аркуші 'Макет відомості...' і 'Словник_підрозділи'):",
            self.template_excel_path, XLSX_TYPES,
        )
        self.add_spinbox_row(
            "Поріг нечіткого пошуку цін (%):", self.price_threshold_var, 50, 100,
            hint="за замовч. 90; підбір цін запускається останнім кроком, після зіставлення назв/синонімів",
        )
        self.add_dir_row("Папка з готовими файлами (опційно, запасне джерело цін):", self.history_folder)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", padx=10, pady=(0, 4))
        self.run_btn = tk.Button(btn_row, text="Генерувати Відомість", command=self.run_clicked,
                                  bg="#2e7d32", fg="white", font=("Segoe UI", 11, "bold"), height=2)
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.test_btn = tk.Button(btn_row, text="Тест", command=self.test_clicked,
                                   bg="#ff9800", fg="white", font=("Segoe UI", 11, "bold"), height=2)
        self.test_btn.pack(side="left", fill="x", expand=True, padx=(4, 4))
        self.stop_btn = tk.Button(btn_row, text="Зупинити", command=self.stop_clicked,
                                   bg="#c62828", fg="white", font=("Segoe UI", 11, "bold"), height=2,
                                   state="disabled")
        self.stop_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def _check_template_sheets(self, template_path):
        if openpyxl is None:
            return True, "openpyxl не встановлено — перевірку аркушів пропущено."
        try:
            wb = openpyxl.load_workbook(template_path, read_only=True)
        except Exception as e:
            return False, f"Не вдалося відкрити шаблон:\n{template_path}\n\n{e}"
        missing = [s for s in self.REQUIRED_SHEETS if s not in wb.sheetnames]
        if missing:
            return False, f"Шаблон не містить аркуш(і): {', '.join(missing)}"
        return True, None

    def run_clicked(self):
        if self.is_running:
            return
        if not all([self.input_txt_dir.get(), self.output_excel_file.get(), self.template_excel_path.get()]):
            messagebox.showwarning("Увага", "Будь ласка, заповніть усі поля.")
            return
        if not EXCEL_GENERATOR_SCRIPT.is_file():
            messagebox.showerror("Помилка", f"Скрипт генератора не знайдено:\n{EXCEL_GENERATOR_SCRIPT}")
            return

        input_dir = Path(self.input_txt_dir.get())
        if not input_dir.is_dir():
            messagebox.showerror("Помилка", f"Вхідна папка не існує:\n{input_dir}")
            return
        if not any(input_dir.glob("*.txt")):
            messagebox.showerror("Помилка", f"У папці {input_dir} немає .txt файлів.")
            return

        output_file = Path(self.output_excel_file.get())
        try:
            output_file.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося створити папку для вихідного файлу:\n{output_file.parent}\n\n{e}")
            return

        template_path = Path(self.template_excel_path.get())
        if not template_path.is_file():
            messagebox.showerror("Помилка", f"Файл шаблону не знайдено:\n{template_path}")
            return

        ok, err = self._check_template_sheets(template_path)
        if not ok:
            messagebox.showerror("Помилка", err)
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.test_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        self.log_panel.open_log_file(prefix="run")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        cmd = [sys.executable, "-u", str(EXCEL_GENERATOR_SCRIPT),
               str(Path(self.input_txt_dir.get())),
               str(Path(self.output_excel_file.get())),
               str(Path(self.template_excel_path.get())),
               "--threshold", str(self.price_threshold_var.get())]
        if self.history_folder.get().strip():
            cmd += ["--history-folder", self.history_folder.get().strip()]
        if self.log_panel.log_file_path:
            self.log_panel.log(f"Лог-файл цього запуску: {self.log_panel.log_file_path}\n")
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")

        exit_code = stream_subprocess(cmd, SCRIPT_DIR, utf8_env(), self.log_panel.log,
                                       proc_holder=self.current_proc)
        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log(f"Генерацію завершено успішно! Результат: {self.output_excel_file.get()}\n")
            open_in_explorer(Path(self.output_excel_file.get()).parent)
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Генерацію завершено з кодом помилки {exit_code}.\n")

        self.log_panel.close_log_file()
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Генерувати Відомість")
        self.test_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._release_memory()

    def test_clicked(self):
        self.test_btn.config(state="disabled", text="Тестування...")
        self.log_panel.clear()
        self.log_panel.open_log_file(prefix="test")
        threading.Thread(target=self._run_test, daemon=True).start()

    def _run_test(self):
        if self.log_panel.log_file_path:
            self.log_panel.log(f"Лог-файл цього тестування: {self.log_panel.log_file_path}\n")
        self.log_panel.log("Тестування...\n")

        if not EXCEL_GENERATOR_SCRIPT.is_file():
            self.log_panel.log(f"Помилка: скрипт не знайдено: {EXCEL_GENERATOR_SCRIPT}\n")
            self._end_test()
            return

        try:
            import rapidfuzz  # noqa: F401
            if openpyxl is None:
                raise ImportError("openpyxl")
            self.log_panel.log("Бібліотеки openpyxl та rapidfuzz встановлені.\n")
        except ImportError as e:
            self.log_panel.log(f"Помилка: {e}\n")
            self._end_test()
            return

        if self.input_txt_dir.get():
            d = Path(self.input_txt_dir.get())
            if d.is_dir():
                self.log_panel.log(f"Вхідна папка: {d}, .txt файлів: {len(list(d.glob('*.txt')))}\n")
            else:
                self.log_panel.log(f"Помилка: папка {d} не існує.\n")

        if self.output_excel_file.get():
            p = Path(self.output_excel_file.get())
            try:
                p.parent.mkdir(parents=True, exist_ok=True)
                self.log_panel.log(f"Вихідний файл можна записати у {p.parent}\n")
            except Exception as e:
                self.log_panel.log(f"Помилка запису: {e}\n")

        if self.template_excel_path.get():
            t = Path(self.template_excel_path.get())
            if t.is_file():
                ok, err = self._check_template_sheets(t)
                self.log_panel.log(f"Шаблон: аркуші присутні: {ok}\n" if ok else f"Помилка шаблону: {err}\n")
            else:
                self.log_panel.log("Помилка: шаблон не знайдено.\n")

        self.log_panel.log("Тестування завершено.\n")
        self._end_test()

    def _end_test(self):
        self.log_panel.close_log_file()
        self.after(0, lambda: self.test_btn.config(state="normal", text="Тест"))


# ──────────────────────────────────────────────────────────────────────
# Вкладка 3: pdf_to_text_multithread.py
# ──────────────────────────────────────────────────────────────────────
class OcrTab(BaseToolTab):
    def __init__(self, parent):
        super().__init__(parent, OCR_SCRIPT)

        self.input_dir = tk.StringVar()
        self.no_magick = tk.BooleanVar(value=False)
        self.save_training = tk.BooleanVar(value=False)
        self.preprocess = tk.BooleanVar(value=False)
        self.correct_names = tk.BooleanVar(value=False)
        self.correction_template = tk.StringVar()
        self.correct_llm = tk.BooleanVar(value=False)
        self.corrected_txt_dir = tk.StringVar(value=str(SCRIPT_DIR / 'output_txt_fixed'))

        self.add_dir_row("Папка з PDF файлами:", self.input_dir)
        self.add_checkbox_row("--no-magick  (без ImageMagick, тільки poppler)", self.no_magick)
        self.add_checkbox_row("--save-training  (зберігати навчальні дані для Tesseract)", self.save_training)
        self.add_checkbox_row(
            "--preprocess  (підготовка сторінок: деварп, довертання, flat-field, CLAHE)",
            self.preprocess,
        )
        self.add_hint(
            "Підготовка коштує ~18-27 сек на сторінку. Вмикай для ФОТО та кривих сканів:\n"
            "на такій сторінці вона витягла фразу, яку звичайний OCR не знайшов.\n"
            "Для рівного цифрового PDF користі немає — якість навіть трохи падає."
        )
        self.add_dir_row("Вихідна папка для виправлених .txt:", self.corrected_txt_dir)
        self.add_checkbox_row("Застосувати корекцію назв після OCR", self.correct_names)
        self.add_open_file_row("Макет для корекції (.xlsm/.xlsx):", self.correction_template, XLSX_TYPES)
        self.add_checkbox_row("LLM для корекції назв після OCR", self.correct_llm)

        self.run_btn, self.stop_btn = self.add_run_stop_row("Запустити OCR", self.run_clicked, color="#1565c0")
        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def run_clicked(self):
        if self.is_running:
            return
        if not self.input_dir.get():
            messagebox.showwarning("Увага", "Вибери папку з PDF файлами.")
            return
        if not os.path.isdir(self.input_dir.get()):
            messagebox.showerror("Помилка", f"Папка не існує:\n{self.input_dir.get()}")
            return
        if not OCR_SCRIPT.is_file():
            messagebox.showerror("Помилка", f"Скрипт не знайдено:\n{OCR_SCRIPT}")
            return
        if self.correct_names.get() and not self.correction_template.get():
            messagebox.showwarning("Увага", "Виберіть макет для корекції назв або вимкніть корекцію.")
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        cmd = [sys.executable, "-u", str(OCR_SCRIPT), self.input_dir.get()]
        if self.no_magick.get():
            cmd.append("--no-magick")
        if self.save_training.get():
            cmd.append("--save-training")
        if self.preprocess.get():
            cmd.append("--preprocess")
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")

        extra = {}
        if POPPLER_BIN.is_dir():
            extra["PATH"] = str(POPPLER_BIN) + os.pathsep + os.environ.get("PATH", "")
        exit_code = stream_subprocess(cmd, SCRIPT_DIR, utf8_env(extra), self.log_panel.log,
                                       proc_holder=self.current_proc)

        if exit_code == 0 and self.correct_names.get():
            raw_dir = SCRIPT_DIR / "output_txt"
            corrected_dir = Path(self.corrected_txt_dir.get() or SCRIPT_DIR / "output_txt_fixed")
            corrected_dir.mkdir(parents=True, exist_ok=True)
            if not raw_dir.is_dir():
                self.log_panel.log(f"\nПОМИЛКА: не знайдено сирі файли OCR у {raw_dir}.\n")
            else:
                corr_cmd = [sys.executable, "-u", str(SCRIPT_DIR / "ocr_text_corrector.py"), str(raw_dir), str(corrected_dir), self.correction_template.get()]
                if self.correct_llm.get():
                    corr_cmd.append("--llm")
                self.log_panel.log(f"\nЗапускаю корекцію назв після OCR: {' '.join(corr_cmd)}\n\n")
                corr_code = stream_subprocess(corr_cmd, SCRIPT_DIR, utf8_env(extra), self.log_panel.log,
                                              proc_holder=self.current_proc)
                if corr_code == 0:
                    self.log_panel.log(f"\nКорекція завершена. Виправлені файли: {corrected_dir}\n")
                elif corr_code is None or corr_code < 0:
                    self.log_panel.log("\nКорекцію зупинено користувачем.\n")
                else:
                    self.log_panel.log(f"\nКорекція завершена з кодом помилки {corr_code}.\n")

        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log("Готово!\n")
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Завершено з кодом помилки {exit_code}.\n")
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Запустити OCR")
        self.stop_btn.config(state="disabled")
        self._release_memory()


class LauncherTab(BaseToolTab):
    DESCRIPTIONS = {
        "run_gui_image_preprocessor.py": (
            "Запускає графічний інструмент для препроцесінгу фото перед OCR: "
            "вирівнювання, обрізку, очищення та покращення якості зображень.")
        ,
        "run_vidomist.py": (
            "Запускає графічний модуль для роботи з відомостями: підготовки, "
            "перевірки та експорту звітів.")
    }

    def __init__(self, parent, script_path, display_name=None):
        super().__init__(parent, script_path)
        self.display_name = display_name or script_path.name
        description = self.DESCRIPTIONS.get(script_path.name,
            f"Запускає графічний інтерфейс для {self.display_name}.")

        self.add_hint(
            description
        )
        self.run_btn, self.stop_btn = self.add_run_stop_row(
            f"Запустити {self.display_name}", self.run_clicked, color="#00796b"
        )
        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def run_clicked(self):
        if self.is_running:
            return
        if not self.script_path.is_file():
            messagebox.showerror("Помилка", f"Скрипт не знайдено:\n{self.script_path}")
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text=f"Запуск {self.display_name}...")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        cmd = [sys.executable, "-u", str(self.script_path)]
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")

        extra = {}
        if POPPLER_BIN.is_dir():
            extra["PATH"] = str(POPPLER_BIN) + os.pathsep + os.environ.get("PATH", "")
        exit_code = stream_subprocess(cmd, SCRIPT_DIR, utf8_env(extra), self.log_panel.log,
                                       proc_holder=self.current_proc)

        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log("Готово!\n")
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Завершено з кодом помилки {exit_code}.\n")
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text=f"Запустити {self.display_name}")
        self.stop_btn.config(state="disabled")
        self._release_memory()


# ──────────────────────────────────────────────────────────────────────
# Вкладка 4: excel_to_word_transfer.py (прямий імпорт, без subprocess)
# ──────────────────────────────────────────────────────────────────────
class TransferTab(BaseToolTab):
    def __init__(self, parent):
        super().__init__(parent, None)  # немає окремого файлу-скрипта для перевірки наявності

        self.template_path = tk.StringVar()
        self.excel_path = tk.StringVar()

        self.add_open_file_row(
            "Шаблон Word (.docx / .doc):", self.template_path,
            [("Word files", "*.docx *.doc"), ("Всі файли", "*.*")],
        )
        self.add_open_file_row(
            "Файл Excel з даними (.xlsx / .xlsm / .xls):", self.excel_path,
            [("Excel files", "*.xlsx *.xlsm *.xls"), ("Всі файли", "*.*")],
        )

        self.run_btn, self.stop_btn = self.add_run_stop_row(
            "Запустити перенесення", self.run_clicked, stop_command=self._stop_transfer_clicked,
        )
        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def run_clicked(self):
        if self.is_running:
            return
        template = self.template_path.get()
        excel = self.excel_path.get()
        if not template or not os.path.isfile(template):
            messagebox.showwarning("Увага", "Спочатку виберіть коректний шаблон Word.")
            return
        if not excel or not os.path.isfile(excel):
            messagebox.showwarning("Увага", "Спочатку виберіть коректний файл Excel.")
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        threading.Thread(target=self._worker, args=(template, excel), daemon=True).start()

    def _stop_transfer_clicked(self):
        """На відміну від інших вкладок, це перенесення виконується прямим
        викликом process_workbook() у потоці цього ж python.exe, а не
        окремим subprocess — тут просто немає окремого PID, який можна
        безпечно вбити (kill_process_tree тут незастосовний). Примусове
        переривання Python-потоку посеред запису .docx могло б лишити
        пошкоджений/напівзаписаний файл, тому чесно попереджаємо, а не
        імітуємо зупинку."""
        if not self.is_running:
            return
        messagebox.showinfo(
            "Зупинка недоступна",
            "Цей інструмент працює всередині самого лаунчера (не окремим "
            "процесом), тому безпечно перервати запис файлу посеред "
            "операції неможливо — це ризикує лишити пошкоджений .docx.\n\n"
            "Інструмент завершить поточний аркуш сам (це швидко). Якщо "
            "справді треба зупинити негайно — закрий програму повністю."
        )

    def _worker(self, template, excel):
        try:
            from excel_to_word_transfer import process_workbook  # локальний імпорт: модуль лежить поруч
        except ImportError as e:
            self.log_panel.log(f"ПОМИЛКА: не вдалося імпортувати excel_to_word_transfer.py: {e}\n")
            self.after(0, lambda: self._finish(False, "Модуль excel_to_word_transfer.py не знайдено."))
            return

        try:
            results = process_workbook(template, excel, log=self.log_panel.log_line)
            ok_count = sum(1 for r in results if r.output_path)
            skipped = [r for r in results if not r.output_path]
            summary = f"Завершено. Успішно: {ok_count} з {len(results)} аркушів."
            self.log_panel.log_line(summary)
            for r in skipped:
                self.log_panel.log_line(f"  Пропущено '{r.sheet_name}': {r.skipped_reason or 'причина невідома'}")
            self.after(0, lambda: self._finish(True, summary))
        except Exception as exc:  # noqa: BLE001
            error_msg = f"ПОМИЛКА: {exc}"
            self.log_panel.log_line(error_msg)
            self.after(0, lambda: self._finish(False, error_msg))

    def _finish(self, success, message):
        self.is_running = False
        self.run_btn.config(state="normal", text="Запустити перенесення")
        self.stop_btn.config(state="disabled")
        self._release_memory()
        if success:
            messagebox.showinfo("Готово", message)
        else:
            messagebox.showerror("Помилка", message)


# Допоміжний метод: process_workbook очікує log(message: str) без "\n" —
# додаємо тонку обгортку поверх LogPanel.log, що сама додає перенесення рядка.
def _log_line(self, message):
    self.log(message + "\n")


LogPanel.log_line = _log_line


# ──────────────────────────────────────────────────────────────────────
# Вкладка 5: universal_dict_extractor.py
# ──────────────────────────────────────────────────────────────────────
class DictExtractorTab(BaseToolTab):
    def __init__(self, parent):
        super().__init__(parent, EXTRACTOR_SCRIPT)

        self.macket_path = tk.StringVar()
        self.source_folder = tk.StringVar()
        self.threshold_var = tk.IntVar(value=85)

        self.add_open_file_row("Макет (xlsm з аркушем 'Словник'):", self.macket_path, XLSX_TYPES)
        self.add_dir_row("Папка з файлами-джерелами:", self.source_folder)
        self.add_hint("Пошук іде рекурсивно — усі підпапки й підпапки підпапок (.xlsx, .xlsm, .docx).")
        self.add_spinbox_row("Поріг збігу (%):", self.threshold_var, 50, 100, hint="рекомендовано 80–95")

        self.run_btn, self.stop_btn = self.add_run_stop_row("Запустити обробку", self.run_clicked)
        self.log_panel = LogPanel(self)
        self.log_panel.pack(fill="both", expand=True, padx=10, pady=6)

    def run_clicked(self):
        if self.is_running:
            return
        if not self.macket_path.get() or not self.source_folder.get():
            messagebox.showwarning("Увага", "Вибери і макет, і папку з файлами-джерелами.")
            return
        if not EXTRACTOR_SCRIPT.is_file():
            messagebox.showerror("Помилка", f"universal_dict_extractor.py не знайдено:\n{EXTRACTOR_SCRIPT}")
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        self.log_panel.clear()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        cmd = [
            sys.executable, "-u", str(EXTRACTOR_SCRIPT),
            self.macket_path.get(), self.source_folder.get(),
            "--threshold", str(self.threshold_var.get()),
        ]
        self.log_panel.log(f"Команда: {' '.join(cmd)}\n\n")

        exit_code = stream_subprocess(cmd, SCRIPT_DIR, utf8_env(), self.log_panel.log,
                                       proc_holder=self.current_proc)

        self.log_panel.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log_panel.log("Готово! Результат записано в обраний макет (бекап оригіналу створено поруч).\n")
            open_in_explorer(Path(self.macket_path.get()).parent)
        elif exit_code is None or exit_code < 0:
            self.log_panel.log("Зупинено користувачем. Пам'ять звільнено.\n")
        else:
            self.log_panel.log(f"Завершено з кодом помилки {exit_code}.\n")
        self.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Запустити обробку")
        self.stop_btn.config(state="disabled")
        self._release_memory()


# ──────────────────────────────────────────────────────────────────────
# Головне вікно
# ──────────────────────────────────────────────────────────────────────
class UnifiedApp:
    def __init__(self, root):
        self.root = root
        root.title("Обробка військової адмін. документації — єдиний лаунчер")
        root.geometry("880x720")
        root.minsize(760, 560)

        main_frame = ttk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=6, pady=6)

        left_panel = ttk.Frame(main_frame)
        left_panel.pack(side="left", fill="y", padx=(0, 6), pady=6)
        tk.Label(left_panel, text="Інструменти", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        self.buttons_frame = ttk.Frame(left_panel)
        self.buttons_frame.pack(fill="y", expand=True)

        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side="left", fill="both", expand=True)

        self.pages = {}
        self.tool_buttons = {}
        self.active_tool = None

        tools = [
            ("Втрати майна", DocProcessorTab),
            ("Відомість (Excel)", ExcelGeneratorTab),
            ("OCR PDF", OcrTab),
            ("Excel → Word", TransferTab),
            ("Словн. екстрактор", DictExtractorTab),
            ("Обробка фото GUI", partial(LauncherTab, script_path=SCRIPT_DIR / "run_gui_image_preprocessor.py", display_name="Обробка фото GUI")),
            ("Відомість GUI", partial(LauncherTab, script_path=SCRIPT_DIR / "run_vidomist.py", display_name="Відомість GUI")),
            ("Інші інструменти", OtherToolsTab),
        ]

        for label, cls in tools:
            btn = tk.Button(
                self.buttons_frame,
                text=label,
                anchor="w",
                relief="flat",
                command=lambda name=label: self.show_tool(name),
                padx=10,
                pady=8,
            )
            btn.pack(fill="x", pady=2)
            self.tool_buttons[label] = btn

            page = cls(right_panel)
            page.pack(fill="both", expand=True)
            page.pack_forget()
            self.pages[label] = page

        if tools:
            self.show_tool(tools[0][0])

    def show_tool(self, label):
        if self.active_tool == label:
            return
        if self.active_tool is not None:
            current_page = self.pages.get(self.active_tool)
            if current_page is not None:
                current_page.pack_forget()
            current_btn = self.tool_buttons.get(self.active_tool)
            if current_btn is not None:
                current_btn.config(bg=self.root.cget("bg"), fg="black")

        self.active_tool = label
        page = self.pages.get(label)
        if page is not None:
            page.pack(fill="both", expand=True)
        btn = self.tool_buttons.get(label)
        if btn is not None:
            btn.config(bg="#d1c4e9", fg="black")


def main():
    root = tk.Tk()
    UnifiedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()