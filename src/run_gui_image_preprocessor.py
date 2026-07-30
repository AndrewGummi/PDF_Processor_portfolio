#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_gui_image_preprocessor.py — GUI-запускатор для image_preprocessor.py.

За зразком run_gui_unified.py: жодного cmd.exe/powershell.exe/bash.exe,
дочірній скрипт запускається напряму як python.exe-процес через
subprocess.Popen (без shell=True), лог стрімиться наживо, є кнопка
"Зупинити" (жорстко завершує процес разом з нащадками через taskkill).

Питає лише два обов'язкові шляхи:
  - яку папку з зображеннями обробляти (рекурсивно, разом з підпапками);
  - куди складати оброблені PNG (структура підпапок зберігається).

Запуск:
    python.exe run_gui_image_preprocessor.py
Поклади поруч з image_preprocessor.py.
"""

import os
import sys
import gc
import subprocess
import threading
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PREPROCESS_SCRIPT = SCRIPT_DIR / "image_preprocessor.py"
LOGS_DIR = SCRIPT_DIR / "Logs"


def utf8_env():
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def kill_process_tree(pid, log_callback=None):
    """Завершує процес разом з усіма нащадками (taskkill /T /F, без
    shell=True) — так само, як у run_gui_unified.py."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if log_callback:
            log_callback(f"  Дерево процесів (PID {pid}) примусово завершено.\n")
    except Exception as e:
        if log_callback:
            log_callback(f"  Увага: не вдалося завершити процес через taskkill: {e}\n")


def open_in_explorer(path):
    try:
        os.startfile(str(path))
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("Підготовка зображень під OCR")
        root.geometry("820x620")
        root.minsize(700, 480)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.make_strips = tk.BooleanVar(value=True)
        self.make_gamma = tk.BooleanVar(value=True)
        self.curve_fix = tk.BooleanVar(value=True)
        self.use_osd = tk.BooleanVar(value=True)
        self.denoise = tk.BooleanVar(value=False)
        self.rotate_choice = tk.StringVar(value="авто (по контенту)")
        self.rows_per_strip = tk.IntVar(value=10)
        self.workers = tk.IntVar(value=3)

        self.is_running = False
        self.current_proc = {"proc": None}

        pad = {"padx": 10, "pady": 6}

        frame = ttk.Frame(root)
        frame.pack(fill="x")
        frame.grid_columnconfigure(1, weight=1)

        tk.Label(frame, text="Папка із зображеннями (шукає рекурсивно, разом з підпапками):").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 0))
        tk.Entry(frame, textvariable=self.input_dir, width=70).grid(row=1, column=0, columnspan=2, sticky="we", **pad)
        tk.Button(frame, text="Вибрати...", command=self.pick_input_dir).grid(row=1, column=2, **pad)

        tk.Label(frame, text="Куди складати оброблені зображення (структура підпапок збережеться):").grid(
            row=2, column=0, columnspan=3, sticky="w", padx=10)
        tk.Entry(frame, textvariable=self.output_dir, width=70).grid(row=3, column=0, columnspan=2, sticky="we", **pad)
        tk.Button(frame, text="Вибрати...", command=self.pick_output_dir).grid(row=3, column=2, **pad)

        tk.Checkbutton(
            frame, text="Нарізати на горизонтальні смуги (кожна зі шапкою колонок)",
            variable=self.make_strips,
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 0))

        tk.Checkbutton(
            frame, text="Додатковий варіант з gamma 0.85 (видно бліді олівцеві правки)",
            variable=self.make_gamma,
        ).grid(row=5, column=0, columnspan=3, sticky="w", padx=10)

        tk.Checkbutton(
            frame, text="Циліндричний деварп (виправляти «горб» від корінця)",
            variable=self.curve_fix,
        ).grid(row=6, column=0, columnspan=3, sticky="w", padx=10)

        tk.Checkbutton(
            frame, text="Визначати орієнтацію по контенту (Tesseract OSD)",
            variable=self.use_osd,
        ).grid(row=7, column=0, columnspan=3, sticky="w", padx=10)

        tk.Checkbutton(
            frame, text="Слабке edge-preserving шумозаглушення (за ТЗ — краще без нього)",
            variable=self.denoise,
        ).grid(row=8, column=0, columnspan=3, sticky="w", padx=10)

        opts_row = ttk.Frame(frame)
        opts_row.grid(row=9, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 8))
        tk.Label(opts_row, text="Поворот:").pack(side="left")
        ttk.Combobox(
            opts_row, textvariable=self.rotate_choice, width=20, state="readonly",
            values=["авто (по контенту)", "0°", "90°", "180°", "270°"],
        ).pack(side="left", padx=(6, 16))
        tk.Label(opts_row, text="Рядків на смугу:").pack(side="left")
        tk.Spinbox(opts_row, from_=8, to=12, increment=1,
                   textvariable=self.rows_per_strip, width=5).pack(side="left", padx=(6, 16))
        tk.Label(opts_row, text="Потоків:").pack(side="left")
        tk.Spinbox(opts_row, from_=1, to=8, increment=1,
                   textvariable=self.workers, width=5).pack(side="left", padx=(6, 0))

        btn_row = ttk.Frame(root)
        btn_row.pack(fill="x", padx=10, pady=6)
        self.run_btn = tk.Button(
            btn_row, text="Обробити зображення", command=self.run_clicked,
            bg="#1565c0", fg="white", font=("Segoe UI", 11, "bold"), height=2,
        )
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.stop_btn = tk.Button(
            btn_row, text="Зупинити", command=self.stop_clicked,
            bg="#c62828", fg="white", font=("Segoe UI", 11, "bold"), height=2,
            state="disabled",
        )
        self.stop_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        tk.Label(root, text="Лог виконання (наживо):").pack(anchor="w", padx=10)
        self.log_box = scrolledtext.ScrolledText(root, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.log_file = None

        if not PREPROCESS_SCRIPT.is_file():
            messagebox.showerror(
                "Помилка",
                f"image_preprocessor.py не знайдено:\n{PREPROCESS_SCRIPT}\n\n"
                "Поклади run_gui_image_preprocessor.py в ту саму папку.",
            )

    def pick_input_dir(self):
        path = filedialog.askdirectory(title="Вибери папку із зображеннями")
        if path:
            self.input_dir.set(path)

    def pick_output_dir(self):
        path = filedialog.askdirectory(title="Вибери папку для оброблених зображень")
        if path:
            self.output_dir.set(path)

    def log(self, text):
        self.log_box.insert(tk.END, text)
        self.log_box.see(tk.END)
        if self.log_file:
            try:
                self.log_file.write(text)
                self.log_file.flush()
            except Exception:
                pass

    def run_clicked(self):
        if self.is_running:
            return
        if not self.input_dir.get():
            messagebox.showwarning("Увага", "Вибери папку із зображеннями.")
            return
        if not os.path.isdir(self.input_dir.get()):
            messagebox.showerror("Помилка", f"Папка не існує:\n{self.input_dir.get()}")
            return
        if not self.output_dir.get():
            messagebox.showwarning("Увага", "Вибери папку, куди складати оброблені зображення.")
            return
        if not PREPROCESS_SCRIPT.is_file():
            messagebox.showerror("Помилка", f"Скрипт не знайдено:\n{PREPROCESS_SCRIPT}")
            return

        out_dir = Path(self.output_dir.get())
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося створити вихідну папку:\n{e}")
            return

        self.is_running = True
        self.run_btn.config(state="disabled", text="Виконується...")
        self.stop_btn.config(state="normal")
        self.log_box.delete("1.0", tk.END)

        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        try:
            self.log_file = open(LOGS_DIR / f"image_preprocess_{ts}.log", "w", encoding="utf-8", buffering=1)
        except Exception:
            self.log_file = None

        threading.Thread(target=self._worker, args=(out_dir,), daemon=True).start()

    def stop_clicked(self):
        if not self.is_running:
            return
        proc = self.current_proc.get("proc")
        if proc is not None and proc.poll() is None:
            self.log("\n>>> Зупинка за запитом користувача...\n")
            kill_process_tree(proc.pid, self.log)
        else:
            self.log("\n>>> Процес уже завершився, нема що зупиняти.\n")
        self.stop_btn.config(state="disabled")

    def _worker(self, out_dir):
        cmd = [sys.executable, "-u", str(PREPROCESS_SCRIPT),
               self.input_dir.get(), str(out_dir),
               "--workers", str(self.workers.get()),
               "--rows-per-strip", str(self.rows_per_strip.get())]
        if not self.make_strips.get():
            cmd.append("--no-strips")
        if not self.make_gamma.get():
            cmd.append("--no-gamma")
        if not self.curve_fix.get():
            cmd.append("--no-curve-fix")
        if not self.use_osd.get():
            cmd.append("--no-osd")
        if self.denoise.get():
            cmd.append("--denoise")
        choice = self.rotate_choice.get()
        if choice.endswith("°"):
            cmd += ["--rotate", choice.rstrip("°")]
        self.log(f"Команда: {' '.join(cmd)}\n\n")

        proc = None
        try:
            proc = subprocess.Popen(
                cmd,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                cwd=str(SCRIPT_DIR),
                env=utf8_env(),
            )
            self.current_proc["proc"] = proc
            for line in proc.stdout:
                self.log(line)
            proc.wait()
            exit_code = proc.returncode
        except Exception as e:
            self.log(f"\nПОМИЛКА ЗАПУСКУ: {e}\n")
            exit_code = -1
        finally:
            self.current_proc["proc"] = None

        self.log(f"\n{'=' * 60}\n")
        if exit_code == 0:
            self.log(f"Готово! Результат у: {out_dir}\n")
            open_in_explorer(out_dir)
        elif exit_code is None or exit_code < 0:
            self.log("Зупинено користувачем.\n")
        else:
            self.log(f"Завершено з кодом помилки {exit_code}.\n")

        if self.log_file:
            try:
                self.log_file.close()
            except Exception:
                pass
            self.log_file = None

        self.root.after(0, self._finish)

    def _finish(self):
        self.is_running = False
        self.run_btn.config(state="normal", text="Обробити зображення")
        self.stop_btn.config(state="disabled")
        gc.collect()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
