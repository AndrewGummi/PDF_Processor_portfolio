#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pipeline_pdf_to_excel.py - full cycle: a folder of PDFs -> a finished Excel
statement.

Runs the three existing stages back to back, feeding each result forward:

  1. pdf_to_text_multithread.py   PDF  -> .txt   (OCR: pdftoppm + Tesseract
                                                  + ImageMagick, consensus)
  2. ocr_text_corrector.py        .txt -> .txt   (name correction against the
                                                  catalogue, optional LLM)
  3. excel_report_generator.py    .txt -> .xlsm  (statement + price matching)

WHY A SEPARATE FILE INSTEAD OF MERGING THE THREE SCRIPTS:
each stage stays independently runnable, which is the normal way this is used
in practice - re-running only the OCR, or only the statement generation. This
file is a thin dispatcher: it computes nothing itself, it only launches the
stages as separate python.exe processes via subprocess (no shell=True, same as
everywhere else in the project) and streams their logs. If a stage fails, the
cycle stops and reports which one.

INTERMEDIATE FOLDERS
Created next to the output file by default:
    <output>_txt_raw/    - raw OCR (so there is something to compare against)
    <output>_txt_fixed/  - after name correction
Both are deliberately left on disk: when an odd row shows up in the statement,
you need to see which stage corrupted it.

USAGE
    python pipeline_pdf_to_excel.py <pdf_folder> <statement.xlsm> <template.xlsm>
        [--llm] [--preprocess] [--threshold N] [--history-folder DIR]
        [--skip-ocr] [--keep-existing-txt]

    --llm               enable the local LLM in stage 2 (slow, more accurate)
    --preprocess        image conditioning in stage 1 (for photos / skewed scans)
    --threshold N       price-matching threshold, 50-100 (default 90)
    --history-folder    folder of finished files used as a fallback price source
    --skip-ocr          skip stage 1 and use the .txt files already present
    --keep-existing-txt do not wipe the raw-OCR folder before the run
"""

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
OCR_SCRIPT = SCRIPT_DIR / "pdf_to_text_multithread.py"
CORRECTOR_SCRIPT = SCRIPT_DIR / "ocr_text_corrector.py"
EXCEL_SCRIPT = SCRIPT_DIR / "excel_report_generator.py"
POPPLER_BIN = SCRIPT_DIR / "poppler" / "poppler-24.08.0" / "Library" / "bin"


def log(msg=""):
    print(msg, flush=True)


def utf8_env():
    """PYTHONUTF8=1 - without it the child python.exe writes stdout in the
    console code page and Cyrillic in the log turns into garbage."""
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    if POPPLER_BIN.is_dir():
        env["PATH"] = str(POPPLER_BIN) + os.pathsep + env.get("PATH", "")
    return env


def run_stage(title, cmd):
    """Runs one stage as a separate process, streaming its output.
    Returns the exit code."""
    log()
    log("=" * 64)
    log(f"[ЕТАП] {title}")
    log("=" * 64)
    log(f"Команда: {' '.join(str(c) for c in cmd)}")
    log()

    t0 = time.time()
    try:
        proc = subprocess.Popen(
            [str(c) for c in cmd],
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
        for line in proc.stdout:
            print(line, end="", flush=True)
        proc.wait()
        code = proc.returncode
    except Exception as e:
        log(f"ПОМИЛКА ЗАПУСКУ ЕТАПУ: {e}")
        return -1

    log()
    log(f"[ЕТАП] {title}: код {code}, {time.time() - t0:.0f} сек")
    return code


def count_txt(folder: Path):
    return len(list(folder.glob("*.txt"))) if folder.is_dir() else 0


def _flag_value(argv, flag, default=None):
    if flag not in argv:
        return default
    i = argv.index(flag)
    if i + 1 >= len(argv):
        print(f"ПОМИЛКА: після {flag} очікується значення.")
        sys.exit(1)
    return argv[i + 1]


def main():
    argv = sys.argv[1:]
    positional = []
    skip_next = False
    value_flags = {"--threshold", "--history-folder"}
    for i, a in enumerate(argv):
        if skip_next:
            skip_next = False
            continue
        if a in value_flags:
            skip_next = True
            continue
        if a.startswith("--"):
            continue
        positional.append(a)

    if len(positional) < 3:
        print(__doc__)
        sys.exit(1)

    pdf_folder = Path(positional[0]).resolve()
    output_file = Path(positional[1]).resolve()
    template = Path(positional[2]).resolve()

    use_llm = "--llm" in argv
    use_preprocess = "--preprocess" in argv
    skip_ocr = "--skip-ocr" in argv
    keep_txt = "--keep-existing-txt" in argv
    threshold = _flag_value(argv, "--threshold")
    history_folder = _flag_value(argv, "--history-folder")

    for path, what in ((pdf_folder, "папка з PDF"), (template, "макет")):
        if not path.exists():
            print(f"ПОМИЛКА: {what} не знайдено: {path}")
            sys.exit(1)
    for script in (OCR_SCRIPT, CORRECTOR_SCRIPT, EXCEL_SCRIPT):
        if not script.is_file():
            print(f"ПОМИЛКА: не знайдено скрипт етапу: {script}")
            sys.exit(1)

    stem = output_file.stem
    raw_dir = output_file.parent / f"{stem}_txt_raw"
    fixed_dir = output_file.parent / f"{stem}_txt_fixed"

    log("=" * 64)
    log("ПОВНИЙ ЦИКЛ: PDF -> відомість Excel")
    log("=" * 64)
    log(f"Папка з PDF:      {pdf_folder}")
    log(f"Макет:            {template.name}")
    log(f"Вихідний файл:    {output_file}")
    log(f"Сирий OCR:        {raw_dir}")
    log(f"Виправлений текст:{fixed_dir}")
    log(f"Підготовка сторінок (--preprocess): {'так' if use_preprocess else 'ні'}")
    log(f"Локальна LLM (--llm):               {'так' if use_llm else 'ні'}")

    t_start = time.time()

    # ---------- STAGE 1: OCR ----------
    if skip_ocr:
        log()
        log(f"[ЕТАП 1] Пропущено (--skip-ocr). Беру готові .txt з {raw_dir}")
        if count_txt(raw_dir) == 0:
            print(f"ПОМИЛКА: у {raw_dir} немає .txt — нічого пропускати.")
            sys.exit(1)
    else:
        if raw_dir.exists() and not keep_txt:
            shutil.rmtree(raw_dir, ignore_errors=True)
        raw_dir.mkdir(parents=True, exist_ok=True)

        # pdf_to_text_multithread.py writes into its own output_txt folder
        # from config.py, and has no env var to redirect that - so we run it
        # as-is and collect the result afterwards.
        #
        # That folder is NOT cleared between runs, so it accumulates results
        # from every previous batch. Copying it wholesale silently mixed
        # unrelated documents into the statement: a run over folder B ended up
        # containing folder A's reports as well (observed: a 13-item demo came
        # out with 25 items, 12 of them from an earlier unrelated run).
        # So we snapshot the folder first and afterwards take only the files
        # this run actually produced - by name and by modification time.
        project_txt = SCRIPT_DIR / "output_txt"
        before = {p.name: p.stat().st_mtime
                  for p in project_txt.glob("*.txt")} if project_txt.is_dir() else {}
        stage_started = time.time()

        cmd = [sys.executable, "-u", OCR_SCRIPT, str(pdf_folder), "--no-llm-recheck"]
        if use_preprocess:
            cmd.append("--preprocess")
        code = run_stage("1/3  OCR: PDF -> текст", cmd)
        if code != 0:
            log("ЦИКЛ ЗУПИНЕНО: етап OCR завершився з помилкою.")
            sys.exit(code)

        moved = skipped_stale = 0
        for src in sorted(project_txt.glob("*.txt")):
            fresh = (src.name not in before
                     or src.stat().st_mtime >= stage_started - 1)
            if not fresh:
                skipped_stale += 1
                continue
            shutil.copy2(src, raw_dir / src.name)
            moved += 1
        log(f"[ЕТАП 1] Скопійовано .txt у {raw_dir}: {moved}")
        if skipped_stale:
            log(f"[ЕТАП 1] Пропущено {skipped_stale} файл(ів) з попередніх "
                f"запусків, що лежали в {project_txt.name} — вони НЕ потрапили "
                f"у відомість.")

    if count_txt(raw_dir) == 0:
        log("ЦИКЛ ЗУПИНЕНО: після OCR немає жодного .txt "
            "(жоден рапорт не пройшов фільтр 'Речова служба').")
        sys.exit(1)

    # ---------- STAGE 2: text correction ----------
    if fixed_dir.exists():
        shutil.rmtree(fixed_dir, ignore_errors=True)
    cmd = [sys.executable, "-u", CORRECTOR_SCRIPT, str(raw_dir), str(fixed_dir), str(template)]
    if use_llm:
        cmd.append("--llm")
    code = run_stage("2/3  Виправлення назв (словник"
                     + (" + LLM)" if use_llm else ")"), cmd)
    if code != 0:
        log("ЦИКЛ ЗУПИНЕНО: етап виправлення тексту завершився з помилкою.")
        sys.exit(code)

    if count_txt(fixed_dir) == 0:
        log("ЦИКЛ ЗУПИНЕНО: етап виправлення не дав жодного файлу.")
        sys.exit(1)

    # ---------- STAGE 3: statement ----------
    cmd = [sys.executable, "-u", EXCEL_SCRIPT, str(fixed_dir), str(output_file), str(template)]
    if threshold:
        cmd += ["--threshold", str(threshold)]
    if history_folder:
        cmd += ["--history-folder", str(history_folder)]
    code = run_stage("3/3  Формування відомості Excel", cmd)
    if code != 0:
        log("ЦИКЛ ЗУПИНЕНО: етап формування відомості завершився з помилкою.")
        sys.exit(code)

    # excel_report_generator.py forces the output extension to match the
    # template's (to preserve .xlsm macros), so the file that actually gets
    # written may not carry the name passed on the command line.
    actual = output_file
    if not actual.is_file():
        candidate = output_file.with_suffix(template.suffix)
        if candidate.is_file():
            actual = candidate

    elapsed = time.time() - t_start
    log()
    log("=" * 64)
    log(f"ЦИКЛ ЗАВЕРШЕНО за {elapsed:.0f} сек ({elapsed / 60:.1f} хв)")
    log(f"  Рапортів розпізнано:  {count_txt(raw_dir)}")
    log(f"  Сирий текст:          {raw_dir}")
    log(f"  Виправлений текст:    {fixed_dir}")
    log(f"  ВІДОМІСТЬ:            {actual}"
        + ("" if actual.is_file() else "  (файл не знайдено — перевір лог етапу 3)"))
    log("=" * 64)
    log("Перевір у відомості позиції, підсвічені червоним — це ті, для яких")
    log("не знайшлось ціни або надійного збігу зі словником.")


if __name__ == "__main__":
    main()
