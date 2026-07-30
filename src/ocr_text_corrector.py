#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ocr_text_corrector.py - the stage between OCR and statement generation.

WHAT IT FIXES
-------------
After pdf_to_text_multithread.py the .txt still carries damage that quietly
corrupts the statement downstream. Measured on real reports:

  "кост м масмув льний зима а чохлум"          -> winter camouflage suit
  "сумка підсумок розвантажувалтна універсальна"
  "рюкзак індив дуальтий бойовий"
  "чохол до кас и будівельної ( ікселаний)"    -> (pixel pattern)
  "каска будівельна (захисна)"                 -> (protective)
  "речова служба ... частини А0О000"           -> А0000

Names in this state do not reach the 82% similarity the catalogue sheet
requires, so excel_report_generator.py drops them into the statement RAW -
no price, lower-case, in a shape somebody then has to fix by hand.

HOW IT FIXES (three tiers, cheapest first)
------------------------------------------
1. Deterministic substitutions: glyphs inside the unit code, glued words, the
   recurring OCR confusions. Instant, no dependencies.
2. Catalogue (the "Словник" sheet of the template, ~760 entries): fuzzy lookup
   via rapidfuzz. This tier does most of the work, because item names repeat
   from document to document.
3. Local LLM (llm_extractor.LlmExtractor -> llama-server + a 7B GGUF model)
   ONLY for names the catalogue could not resolve. resolve_name already applies
   a Levenshtein pre-filter, so the model sees roughly ten relevant candidates
   instead of all 760. Enabled with --llm.

Every LLM substitution is cached in Data/ocr_corrections.json, so a repeat run
over the same names does not pay for the model twice.

USAGE
-----
    python ocr_text_corrector.py <txt_folder> <output_folder> <template.xlsm> [--llm]

The original .txt files are never modified - corrected copies go to the output
folder, so a before/after comparison is always possible.
"""

import json
import os
import re
import sys
import time
from pathlib import Path

try:
    from rapidfuzz import fuzz
except ImportError:
    print("ПОМИЛКА: потрібен rapidfuzz (pip install rapidfuzz)", flush=True)
    raise

try:
    import openpyxl
except ImportError:
    print("ПОМИЛКА: потрібен openpyxl (pip install openpyxl)", flush=True)
    raise

PROJECT_ROOT = Path(__file__).resolve().parent
CACHE_FILE = PROJECT_ROOT / "Data" / "ocr_corrections.json"

# Threshold for AUTOMATIC catalogue substitution. Deliberately high.
#
# The first version of this module used 72, and on real reports that produced
# semantically wrong substitutions which looked "successful" in the log:
#   "рюкзак індив дуальтий бойовий" -> "Казанок індивідуальний польовий" (77%)
#     (a backpack replaced by a mess tin)
#   "чоботи гумові"                 -> "Чоботи хромові"                  (81%)
#     (rubber boots replaced by chrome-leather boots)
#   "чохол ... ( ікселаний)"        -> an unrelated "Чохол ..." entry  (77%)
# Fuzzy similarity measures SHARED CHARACTERS, not "the same object": entries
# for different products differ by a single word, and Levenshtein cannot see
# that difference. So above 92 we accept on our own (there the difference
# really is only OCR noise), and the whole ambiguous band goes to the LLM,
# which is asked "is this THE SAME item?" rather than "does the spelling look
# alike?".
#
# With the LLM disabled the line is left untouched. That is not a loss:
# excel_report_generator.py performs its own catalogue matching downstream, so
# nothing breaks - we simply refrain from guessing.
DICT_ACCEPT_SCORE = 92

# Below this we do not touch the name even with the LLM enabled: if it
# resembles nothing in the catalogue, it is most likely a genuinely new item
# rather than a corrupted one.
DICT_LLM_FLOOR_SCORE = 45

# A token longer than this with no spaces is the signature of "glued
# columns" (see _looks_glued). Such lines are flagged, not repaired:
# inventing word boundaries means inventing data.
GLUED_TOKEN_LEN = 28

ITEM_LINE_RE = re.compile(
    r'^(?P<name>.+?)\s*[-–—]{1,2}\s*'
    r'(?P<tail>[^-–—]*)$',
    re.UNICODE
)


def log(msg):
    print(msg, flush=True)


def read_utf8(path):
    try:
        with open(path, encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


def write_utf8(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


# ================================================================
# CORRECTION CACHE
# ================================================================
def load_cache():
    try:
        with open(CACHE_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_cache(cache):
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2, sort_keys=True)


# ================================================================
# TIER 1: DETERMINISTIC SUBSTITUTIONS
# ================================================================
# Unit code: OCR reads "А0000" as "А0О000", "Аб0000", "доОг000" and so on.
# A letter inside the numeric part is always an error, because the format is
# fixed: one letter followed by four digits.
_UNIT_CODE_RE = re.compile(r'\b[АAД][\dОоОоOo0бБгГ]{4,7}\b')
_CODE_GLYPH = {'О': '0', 'о': '0', 'O': '0', 'o': '0', 'б': '6', 'Б': '6',
               'г': '', 'Г': '', 'д': '', 'Д': ''}


def fix_unit_code(text):
    """Normalises corrupted unit codes to the "А0000" shape."""
    def repl(m):
        raw = m.group(0)
        digits = ''.join(_CODE_GLYPH.get(ch, ch) for ch in raw[1:])
        digits = re.sub(r'\D', '', digits)
        if len(digits) >= 4:
            return 'А' + digits[:4]
        return raw
    return _UNIT_CODE_RE.sub(repl, text)


def _looks_glued(line):
    """True if the line contains an unnaturally long space-free token - the
    trace of table columns glued together by OCR."""
    for token in line.split():
        if len(token) >= GLUED_TOKEN_LEN and token.isalpha():
            return True
    return False


def fix_common_ocr(text):
    """Substitutions that are unambiguous from context alone, no catalogue."""
    text = fix_unit_code(text)
    # "мм?" / "мм:" instead of "мм²" - the usual superscript misread
    text = re.sub(r'\bмм[?:]', 'мм²', text)
    # A period where OCR clearly meant a comma: "павук. виносна"
    text = re.sub(r'([а-яіїєґ])\.\s+([а-яіїєґ])', r'\1, \2', text)
    return text


# ================================================================
# TIER 2: CATALOGUE
# ================================================================
def load_dictionary(template_path):
    """Reads the template catalogue sheet: column A - name, B - unit."""
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    if "Словник" not in wb.sheetnames:
        wb.close()
        raise SystemExit(f"ПОМИЛКА: у макеті немає аркуша 'Словник': {template_path}")
    ws = wb["Словник"]
    entries = []
    for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        name = row[0]
        if name and str(name).strip():
            entries.append(str(name).strip())
    wb.close()
    return entries


def dict_best_match(name, dictionary):
    """Best catalogue match -> (name, score 0-100)."""
    low = name.strip().lower()
    best, best_score = None, -1.0
    for entry in dictionary:
        score = max(fuzz.ratio(low, entry.lower()),
                    fuzz.token_sort_ratio(low, entry.lower()))
        if score > best_score:
            best_score, best = score, entry
    return best, best_score


# ================================================================
# ITEM-LINE PARSING
# ================================================================
def split_item_line(line):
    """Splits "name - 2 pcs.;" into (name, tail). The tail (quantity, unit,
    punctuation) is left alone - excel_report_generator parses it, and the
    quantity-glyph decoding already lives there."""
    m = ITEM_LINE_RE.match(line.strip())
    if not m:
        return None, None
    name = m.group("name").strip()
    tail = m.group("tail").strip()
    if len(name) < 3:
        return None, None
    # The tail must look like quantity + unit; otherwise this is not an item
    # line but an ordinary sentence that happens to contain a dash.
    if not re.match(r'^[\d|IІіlЇїбБзЗОоOo]+[\s.,]*[А-Яа-яA-Za-z.\-/]*[\s.,;]*$', tail):
        return None, None
    return name, tail


def is_item_line(line):
    name, tail = split_item_line(line)
    return name is not None


# ================================================================
# MAIN PASS
# ================================================================
class Corrector:
    def __init__(self, dictionary, use_llm=False):
        self.dictionary = dictionary
        self.use_llm = use_llm
        self.cache = load_cache()
        self.extractor = None
        self.stats = {"dict": 0, "llm": 0, "cache": 0, "kept": 0, "glued": 0}

    # ---- The LLM starts LAZILY: if the catalogue resolved everything,
    # ---- the 4 GB model is never loaded into memory at all.
    def _ensure_llm(self):
        if self.extractor is not None:
            return True
        try:
            import llm_extractor as llm_mod
        except Exception as e:
            log(f"  [LLM] недоступна ({e}) — лишаю назви як є")
            self.use_llm = False
            return False
        try:
            log("  [LLM] Піднімаю llama-server (перший раз ~30-60 сек)...")
            self.extractor = llm_mod.LlmExtractor()
            self.extractor.start()
            return True
        except Exception as e:
            log(f"  [LLM] Не вдалося запустити: {e} — лишаю назви як є")
            self.extractor = None
            self.use_llm = False
            return False

    def stop(self):
        if self.extractor is not None:
            try:
                self.extractor.stop()
            except Exception:
                pass
            self.extractor = None

    def correct_name(self, raw_name):
        """Returns (corrected_name, source)."""
        cleaned = fix_common_ocr(raw_name)

        best, score = dict_best_match(cleaned, self.dictionary)
        if best is not None and score >= DICT_ACCEPT_SCORE:
            self.stats["dict"] += 1
            return best, f"словник {score:.0f}%"

        key = cleaned.strip().lower()
        if key in self.cache:
            self.stats["cache"] += 1
            cached = self.cache[key]
            return (cached, "кеш") if cached else (cleaned, "кеш: без змін")

        # Too unlike anything in the catalogue - more likely a new item than a
        # corrupted one. Not worth an LLM call.
        if best is None or score < DICT_LLM_FLOOR_SCORE:
            self.stats["kept"] += 1
            return cleaned, f"лишено (макс. схожість {score:.0f}%)"

        if not self.use_llm or not self._ensure_llm():
            self.stats["kept"] += 1
            return cleaned, f"лишено (словник {score:.0f}%, LLM вимкнена)"

        try:
            resolved = self.extractor.resolve_name(cleaned, self.dictionary)
        except Exception as e:
            log(f"    [LLM] помилка на '{cleaned[:40]}': {e}")
            resolved = None

        self.cache[key] = resolved or ""
        save_cache(self.cache)

        if resolved:
            self.stats["llm"] += 1
            return resolved, "LLM"
        self.stats["kept"] += 1
        return cleaned, f"лишено (LLM не впізнала, словник {score:.0f}%)"

    def correct_text(self, text, file_label=""):
        out_lines = []
        for raw in text.split("\n"):
            line = raw.rstrip()
            if not line.strip():
                out_lines.append(line)
                continue

            if _looks_glued(line):
                self.stats["glued"] += 1
                log(f"    УВАГА склеєні колонки, лишаю як є: «{line.strip()[:70]}»")
                out_lines.append(fix_common_ocr(line))
                continue

            name, tail = split_item_line(line)
            if name is None:
                out_lines.append(fix_common_ocr(line))
                continue

            fixed, source = self.correct_name(name)
            if fixed != name:
                log(f"    {source:<26} «{name[:44]}» -> «{fixed[:44]}»")
            out_lines.append(f"{fixed} - {tail}")
        return "\n".join(out_lines)


def main():
    if len(sys.argv) < 4:
        print("Використання: python ocr_text_corrector.py <папка_txt> "
              "<вихідна_папка> <макет.xlsm|.xlsx> [--llm]")
        print("  --llm   вмикати локальну LLM для назв, які не взяв словник")
        print("          (повільно на CPU, але результат кешується)")
        sys.exit(1)

    in_dir = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    template = Path(sys.argv[3])
    use_llm = "--llm" in sys.argv

    if not in_dir.is_dir():
        print(f"ПОМИЛКА: вхідна папка не існує: {in_dir}")
        sys.exit(1)
    if not template.is_file():
        print(f"ПОМИЛКА: макет не знайдено: {template}")
        sys.exit(1)

    dictionary = load_dictionary(template)
    log(f"Словник: {len(dictionary)} позицій з {template.name}")
    log(f"LLM: {'увімкнена' if use_llm else 'вимкнена (--llm щоб увімкнути)'}")

    files = sorted(in_dir.glob("*.txt"))
    if not files:
        print(f"ПОМИЛКА: у {in_dir} немає .txt файлів")
        sys.exit(1)
    log(f"Файлів: {len(files)}")

    corrector = Corrector(dictionary, use_llm=use_llm)
    t0 = time.time()
    try:
        for i, path in enumerate(files, 1):
            log(f"[{i}/{len(files)}] {path.name}")
            fixed = corrector.correct_text(read_utf8(path), path.name)
            write_utf8(str(out_dir / path.name), fixed)
    finally:
        corrector.stop()

    s = corrector.stats
    log("")
    log("=" * 52)
    log(f"ГОТОВО за {time.time() - t0:.0f} сек")
    log(f"Виправлено словником:  {s['dict']}")
    log(f"Виправлено LLM:        {s['llm']}")
    log(f"Взято з кешу:          {s['cache']}")
    log(f"Лишено без змін:       {s['kept']}")
    if s["glued"]:
        log(f"Склеєних рядків:       {s['glued']} (потребують ручної перевірки)")
    log(f"Результат:             {out_dir}")
    log("=" * 52)


if __name__ == "__main__":
    main()
